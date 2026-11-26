#!/usr/bin/env python3
"""Пошук протоколів автоматизованого розподілу прямо у відкритих даних (без покажчика на Диску).

Автор: В'ячеслав Плескач. Версія 2.2 (26.11.2026): --from / --to / --around — РРРР-ММ-ДД або ДД.ММ.РРРР, як в інших скриптах (раніше
ДД.ММ.РРРР давало traceback). Версія 2.1 (25.11.2026): мітки полів усередині квадратних дужок («[Примітки: Наказ…]» у причині
виключення судді) більше не обривають перелік виключених суддів і не занижують excluded_count; номер справи порівнюється
без пробілів, регістру й кінцевого суфікса (-к, -ц, -а), як в інших скриптах; відсутні в наборі дні й помилки
завантаження друкуються і в результаті (stdout), а не лише в stderr; вивід понад 200 рядків позначається. 2.0 (21.11.2026): джерело — data.gov.ua і сайт ДСА, погодинні архіви,
звіт про автоматизований розподіл (поле LOG_JUDGE_ASSIGN).

Набір: «Протоколи (та звіти) автоматизованого розподілу судових справ між суддями».
  - data.gov.ua 0eb7722b-2d0e-4d28-86b8-ba0a45cae12d — архіви на день ПУБЛІКАЦІЇ, фактично 16.11.2023–08.02.2026;
  - сайт ДСА https://dsa.court.gov.ua/dsa/inshe/oddata/757/ (паспорт open_data_json.php?json=757) — той самий набір,
    ширше: з 28.04.2021; з 17.11.2026 ДСА дозавантажує період після 08.02.2026. 09.02–03.03.2026 — погодинні
    архіви, з 04.03.2026 — добові з 24 погодинними JSON усередині.
--dzherelo auto (типово): дні, які є на data.gov.ua, — з порталу, решта — із сайту ДСА; якщо портал не
відповідає — усе з ДСА. --dzherelo portal | dsa — лише одне джерело.
Скрипт завантажує архіви лише за заданий проміжок днів, відбирає протоколи за номером справи / судом /
суддею і розбирає їх текст. Повтори одного протоколу (той самий DOC_DBID + DOC_ID у кількох архівах) відкидає.

Звіт про автоматизований розподіл (з 09.02.2026, ≈90 % записів): колонки zvit_* — дата й час початку розподілу,
виключені судді з причинами, судді-учасники з навантаженням і коефіцієнтами, випадкове число, обраний суддя.
--zvit-dir ТЕКА зберігає повний текст кожного звіту у файл <doc_dbid>_<doc_id>.txt.

Оскільки архіви — за днями, проміжок треба задати: --from/--to (дні публікації) або --around ДАТА --days N.
Протокол зазвичай публікується в день розподілу або наступного дня; дату реєстрації справи беріть
з ЄДРСР (edrsr_vybirka.py) чи з набору стану розгляду справ.

Приклади:
  python3 protokoly_portal.py --case "757/30191/25-к" --around 2025-06-26 --days 3
  python3 protokoly_portal.py --case "490/13/26" --from 2026-03-09 --to 2026-03-11 --zvit-dir zvity
  python3 protokoly_portal.py --court 2606 --judge "Вовк" --from 2025-06-23 --to 2025-06-27 --out vovk.tsv
Швидкість: ~10–25 с на день (завантаження + відбір), 2 дні паралельно.
"""
import argparse, csv, json, os, subprocess, sys, time, datetime as dt
from concurrent.futures import ThreadPoolExecutor
PKG = '0eb7722b-2d0e-4d28-86b8-ba0a45cae12d'
DSA_JSON = 'https://dsa.court.gov.ua/open_data_json.php?json=757'
import base64, zipfile, io, re, html, datetime
from zoneinfo import ZoneInfo
KYIV = ZoneInfo('Europe/Kyiv')

F = ['day_file', 'court_code', 'court_name', 'case_number', 'proc_number', 'doc_form', 'protocol_type',
     'distribution_start', 'distribution_end', 'entry_date', 'ground', 'category', 'subject', 'reporting_judge',
     'collegium', 'excluded_count', 'participating_count', 'participating_judges', 'excluded_judges',
     'replaced_judges', 'doc_create_date', 'doc_dbid', 'doc_id', 'judges_field',
     'dzherelo', 'zvit_rozpodil_start', 'zvit_vykliucheni', 'zvit_uchasnyky', 'zvit_vypadkove_chyslo', 'zvit_obrano']


def text(b64):
    z = zipfile.ZipFile(io.BytesIO(base64.b64decode(b64)))
    s = z.read(z.namelist()[0]).decode('cp1251', 'replace')
    s = re.sub(r'(?is)<(style|script)[^>]*>.*?</\1>', ' ', s)
    s = re.sub(r'<[^>]+>', ' ', s)
    s = html.unescape(s)
    return re.sub(r'\s+', ' ', s).strip()


# label -> canonical key
LAB = [
    (r'Найменування суду', 'court'),
    (r'Єдиний унікальний номер судової справи', 'case'),
    (r'Вхідний номер справи', 'inno'),
    (r'Номер провадження(?: \(за наявності\))?', 'proc'),
    (r'Суть судової справи', 'subject'),
    (r'Категорія справи', 'category'),
    (r'Складність категорії справи', 'complex'),
    (r'Коефіцієнт складності', 'complex'),
    (r'Учасники (?:процесу|справи)', 'parties'),
    (r'Підстави здійснення автоматизованого розподілу', 'ground'),
    (r'Підстави проведення повторного автоматизованого (?:розподілу судової справи|визначення слідчого судді)', 'ground'),
    (r'Підстави (?:автоматизованого|автоматичного) визначення складу колегії', 'ground'),
    (r'Підстави передачі', 'ground'),
    (r'Обґрунтування призначення проведення повторного автоматизованого розподілу судової справи', 'ground2'),
    (r'Примітка проведення перерозподілу', 'ground2'),
    (r'Дата та час початку розподілу|Дата, час початку автоматизованого розподілу', 'start'),
    (r'Дата та час закінчення розподілу|Дата, час закінчення автоматизованого розподілу', 'end'),
    (r'Судді виключені з розподілу', 'excl'),
    (r'Були виключені наступні судді', 'excl'),
    (r'Судді, які (?:приймали|брали) участь (?:в|у) розподілі', 'part'),
    (r'Вибір відбувався серед суддів', 'part'),
    (r'Призначені судді', 'appointed'),
    (r'Головуючий суддя \(суддя-доповідач\)', 'rj'),
    (r'Суддя-доповідач', 'rj'),
    (r'Слідчий суддя', 'rj'),
    (r'Склад колегії суддів(?: \(за наявності\))?', 'coll'),
    (r'Вилучені \(замінені\) судді|Замінені \(вилучені\) судді', 'removed'),
    (r'Нові судді', 'new'),
    (r'Попередньо відібрані присяжні|Основні присяжні|Запасні присяжні', 'jurors'),
    (r'Примітки', 'notes'),
    (r'Призначення не відбулося', 'fail'),
    (r'Інформацію внесено', 'end_marker'),
]
LRE = re.compile('|'.join(f'(?P<g{i}>{p})' for i, (p, _) in enumerate(LAB)) + r'\s*:?')
HEAD = re.compile(r'(Витяг з протоколу|Протокол)\s+([а-яіїєґ\'’ ]+?)(?=\s+(?:[А-ЯІЇЄҐ0-9(№]|$))')
DT_END = re.compile(r'(?<!\d)(\d{2}\.\d{2}\.(?:\d{4}|\d{2}))(?!\d)(?:\s+(\d{1,2}:\d{2}(?::\d{2})?))?')
NAME_SEP = re.compile(r'(?:^|[;,]\s*)([А-ЯІЇЄҐ][^;:,]{2,80}?)\s+[-:]\s+')


def in_brackets(t):
    """Проміжки тексту всередині парних квадратних дужок: там мітки полів — частина значення, а не нове поле."""
    spans, stack = [], []
    for i, ch in enumerate(t):
        if ch == '[':
            stack.append(i)
        elif ch == ']' and stack:
            a = stack.pop()
            if not stack:
                spans.append((a, i))
    return spans


def case_key(s):
    return re.sub(r'-[^/0-9-]{1,4}$', '', re.sub(r'\s+', '', s or '').lower())


def fields(t):
    out = {}
    br = in_brackets(t)
    ms = [m for m in LRE.finditer(t) if not any(a < m.start() < b for a, b in br)]
    for i, m in enumerate(ms):
        k = LAB[int(m.lastgroup[1:])][1]
        v = t[m.end():ms[i + 1].start() if i + 1 < len(ms) else len(t)].strip(' :')
        if k in out and out[k]:
            continue
        out[k] = v
    return out


def clean(v, n=400):
    return (v or '').replace('\t', ' ').replace('\r', ' ').replace('\n', ' ').strip()[:n]


def count_people(s):
    if not s:
        return 0
    ini = re.findall(r'[А-ЯІЇЄҐ][\w\'’-]+\s+[А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ]\.', s)
    full = re.findall(r'[А-ЯІЇЄҐ][\w\'’-]+\s+[А-ЯІЇЄҐ][\w\'’-]+\s+[А-ЯІЇЄҐ][\w\'’-]+(?:ович|евич|івна|ївна|овна|ич|ична|інична)\b', s)
    return max(len(ini), len(full))


def norm_dt(s):
    s = (s or '').strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%d.%m.%Y %H:%M:%S', '%d.%m.%Y %H:%M', '%d.%m.%Y', '%d.%m.%y %H:%M:%S', '%d.%m.%y'):
        try:
            d = datetime.datetime.strptime(s[:19] if '%S' in f else s[:16] if '%M' in f else s[:10 if '%Y' in f else 8], f)
            return d.strftime('%Y-%m-%d %H:%M:%S' if '%H' in f else '%Y-%m-%d')
        except Exception:
            pass
    return s[:40]


def zvit_text(b64):
    """Текст «Звіту про автоматизований розподіл» з поля LOG_JUDGE_ASSIGN (base64-zip, .log, cp1251)."""
    if not b64:
        return ''
    try:
        z = zipfile.ZipFile(io.BytesIO(base64.b64decode(b64)))
        return z.read(z.namelist()[0]).decode('cp1251', 'replace')
    except Exception:
        return ''


def zvit_fields(t):
    """→ (початок розподілу, виключені з причинами, учасники, випадкове число, обраний суддя)."""
    if not t:
        return '', '', '', '', ''
    m = re.search(r'Дата та час початку автоматизованого розподілу:\s*([\d.]+\s+[\d:.]+)', t)
    start = m.group(1).strip() if m else ''
    excl = [re.sub(r'\s+', ' ', x).strip() for x in re.findall(r'(?m)^\s*Виключено\s+(.+?)\s*$', t)]
    part = [re.sub(r'\s+', ' ', x).strip() for x in re.findall(r'(?m)^\s*\d+\s{1,3}(\S.*?період з.*?)\s*$', t)]
    m = re.search(r'випадковим числом:\s*(\S+)', t)
    rnd = m.group(1) if m else ''
    chosen = []
    for m in re.finditer(r'Обрано суддю за випадковим числом:[^\n]*\n\s*(.+?)\s+період з', t):
        chosen.append(re.sub(r'\s+', ' ', m.group(1)).strip())
    return start, '; '.join(excl), '; '.join(part), rnd, '; '.join(dict.fromkeys(chosen))


def ts(v):
    """DOC_CREATE_DATE: Unix-час у секундах або мілісекундах → РРРР-ММ-ДД гг:хх:сс (Київ)."""
    try:
        n = int(v)
        return datetime.datetime.fromtimestamp(n / 1000 if n > 10 ** 11 else n, KYIV).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return v or ''


def row(x, day, src='', zdir=None):
    try:
        t = text(x['DOC_HTML'])
    except Exception:
        t = ''
    f = fields(t)
    h = HEAD.search(t)
    form = h.group(1) if h else ''
    ptype = h.group(2).strip() if h else ''
    court = f.get('court', '')
    if not court and h:
        m = re.match(r'\s*(.{5,200}?)\s+Єдиний унікальний', t[h.end():])
        court = m.group(1) if m else ''
    # date of entry: tail "dd.mm.yyyy /дата внесення.../" or "dd.mm.yy hh:mm:ss"
    tail = f.get('end_marker', '') or t[-200:]
    ms = list(DT_END.finditer(tail)); m = ms[-1] if ms else None
    entry = norm_dt(m.group(1) + (' ' + m.group(2) if m.group(2) else '')) if m else ''
    ex = f.get('excl', '')
    pa = f.get('part', '')
    exn = len(NAME_SEP.findall(ex)) if ex else 0
    if ex and exn == 0:
        exn = count_people(ex)
    pan = (len([p for p in pa.split(',') if p.strip()]) if ',' in pa else count_people(pa)) if pa else 0
    rj = f.get('rj', '') or f.get('appointed', '')
    ground = f.get('ground', '') or f.get('ground2', '') or ('Призначення не відбулося' if 'fail' in f else '')
    repl = ''
    if f.get('removed') or f.get('new'):
        repl = f"вилучені: {f.get('removed', '')} | нові: {f.get('new', '')}"
    cd = ts(x.get('DOC_CREATE_DATE', ''))
    zt = zvit_text(x.get('LOG_JUDGE_ASSIGN', ''))
    zs, zx, zp, zr, zo = zvit_fields(zt)
    if zt and zdir:
        os.makedirs(zdir, exist_ok=True)
        nm = re.sub(r'[^\w.-]', '_', f"{x.get('DOC_DBID', '')}_{x.get('DOC_ID', '')}")
        with open(os.path.join(zdir, nm + '.txt'), 'w', encoding='utf-8') as fz:
            fz.write(zt)
    return [day, x.get('EDRSR_CODE_2007', ''), clean(court, 200), clean(x.get('CAUSE_NUMBER', ''), 60), clean(x.get('CAUSE_PROC_NUM', ''), 60),
            form, clean(ptype, 120), norm_dt(f.get('start', '')) if f.get('start') else '', norm_dt(f.get('end', '')) if f.get('end') else '',
            entry, clean(ground, 300), clean(f.get('category', ''), 250), clean(f.get('subject', ''), 250), clean(rj, 150),
            clean(f.get('coll', ''), 400), exn, pan, clean(pa, 3000), clean(ex, 8000), clean(repl, 600), cd,
            x.get('DOC_DBID', ''), x.get('DOC_ID', ''), clean(x.get('JUDGES', ''), 400),
            src, norm_dt(zs) if zs else '', clean(zx, 8000), clean(zp, 4000), zr, clean(zo, 200)]


def _span(n):
    """Назва архіву РРРРММДДггххсс_РРРРММДДггххсс.zip → (день публікації РРРРММДД, тривалість у годинах) або None."""
    m = re.match(r'(\d{8})(\d{6})_(\d{8})(\d{6})\.zip$', n)
    if not m:
        return None
    try:
        a = dt.datetime.strptime(m[1] + m[2], '%Y%m%d%H%M%S'); b = dt.datetime.strptime(m[3] + m[4], '%Y%m%d%H%M%S')
    except ValueError:
        return None
    h = (b - a).total_seconds() / 3600
    return (m[1], h) if 0 < h <= 24 else None


def _pick(files):
    """files: [(назва, url, розмір)] → {день: [(url, назва)]}: добовий архів, а якщо його немає — погодинні;
    з кількох файлів з однаковою назвою — більший."""
    by = {}
    for n, u, size in files:
        sp = _span(n)
        if not sp:
            continue
        k = by.setdefault(sp[0], {})
        if n not in k or (size or 0) > (k[n][1] or 0):
            k[n] = (u, size)
    out = {}
    for d, k in by.items():
        daily = [n for n in k if _span(n)[1] == 24]
        names = daily[:1] if daily else sorted(k)
        out[d] = [(k[n][0], n) for n in names]
    return out


def resources_portal():
    j = json.loads(subprocess.run(['curl', '-sS', '-f', '-m', '120', f'https://data.gov.ua/api/3/action/package_show?id={PKG}'],
                                  check=True, capture_output=True).stdout)['result']
    return _pick([(r['name'], r['url'], r.get('size') or 0) for r in j['resources']])


def resources_dsa():
    j = json.loads(subprocess.run(['curl', '-sS', '-f', '-m', '180', '--retry', '2', DSA_JSON],
                                  check=True, capture_output=True).stdout)
    # розміру в паспорті немає: дублі назв розрізняються порядком (пізніший у паспорті — новіший)
    files = [(k, v, i) for i, f in enumerate(j.get('Файли', [])) for k, v in f.items()]
    return _pick(files)


def resources(src):
    """→ {день: [(url, назва, джерело)]}, примітка про джерело."""
    portal, dsa, note = {}, {}, ''
    if src in ('auto', 'portal'):
        try:
            portal = resources_portal()
        except Exception as e:
            if src == 'portal':
                raise SystemExit(f'data.gov.ua не відповідає: {e}')
            note = 'data.gov.ua не відповідає — дані із сайту ДСА (oddata/757)'
    if src in ('auto', 'dsa'):
        try:
            dsa = resources_dsa()
        except Exception as e:
            if not portal:
                raise SystemExit(f'Недоступні і data.gov.ua, і сайт ДСА: {e}')
            note = f'сайт ДСА не відповідає ({str(e)[:80]}) — лише data.gov.ua'
    res = {d: [(u, n, 'data.gov.ua') for u, n in v] for d, v in portal.items()}
    for d, v in dsa.items():
        if d not in res:
            res[d] = [(u, n, 'сайт ДСА') for u, n in v]
    return res, note


def pdate(s, name):
    """РРРР-ММ-ДД або ДД.ММ.РРРР → date (2.2)."""
    for f in ('%Y-%m-%d', '%d.%m.%Y'):
        try:
            return dt.datetime.strptime(s.strip(), f).date()
        except ValueError:
            pass
    sys.exit(f'{name}: не розпізнано дату «{s}» (потрібно РРРР-ММ-ДД або ДД.ММ.РРРР).')


def match(x, a):
    if a.case and case_key(x.get('CAUSE_NUMBER', '')) not in a.case:
        return False
    if a.court and str(x.get('EDRSR_CODE_2007', '')).lstrip('0') not in a.court:
        return False
    if a.judge and a.judge.lower() not in (x.get('JUDGES', '') or '').lower():
        return False
    return True


def work(day, files, a):
    """Усі архіви дня (добовий або погодинні) → рядки протоколів, що підходять під умови."""
    out, errs = [], []
    for url, name, src in files:
        z = os.path.join(a.cache, name)
        for att in range(3):
            rc = subprocess.run(['curl', '-sS', '-L', '-f', '-m', '900', '-o', z, url], capture_output=True, text=True)
            if rc.returncode == 0:
                break
            time.sleep(5 * (att + 1))
        else:
            errs.append(f'{name}: {rc.stderr[:100]}')
            continue
        try:
            with zipfile.ZipFile(z) as zf:
                for nm in zf.namelist():
                    if nm.endswith('.json'):
                        for x in json.loads(zf.read(nm)).get('items', []):
                            if match(x, a):
                                out.append(row(x, day, src, a.zvit_dir))
        except Exception as e:
            errs.append(f'{name}: {repr(e)[:100]}')
        finally:
            if not a.keep and os.path.exists(z):
                os.remove(z)
    if errs and not out and len(errs) == len(files):
        return day, None, '; '.join(errs)[:300]
    return day, out, '; '.join(errs)[:300]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--case', help='номер(и) справи через кому (точно)')
    ap.add_argument('--court', help='код(и) суду ЄДРСР через кому')
    ap.add_argument('--judge', help='частина ПІБ судді (поле JUDGES: «Прізвище Ім’я По батькові»)')
    ap.add_argument('--from', dest='dfrom'); ap.add_argument('--to', dest='dto')
    ap.add_argument('--around', help='дата, навколо якої шукати (РРРР-ММ-ДД)'); ap.add_argument('--days', type=int, default=3)
    ap.add_argument('--threads', type=int, default=2)
    ap.add_argument('--cache', default='protokoly_portal_cache')
    ap.add_argument('--keep', action='store_true', help='не видаляти завантажені архіви днів')
    ap.add_argument('--dzherelo', choices=['auto', 'portal', 'dsa'], default='auto',
                    help='auto — data.gov.ua, а дні, яких там немає (чи весь набір, якщо портал не відповідає), — сайт ДСА')
    ap.add_argument('--zvit-dir', help='тека для повних текстів звітів про автоматизований розподіл')
    ap.add_argument('--out')
    a = ap.parse_args()
    if not (a.case or a.court or a.judge):
        sys.exit('Задайте --case, --court або --judge.')
    if a.around:
        c = pdate(a.around, '--around')
        d1, d2 = c - dt.timedelta(days=1), c + dt.timedelta(days=a.days)
    elif a.dfrom:
        d1 = pdate(a.dfrom, '--from'); d2 = pdate(a.dto or a.dfrom, '--to')
    else:
        sys.exit('Задайте проміжок днів: --from/--to або --around.')
    if d2 < d1:
        sys.exit(f'Кінець проміжку ({d2}) раніше за початок ({d1}).')
    a.case = set(case_key(c) for c in a.case.split(',') if c.strip()) if a.case else None
    a.court = set(c.strip().lstrip('0') for c in a.court.split(',')) if a.court else None
    os.makedirs(a.cache, exist_ok=True)
    res, note = resources(a.dzherelo)
    if note:
        print(note, file=sys.stderr)
    days = [d.strftime('%Y%m%d') for d in (d1 + dt.timedelta(n) for n in range((d2 - d1).days + 1))]
    have = [d for d in days if d in res]
    srcs = sorted({f[2] for d in have for f in res[d]})
    print('Джерело: ' + (', '.join(srcs) or '—') + '; погодинні архіви: ' +
          str(sum(1 for d in have if len(res[d]) > 1)) + ' дн.', file=sys.stderr)
    miss = [d for d in days if d not in res]
    print(f'Днів у проміжку: {len(days)}, є в наборі: {len(have)}' + (f'; відсутні: {", ".join(miss[:10])}' + ('…' if len(miss) > 10 else '') if miss else ''), file=sys.stderr)
    if len(have) > 60:
        print('Увага: понад 60 днів — це довго (~10–25 с на день). Звузьте проміжок або скористайтеся покажчиком на Диску.', file=sys.stderr)
    rows, t, failed = [], time.time(), []
    with ThreadPoolExecutor(a.threads) as ex:
        for day, out, err in ex.map(lambda d: work(d, res[d], a), have):
            if out is None:
                failed.append(day)
                print(f'  {day}: ПОМИЛКА {err}', file=sys.stderr)
            else:
                rows += out
                if err:
                    failed.append(day + ' (частково)')
                print(f'  {day}: знайдено {len(out)}' + (f' (частково, помилки: {err})' if err else ''), file=sys.stderr)
    seen, uniq = set(), []
    for r in rows:
        k = (r[F.index('doc_dbid')], r[F.index('doc_id')])
        if k in seen and k != ('', ''):
            continue
        seen.add(k); uniq.append(r)
    if len(uniq) < len(rows):
        print(f'Повторів одного протоколу відкинуто: {len(rows) - len(uniq)}', file=sys.stderr)
    rows = uniq
    rows.sort(key=lambda r: (r[7] or r[9] or r[20], r[20]))
    print(f'Разом {len(rows)} протоколів за {time.time() - t:.0f} с', file=sys.stderr)
    if a.out:
        with open(a.out, 'w', encoding='utf-8', newline='') as f:
            w = csv.writer(f, delimiter='\t', lineterminator='\n'); w.writerow(F); w.writerows(rows)
        print(f'  → {a.out}', file=sys.stderr)
    else:
        show = ['day_file', 'court_code', 'case_number', 'proc_number', 'doc_form', 'protocol_type', 'distribution_start', 'entry_date',
                'ground', 'reporting_judge', 'judges_field', 'excluded_judges', 'dzherelo', 'zvit_vykliucheni', 'zvit_obrano']
        idx = [F.index(c) for c in show]
        print('\t'.join(show))
        for r in rows[:200]:
            print('\t'.join(str(r[i])[:300] for i in idx))
        if len(rows) > 200:
            print(f'… показано 200 з {len(rows)}; повний результат — з параметром --out файл.tsv')
    # неповнота результату — у stdout: «0 протоколів» не має читатися як «протоколів не було»
    if miss:
        last = max(res) if res else '—'
        print(f'УВАГА: {len(miss)} з {len(days)} днів проміжку немає в наборі ({", ".join(miss[:6])}{"…" if len(miss) > 6 else ""}); '
              f'останній день у наборі — {last}. Протоколи, опубліковані у відсутні дні, не перевірено.')
    if failed:
        print(f'УВАГА: архіви не завантажено або завантажено частково: {", ".join(failed[:10])} — результат неповний, повторіть запит.')


if __name__ == '__main__':
    main()
