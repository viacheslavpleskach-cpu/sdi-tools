#!/usr/bin/env python3
"""Пошук протоколів автоматизованого розподілу прямо в наборі data.gov.ua (без покажчика на Диску).

Автор: В'ячеслав Плескач

Набір: «Протоколи автоматизованого розподілу судових справ між суддями»
https://data.gov.ua/dataset/0eb7722b-2d0e-4d28-86b8-ba0a45cae12d — один архів на кожен день ПУБЛІКАЦІЇ
(фактично 16.11.2023–08.02.2026; ~30–90 МБ на будній день). Скрипт завантажує архіви лише за заданий
проміжок днів, відбирає протоколи за номером справи / судом / суддею і розбирає їх текст.

Оскільки архіви — за днями, проміжок треба задати: --from/--to (дні публікації) або --around ДАТА --days N.
Протокол зазвичай публікується в день розподілу або наступного дня; дату реєстрації справи беріть
з ЄДРСР (edrsr_vybirka.py) чи з набору стану розгляду справ.

Приклади:
  python3 protokoly_portal.py --case "757/30191/25-к" --around 2025-06-26 --days 3
  python3 protokoly_portal.py --court 2606 --judge "Вовк" --from 2025-06-23 --to 2025-06-27 --out vovk.tsv
Швидкість: ~10–25 с на день (завантаження + відбір), 2 дні паралельно.
"""
import argparse, csv, json, os, subprocess, sys, time, datetime as dt
from concurrent.futures import ThreadPoolExecutor
PKG = '0eb7722b-2d0e-4d28-86b8-ba0a45cae12d'
import base64, zipfile, io, re, html, datetime
from zoneinfo import ZoneInfo
KYIV = ZoneInfo('Europe/Kyiv')

F = ['day_file', 'court_code', 'court_name', 'case_number', 'proc_number', 'doc_form', 'protocol_type',
     'distribution_start', 'distribution_end', 'entry_date', 'ground', 'category', 'subject', 'reporting_judge',
     'collegium', 'excluded_count', 'participating_count', 'participating_judges', 'excluded_judges',
     'replaced_judges', 'doc_create_date', 'doc_dbid', 'doc_id', 'judges_field']


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


def fields(t):
    out = {}
    ms = list(LRE.finditer(t))
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


def row(x, day):
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
    cd = x.get('DOC_CREATE_DATE', '')
    try:
        cd = datetime.datetime.fromtimestamp(int(cd) / 1000, KYIV).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        pass
    return [day, x.get('EDRSR_CODE_2007', ''), clean(court, 200), clean(x.get('CAUSE_NUMBER', ''), 60), clean(x.get('CAUSE_PROC_NUM', ''), 60),
            form, clean(ptype, 120), norm_dt(f.get('start', '')) if f.get('start') else '', norm_dt(f.get('end', '')) if f.get('end') else '',
            entry, clean(ground, 300), clean(f.get('category', ''), 250), clean(f.get('subject', ''), 250), clean(rj, 150),
            clean(f.get('coll', ''), 400), exn, pan, clean(pa, 3000), clean(ex, 8000), clean(repl, 600), cd,
            x.get('DOC_DBID', ''), x.get('DOC_ID', ''), clean(x.get('JUDGES', ''), 400)]


def resources():
    j = json.loads(subprocess.run(['curl', '-sS', '-f', '-m', '120', f'https://data.gov.ua/api/3/action/package_show?id={PKG}'],
                                  check=True, capture_output=True).stdout)['result']
    best = {}
    for r in j['resources']:
        n = r['name']
        try:
            a = dt.date(int(n[:4]), int(n[4:6]), int(n[6:8])); b = dt.date(int(n[15:19]), int(n[19:21]), int(n[21:23]))
        except ValueError:
            continue
        if (b - a).days != 1:
            continue
        if n[:8] not in best or (r.get('size') or 0) > (best[n[:8]].get('size') or 0):
            best[n[:8]] = r
    return {k: v['url'] for k, v in best.items()}


def match(x, a):
    if a.case and x.get('CAUSE_NUMBER', '') not in a.case:
        return False
    if a.court and str(x.get('EDRSR_CODE_2007', '')).lstrip('0') not in a.court:
        return False
    if a.judge and a.judge.lower() not in (x.get('JUDGES', '') or '').lower():
        return False
    return True


def work(day, url, a):
    z = os.path.join(a.cache, f'{day}.zip')
    for att in range(3):
        rc = subprocess.run(['curl', '-sS', '-L', '-f', '-m', '900', '-o', z, url], capture_output=True, text=True)
        if rc.returncode == 0:
            break
        time.sleep(5 * (att + 1))
    else:
        return day, None, rc.stderr[:120]
    import zipfile
    out = []
    try:
        with zipfile.ZipFile(z) as zf:
            for nm in zf.namelist():
                if nm.endswith('.json'):
                    for x in json.loads(zf.read(nm)).get('items', []):
                        if match(x, a):
                            out.append(row(x, day))
    except Exception as e:
        return day, None, repr(e)[:120]
    finally:
        if not a.keep and os.path.exists(z):
            os.remove(z)
    return day, out, ''


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
    ap.add_argument('--out')
    a = ap.parse_args()
    if not (a.case or a.court or a.judge):
        sys.exit('Задайте --case, --court або --judge.')
    if a.around:
        c = dt.date.fromisoformat(a.around)
        d1, d2 = c - dt.timedelta(days=1), c + dt.timedelta(days=a.days)
    elif a.dfrom:
        d1 = dt.date.fromisoformat(a.dfrom); d2 = dt.date.fromisoformat(a.dto or a.dfrom)
    else:
        sys.exit('Задайте проміжок днів: --from/--to або --around.')
    a.case = set(a.case.split(',')) if a.case else None
    a.court = set(c.strip().lstrip('0') for c in a.court.split(',')) if a.court else None
    os.makedirs(a.cache, exist_ok=True)
    res = resources()
    days = [d.strftime('%Y%m%d') for d in (d1 + dt.timedelta(n) for n in range((d2 - d1).days + 1))]
    have = [d for d in days if d in res]
    miss = [d for d in days if d not in res]
    print(f'Днів у проміжку: {len(days)}, є в наборі: {len(have)}' + (f'; відсутні: {", ".join(miss[:10])}' + ('…' if len(miss) > 10 else '') if miss else ''), file=sys.stderr)
    if len(have) > 60:
        print('Увага: понад 60 днів — це довго (~10–25 с на день). Звузьте проміжок або скористайтеся покажчиком на Диску.', file=sys.stderr)
    rows, t = [], time.time()
    with ThreadPoolExecutor(a.threads) as ex:
        for day, out, err in ex.map(lambda d: work(d, res[d], a), have):
            if out is None:
                print(f'  {day}: ПОМИЛКА {err}', file=sys.stderr)
            else:
                rows += out
                print(f'  {day}: знайдено {len(out)}', file=sys.stderr)
    rows.sort(key=lambda r: (r[7] or r[9] or r[20], r[20]))
    print(f'Разом {len(rows)} протоколів за {time.time() - t:.0f} с', file=sys.stderr)
    if a.out:
        with open(a.out, 'w', encoding='utf-8', newline='') as f:
            w = csv.writer(f, delimiter='\t', lineterminator='\n'); w.writerow(F); w.writerows(rows)
        print(f'  → {a.out}', file=sys.stderr)
    else:
        show = ['day_file', 'court_code', 'case_number', 'proc_number', 'doc_form', 'protocol_type', 'distribution_start', 'entry_date',
                'ground', 'reporting_judge', 'judges_field', 'excluded_judges']
        idx = [F.index(c) for c in show]
        print('\t'.join(show))
        for r in rows[:200]:
            print('\t'.join(str(r[i])[:300] for i in idx))


if __name__ == '__main__':
    main()
