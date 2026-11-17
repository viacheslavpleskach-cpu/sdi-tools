#!/usr/bin/env python3
"""Пошук судді у «Списку суддів» ВККСУ (суд призначення, стать, досьє) і перевірка статусу судді.

Автор: В'ячеслав Плескач. Версія 1.3 (17.11.2026): суддю шукає за будь-яким написанням ПІБ у скарзі (порядок слів,
ініціали, відмінок, російське написання, описки, частина подвійного прізвища, лише прізвище) і за судом, якщо прізвище
не збігається; рядки «ПІБ за списком ВККСУ» і «✓ / ⚠ ПІБ». 1.2 (14–15.11.2026): сумісність з edrsr_vybirka.py 3.0,
статус, коли рішення знайдено лише в іншому суді.

Використання:
    python3 suddi_poshuk.py "Прізвище [І.Б. | Ім'я По батькові]" ["суд, указаний у скарзі / рішенні"]
        [--url URL] [--bez-edrsr] [--dni 60]
Приклади:
    python3 suddi_poshuk.py "Свіягіна І.М." "Дніпровський апеляційний суд"
    python3 suddi_poshuk.py "Шевчук Юлія Валеріївна" --bez-edrsr
    python3 suddi_poshuk.py "судді Свиягиной И.Н." "Дніпровський апеляційний суд"   # будь-яке написання зі скарги

Що робить (одним запуском):
0. ПІБ передається так, як його написав скаржник (без виправлень). Спершу — точний збіг; якщо його немає —
   нечіткий: інший порядок слів, ініціали замість імені, непрямий відмінок («Свіягіної І.М.»), російське написання
   («Свиягина»), описка, одна частина подвійного прізвища; а якщо вказано суд — особи цього суду з тим самим
   ім'ям та по батькові (можлива зміна прізвища). Підсумок: «ПІБ за списком ВККСУ: …» і рядок «✓ ПІБ» (збіг) або
   «⚠ ПІБ» (написання в скарзі інше / кілька кандидатів / не знайдено) — у висновку ПІБ і суд лише за списком.
1. Знаходить найсвіжіший «Список суддів» ВККСУ (spysok_suddiv_na_ДД.ММ.РРРР.pdf; перебирає дати назад від
   сьогодні, до 120 днів) і шукає в ньому суддю: суд призначення, стать, досьє. Якщо суд зі скарги інший, але
   особу знайдено в суді того самого виду, — ймовірне відрядження (у рішеннях і скаргах воно не зазначається).
2. СТАТУС СУДДІ:
   - якщо в найсвіжішому списку судді немає — шукає в попередніх списках і повідомляє, з якого списку суддя зник;
   - дата останнього рішення судді в ЄДРСР (набір data.gov.ua поточного року, а якщо там нічого — і попереднього).
     Використовує edrsr_vybirka.py (має лежати поруч) і його кеш ~/.cache/edrsr — спільний з усіма вибірками ЄДРСР.
     Якщо остання дата давніша за --dni днів (типово 60) — попередження.
   Підсумковий рядок «⚠ СТАТУС: …» означає, що суддя міг бути звільнений (або тривало не здійснює правосуддя:
   відпустка, відсторонення, мобілізація, закінчення строку повноважень) — це треба перевірити до розгляду скарги.
   Рядок «✓ СТАТУС: …» — підстав для такого припущення немає.

Списки кешуються у suddi_<дата>.csv поруч зі скриптом; перелік доступних списків — у suddi_spysky_<сьогодні>.txt.
"""
import argparse, csv, datetime as dt, difflib, os, re, subprocess, sys
from concurrent.futures import ThreadPoolExecutor

URL_TPL = 'https://vkksu.gov.ua/sites/default/files/field/pdf/spysok_suddiv_na_{}.pdf'
KNOWN = ['09.11.2026', '02.11.2026', '26.10.2026', '19.10.2026']   # перевірені списки (13.11.2026)
HERE = os.path.dirname(os.path.abspath(__file__))


def norm(s):
    s = s.lower().replace('’', "'").replace('`', "'").replace('ʼ', "'")
    return re.sub(r'\s+', ' ', s).strip()


def load(url):
    tag = re.search(r'(\d{2}\.\d{2}\.\d{4})', url)
    cache = os.path.join(HERE, f'suddi_{tag.group(1) if tag else "list"}.csv')
    if os.path.exists(cache):
        with open(cache, encoding='utf-8') as f:
            return list(csv.DictReader(f)), cache
    pdf = cache[:-4] + '.pdf'
    subprocess.run(['curl', '-sS', '-f', '-m', '180', '-o', pdf, url], check=True)
    txt = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, check=True).stdout.decode()
    rows = []
    pat = re.compile(r"^\s*(\d+)\s+(\d{5})\s+(.+?)\s{2,}([ЖЧ])\s{2,}(.+?)\s*$")
    last = None
    for line in txt.splitlines():
        m = pat.match(line)
        if m:
            last = {'n': m[1], 'dosie': m[2], 'pib': m[3].strip(), 'stat': m[4], 'sud': m[5].strip()}
            rows.append(last)
        elif last and line.strip() and line.startswith(' ' * 60) and not re.search(r'СПИСОК|досьє|Найменування', line):
            last['sud'] += ' ' + line.strip()   # перенесена назва суду
    with open(cache, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['n', 'dosie', 'pib', 'stat', 'sud'])
        w.writeheader(); w.writerows(rows)
    os.remove(pdf)
    return rows, cache


def lists_available(days=120):
    """Дати опублікованих списків ВККСУ (від найновішої). Результат кешується на добу."""
    today = dt.date.today()
    cf = os.path.join(HERE, f'suddi_spysky_{today:%Y%m%d}.txt')
    if os.path.exists(cf):
        return open(cf).read().split()
    dates = [(today - dt.timedelta(d)).strftime('%d.%m.%Y') for d in range(days + 1)]

    def ok(d):
        r = subprocess.run(['curl', '-sS', '-I', '-m', '20', '-o', '/dev/null', '-w', '%{http_code}', URL_TPL.format(d)],
                           capture_output=True, text=True)
        return d if r.stdout.strip() == '200' else None
    with ThreadPoolExecutor(16) as ex:
        found = [d for d in ex.map(ok, dates) if d]
    found += [d for d in KNOWN if d not in found]
    found.sort(key=lambda d: dt.datetime.strptime(d, '%d.%m.%Y'), reverse=True)
    if found:
        open(cf, 'w').write('\n'.join(found))
    return found


def kind(sud):
    s = norm(sud)
    for k, keys in [('Верховний Суд', ['верховного суду', 'касаційний']),
                    ('апеляційний адміністративний', ['апеляційний адміністративний']),
                    ('апеляційний господарський', ['апеляційний господарський']),
                    ('окружний адміністративний', ['окружний адміністративний']),
                    ('господарський', ['господарський суд']),
                    ('апеляційний', ['апеляційний суд']),
                    ('вищий спеціалізований', ['вищий антикорупційний', 'вищий суд з питань']),
                    ('місцевий загальний', ['районний', 'міськрайонний', 'міський суд'])]:
        if any(x in s for x in keys):
            return k
    return 'інший'


def parse_query(query):
    q = norm(query).replace('.', '. ').split()
    return q[0], [x.strip('.') for x in q[1:] if x.strip('.')]


def name_ok(pib, surname, rest):
    p = norm(pib).replace('.', '. ').split()
    p = [x.strip('.') for x in p if x.strip('.')]
    if not p or p[0] != surname:
        return False
    for i, part in enumerate(rest[:2]):
        if i + 1 >= len(p) or not (p[i + 1].startswith(part) or part.startswith(p[i + 1])):
            return False
    return True


def match(rows, query):
    s, rest = parse_query(query)
    return [r for r in rows if name_ok(r['pib'], s, rest)]


# ---- нечіткий пошук (версія 1.3): ПІБ так, як його написав скаржник ----
SLUZHBOVI = {'суддя', 'судді', 'суддею', 'суддю', 'суддів', 'суддям', 'суддями', 'головуючий', 'головуюча',
             'головуючого', 'головуючої', 'головуючим', 'головуючою', 'доповідач', 'доповідача', 'доповідачем',
             'гр', 'громадянина', 'громадянки', 'пан', 'пані', 'та', 'щодо', 'стосовно',
             'на', 'дії', 'дій', 'скарга', 'суддей', 'судья', 'судьи', 'судьей', 'судью'}
RU_UK = str.maketrans({'ы': 'и', 'э': 'е', 'ё': 'е', 'ъ': "'"})


def key(w):
    """Ключ для порівняння без урахування українського / російського написання: є→е, ї→і, и→і, ґ→г, без апострофа."""
    w = norm(w).translate(RU_UK)
    w = w.replace('є', 'е').replace('ї', 'і').replace('и', 'і').replace('ґ', 'г').replace("'", '')
    return re.sub(r'[^a-zа-яі-]', '', w)


def tokens(query):
    q = norm(query).replace('.', '. ').replace(',', ' ').replace('«', ' ').replace('»', ' ').replace('"', ' ')
    out = []
    for t in q.split():
        t = t.strip('.()')
        if t and key(t) and t not in SLUZHBOVI:
            out.append(t)
    return out


def word_sim(q, part):
    """Схожість слова зі скарги з частиною ПІБ у списку (1.0 — точно; ~0.9 — відмінок / написання; далі — описки)."""
    if norm(q) == norm(part):
        return 1.0
    a, b = key(q), key(part)
    if not a or not b:
        return 0.0
    if a == b:
        return 0.95
    pre = len(os.path.commonprefix([a, b]))
    if pre >= 3 and pre >= len(b) - 2 and pre >= len(a) - 3:   # відмінкові закінчення: Свіягіної / Свіягіна, Ткача / Ткач
        return 0.9
    return difflib.SequenceMatcher(None, a, b).ratio()


def surname_sim(q, surname):
    best = word_sim(q, surname)
    if '-' in surname:   # подвійне прізвище: у скарзі часто лише одна частина
        best = max([best] + [0.88 * word_sim(q, x) for x in surname.split('-') if x])
    return best


def parts(pib):
    p = [x.strip('.') for x in norm(pib).replace('.', '. ').split() if x.strip('.')]
    return (p + ['', '', ''])[:3] if p else ['', '', '']


def given_sim(toks, name, patr):
    """Ім'я та по батькові: ініціали чи повні слова в порядку «ім'я, по батькові» (будь-яка форма, відмінок)."""
    sims = []
    for t, part in zip(toks, [name, patr]):
        if not part:
            return 0.0
        if len(key(t)) == 1:
            sims.append(1.0 if key(part).startswith(key(t)) else 0.0)
        else:
            sims.append(word_sim(t, part))
    if len(toks) > 2:
        sims.append(0.0)
    return min(sims) if sims else 1.0


def fuzzy(rows, query, sud=None, porig=0.8):
    """[(оцінка, рядок, причина)] за спаданням оцінки. Прізвище — будь-яке слово запиту (не ініціал)."""
    toks = tokens(query)
    words = [i for i, t in enumerate(toks) if len(key(t)) > 1]
    out = {}
    for r in rows:
        sur, name, patr = parts(r['pib'])
        best = None
        for i in words:
            ss = surname_sim(toks[i], sur)
            if ss < porig:
                continue
            g = given_sim([t for j, t in enumerate(toks) if j != i], name, patr)
            if g < porig:
                continue
            sc = min(ss, g)
            if best is None or sc > best[0]:
                best = (sc, 'збіг ПІБ' if sc == 1.0 else 'інше написання ПІБ (відмінок, порядок слів, описка, мова)')
        if best:
            out[r['dosie'] + r['sud']] = (best[0], r, best[1])
    if sud:   # прізвище не збігається, але в суді зі скарги є особа з тим самим ім'ям та по батькові
        for r in rows:
            if not same_court(sud, r['sud']):
                continue
            sur, name, patr = parts(r['pib'])
            for i in words:
                rest = [t for j, t in enumerate(toks) if j != i]
                if len(rest) >= 2 and all(len(key(t)) > 1 for t in rest[:2]) and given_sim(rest, name, patr) >= 0.9:
                    k = r['dosie'] + r['sud']
                    if k not in out:
                        out[k] = (0.7, r, 'у суді зі скарги — те саме ім\'я та по батькові, інше прізвище (зміна прізвища?)')
            if not words or len(toks) == 2 and len(words) == 2:   # прізвища в скарзі немає («Наталія Сергіївна»)
                if len(toks) >= 2 and given_sim(toks[:2], name, patr) >= 0.9:
                    out.setdefault(r['dosie'] + r['sud'], (0.7, r, 'у скарзі немає прізвища; ім\'я та по батькові збігаються'))
    if not out:   # останній рівень: збігається лише прізвище (ініціали інші — описка, російські ініціали «А.» замість «О.»)
        for r in rows:
            sur = parts(r['pib'])[0]
            if any(surname_sim(toks[i], sur) >= 0.9 for i in words):
                out[r['dosie'] + r['sud']] = (0.6, r, 'збігається лише прізвище, ім\'я / по батькові інші')
    res = sorted(out.values(), key=lambda x: (-x[0], x[1]['pib']))
    if sud:   # спершу — суд зі скарги, потім суди того самого виду
        res.sort(key=lambda x: (not same_court(sud, x[1]['sud']), kind(x[1]['sud']) != kind(sud), -x[0]))
    return res


def find(rows, query, sud=None):
    """(знайдені рядки, примітка про ПІБ або None). Спершу точний пошук, потім нечіткий."""
    toks = tokens(query)
    exact = match(rows, ' '.join(toks)) if toks and len(key(toks[0])) > 1 else []
    if exact:
        return exact, None
    fz = fuzzy(rows, query, sud)
    if not fz:
        return [], None
    full = [x for x in fz if x[0] == 1.0]   # точний збіг, лише інший порядок слів
    if full and (not sud or any(same_court(sud, x[1]['sud']) for x in full) or len(full) == 1):
        here = [x[1] for x in full if sud and same_court(sud, x[1]['sud'])]
        return here or [x[1] for x in full], None
    if sud:
        here = [x for x in fz if same_court(sud, x[1]['sud'])]
        if len(here) == 1:
            return [here[0][1]], here[0][2]
        if len(here) > 1:
            return [x[1] for x in here], 'кілька кандидатів у суді зі скарги'
    if sud and fz[0][0] < 0.7:   # лише прізвище: брати тільки суди того самого виду
        fz = [x for x in fz if kind(x[1]['sud']) == kind(sud)] or fz
    top = [x for x in fz if x[0] >= 0.85]
    if len(top) == 1:
        return [top[0][1]], top[0][2]
    return [x[1] for x in (top or fz)[:8]], 'кілька можливих кандидатів' if len(top or fz) > 1 else (top or fz)[0][2]


def court_key(s):
    s = norm(s).replace('"', '').replace('«', '').replace('»', '')
    s = re.sub(r'\bм\.\s*', 'міста ', s)
    return re.sub(r'[^a-zа-яіїєґ\' ]', ' ', s)


def same_court(a, b):
    a, b = court_key(a), court_key(b)
    if a in b or b in a:
        return True
    # перше слово назви (Шевченківський / Печерський) має збігатися: спільний хвіст «районний суд міста Києва» не рахується
    fa, fb = (a.split() or [''])[0], (b.split() or [''])[0]
    return difflib.SequenceMatcher(None, fa, fb).ratio() > 0.8 and difflib.SequenceMatcher(None, a, b).ratio() > 0.88


def edrsr_last(surname, rest):
    """[(суддя, суд, остання дата, к-сть рішень)] за роком надходження: поточний, а якщо порожньо — попередній.
    Працює з edrsr_vybirka.py 3.x: year_dir() → тека з documents.parquet і довідниками CSV."""
    sys.path.insert(0, HERE)
    import edrsr_vybirka as ev
    if not hasattr(ev, 'year_dir'):
        sys.exit('edrsr_vybirka.py застарів (потрібна версія 3.x з year_dir). Запустіть get.sh з sdi-tools.')
    duckdb = ev.need('duckdb')
    y = dt.date.today().year
    out, used = [], []
    for year in (y, y - 1):
        d, src = ev.year_dir(year)
        used.append(f'{year}: {src}')
        con = duckdb.connect()
        rows = con.execute(f"""SELECT x.judge, coalesce(c.name, x.court_code), max(substr(x.adjudication_date,1,10)), count(*)
            FROM read_parquet('{d}/documents.parquet') x LEFT JOIN {ev.rd(d, 'courts.csv')} c ON c.court_code = x.court_code
            WHERE lower(x.judge) LIKE ? GROUP BY 1, 2""", [surname.replace("'", "_") + ' %']).fetchall()
        out = [r for r in rows if name_ok(r[0], surname, rest)]
        if out:
            break
    return sorted(out, key=lambda r: r[2], reverse=True), used


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pib')
    ap.add_argument('sud', nargs='?')
    ap.add_argument('--url', help='конкретний список ВККСУ (інакше — найсвіжіший)')
    ap.add_argument('--bez-edrsr', action='store_true', help='не перевіряти дату останнього рішення в ЄДРСР')
    ap.add_argument('--dni', type=int, default=60, help='поріг «давно не ухвалює рішень», днів (типово 60)')
    ap.add_argument('--cache', help=argparse.SUPPRESS)   # сумісність з версією 1.1; кеш ЄДРСР — ~/.cache/edrsr
    a = ap.parse_args()

    dates = [re.search(r'(\d{2}\.\d{2}\.\d{4})', a.url).group(1)] if a.url else lists_available()
    url = a.url or URL_TPL.format(dates[0])
    rows, cache = load(url)
    found, pib_note = find(rows, a.pib, a.sud)
    warn = []
    print(f'Список ВККСУ: {len(rows)} суддів ({os.path.basename(cache)})')
    print(f'ПІБ у скарзі: «{a.pib}»')
    if not found:
        print(f'«{a.pib}» у списку не знайдено ні точно, ні за іншим написанням.')
        if not a.url:
            for d in dates[1:6]:
                try:
                    older, note = find(load(URL_TPL.format(d))[0], a.pib, a.sud)
                except Exception:
                    continue
                if older:
                    print(f'  У списку на {d} суддя був: ' + '; '.join(f"{r['pib']} — {r['sud']}" for r in older) + '.')
                    warn.append(f'суддя зник зі списку ВККСУ (є у списку на {d}, немає у списку на {dates[0]})')
                    found, pib_note = older, note
                    break
            else:
                warn.append(f'судді немає у списку ВККСУ на {dates[0]} і в попередніх списках')
        else:
            warn.append('судді немає у вказаному списку ВККСУ')
    if pib_note:
        print(f'≈ Точного збігу немає — {pib_note}:')
    for r in found:
        stat = 'жінка' if r['stat'] == 'Ж' else 'чоловік'
        print(f"- {r['pib']} — {r['sud']} (стать: {stat}; досьє № {r['dosie']})")
    if a.sud and found and not warn:
        same = [r for r in found if same_court(a.sud, r['sud'])]
        if same:
            print(f'✓ Суддя значиться у суді «{a.sud}».')
        else:
            k = kind(a.sud)
            cand = [r for r in found if kind(r['sud']) == k]
            print(f'⚠ У суді «{a.sud}» особи з таким ПІБ немає.')
            if cand:
                print('  Ймовірне відрядження. Суд призначення: ' + '; '.join(r['sud'] for r in cand) + '.' +
                      ' У таблиці висновку: «[ПІБ] - [суд призначення] (відряджена / відряджений до [суд зі скарги])».')
            else:
                print('  Особи з таким ПІБ у суді того самого виду немає — перевірте вручну.')

    if not a.bez_edrsr:
        s, rest = parse_query(found[0]['pib'] if len(found) == 1 else ' '.join(tokens(a.pib)) or a.pib)
        hits, used = edrsr_last(s, rest)
        courts = [r['sud'] for r in found] + ([a.sud] if a.sud else [])
        mine = [h for h in hits if any(same_court(h[1], c) for c in courts)] if courts else hits
        print('ЄДРСР (' + '; '.join(used) + '):')
        for h in hits[:6]:
            mark = '*' if h in mine else ' '
            print(f'  {mark} {h[0]} — {h[1]}: останнє рішення {dt.date.fromisoformat(h[2]):%d.%m.%Y}, рішень {h[3]}')
        if len(hits) > 6:
            print(f'    … ще {len(hits) - 6} збігів (тезки)')
        if courts and hits and not mine:
            print('  (у суді призначення / суді зі скарги рішень немає — збіги вище можуть бути тезками)')
        if not mine and hits:
            last = max(h[2] for h in hits)
            warn.append(f'у суді призначення рішень у ЄДРСР немає; збіги ПІБ в інших судах (відрядження або тезки) — '
                        f'останнє рішення {dt.date.fromisoformat(last):%d.%m.%Y}')
        elif not mine:
            warn.append('рішень судді в ЄДРСР за поточний і попередній роки не знайдено')
        else:
            last = max(h[2] for h in mine)
            days = (dt.date.today() - dt.date.fromisoformat(last)).days
            print(f'  Останнє рішення судді (* — суд призначення / суд зі скарги): {dt.date.fromisoformat(last):%d.%m.%Y}, {days} дн. тому.')
            if days > a.dni:
                warn.append(f'останнє рішення в ЄДРСР — {dt.date.fromisoformat(last):%d.%m.%Y} ({days} дн. тому, поріг {a.dni})')

    if len(found) > 1:
        print('⚠ У списку кілька осіб з таким ПІБ — статус нижче оцінено сукупно; уточніть ініціали або суд.')
    if len(found) == 1 and not pib_note:
        print(f"✓ ПІБ: «{found[0]['pib']}» — збігається зі списком ВККСУ.")
    elif len(found) == 1:
        print(f"⚠ ПІБ: у скарзі «{a.pib}», у списку ВККСУ — «{found[0]['pib']}» ({pib_note}). "
              'У висновку — ПІБ за списком; розбіжність — у звіт.')
    elif found:
        print(f'⚠ ПІБ: однозначно встановити суддю за написанням у скарзі не вдалося ({pib_note or "кілька осіб"}) — '
              'уточнити за текстами рішень у справі (ЄДРСР) або запитати інспектора.')
    else:
        print(f'⚠ ПІБ: «{a.pib}» не знайдено в списку ВККСУ за жодним написанням.')
    if len(found) == 1:
        print(f"ПІБ за списком ВККСУ: {found[0]['pib']} — {found[0]['sud']}")
    if warn:
        print('⚠ СТАТУС: суддя міг бути звільнений або тривало не здійснює правосуддя — ' + '; '.join(warn) + '.')
        print('  Перевірити (рішення ВРП про звільнення, укази, hcj.gov.ua) і запитати інспектора, чи продовжувати.')
    else:
        print('✓ СТАТУС: суддя є в чинному списку ВККСУ' + ('' if a.bez_edrsr else ' і ухвалює рішення') + '.')


if __name__ == '__main__':
    main()
