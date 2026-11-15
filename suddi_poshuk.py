#!/usr/bin/env python3
"""Пошук судді у «Списку суддів» ВККСУ (суд призначення, стать, досьє) і перевірка статусу судді.

Автор: В'ячеслав Плескач. Версія 1.2 (14.11.2026): сумісність з edrsr_vybirka.py 3.0 (кеш ~/.cache/edrsr, parquet); 15.11.2026 — уточнено
статус, коли рішення знайдено лише в іншому суді.

Використання:
    python3 suddi_poshuk.py "Прізвище [І.Б. | Ім'я По батькові]" ["суд, указаний у скарзі / рішенні"]
        [--url URL] [--bez-edrsr] [--dni 60]
Приклади:
    python3 suddi_poshuk.py "Свіягіна І.М." "Дніпровський апеляційний суд"
    python3 suddi_poshuk.py "Шевчук Юлія Валеріївна" --bez-edrsr

Що робить (одним запуском):
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


def court_key(s):
    s = norm(s).replace('"', '').replace('«', '').replace('»', '')
    s = re.sub(r'\bм\.\s*', 'міста ', s)
    return re.sub(r'[^a-zа-яіїєґ\' ]', ' ', s)


def same_court(a, b):
    a, b = court_key(a), court_key(b)
    return a in b or b in a or difflib.SequenceMatcher(None, a, b).ratio() > 0.88


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
    found = match(rows, a.pib)
    warn = []
    print(f'Список ВККСУ: {len(rows)} суддів ({os.path.basename(cache)})')
    if not found:
        print(f'«{a.pib}» у списку не знайдено (перевірте написання, відмінок, зміну прізвища).')
        if not a.url:
            for d in dates[1:6]:
                try:
                    older = match(load(URL_TPL.format(d))[0], a.pib)
                except Exception:
                    continue
                if older:
                    print(f'  У списку на {d} суддя був: ' + '; '.join(f"{r['pib']} — {r['sud']}" for r in older) + '.')
                    warn.append(f'суддя зник зі списку ВККСУ (є у списку на {d}, немає у списку на {dates[0]})')
                    found = older
                    break
            else:
                warn.append(f'судді немає у списку ВККСУ на {dates[0]} і в попередніх списках')
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
        s, rest = parse_query(found[0]['pib'] if len(found) == 1 else a.pib)
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
    if warn:
        print('⚠ СТАТУС: суддя міг бути звільнений або тривало не здійснює правосуддя — ' + '; '.join(warn) + '.')
        print('  Перевірити (рішення ВРП про звільнення, укази, hcj.gov.ua) і запитати інспектора, чи продовжувати.')
    else:
        print('✓ СТАТУС: суддя є в чинному списку ВККСУ' + ('' if a.bez_edrsr else ' і ухвалює рішення') + '.')


if __name__ == '__main__':
    main()
