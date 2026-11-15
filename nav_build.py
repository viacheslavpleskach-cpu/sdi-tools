#!/usr/bin/env python3
"""Довідник «Навантаження_судів.csv» (проєкт «СДІ»): збирання і поповнення.

Повне перезбирання (архів ДСА «Навантаження на суддів» + таблиці rsu.gov.ua для 2022 р.,
І півріччя і року 2023, І кварталу 2024, де в архіві лише скани):
    python3 nav_build.py <тека_архіву_ДСА> <тека_rsu_csv> <тека_для_довідника>
Довідник — три файли: Навантаження_суди.csv (id, вид суду, регіон, назва), Навантаження_періоди.csv
(період, дата стану, джерело — лист ДСА) і Навантаження_дані.csv (сирі показники ДСА).
Дописати новий період з нового листа ДСА (xlsx у форматі 2024–2026):
    python3 nav_build.py --append <файл.xlsx> 2026_3 "9 місяців 2026 року" 9 30.09.2026 \\
        "інформація ДСА України (вх. ВРП № … від …)" <тека з довідником>
Методика ДСА (перевірено): днів на суддю = нормативний час / (суддів з повноваженнями × 8);
середньомісячне надходження = надійшло / (місяців × 11/12) / суддів з повноваженнями.
"""
import csv, glob, os, re, sys
import openpyxl

SECTIONS = ['Місцеві загальні суди', 'Апеляційні загальні суди', 'Окружні адміністративні суди',
            'Апеляційні адміністративні суди', 'Місцеві господарські суди', 'Апеляційні господарські суди']
RSU_KIND = {'mzs': 'Місцеві загальні суди', 'as': 'Апеляційні загальні суди',
            'oas': 'Окружні адміністративні суди', 'aas': 'Апеляційні адміністративні суди',
            'gs': 'Місцеві господарські суди', 'gas': 'Апеляційні господарські суди'}

# період: (код, назва, місяців, дата стану, джерело-файл / rsu-період, лист)
PERIODS = [
    ('2022_4', '2022 рік', 12, '31.12.2022', 'rsu:2022',
     'інформація ДСА України (вх. ВРП № 621/0/8-23 від 17.02.2023)'),
    ('2023_1', 'І квартал 2023 року', 3, '31.03.2023',
     '2023 рік/навантаження суддів за І квартал 2023 року/навантаження І квартал 2023.xlsx',
     'інформація ДСА України (вх. ВРП № 2876/0/8-23 від 18.05.2023)'),
    ('2023_2', 'І півріччя 2023 року', 6, '30.06.2023', 'rsu:2023_2',
     'лист ДСА України від 01.08.2023, вх. ВРП № 4031/0/8-23'),
    ('2023_3', '9 місяців 2023 року', 9, '30.09.2023',
     '2023 рік/навантаження суддів за 9 місяців 2023/навантаження 9 місяців в таблиці.xlsx',
     'лист ДСА України від 07.11.2023 № 15-13680/23'),
    ('2023_4', '2023 рік', 12, '31.12.2023', 'rsu:2023_4',
     'лист ДСА України від 23.01.2024 № 15-2403/24'),
    ('2024_1', 'І квартал 2024 року', 3, '31.03.2024', 'rsu:2024_1',
     'інформація ДСА України (вх. ВРП № 6142/0/8-24 від 24.04.2024)'),
    ('2024_2', 'І півріччя 2024 року', 6, '30.06.2024', '2024 рік/навантаження І півріччя 2024 таблиця.xlsx',
     'інформація ДСА України (вх. ВРП № 9698/0/8-24 від 19.07.2024)'),
    ('2024_3', '9 місяців 2024 року', 9, '30.09.2024',
     '2024 рік/навантаження за 9 міс 2024 р. оновлена лист 13552_0_8-24.xlsx',
     'інформація ДСА України (вх. ВРП № 13552/0/8-24 від 29.11.2024)'),
    ('2024_4', '2024 рік', 12, '31.12.2024', '2024 рік/навантаження за 2024 рік.xlsx',
     'інформація ДСА України (вх. ВРП № 985/0/8-25 від 28.01.2025)'),
    ('2025_1', 'І квартал 2025 року', 3, '31.03.2025', '2025 рік/навантаження за І квартал 2025.xlsx',
     'інформація ДСА України (вх. ВРП № 4264/0/8-25 від 25.04.2025)'),
    ('2025_2', 'І півріччя 2025 року', 6, '30.06.2025', '2025 рік/Навантаження І півріччя 2025.xlsx',
     'інформація ДСА України (вх. ВРП № 8058/0/8-25 від 23.07.2025)'),
    ('2025_3', '9 місяців 2025 року', 9, '30.09.2025', '2025 рік/навантаження за 9 місяців 2025 року.xlsx',
     'інформація ДСА України (вх. ВРП № 11376/1/8-25 від 03.11.2025)'),
    ('2025_4', '2025 рік', 12, '31.12.2025', '2025 рік/навантаження за 2025 рік.xlsx',
     'лист ДСА України від 29.01.2026 № 15-2131/26'),
    ('2026_1', 'І квартал 2026 року', 3, '31.03.2026', '2026 рік/навантаження за І квартал 2026 року.xlsx',
     'інформація ДСА України (вх. ВРП № 4244/1/8-26 від 01.05.2026)'),
    ('2026_2', 'І півріччя 2026 року', 6, '30.06.2026', '2026 рік/навантаження за І півріччя 2026.xlsx',
     'інформація ДСА України (вх. ВРП № 7801/0/8-26 від 23.07.2026)'),
]

FIELDS = ['період_код', 'період', 'стан_на', 'судді_станом_на', 'вид_суду', 'регіон', 'суд', 'надійшло', 'норм_год',
          'суддів_ВРП', 'суддів_з_повноваженнями', 'суддів_за_нормативами',
          'на_суддю_на_місяць', 'днів_на_суддю', 'днів_якщо_мінус_1_суддя', 'днів_якщо_плюс_1_суддя',
          'джерело', 'примітка']

REGION_RE = re.compile(r'(область|Автономна Республіка Крим|^м\. Київ$|^місто Київ$|Севастополь$)')


def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace('\xa0', '').replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def r1(x):
    return '' if x is None else (str(int(round(x))) if abs(x - round(x)) < 1e-9 else f'{x:.1f}')


def derived(m, inc, nh, act):
    """Методика ДСА (перевірено на 2024–2026): на місяць — надійшло / (місяців·11/12) / суддів
    з повноваженнями; днів — норм. год / (суддів з повноваженнями · 8)."""
    per_m = inc / (m * 11 / 12) / act if inc is not None and act else None
    days = nh / (act * 8) if nh is not None and act else None
    minus = nh / ((act - 1) * 8) if nh is not None and act and act > 1 else None
    plus = nh / ((act + 1) * 8) if nh is not None and act is not None else None
    return per_m, days, minus, plus


def region_of(name):
    m = re.search(r'([А-ЯІЇЄҐ][а-яіїєґ\'’-]+ської області|Автономної Республіки Крим|міста Києва|м\. Києва|'
                  r'міста Севастополя|м\. Севастополя)', name)
    if not m:
        return ''
    g = m.group(1)
    if 'Крим' in g:
        return 'Автономна Республіка Крим'
    if 'Києва' in g:
        return 'м. Київ'
    if 'Севастопол' in g:
        return 'м. Севастополь'
    return g.replace('ської області', 'ська область')


def from_xlsx(path, p):
    code, title, months, state, _, src = p
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    # знайти рядок з номерами колонок «А, 1, 2, ...»
    i0 = next(i for i, r in enumerate(rows) if r and r[0] == 'А')
    nums = {int(c): j for j, c in enumerate(rows[i0]) if isinstance(c, (int, float))}
    wide = 7 in nums and 15 in nums
    out, section, region = [], None, ''
    for r in rows[i0 + 1:]:
        if not r or r[0] is None:
            continue
        name = re.sub(r'\s+', ' ', str(r[0])).strip()
        if name in SECTIONS:
            section, region = name, ''
            continue
        if name.startswith(('Усього', 'Разом')) or section is None:
            if name.startswith('Разом'):
                break
            continue
        vals = [num(c) for c in r]
        if REGION_RE.search(name) and 'суд' not in name.lower():
            region = 'м. Київ' if name in ('м. Київ', 'місто Київ') else name
            continue
        inc, nh, vrp, act, norm = vals[1], vals[2], vals[3], vals[4], vals[5]
        if not inc and not act and not vrp:
            continue  # суди на тимчасово окупованих територіях / зі зміненою підсудністю
        per_m, days, minus, plus = derived(months, inc, nh, act)
        note = []
        if wide:   # формат 2024–2026: середньомісячне надходження — значення ДСА
            per_m = vals[nums[7]] if vals[nums[7]] is not None else per_m
        else:
            note.append('на місяць — розраховано за методикою ДСА')
        if 9 in nums and 13 in nums:   # 9 міс. 2023 і 2024–2026: дні — значення ДСА
            days = vals[nums[9]] if vals[nums[9]] is not None else days
            minus = vals[nums[10]] or None
            plus = vals[nums[13]] or None
        else:
            note.append('дні — розраховано за методикою ДСА')
        note = '; '.join(note)
        out.append({'період_код': code, 'період': title, 'стан_на': state, 'судді_станом_на': state,
                    'вид_суду': section,
                    'регіон': region if section == 'Місцеві загальні суди' else region_of(name),
                    'суд': name, 'надійшло': r1(inc), 'норм_год': r1(nh), 'суддів_ВРП': r1(vrp),
                    'суддів_з_повноваженнями': r1(act), 'суддів_за_нормативами': r1(norm),
                    'на_суддю_на_місяць': r1(per_m), 'днів_на_суддю': r1(days),
                    'днів_якщо_мінус_1_суддя': r1(minus), 'днів_якщо_плюс_1_суддя': r1(plus),
                    'джерело': src, 'примітка': note})
    return out


def from_rsu(rsu_dir, p):
    code, title, months, state, spec, src = p
    per = spec.split(':')[1]
    out, seen = [], set()
    for kind, section in RSU_KIND.items():
        f = os.path.join(rsu_dir, f'rsu_{kind}_{per}.csv')
        with open(f, encoding='utf-8') as fh:
            rd = csv.reader(fh)
            header = next(rd)
            hs = ' '.join(header)
            mm = re.search(r'з повноваженнями станом на (\d\d\.\d\d\.\d{4})', hs)
            jstate = mm.group(1) if mm else state
            for c in rd:
                h = dict(zip(header, c))
                g = lambda pref: next((num(v) for k, v in h.items() if k.startswith(pref)), None)
                inc = g('Надійшло')
                nh = g('Нормативний час')
                act = next((num(v) for k, v in h.items() if 'з повноваженнями' in k), None)
                vrp = next((num(v) for k, v in h.items() if 'рішення ВРП' in k or 'наказів ДСА' in k), None)
                norm = next((num(v) for k, v in h.items() if 'за нормативами' in k), None)
                if not inc and not act and not vrp:
                    continue
                per_m, days, minus, plus = derived(months, inc, nh, act)
                name = re.sub(r'\s+', ' ', c[1]).strip()
                if (section, name) in seen:
                    continue   # на сайті таблиця дублюється
                seen.add((section, name))
                out.append({'період_код': code, 'період': title, 'стан_на': state, 'судді_станом_на': jstate,
                            'вид_суду': section,
                            'регіон': ({'місто Київ': 'м. Київ'}.get(c[0], c[0]) if section == 'Місцеві загальні суди' else region_of(name)),
                            'суд': name, 'надійшло': r1(inc), 'норм_год': r1(nh), 'суддів_ВРП': r1(vrp),
                            'суддів_з_повноваженнями': r1(act), 'суддів_за_нормативами': r1(norm),
                            'на_суддю_на_місяць': r1(per_m), 'днів_на_суддю': r1(days),
                            'днів_якщо_мінус_1_суддя': r1(minus), 'днів_якщо_плюс_1_суддя': r1(plus),
                            'джерело': src,
                            'примітка': 'дані rsu.gov.ua; на місяць і днів — розраховано за методикою ДСА'})
    return out


PERIOD_FIELDS = ['код', 'період', 'місяців', 'стан_на', 'судді_станом_на', 'джерело', 'примітка']
KIND_CODE = {k: str(i) for i, k in enumerate(SECTIONS, 1)}


def norm_name(s):
    s = s.lower().replace('’', "'").replace('`', "'").replace('ʼ', "'")
    s = re.sub(r'\bм\.\s*', 'міста ', s)
    return re.sub(r'\s+', ' ', s).strip()


def load_compact(d):
    def rd(n):
        p = os.path.join(d, n)
        if not os.path.exists(p):
            return []
        with open(p, encoding='utf-8') as fh:
            return list(csv.DictReader(fh))
    return rd('Навантаження_суди.csv'), rd('Навантаження_періоди.csv'), rd('Навантаження_дані.csv')


def save_compact(d, courts, periods, data):
    for name, fields, rows in [('Навантаження_суди.csv', ['id', 'вид', 'регіон', 'суд'], courts),
                               ('Навантаження_періоди.csv', PERIOD_FIELDS, periods),
                               ('Навантаження_дані.csv', ['період', 'id', 'надійшло', 'норм_год', 'ВРП',
                                                          'повн', 'норм'], data)]:
        with open(os.path.join(d, name), 'w', newline='', encoding='utf-8') as fh:
            w = csv.DictWriter(fh, fieldnames=fields, lineterminator='\n')
            w.writeheader(); w.writerows(rows)


def add_rows(courts, data, rows):
    idx = {(c['вид'], norm_name(c['суд'])): c for c in courts}
    nxt = max([int(c['id']) for c in courts] or [0]) + 1
    for r in rows:
        key = (KIND_CODE[r['вид_суду']], norm_name(r['суд']))
        c = idx.get(key)
        if c is None:
            c = {'id': str(nxt), 'вид': key[0], 'регіон': r['регіон'], 'суд': r['суд']}
            nxt += 1; courts.append(c); idx[key] = c
        else:
            c['суд'] = r['суд']            # остання назва (з ДСА — «міста», а не «м.»)
            c['регіон'] = c['регіон'] or r['регіон']
        nh = r['норм_год']
        data.append({'період': r['період_код'], 'id': c['id'], 'надійшло': r['надійшло'],
                     'норм_год': str(int(round(float(nh)))) if nh else '', 'ВРП': r['суддів_ВРП'],
                     'повн': r['суддів_з_повноваженнями'], 'норм': r['суддів_за_нормативами']})


def main():
    dsa_dir, rsu_dir, out_dir = sys.argv[1:4]
    courts, periods, data = [], [], []
    for p in PERIODS:
        spec = p[4]
        rows = from_rsu(rsu_dir, p) if spec.startswith('rsu:') else from_xlsx(os.path.join(dsa_dir, spec), p)
        print(p[0], len(rows))
        jst = rows[0]['судді_станом_на'] if rows else p[3]
        note = 'дані rsu.gov.ua (у теці ДСА — лише скан листа)' if spec.startswith('rsu:') else ''
        periods.append({'код': p[0], 'період': p[1], 'місяців': p[2], 'стан_на': p[3],
                        'судді_станом_на': jst, 'джерело': p[5], 'примітка': note})
        add_rows(courts, data, rows)
    save_compact(out_dir, courts, periods, data)
    print('судів:', len(courts), 'рядків даних:', len(data))


def append(argv):
    """Дописати новий період з таблиці ДСА (xlsx у форматі 2024–2026):
    python3 nav_build.py --append <файл.xlsx> <код 2026_3> "<9 місяців 2026 року>" <місяців 9>
        <стан_на 30.09.2026> "<джерело>" <тека з довідником>"""
    xlsx, code, title, months, state, src, d = argv
    courts, periods, data = load_compact(d)
    rows = from_xlsx(xlsx, (code, title, int(months), state, xlsx, src))
    data = [x for x in data if x['період'] != code]
    periods = [x for x in periods if x['код'] != code] + [
        {'код': code, 'період': title, 'місяців': months, 'стан_на': state, 'судді_станом_на': state,
         'джерело': src, 'примітка': ''}]
    periods.sort(key=lambda x: x['код'])
    add_rows(courts, data, rows)
    save_compact(d, courts, periods, data)
    print(f'{code}: дописано {len(rows)} судів; усього рядків {len(data)}')


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--append':
        append(sys.argv[2:])
    else:
        main()
