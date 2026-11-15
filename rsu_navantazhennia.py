#!/usr/bin/env python3
"""Завантажує таблиці навантаження судів з rsu.gov.ua (розділ «Нормативи навантаження»).

Використання:
    python3 rsu_navantazhennia.py <тип> <період> [фільтр_назви_суду]
    тип:    mzs — місцеві загальні; oas — окружні адміністративні; gs — місцеві господарські;
            as — апеляційні загальні; aas — апеляційні адміністративні; gas — апеляційні господарські
    період: 2026_2 (I півріччя 2026), 2026_1 (I квартал), 2025_4 (рік), 2025_3 (9 міс.) ...
            для 2019–2022 і частини 2023 — просто рік: 2022
Приклад:
    python3 rsu_navantazhennia.py mzs 2026_2 "Голосіївський"
Результат: CSV rsu_<тип>_<період>.csv (усі суди) і друк рядків, що відповідають фільтру,
з розрахованими показниками (справ на суддю з повноваженнями, % укомплектованості,
співвідношення нормативної та фактичної кількості суддів, місце в області за надходженням на суддю).
"""
import csv, html, re, subprocess, sys, urllib.parse

BASE = 'https://rsu.gov.ua'


def get(url):
    return subprocess.run(['curl', '-sS', '-m', '60', '-L', url], capture_output=True,
                          check=True).stdout.decode('utf-8', 'replace')


def cells(row):
    out = []
    for c in re.findall(r'<td[^>]*>(.*?)</td>', row, re.S):
        t = re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]+>', '', c))).strip()
        if t:
            out.append(t)
    return out


def num(s):
    s = s.replace('\xa0', '').replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def load(kind, period):
    page = get(f'{BASE}/{kind}-{period}-maps')
    m = re.search(r'<iframe[^>]+src="([^"]+)"', page)
    if not m:
        sys.exit('На сторінці немає таблиці (iframe) — перевірте тип і період')
    frame_url = urllib.parse.urljoin(BASE + '/', urllib.parse.quote(m.group(1)))
    frame = get(frame_url)
    sheets = re.findall(r'<link id="shLink" href="([^"]+)"', frame)
    names = re.findall(r'<x:Name>([^<]+)', frame)
    header, rows = None, []
    for i, sh in enumerate(sheets):
        region = names[i] if i < len(names) else ''
        if region.strip() == 'Україна':
            continue  # зведений аркуш дублює області
        s = get(urllib.parse.urljoin(frame_url, sh))
        for r in re.findall(r'<tr[^>]*>(.*?)</tr>', s, re.S):
            c = cells(r)
            if not c:
                continue
            if c[0] == 'Регіон' and header is None:
                header = c
            elif header and len(c) == len(header) and c[0] != 'Регіон':
                rows.append(c)
    return header, rows


def main():
    kind, period = sys.argv[1], sys.argv[2]
    flt = sys.argv[3].lower() if len(sys.argv) > 3 else None
    header, rows = load(kind, period)
    out = f'rsu_{kind}_{period}.csv'
    with open(out, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f'{len(rows)} судів → {out}')
    print('Колонки:', ' | '.join(header))
    # розрахунок «справ на суддю з повноваженнями» і місця в області
    i_in = next(i for i, h in enumerate(header) if h.startswith('Надійшло'))
    i_act = next(i for i, h in enumerate(header) if 'з повноваженнями' in h)
    i_norm = next((i for i, h in enumerate(header) if 'за нормативами' in h), None)
    i_vrp = next((i for i, h in enumerate(header) if 'ВРП' in h), None)
    by_region = {}
    for r in rows:
        a, n = num(r[i_in]), num(r[i_act])
        r.append(round(a / n, 1) if a and n else None)
        by_region.setdefault(r[0], []).append(r)
    L = len(header)
    for r in rows:
        a, nm = num(r[i_act]), (num(r[i_norm]) if i_norm is not None else None)
        r.append(round(nm / a, 2) if a and nm else None)   # L+1: нормативна / фактична кількість суддів
    for reg, rs in by_region.items():
        for col in (L, L + 1):
            ok = sorted([r for r in rs if r[col]], key=lambda r: -r[col])
            for k, r in enumerate(ok, 1):
                r.append((col, f'{k} з {len(ok)}'))
    nat = {}
    for col in (L, L + 1):
        ok = sorted([r for r in rows if r[col]], key=lambda r: -r[col])
        for k, r in enumerate(ok, 1):
            nat[(id(r), col)] = f'{k} з {len(ok)}'
    if flt:
        for r in rows:
            if flt in r[1].lower():
                print('\n' + r[1] + ' (' + r[0] + ')')
                for h, v in zip(header, r):
                    print(f'  {h}: {v}')
                ranks = dict(x for x in r[L + 2:] if isinstance(x, tuple))
                print(f'  Справ на одного суддю з повноваженнями: {r[L]} (місце в регіоні: {ranks.get(L)}; в Україні: {nat.get((id(r), L))})')
                print(f'  Нормативна / фактична кількість суддів: {r[L + 1]} (місце в регіоні: {ranks.get(L + 1)}; в Україні: {nat.get((id(r), L + 1))})')
                if i_vrp is not None and num(r[i_vrp]):
                    print(f'  Укомплектованість (з повноваженнями / за рішенням ВРП): '
                          f'{round(100 * num(r[i_act]) / num(r[i_vrp]))} %')



if __name__ == '__main__':
    main()
