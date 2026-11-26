#!/usr/bin/env python3
"""Пошук у покажчику протоколів автоматизованого розподілу (дані ДСА, 16.11.2023–08.02.2026).

Автор: В'ячеслав Плескач. Версія 1.3 (26.11.2026): межі покажчика (до 08.02.2026) виводяться щоразу, коли проміжок пошуку не закінчується
до 08.02.2026 (зокрема --case без дат): раніше за знайденими протоколами мовчки не було видно, що пізніші не перевірено.
Версія 1.2 (25.11.2026): дати — РРРР-ММ-ДД або ДД.ММ.РРРР; для номера справи з роком
(«…/25») файли ранніх років не завантажуються; завантажений файл перевіряється; межі покажчика і обрізання виводу —
у результаті. УВАГА: у покажчику поле excluded_judges обірване на першому судді, причина виключення якого містить
«[Примітки: …]» (і excluded_count занижений) — для таких протоколів звіряйте з protokoly_portal.py 2.1.
Версія 1.1 (14.11.2026): номер справи порівнюється без пробілів, регістру й кінцевого
суфікса (-к, -ц, -а), як у edrsr_vybirka.py і stan_poshuk.py.

Файли покажчика лежать на Google Диску (тека https://drive.google.com/drive/folders/1EWOTy7D5kjT3nu2md0p9jEpWQX7aVscx,
доступ за посиланням); скрипт сам завантажує потрібні роки (рік = рік ПУБЛІКАЦІЇ протоколу) і кешує їх.

Приклади:
  python3 protokoly_poshuk.py --case "757/30191/25-к"                     # усі протоколи у справі
  python3 protokoly_poshuk.py --case "757/30191/25-к" --details           # + склад, виключені судді з причинами
  python3 protokoly_poshuk.py --court 2606 --judge "Вовк" --from 2025-06-01 --to 2025-06-30 --out vovk.tsv
  python3 protokoly_poshuk.py --court 2606 --type "повторного" --years 2024-2025 --out povtorni.tsv

Фільтри (поєднуються через І): --case (номер справи або шаблон із %), --court (код ЄДРСР, 3 чи 4 цифри),
--judge (частина ПІБ судді-доповідача; шукає і в повному ПІБ поля judges_field — потребує --details),
--type (частина виду протоколу), --from/--to (дата розподілу або внесення, РРРР-ММ-ДД), --years (роки файлів).
Без --out друкує результат на екран (до 200 рядків).
"""
import argparse, csv, datetime, os, re, subprocess, sys, time

POKRYTTIA = 'Покажчик охоплює протоколи, опубліковані 16.11.2023–08.02.2026; пізніші — protokoly_portal.py.'

FILES = {  # рік: (core, det) — ID файлів на Диску
    2022: ('1b9pJfmypXpiMrAKTVIeZeUwveV2-vqMb', '18_6XnPDXJDmvOeZJ9Zcdruiyhya6pW-N'),
    2023: ('1sQz0BmfvfkrTfAAZp90ot1cH5fAnxAjD', '1VSmTWfslPF06YeoscQqRLHqceHl3OSLN'),
    2024: ('1I9XZYp1f0L45Ukk48Si9SQXiIyI_EW7M', '1tLHwztGQuLGL2reGYiYnov7JPDfCkhKq'),
    2025: ('1ERezgVoZXmkpLeeyUefPdd1K53rF14vx', '1SvH1Z4AU_KM6tF4lDi7Dt5ZXFTMPdoHs'),
    2026: ('1Tm3fJdwG6pJd37nQADF3KDIwgDDLtruv', '1Iq7eJlxJpbt-ZO9MxqrLon8MZXzF0lBC'),
}
COURTS_ID = '1GM-aLV9a1pIJ5zTmFht-BdDn5c7tMY8A'  # sudy_kody.csv


def need(mod):
    try:
        return __import__(mod)
    except ImportError:
        subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', '--break-system-packages', mod], check=True)
        return __import__(mod)


def fetch(fid, path):
    if not os.path.exists(path):
        t = time.time()
        rc = subprocess.run(['curl', '-sS', '-L', '-f', '--retry', '3', '-m', '1800', '-o', path + '.part',
                             f'https://drive.usercontent.google.com/download?id={fid}&export=download&confirm=t']).returncode
        ok = False
        if not rc:
            with open(path + '.part', 'rb') as fh:     # Диск замість файла може віддати сторінку (HTTP 200)
                ok = fh.read(2) == b'\x1f\x8b'
        if not ok:
            if os.path.exists(path + '.part'):
                os.remove(path + '.part')
            sys.exit(f'Не вдалося завантажити {os.path.basename(path)} з Диска (id {fid}): мережа, закритий доступ за посиланням '
                     'або ліміт завантажень.')
        os.replace(path + '.part', path)
        print(f'  завантажено {os.path.basename(path)} за {time.time() - t:.0f} с', file=sys.stderr)
    return path


def case_key(s):
    return re.sub(r'-[^/0-9-]{1,4}$', '', re.sub(r'\s+', '', s or '').lower())


def iso(s, name):
    for f in ('%Y-%m-%d', '%d.%m.%Y'):
        try:
            return datetime.datetime.strptime(s.strip(), f).date().isoformat()
        except ValueError:
            pass
    sys.exit(f'{name}: не розпізнано дату «{s}» (потрібно РРРР-ММ-ДД або ДД.ММ.РРРР).')


CASE_SQL = "regexp_replace(lower(replace(coalesce(case_number,''),' ','')), '-[^/0-9-]{1,4}$', '')"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--case'); ap.add_argument('--court'); ap.add_argument('--judge'); ap.add_argument('--type')
    ap.add_argument('--from', dest='dfrom'); ap.add_argument('--to', dest='dto')
    ap.add_argument('--years', default='2022-2026')
    ap.add_argument('--details', action='store_true', help='приєднати деталі (склад, виключені судді, ПІБ повністю)')
    ap.add_argument('--cache', default=os.path.expanduser('~/.cache/protokoly'))
    ap.add_argument('--out')
    a = ap.parse_args()
    if not any([a.case, a.court, a.judge, a.type, a.dfrom]):
        sys.exit('Задайте хоча б один фільтр (--case, --court, --judge, --type, --from).')
    duckdb = need('duckdb')
    os.makedirs(a.cache, exist_ok=True)
    a.dfrom = iso(a.dfrom, '--from') if a.dfrom else None
    a.dto = iso(a.dto, '--to') if a.dto else None
    y1, _, y2 = a.years.partition('-')
    years = [y for y in range(int(y1), int(y2 or y1) + 1) if y in FILES]
    if a.dfrom:  # файл року публікації не раніше року розподілу
        years = [y for y in years if y >= int(a.dfrom[:4])]
    m = re.search(r'/(\d{2})$', case_key(a.case)) if a.case and '%' not in a.case else None
    if m and 20 <= int(m.group(1)) <= datetime.date.today().year % 100:   # «757/30191/25»: протоколів до 2025 року бути не може
        years = [y for y in years if y >= 2000 + int(m.group(1))]
    if not years:
        sys.exit(f'Для заданих умов у покажчику немає файлів (роки файлів: {min(FILES)}–{max(FILES)}). ' + POKRYTTIA)
    con = duckdb.connect()
    w = []
    esc = lambda s: s.replace("'", "''")
    if a.case:
        w.append(f"case_number LIKE '{esc(a.case)}'" if '%' in a.case else f"{CASE_SQL} = '{esc(case_key(a.case))}'")
    if a.court:
        w.append('court_code IN (' + ','.join(f"'{c.strip().lstrip('0') or '0'}'" for c in a.court.split(',') if c.strip()) + ')')
    if a.type:
        w.append(f"protocol_type ILIKE '%{esc(a.type)}%'")
    dexpr = "coalesce(nullif(distribution_start,''), nullif(entry_date,''), doc_create_date)"
    if a.dfrom: w.append(f"substr({dexpr},1,10) >= '{a.dfrom}'")
    if a.dto: w.append(f"substr({dexpr},1,10) <= '{a.dto}'")
    jw = f"(reporting_judge ILIKE '%{esc(a.judge)}%'" + (f" OR judges_field ILIKE '%{esc(a.judge)}%')" if a.details else ')') if a.judge else None
    parts = []
    for y in years:
        core = fetch(FILES[y][0], os.path.join(a.cache, f'protokoly_core_{y}.tsv.gz'))
        rc = f"read_csv('{core}', delim='\t', header=true, quote='\"', all_varchar=true, max_line_size=20000000)"
        if a.details:
            det = fetch(FILES[y][1], os.path.join(a.cache, f'protokoly_det_{y}.tsv.gz'))
            rd = f"read_csv('{det}', delim='\t', header=true, quote='\"', all_varchar=true, max_line_size=20000000)"
            sel = f"SELECT c.*, d.court_name, d.subject, d.distribution_end, d.collegium, d.participating_judges, d.excluded_judges, d.replaced_judges, d.judges_field FROM {rc} c JOIN {rd} d USING (doc_dbid, doc_id)"
        else:
            sel = f"SELECT * FROM {rc}"
        cond = w + ([jw] if jw else [])
        parts.append(f"SELECT * FROM ({sel}) WHERE " + ' AND '.join(cond))
    q = f"SELECT *, {dexpr} AS date_any FROM ({' UNION ALL '.join(parts)}) ORDER BY date_any, doc_create_date"
    t = time.time()
    res = con.execute(q); cols = [c[0] for c in res.description]; rows = res.fetchall()
    print(f'Знайдено {len(rows)} протоколів за {time.time() - t:.1f} с', file=sys.stderr)
    if not rows:
        print('Протоколів не знайдено. ' + POKRYTTIA)
    elif not a.dto or a.dto > '2026-02-08':           # 1.3: і для --case без дат
        print('УВАГА: ' + POKRYTTIA + ' Протоколи, опубліковані після 08.02.2026, тут не перевірено.')
    if a.out:
        with open(a.out, 'w', encoding='utf-8', newline='') as f:
            wr = csv.writer(f, delimiter='\t', lineterminator='\n'); wr.writerow(cols); wr.writerows(rows)
        print(f'  → {a.out}', file=sys.stderr)
    else:
        show = ['date_any', 'court_code', 'case_number', 'proc_number', 'doc_form', 'protocol_type', 'ground', 'reporting_judge', 'excluded_count', 'participating_count']
        if a.details: show += ['judges_field', 'excluded_judges']
        idx = [cols.index(c) for c in show]
        print('\t'.join(show))
        for r in rows[:200]:
            print('\t'.join(str(r[i])[:300] for i in idx))
        if len(rows) > 200:
            print(f'… показано 200 з {len(rows)}; повний результат — з параметром --out файл.tsv')


if __name__ == '__main__':
    main()
