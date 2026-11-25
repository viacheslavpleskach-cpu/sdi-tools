#!/usr/bin/env python3
"""Пошук в індексі історії «Стан розгляду справ» (01.01.2023 – 30.10.2026) для висновків.

Версія 1.3 (25.11.2026): `spysok` працює на чистому кеші (перелік теки Диска читається один раз, до паралельного
завантаження шардів); невідомий параметр — помилка, а не мовчазне ігнорування; завантажений файл перевіряється
(parquet, а не сторінка Диска); дати `--vid/--do` — РРРР-ММ-ДД або ДД.ММ.РРРР; медіана інтервалів — справжня.
Версія 1.2 (11.11.2026): застереження про дати стадій уточнено за звіркою з реальними засіданнями
(«Довідник_відкритих_даних_судів.md», розд. 4.4); у підсумку `sprava` з'явився блок «Дати, що можуть відповідати
проведеним засіданням (орієнтовно)».

Індекс — тека Google Диска «Індекс історії стану розгляду 2023–2026»
https://drive.google.com/drive/folders/145g1zR_CA36DwwBLhKXnYKPrI2ls-rTi (доступ за посиланням):
  stan_istoriia_00..31.parquet — усі записи ретроархіву 01.2023–09.2026 + вихідний знімок на 30.10.2026,
      розкладені за номером справи (шард = md5(ключ) % 32) і відсортовані; ~65 МБ кожен (62,4 млн записів);
  stan_suddi_misiats.parquet   — помісячні показники по суддях і судах (суддя «*» — увесь суд), 01.2023–09.2026;
  stan_suddi_intervaly.parquet — по суддях і судах за видом провадження: медіана й 90-й перцентиль інтервалу між
      датами (пере)призначення до розгляду; медіана днів від реєстрації до «Розглянуто».
Ключ справи: нижній регістр, без пробілів, без кінцевого суфікса «-к», «-ц», «-а»… («757/28368/24-к» → «757/28368/24»),
тож різні записи того самого номера знаходяться разом.

Записи після 30.10.2026 (жовтневий шар і щоденні зміни) — у GitHub stan-rozghliadu-arkhiv; якщо репозиторій
склоновано, передайте --repo <тека>, і скрипт додасть їх до хронології.

Команди:
  stan_poshuk.py sprava "757/28368/24-к" [--sud "Печерський"] [--repo DIR] [--tsv out.tsv]
        хронологія кожного провадження у справі + підсумок: реєстрація, розподіл і зміна суддів, відкриття,
        без руху, дати стадії «Призначено до судового розгляду» та інтервали між ними, розгляд, набрання сили,
        направлення до апеляції; орієнтовні дати проведених засідань
  stan_poshuk.py suddia "Гридасова А.М." --sud "Печерський" [--vyd "кримінальне"]
        помісячні показники судді поряд із судом і інтервали за видами проваджень (без завантаження шардів)
  stan_poshuk.py sud "Печерський районний суд міста Києва"            # помісячні показники суду, перелік суддів
  stan_poshuk.py spysok --sud "Печерський" --suddia "Гридасова" [--etap "Призначено до"] [--vid 2025-01-01 --do 2025-12-31]
        [--uchasnyk "Іваненко"] [--tsv out.tsv]     # провадження за умовами; читає ВСІ шарди (~2 ГБ, 2–3 хв)
Загальні: --dir <тека з файлами індексу> (якщо вже завантажені), --kesh <тека кешу> (типово ~/.cache/stan_indeks).

Застереження (виводяться й у звіті; докладно — «Довідник_відкритих_даних_судів.md», розд. 4.4):
  * stage_date — дата, коли стадію зафіксовано в системі суду. Звірка з реальними засіданнями (11.11.2026):
    - розподіл («Призначено склад суду») і остаточний розгляд («Розглянуто») збігаються з реальними датами майже завжди;
    - «Призначено до судового розгляду»: дата — або день засідання (якщо запис внесли в день засідання), або день
      внесення запису про (пере)призначення (зсув від 1 до 8 днів залежить від суду); у даних є лише частина засідань
      (орієнтовно 40–60 %);
    - засідання, зняті з розгляду або такі, що не відбулися, під своєю датою у даних відсутні; натомість через 2–10 днів
      може з'явитися запис про нове призначення; зупинення й інші процесуальні ухвали не фіксуються.
    Тому відсутність дати в даних НЕ доводить, що засідання не було, кількість засідань за даними не рахується, а
    інтервали між датами — лише орієнтир. Дата, НА ЯКУ призначено засідання, — лише в «Списку справ, призначених до
    розгляду» (архів з 31.10.2026).
  * snapshot_date — день, коли стан уперше видно у вивантаженні ДСА; пропущені дні перелічено в примітці
    (зміни видно наступного дня), червень–липень 2025 вивантаження здебільшого неповні.
  * 01–07.2023 у вивантаженнях ДСА лише події дня (стадії, зафіксовані в день файла): відсутність запису за день
    означає, що подій не опубліковано, а не що їх не було. Дні 0 КБ на сайті ДСА (серпень–листопад 2023) даних не мають.
  * Ретроархів містить лише зміни стадії, судді, результату, дати реєстрації; зміни учасників не зберігалися.
Автор: В'ячеслав Плескач
"""
import argparse, collections, concurrent.futures as cf, datetime as dt, hashlib, os, re, statistics, subprocess, sys, threading

FOLDER = '145g1zR_CA36DwwBLhKXnYKPrI2ls-rTi'
NSHARD = 32
KNOWN = {'stan_suddi_intervaly.parquet': '1Fg8NpdutY2llVS0LHyaxNPs8TtzdbrOZ',
         'stan_suddi_misiats.parquet': '12aatVD7obuK8A7_JS1l-pop8HJyG4aQi'}
SKIPPED = ['2024-12-04', '2024-12-11', '2024-12-27', '2024-12-30',
           '2025-01-11', '2025-02-11', '2025-02-19', '2025-03-03', '2025-03-05', '2025-03-08', '2025-03-23',
           '2025-04-12', '2025-04-29', '2025-05-27', '2025-05-30']
EMPTY_2023 = ('0 КБ на сайті ДСА: 10, 12, 19.08.2023; 11, 15, 18.09.2023; 11, 16, 20, 23–30.10.2023; '
              '01–03, 06–10, 13–17.11.2023 (20.11.2023 файла немає)')
NOTE = ('Примітка: дата стадії — дата, коли стадію зафіксовано в системі суду. Для «Призначено до судового розгляду» '
        'це або день засідання, або день внесення запису про (пере)призначення (зсув 1–8 днів залежно від суду); у даних '
        'є лише частина засідань (орієнтовно 40–60 %), а засідання, зняті з розгляду, під своєю датою відсутні. '
        'Відсутність дати не доводить, що засідання не було; кількість засідань за даними не рахується; дата, на яку '
        'призначено засідання, у даних відсутня. «Перша поява» — день, коли стан уперше видно у вивантаженні ДСА. '
        '01–07.2023 — лише події дня; ' + EMPTY_2023 + '. Червень–липень 2025 — неповні вивантаження; '
        'пропущені (не завантажені) дні: ' + ', '.join(SKIPPED) + '.')


def need_duckdb():
    try:
        import duckdb  # noqa
    except ImportError:
        subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', '--break-system-packages', 'duckdb'], check=True)
    import duckdb
    return duckdb


def key(case):
    s = re.sub(r'\s+', '', (case or '')).lower()
    return re.sub(r'-[^/0-9-]{1,4}$', '', s)


def shard(k):
    return int(hashlib.md5(k.encode('utf-8')).hexdigest()[:8], 16) % NSHARD


# ---------- файли індексу ----------
_ids = None
_ids_lock = threading.Lock()
def folder_ids():
    """Назва файла → ID на Диску. Під замком: шарди завантажуються в кілька потоків, і до 1.3 решта потоків
    бачила ще не заповнений перелік («Не знайдено stan_istoriia_NN.parquet»)."""
    global _ids
    with _ids_lock:
        if _ids is None:
            ids = dict(KNOWN)
            r = subprocess.run(['curl', '-sS', '-L', '-m', '60', f'https://drive.google.com/embeddedfolderview?id={FOLDER}'],
                               capture_output=True)
            html = r.stdout.decode('utf-8', 'replace')
            for m in re.finditer(r'/file/d/([A-Za-z0-9_-]+).*?flip-entry-title">([^<]+)', html, re.S):
                ids[m.group(2).strip()] = m.group(1)
            if r.returncode or len(ids) == len(KNOWN):
                print(f'УВАГА: перелік теки Диска {FOLDER} не отримано (мережа або доступ за посиланням) — '
                      'відомі лише файли показників суддів.', file=sys.stderr)
            _ids = ids
    return _ids


def get(name, a):
    if a.dir and os.path.exists(os.path.join(a.dir, name)):
        return os.path.join(a.dir, name)
    os.makedirs(a.kesh, exist_ok=True)
    p = os.path.join(a.kesh, name)
    if os.path.exists(p) and os.path.getsize(p) > 0:
        return p
    fid = folder_ids().get(name)
    if not fid:
        sys.exit(f'Не знайдено {name} у теці Диска {FOLDER}. Завантажте файл через конектор Google Drive '
                 f'(search_files → download_file_content) у теку й передайте --dir.')
    rc = subprocess.run(['curl', '-sS', '-L', '-f', '--retry', '3', '-m', '900', '-o', p + '.part',
                         f'https://drive.usercontent.google.com/download?id={fid}&export=download&confirm=t']).returncode
    if rc:
        sys.exit(f'Не вдалося завантажити {name} (id {fid}).')
    with open(p + '.part', 'rb') as fh:                # Диск замість файла може віддати сторінку (HTTP 200)
        ok = fh.read(4) == b'PAR1'
    if not ok:
        os.remove(p + '.part')
        sys.exit(f'Замість {name} Диск повернув не parquet (закрито доступ за посиланням або вичерпано ліміт завантажень). '
                 'Завантажте файл через конектор Google Drive і передайте --dir.')
    os.replace(p + '.part', p)
    return p


def all_shards(a):
    names = [f'stan_istoriia_{i:02d}.parquet' for i in range(NSHARD)]
    folder_ids()                                       # один раз, до потоків
    with cf.ThreadPoolExecutor(4) as ex:
        return list(ex.map(lambda n: get(n, a), names))


# ---------- допоміжне ----------
def d(s):
    try:
        return dt.datetime.strptime(s, '%d.%m.%Y').date()
    except (TypeError, ValueError):
        return None


def fmt(x):
    return x.strftime('%d.%m.%Y') if x else ''


def jname(j):
    m = re.search(r':\s*([^,;]+)', j or '')
    return m.group(1).strip() if m else (j or '').strip()


COLS = ['snapshot_date', 'change', 'court_name', 'case_number', 'case_proc', 'type', 'registration_date', 'judge',
        'judges', 'stage_date', 'stage_name', 'cause_result', 'description', 'participants']


def repo_rows(repo, case):
    """Записи з архіву GitHub (жовтневий шар і щоденні файли) для справи."""
    import glob, lzma, csv
    k = key(case)
    out = []
    files = sorted(glob.glob(os.path.join(repo, 'data', '*', '*', '*.tsv.xz')))
    for f in files:
        r = subprocess.run(['xzgrep', '-h', '-i', '-F', k.split('/')[1] if '/' in k else k, f], capture_output=True)
        if not r.stdout:
            continue
        with lzma.open(f, 'rt', encoding='utf-8') as fh:
            h = fh.readline().rstrip('\n').split('\t')
        for line in r.stdout.decode('utf-8', 'replace').splitlines():
            x = next(csv.reader([line], delimiter='\t'))
            row = dict(zip(h, x))
            if key(row.get('case_number', '')) == k and row.get('stage_name'):
                out.append({c: row.get(c, '') for c in COLS})
    return out


# ---------- sprava ----------
def cmd_sprava(a):
    duckdb = need_duckdb()
    k = key(a.case)
    f = get(f'stan_istoriia_{shard(k):02d}.parquet', a)
    con = duckdb.connect()
    q = f"select {', '.join(COLS)} from read_parquet('{f}') where k = ?"
    rows = [dict(zip(COLS, r)) for r in con.execute(q, [k]).fetchall()]
    if a.repo:
        rows += repo_rows(a.repo, a.case)
    if a.sud:
        rows = [r for r in rows if a.sud.lower() in (r['court_name'] or '').lower()]
    if not rows:
        print(f'Справу {a.case} (ключ {k}) в індексі не знайдено.'); return
    if a.tsv:
        import csv
        with open(a.tsv, 'w', encoding='utf-8', newline='') as fh:
            w = csv.writer(fh, delimiter='\t'); w.writerow(COLS)
            for r in sorted(rows, key=lambda r: (r['court_name'], r['case_proc'] or '', r['snapshot_date'])):
                w.writerow([r[c] or '' for c in COLS])
    procs = collections.defaultdict(list)
    for r in rows:
        procs[(r['court_name'], r['case_proc'] or '', r['type'] or '')].append(r)
    order = sorted(procs, key=lambda p: min((d(r['registration_date']) or dt.date(2100, 1, 1)) for r in procs[p]))
    for p in order:
        rs = sorted(procs[p], key=lambda r: (r['snapshot_date'], d(r['stage_date']) or dt.date(1900, 1, 1)))
        print('=' * 100)
        print(f'{p[0]} | справа {rs[0]["case_number"]} | провадження {p[1] or "—"} | {p[2]}')
        regs = sorted({r['registration_date'] for r in rs if r['registration_date']}, key=lambda s: d(s) or dt.date(1900, 1, 1))
        desc = next((r['description'] for r in rs if r['description']), '')
        print(f'Дата реєстрації: {", ".join(regs) or "—"}' + (f' | предмет: {desc[:150]}' if desc else ''))
        part = next((r['participants'] for r in reversed(rs) if r['participants']), '')
        if part:
            print(f'Учасники (останній запис): {part[:300]}')
        # події: стадія + дата стадії, найраніша поява
        ev = {}
        for r in rs:
            kk = (r['stage_name'], r['stage_date'], r['cause_result'] or '')
            if kk not in ev or r['snapshot_date'] < ev[kk]['snapshot_date']:
                ev[kk] = r
        evl = sorted(ev.values(), key=lambda r: (d(r['stage_date']) or dt.date(1900, 1, 1), r['snapshot_date']))
        print(f'\n{"дата стадії":<12} {"перша поява":<12} {"стадія":<45} {"результат":<35} суддя')
        last_j = None
        for r in evl:
            j = jname(r['judge'])
            print(f'{r["stage_date"] or "":<12} {str(r["snapshot_date"]):<12} {(r["stage_name"] or "")[:45]:<45} '
                  f'{(r["cause_result"] or "")[:35]:<35} {j if j != last_j else "〃"}')
            last_j = j
        # підсумок
        print('\nПідсумок:')
        judges = []
        for r in sorted(rs, key=lambda r: r['snapshot_date']):
            j = jname(r['judge'])
            if j and (not judges or judges[-1][0] != j):
                judges.append((j, r['stage_date'] or str(r['snapshot_date'])))
        if judges:
            print('  судді (у порядку появи): ' + '; '.join(f'{j} (з {s})' for j, s in judges))
            if len(judges) > 1:
                print('  ! зміна судді / складу — перевірити протоколи розподілу (protokoly_poshuk.py) і рішення')
        def first(pred):
            x = [d(r['stage_date']) for r in evl if pred(r['stage_name'] or '') and d(r['stage_date'])]
            return min(x) if x else None
        reg = min([d(s) for s in regs if d(s)] or [None]) if regs else None
        milestones = [('склад суду / передано судді', first(lambda s: s in ('Призначено склад суду', 'Передано судді', 'Передано в доповідь'))),
                      ('залишено без руху', first(lambda s: s == 'Залишено без руху')),
                      ('відкрито провадження', first(lambda s: s == 'Відкрито провадження')),
                      ('перший запис «Призначено до судового розгляду»', first(lambda s: s == 'Призначено до судового розгляду')),
                      ('зупинено', first(lambda s: 'упинено' in s)),
                      ('розглянуто', first(lambda s: s.startswith('Розглянуто') and 'законної сили' not in s)),
                      ('набрало законної сили', first(lambda s: 'законної сили' in s)),
                      ('направлено до апеляційного / касаційного суду', first(lambda s: s.startswith('Направлено до'))),
                      ('повернуто / відмовлено / закрито / без розгляду', first(lambda s: s.startswith(('Повернуто', 'Відмовлено', 'Закрито', 'Залишено без розгляду')))),
                      ('направлено за підсудністю', first(lambda s: s.startswith('Направлено за підсудністю')))]
        if reg:
            print(f'  реєстрація: {fmt(reg)}')
        for n, x in milestones:
            if x:
                print(f'  {n}: {fmt(x)}' + (f' (через {(x - reg).days} дн. після реєстрації)' if reg and x >= reg else ''))
        pr = sorted({d(r['stage_date']) for r in evl if r['stage_name'] == 'Призначено до судового розгляду' and d(r['stage_date'])})
        if pr:
            gaps = [(pr[i] - pr[i - 1]).days for i in range(1, len(pr))]
            print(f'  дат стадії «Призначено до судового розгляду»: {len(pr)}: ' + ', '.join(fmt(x) for x in pr))
            if gaps:
                mx = max(gaps); i = gaps.index(mx)
                print(f'  інтервали між ними, днів: {gaps}; медіана {statistics.median(gaps):g}; найбільший {mx} '
                      f'({fmt(pr[i])} → {fmt(pr[i + 1])})')
            print('  ! це дні засідань або дні внесення запису про (пере)призначення; частину засідань у даних немає, '
                  'зняті з розгляду не видно — інтервали орієнтовні')
        # орієнтовні дати проведених засідань: записи «Призначено до судового розгляду» + остаточний розгляд
        fin = sorted({d(r['stage_date']) for r in evl if (r['stage_name'] or '').startswith('Розглянуто') and
                      'законної сили' not in (r['stage_name'] or '') and 'апеляційній' not in (r['stage_name'] or '') and d(r['stage_date'])})
        if pr or fin:
            allx = sorted(set(pr) | set(fin))
            print(f'  Дати, що можуть відповідати проведеним засіданням (орієнтовно; {len(allx)}): '
                  + ', '.join(fmt(x) + ('*' if x in fin else '') for x in allx))
            print('    * — «Розглянуто» (остаточний розгляд; збігається з реальною датою майже завжди). Решта — дні засідань '
                  'або дні внесення запису (зсув 1–8 днів); у даних немає орієнтовно 40–60 % засідань, зняті з розгляду не видно. '
                  'У висновок — лише дати, підтверджені ухвалами чи матеріалами справи.')
        last = evl[-1]
        print(f'  останній відомий стан: {last["stage_name"]} від {last["stage_date"]} (видно з {last["snapshot_date"]})'
              + (f', {last["cause_result"]}' if last['cause_result'] else ''))
    print('\n' + NOTE)
    if not a.repo:
        print('Записи після 30.10.2026 тут не враховано: додайте --repo <клон stan-rozghliadu-arkhiv> або перевірте поточний файл data.gov.ua.')


# ---------- suddia / sud ----------
def cmd_suddia(a):
    duckdb = need_duckdb()
    m = get('stan_suddi_misiats.parquet', a); g = get('stan_suddi_intervaly.parquet', a)
    con = duckdb.connect()
    sud = f'%{a.sud}%' if a.sud else '%'
    jl = re.sub(r'\s+', ' ', a.suddia.strip())
    courts = con.execute(f"select distinct court_name from '{m}' where court_name ilike ? and suddia ilike ? order by 1",
                         [sud, jl + '%']).fetchall()
    if not courts:
        print('Суддю не знайдено (формат: «Прізвище І.Б.»; уточніть --sud).'); return
    for (c,) in courts:
        js = [x[0] for x in con.execute(f"select distinct suddia from '{m}' where court_name=? and suddia ilike ?", [c, jl + '%']).fetchall()]
        for j in js:
            print('=' * 100); print(f'{c} | суддя {j}')
            df = con.execute(f"""select a.ym місяць, a.provadzhen_z_rukhom "проваджень з рухом", a.rozpodileno розподілено,
                 a.pryznachen "(пере)призначень", a.rozghlianuto розглянуто, a.bez_rukhu "без руху", a.povernuto повернуто,
                 b.provadzhen_z_rukhom "суд: з рухом", b.rozpodileno "суд: розподілено", b.rozghlianuto "суд: розглянуто"
                 from '{m}' a left join '{m}' b on b.court_name=a.court_name and b.suddia='*' and b.ym=a.ym
                 where a.court_name=? and a.suddia=? order by 1""", [c, j]).fetchdf()
            print(df.to_string(index=False))
            vyd = f'%{a.vyd}%' if a.vyd else '%'
            df = con.execute(f"""select a.type вид, a.intervaliv "інтервалів", a.med_interval "медіана інт.", a.p90_interval "p90 інт.",
                 b.med_interval "суд: медіана інт.", b.p90_interval "суд: p90", a.rozghl_n "розглянуто", a.med_dniv_reiestr_rozghl "днів реєстр.→розгляд",
                 b.med_dniv_reiestr_rozghl "суд: днів реєстр.→розгляд"
                 from '{g}' a left join '{g}' b on b.court_name=a.court_name and b.suddia='*' and b.type=a.type
                 where a.court_name=? and a.suddia=? and a.type ilike ? and coalesce(a.intervaliv,0)+coalesce(a.rozghl_n,0) >= 5
                 order by coalesce(a.intervaliv,0)+coalesce(a.rozghl_n,0) desc""", [c, j, vyd]).fetchdf()
            print('\nІнтервали між датами стадії «Призначено до судового розгляду» та тривалість (днів), суддя поряд із судом:')
            print(df.to_string(index=False))
    print('\nПримітка: показники обчислено за змінами стану розгляду (01.2023–09.2026, для інтервалів і тривалості — до 30.10.2026); '
          'це непрямі дані про рух справ, а не офіційна статистика навантаження. «Розподілено» — провадження з записом '
          '«Призначено склад суду / Передано судді (в доповідь)» у місяці (охоплює й повторні розподіли). Інтервали — між '
          'датами стадії «Призначено до судового розгляду» (дні засідань або внесення запису): у даних є лише частина '
          'засідань, тож інтервали завищені, а кількість (пере)призначень занижена; для порівняння судді із судом придатні, '
          'як абсолютні величини — лише орієнтир (розд. 4.4 Довідника). До 07.2023 у вивантаженнях '
          'лише події дня, серпень–листопад 2023 містять дні 0 КБ, червень–липень 2025 неповні.')


def cmd_sud(a):
    duckdb = need_duckdb()
    m = get('stan_suddi_misiats.parquet', a)
    con = duckdb.connect()
    for (c,) in con.execute(f"select distinct court_name from '{m}' where court_name ilike ? order by 1", [f'%{a.sud}%']).fetchall():
        print('=' * 100); print(c)
        print(con.execute(f"""select ym місяць, provadzhen_z_rukhom "з рухом", rozpodileno розподілено, pryznachen "(пере)призначень",
              rozghlianuto розглянуто, bez_rukhu "без руху", povernuto повернуто from '{m}' where court_name=? and suddia='*' order by 1""", [c]).fetchdf().to_string(index=False))
        print('\nСудді (сума проваджень з рухом по місяцях, перші місяць і останній місяць):')
        print(con.execute(f"""select suddia, sum(provadzhen_z_rukhom) "з рухом", min(ym) "з", max(ym) "по" from '{m}'
              where court_name=? and suddia not in ('*','') group by 1 order by 2 desc""", [c]).fetchdf().to_string(index=False))


# ---------- spysok ----------
def cmd_spysok(a):
    duckdb = need_duckdb()
    fs = all_shards(a)
    con = duckdb.connect(); con.execute('SET threads=4')
    w, p = ['1=1'], []
    if a.sud: w.append('court_name ilike ?'); p.append(f'%{a.sud}%')
    if a.suddia: w.append('(judge ilike ? or judges ilike ?)'); p += [f'%{a.suddia}%'] * 2
    if a.etap: w.append('stage_name ilike ?'); p.append(f'%{a.etap}%')
    if a.uchasnyk: w.append('participants ilike ?'); p.append(f'%{a.uchasnyk}%')
    if a.vyd: w.append('type ilike ?'); p.append(f'%{a.vyd}%')
    sd = "try_strptime(stage_date,'%d.%m.%Y')::date"

    def iso(s, name):
        for f in ('%Y-%m-%d', '%d.%m.%Y'):
            try:
                return dt.datetime.strptime(s.strip(), f).date().isoformat()
            except ValueError:
                pass
        sys.exit(f'{name}: не розпізнано дату «{s}» (потрібно РРРР-ММ-ДД або ДД.ММ.РРРР).')
    if a.vid: w.append(f'{sd} >= ?'); p.append(iso(a.vid, '--vid'))
    if a.do: w.append(f'{sd} <= ?'); p.append(iso(a.do, '--do'))
    q = f"""select court_name, case_number, case_proc, type, min(registration_date) registration, count(*) zapysiv,
            count(distinct stage_date) filter (where stage_name='Призначено до судового розгляду') pryznachen,
            min({sd}) persha_stadiia, max({sd}) ostannia_stadiia, any_value(judge) judge
            from read_parquet({fs}) where {' and '.join(w)} group by all order by court_name, persha_stadiia"""
    df = con.execute(q, p).fetchdf()
    print(df.to_string(index=False, max_rows=a.max))
    print(f'\nПроваджень: {len(df)}')
    print('Примітка: «pryznachen» — кількість різних дат стадії «Призначено до судового розгляду»; це не кількість засідань '
          '(у даних є лише частина засідань; розд. 4.4 Довідника).')
    if a.tsv:
        df.to_csv(a.tsv, sep='\t', index=False); print('Збережено:', a.tsv)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--dir'); ap.add_argument('--kesh', default=os.path.expanduser('~/.cache/stan_indeks'))
    sp = ap.add_subparsers(dest='cmd', required=True)
    s = sp.add_parser('sprava'); s.add_argument('case'); s.add_argument('--sud'); s.add_argument('--repo'); s.add_argument('--tsv')
    s = sp.add_parser('suddia'); s.add_argument('suddia'); s.add_argument('--sud'); s.add_argument('--vyd')
    s = sp.add_parser('sud'); s.add_argument('sud')
    s = sp.add_parser('spysok')
    for o in ('--sud', '--suddia', '--etap', '--uchasnyk', '--vyd', '--vid', '--do', '--tsv'):
        s.add_argument(o)
    s.add_argument('--max', type=int, default=200)
    a, rest = ap.parse_known_args()
    for o in ('dir', 'kesh'):  # дозволити --dir / --kesh після назви команди
        if f'--{o}' in rest:
            i = rest.index(f'--{o}')
            if i + 1 >= len(rest):
                ap.error(f'--{o} потребує значення')
            setattr(a, o, rest[i + 1]); del rest[i:i + 2]
    if rest:                   # до 1.3 зайві параметри мовчки ігнорувалися — фільтр з помилкою в назві не діяв
        ap.error(f'невідомі параметри для «{a.cmd}»: {" ".join(rest)} (див. --help)')
    {'sprava': cmd_sprava, 'suddia': cmd_suddia, 'sud': cmd_sud, 'spysok': cmd_spysok}[a.cmd](a)


if __name__ == '__main__':
    main()
