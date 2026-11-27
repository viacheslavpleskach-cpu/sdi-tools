#!/usr/bin/env python3
"""Відкриті дані Вищої ради правосуддя на data.gov.ua: попередні скарги скаржника чи щодо судді, склад
Дисциплінарної палати на дату, звірка «Реєстру_рішень_ВРП.csv», пошук у текстах актів, одноосібні акти.

Автор: В'ячеслав Плескач. Версія 1.1 (27.11.2026): `zvirka` звіряє також «номер без суфікса + дата», тож описки в номерах
(1175/0/15-19 замість 15-18) не дають хибних пропусків. 1.0 (27.11.2026) — перша версія.

Набори (розпорядник — ВРП; дзеркала на сайті ДСА немає):
  rishennia — https://data.gov.ua/dataset/16a4fb23-e65d-4a4a-956d-32f5f2d14dd4
      «Рішення ВРП, в тому числі її органів (з фіксацією поіменного голосування)»: JSON (~45 МБ) і архів текстів
      (~530 МБ; розширення .zip, насправді 7z; ~7 500 файлів DOCX/DOC/RTF/PDF, окремі думки — «_od»).
      Охоплення — з 23.11.2023; оновлюється приблизно раз на 1–2 місяці (на 27.11.2026 — до 22.09.2026).
      ОДИН ЗАПИС JSON = ОДНЕ ГОЛОСУВАННЯ, а не документ: ухвала палати про відмову чи повернення має по голосуванню
      на кожну скаргу (до 80+), у votingTitle — скаржник, суддя і суд. Документ = номер + дата.
  akty — https://data.gov.ua/dataset/04ccf0ad-c7e5-4e13-99f5-1b08498d1255
      «Акти ВРП»: переліки XLSX колегіальних актів 2017–27.11.2024 (номер, дата, назва, посилання на hcj.gov.ua),
      одноосібних актів 2019–2026 (з 2025 р. — ухвали дисциплінарних інспекторів, хоч стовпець зветься «Член ВРП»),
      тексти лише за 2017 – січень 2021 р. (тут не використовуються), окремі думки 2016–2023.

Команди:
  stan                                   — ресурси наборів, дати оновлення, охоплення JSON
  skarzhnyk "Прізвище І.Б." [--suddia "Прізвище"] [--teksty]
                                         — попередні рішення палат і ВРП за скаргами цієї особи
  suddia "Прізвище І.Б." [--sud "частина назви"] [--teksty]
                                         — рішення палат і ВРП щодо судді
  sklad --data РРРР-ММ-ДД [--palata 1|2|3|vrp]
                                         — хто голосував у палаті в цей день (або найближчі засідання)
  zvirka [--reiestr Реєстр_рішень_ВРП.csv] [--xlsx] [--out файл.csv]
                                         — рішення по суті, рішення за скаргами й окремі думки, яких немає в реєстрі
                                           практики (--xlsx — також переліки 2017–2024 набору «Акти ВРП»)
  poshuk "регулярний вираз" [--organ 1дп|2дп|3дп|vrp] [--vid РРРР-ММ-ДД] [--do РРРР-ММ-ДД] [--context 300]
                                         — пошук у текстах актів (наприклад, номер скарги «К-2972/6/7-24»)
  tekst <номер> [--data РРРР-ММ-ДД] [--dir тека]
                                         — витягти файл(и) акта з архіву
  odnoosibni [--khto "Прізвище"] [--rik 2026] [--nazva "регулярний вираз"]
                                         — одноосібні акти (з 2025 р. — ухвали інспекторів) за переліками XLSX
Спільні: --out файл.tsv|csv, --onovyty (завантажити заново, без кешу).

Кеш — ~/.cache/vrp_od/ (метадані — 12 год; файли — доки на порталі не з'явиться новий ресурс). Архів текстів
розпаковується один раз (~0,9 ГБ) лише для --teksty, poshuk, tekst і перевірки окремих думок у zvirka.
Перший рядок виводу — джерело й дата даних: зазначати у звіті в чаті (у тексті висновку — ні).

Застереження: пошук за прізвищем дає й тезок — перевіряти ініціали, суд і текст ухвали; скарги, ще не розглянуті
палатами, у наборі відсутні; акти новіші за останню дату JSON шукати на hcj.gov.ua (з робочого середовища Claude
сайт недоступний — помилка TLS-сертифіката, обходити не можна) або в «Реєстрі_рішень_ВРП.csv».
"""
import argparse, collections, csv, datetime as dt, glob, json, os, re, shutil, subprocess, sys, time, zipfile

PKG = {'rishennia': '16a4fb23-e65d-4a4a-956d-32f5f2d14dd4', 'akty': '04ccf0ad-c7e5-4e13-99f5-1b08498d1255'}
CACHE = os.path.expanduser('~/.cache/vrp_od')
META_AGE = 12 * 3600
ORGAN = {'Вища рада правосуддя': 'ВРП', 'Перша дисциплінарна палата': '1ДП',
         'Друга дисциплінарна палата': '2ДП', 'Третя дисциплінарна палата': '3ДП'}
PAL = {'1': 'Перша дисциплінарна палата', '1дп': 'Перша дисциплінарна палата',
       '2': 'Друга дисциплінарна палата', '2дп': 'Друга дисциплінарна палата',
       '3': 'Третя дисциплінарна палата', '3дп': 'Третя дисциплінарна палата',
       'vrp': 'Вища рада правосуддя', 'врп': 'Вища рада правосуддя', '0': 'Вища рада правосуддя'}
SUB = re.compile(r'^\s*(?:про|прo)\s+(притягнення|відмову\s+(?:у|в)\s+притягненні|залишення\s+без\s+змін\s+рішення|'
                 r'скасування\s+(?:повністю\s+|частково\s+)?рішення|часткове\s+скасування\s+рішення|зміну\s+рішення|'
                 r'змін\w*\s+рішення|розгляд\s+скарги\s+.{0,120}?на\s+рішення)', re.I)
APO = str.maketrans({'’': "'", 'ʼ': "'", '`': "'", '´': "'"})


def say(*a):
    print(*a, file=sys.stderr)


def curl(url, out=None, timeout=1800):
    cmd = ['curl', '-sSfL', '--retry', '2', '-m', str(timeout), url]
    if out:
        cmd[1:1] = ['-o', out]
    r = subprocess.run(cmd, capture_output=True)
    if r.returncode:
        raise RuntimeError(f'curl {url}: {r.stderr.decode(errors="ignore").strip()[:300]}')
    return r.stdout


def meta(kind, fresh=False):
    os.makedirs(CACHE, exist_ok=True)
    p = os.path.join(CACHE, f'meta_{kind}.json')
    if not fresh and os.path.exists(p) and time.time() - os.path.getmtime(p) < META_AGE:
        return json.load(open(p))
    try:
        m = json.loads(curl(f'https://data.gov.ua/api/3/action/package_show?id={PKG[kind]}', timeout=90))['result']
        json.dump(m, open(p, 'w'), ensure_ascii=False)
        return m
    except Exception as e:
        if os.path.exists(p):
            say(f'! data.gov.ua не відповідає ({e}); беру метадані з кешу від '
                f'{dt.datetime.fromtimestamp(os.path.getmtime(p)):%d.%m.%Y %H:%M}')
            return json.load(open(p))
        sys.exit(f'data.gov.ua не відповідає і кешу немає: {e}')


def resource(kind, fmt, fresh=False):
    """Найсвіжіший ресурс формату fmt ('JSON'/'ZIP') → (шлях, ресурс). Кеш за назвою файлу ресурсу."""
    m = meta(kind, fresh)
    rs = [r for r in m['resources'] if (r.get('format') or '').upper() == fmt]
    if not rs:
        sys.exit(f'У наборі {PKG[kind]} немає ресурсу {fmt}')
    r = max(rs, key=lambda r: r.get('last_modified') or r.get('created') or '')
    name = r['url'].rsplit('/', 1)[-1]
    path = os.path.join(CACHE, name)
    stamp = r.get('last_modified') or r.get('created') or ''
    sp = path + '.stamp'
    if fresh or not os.path.exists(path) or (os.path.exists(sp) and open(sp).read() != stamp):
        say(f'Завантажую {name} ({(r.get("size") or 0) / 1e6:.0f} МБ)…')
        tmp = path + '.part'
        curl(r['url'], tmp)
        os.replace(tmp, path)
        open(sp, 'w').write(stamp)
    return path, r


def norm_num(s):
    return re.sub(r'\s|№', '', str(s or '')).lower().replace('дп', 'дп')


def load_json(a):
    path, r = resource('rishennia', 'JSON', a.onovyty)
    d = json.load(open(path, encoding='utf-8'))
    for x in d:
        x['proceedingId'] = [p for p in (x.get('proceedingId') or []) if p]
        x['texts'] = x.get('texts') or []
        x['votingResult'] = x.get('votingResult') or []
        x['decisionNum'] = re.sub(r'\s+', '', (x.get('decisionNum') or '').replace('№', ''))
        x['day'] = (x.get('dateTime') or '')[:10]
        x['org'] = ORGAN.get(x.get('authority'), x.get('authority'))
    last = max(x['day'] for x in d)
    first = min(x['day'] for x in d)
    print(f'# Джерело: data.gov.ua {PKG["rishennia"]}, ресурс «{r.get("name")}» ({path.rsplit("/", 1)[-1]}), '
          f'оновлено {(r.get("last_modified") or "")[:10]}; голосування {first} – {last} ({len(d)} записів). '
          f'Акти після {last} у наборі відсутні.')
    return d


def ddmm(day):
    return f'{day[8:10]}.{day[5:7]}.{day[:4]}' if len(day) >= 10 else day


# ---------- імена ----------

def name_rx(q):
    """«Стояновський В.В.» / «Стояновський Валерій Володимирович» / «Стояновський» → regex для будь-якого відмінка
    й запису (ініціали чи повне ім'я)."""
    q = q.translate(APO).strip()
    parts = [p for p in re.split(r'[\s.]+', q) if p]
    if not parts:
        sys.exit('Порожнє ім\'я')
    sur = parts[0]
    low = sur.lower()
    # основа: відкидаємо закінчення називного відмінка
    for end in ('ський', 'цький', 'зький', 'ська', 'цька', 'зька', 'ий', 'ій', 'ая', 'а', 'я', 'о', 'ь', 'е', 'і', 'ї'):
        if low.endswith(end) and len(low) - len(end) >= 3:
            stem = sur[:len(sur) - len(end)]
            if end in ('ський', 'цький', 'зький', 'ська', 'цька', 'зька'):
                stem = sur[:len(sur) - len(end) + 2]   # «Стояновськ»
            break
    else:
        stem = sur if len(sur) <= 4 else sur[:-1]
    stem = re.escape(stem).replace("'", "['’ʼ`]?")
    rest = parts[1:]
    if not rest:
        return re.compile(rf'(?<![А-ЯІЇЄҐа-яіїєґ]){stem}\w*', re.I)
    ini = [p[0] for p in rest[:2]]
    tail = r'\s*'.join(rf'(?:{i}\.|{i}\w+)' for i in ini)
    return re.compile(rf'(?<![А-ЯІЇЄҐа-яіїєґ]){stem}\w*\s+{tail}', re.I)


VT = re.compile(r'(?is)скарг\w*\s+(?P<who>.{3,220}?)\s+(?:стосовно|щодо)\s+(?:судді|суддів)\s+(?P<judge>.{5,260}?)'
                r'(?=\s+(?:залишити|та\s+повернути|повернути|відмовити)|\s*[;]|\s*$)')


def vote_parts(x):
    t = re.sub(r'\s+', ' ', x.get('votingTitle') or '')
    m = VT.search(t)
    who, judge = (m['who'].strip(' ,'), m['judge'].strip(' ,.')) if m else ('', '')
    who = re.split(r'\s+(?:від\s+\d|в\s+частині|за\s+вхідним|№)', who)[0].strip(' ,')
    low = t.lower()
    if 'відмовити у відкритті' in low or 'відмовити в відкритті' in low:
        res = 'відмова у відкритті'
    elif 'залишити без розгляду' in low and 'повернути' in low:
        res = 'залишено без розгляду, повернуто'
    elif 'залишити без розгляду' in low:
        res = 'залишено без розгляду'
    elif re.search(r'відкрити\s+дисциплінарн', low):
        res = 'відкрито справу'
    elif re.search(r'повернути', low):
        res = 'повернуто'
    elif re.search(r'притягнути', low):
        res = 'притягнуто'
    elif re.search(r'відмовити\s+(у|в)\s+притягненні', low):
        res = 'відмовлено у притягненні'
    else:
        res = re.sub(r'^\s*Про\s+', '', x.get('title') or '')[:80]
    return who, judge, res, t


# ---------- тексти ----------

def texts_dir(a):
    path, r = resource('rishennia', 'ZIP', a.onovyty)
    d = os.path.join(CACHE, 'teksty_' + os.path.basename(path).rsplit('.', 1)[0])
    if not os.path.exists(os.path.join(d, '.ok')):
        say(f'Розпаковую {os.path.basename(path)} (один раз, ~1–3 хв)…')
        shutil.rmtree(d, ignore_errors=True)
        os.makedirs(d)
        with open(path, 'rb') as f:
            sig = f.read(6)
        if sig.startswith(b'7z\xbc\xaf'):
            try:
                import py7zr
            except ImportError:
                subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', '--break-system-packages', 'py7zr'], check=True)
                import py7zr
            with py7zr.SevenZipFile(path) as z:
                z.extractall(d)
        else:
            zipfile.ZipFile(path).extractall(d)
        open(os.path.join(d, '.ok'), 'w').write(r.get('last_modified') or '')
    return d


def to_text(fp):
    """Текст файла акта; кеш поруч (.txt)."""
    tp = fp + '.txt'
    if os.path.exists(tp):
        return open(tp, encoding='utf-8', errors='ignore').read()
    ext = fp.rsplit('.', 1)[-1].lower()
    t = ''
    try:
        if ext == 'docx':
            import docx
            t = '\n'.join(p.text for p in docx.Document(fp).paragraphs)
        elif ext == 'rtf':
            try:
                from striprtf.striprtf import rtf_to_text
            except ImportError:
                subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', '--break-system-packages', 'striprtf'], check=True)
                from striprtf.striprtf import rtf_to_text
            t = rtf_to_text(open(fp, encoding='latin-1', errors='ignore').read(), errors='ignore')
        elif ext == 'pdf':
            t = subprocess.run(['pdftotext', '-layout', fp, '-'], capture_output=True, timeout=120).stdout.decode('utf-8', 'ignore')
        elif ext == 'doc':
            out = os.path.dirname(fp)
            subprocess.run(['soffice', '--headless', '--convert-to', 'txt:Text', '--outdir', out + '/.soffice', fp],
                           capture_output=True, timeout=180)
            q = os.path.join(out, '.soffice', os.path.basename(fp).rsplit('.', 1)[0] + '.txt')
            if os.path.exists(q):
                t = open(q, encoding='utf-8', errors='ignore').read()
    except Exception as e:
        say(f'! {os.path.basename(fp)}: {e}')
    open(tp, 'w', encoding='utf-8').write(t)
    return t


def file_path(tdir, name):
    p = os.path.join(tdir, name)
    if os.path.exists(p):
        return p
    hits = glob.glob(os.path.join(tdir, '**', glob.escape(name)), recursive=True)
    return hits[0] if hits else None


def complaint_block(text, who_rx):
    """Абзац «N) До ВРП надійшла скарга …» про цього скаржника: номер скарги, інспектор, пропозиція, підстава."""
    tt = re.sub(r'\s+', ' ', text)
    parts = re.split(r'(?i)(?=(?:\b\d+\)\s*)?до\s+вищої\s+ради\s+правосуддя\s+надійшл)', tt)
    out = []
    for p in parts[1:]:
        p = p[:2500]
        if not who_rx.search(p[:500]):
            continue
        nums = re.findall(r'№\s*([А-ЯІЇЄҐ]{1,3}-\d+/\d+/\d+-\d+)', p[:900])
        ins = re.search(r'інспектор\w*\s+Вищої ради правосуддя\s*[–-]\s*доповідач\w*\s+([А-ЯІЇЄҐ][а-яіїєґ’\'-]+\s+'
                        r'[А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ]\.)\s+скла\w*\s+висновок\s+про\s+(.{5,140}?)(?:,|\.)', p)
        gr = re.findall(r'пункт\w*\s+(\d+)\s+частини\s+першої\s+статті\s+(4[345])', p)
        out.append({'skargy': ', '.join(dict.fromkeys(nums)), 'inspektor': ins[1] if ins else '',
                    'vysnovok': ins[2] if ins else '',
                    'pidstava': '; '.join(dict.fromkeys(f'п. {g[0]} ч. 1 ст. {g[1]}' for g in gr))})
    return out


# ---------- команди ----------

def write(rows, cols, a, sep='\t'):
    if a.out:
        sep = ',' if a.out.endswith('.csv') else '\t'
        with open(a.out, 'w', encoding='utf-8-sig' if sep == ',' else 'utf-8', newline='') as f:
            w = csv.writer(f, delimiter=sep)
            w.writerow(cols)
            w.writerows([[r.get(c, '') for c in cols] for r in rows])
        print(f'# Записано {len(rows)} рядків у {a.out}')
    else:
        print('\t'.join(cols))
        for r in rows:
            print('\t'.join(str(r.get(c, '')).replace('\t', ' ') for c in cols))


def cmd_stan(a):
    for k in PKG:
        m = meta(k, a.onovyty)
        print(f'== {m["title"][:110]}\n   https://data.gov.ua/dataset/{PKG[k]}  змінено {m["metadata_modified"][:10]}')
        for r in sorted(m['resources'], key=lambda r: r.get('name') or ''):
            print(f'   {r.get("format","?"):5} {(r.get("size") or 0) / 1e6:8.1f} МБ  {(r.get("last_modified") or r.get("created") or "")[:10]}  {r.get("name")}')
    d = load_json(a)
    c = collections.Counter((x['org'], x['day'][:4]) for x in d)
    docs = {(x['decisionNum'], x['day']) for x in d}
    print(f'# Документів (номер + дата): {len(docs)}; голосувань: {len(d)}')
    for o in ('ВРП', '1ДП', '2ДП', '3ДП'):
        print(f'   {o}: ' + ', '.join(f'{y} — {c[(o, y)]}' for y in sorted({x["day"][:4] for x in d})))


def search_votes(d, rx, field):
    rows = []
    for x in d:
        who, judge, res, t = vote_parts(x)
        hay = who if field == 'who' else judge
        if not hay and field == 'who':      # рішення ВРП і по суті: скаржник у назві чи тексті голосування
            hay = (x.get('title') or '') + ' ' + t
        if not hay and field == 'judge':
            hay = (x.get('title') or '') + ' ' + t
        if rx.search(hay):
            rows.append({'data': ddmm(x['day']), 'organ': x['org'], 'nomer': x['decisionNum'], 'rezultat': res,
                         'skarzhnyk': who, 'suddia': judge, 'fail': '; '.join(x['texts']),
                         'nazva': re.sub(r'\s+', ' ', x.get('title') or '')[:200], '_x': x})
    rows.sort(key=lambda r: (r['_x']['day'], r['nomer']))
    return rows


def add_texts(rows, a, rx_who):
    tdir = texts_dir(a)
    for r in rows:
        blocks = []
        for n in r['_x']['texts']:
            fp = file_path(tdir, n)
            if fp:
                blocks += complaint_block(to_text(fp), rx_who)
        if blocks:
            for k in ('skargy', 'inspektor', 'vysnovok', 'pidstava'):
                r[k] = ' | '.join(dict.fromkeys(b[k] for b in blocks if b[k]))


def variants(rows, key):
    c = collections.Counter(re.sub(r'\s+', ' ', r[key]) for r in rows if r[key])
    if len(c) > 1:
        print(f'# Варіанти запису ({key}) — перевірте тезок: ' + '; '.join(f'{k} — {v}' for k, v in c.most_common(8)))


def cmd_skarzhnyk(a):
    d = load_json(a)
    rx = name_rx(a.imia)
    rows = search_votes(d, rx, 'who')
    if a.suddia:
        rj = name_rx(a.suddia)
        rows = [r for r in rows if rj.search(r['suddia'] or r['nazva'])]
    if a.teksty and rows:
        add_texts(rows, a, rx)
    print(f'# Скаржник «{a.imia}»{" щодо судді «" + a.suddia + "»" if a.suddia else ""}: {len(rows)} голосувань у '
          f'{len({(r["nomer"], r["data"]) for r in rows})} актах')
    variants(rows, 'skarzhnyk')
    cols = ['data', 'organ', 'nomer', 'rezultat', 'skarzhnyk', 'suddia', 'fail']
    if a.teksty:
        cols += ['skargy', 'inspektor', 'vysnovok', 'pidstava']
    write(rows, cols, a)


def cmd_suddia(a):
    d = load_json(a)
    rx = name_rx(a.imia)
    rows = search_votes(d, rx, 'judge')
    if a.sud:
        s = a.sud.lower()
        rows = [r for r in rows if s in (r['suddia'] + ' ' + r['nazva']).lower()]
    if a.teksty and rows:
        add_texts(rows, a, re.compile('.'))
    print(f'# Суддя «{a.imia}»{" (" + a.sud + ")" if a.sud else ""}: {len(rows)} голосувань у '
          f'{len({(r["nomer"], r["data"]) for r in rows})} актах')
    variants(rows, 'suddia')
    cols = ['data', 'organ', 'nomer', 'rezultat', 'skarzhnyk', 'suddia', 'fail']
    write(rows, cols + (['skargy', 'inspektor', 'vysnovok', 'pidstava'] if a.teksty else []), a)


def short(v):
    p = v.split()
    if len(p) >= 3:
        return f'{p[0].capitalize()} {p[1][0]}.{p[2][0]}.'
    return v.title()


def cmd_sklad(a):
    d = load_json(a)
    orgs = [PAL[a.palata.lower()]] if a.palata else ['Перша дисциплінарна палата', 'Друга дисциплінарна палата',
                                                       'Третя дисциплінарна палата']
    for o in orgs:
        days = sorted({x['day'] for x in d if x['authority'] == o})
        if a.data in days:
            pick = [a.data]
            note = ''
        else:
            before = [x for x in days if x < a.data][-1:]
            after = [x for x in days if x > a.data][:1]
            pick = before + after
            note = ' (засідання цього дня в наборі немає; найближчі: ' + ', '.join(ddmm(p) for p in pick) + ')'
        print(f'== {o}{note}')
        for p in pick:
            c = collections.Counter()
            n = 0
            for x in d:
                if x['authority'] == o and x['day'] == p:
                    n += 1
                    for v in x['votingResult']:
                        c[v.get('voterName', '').strip()] += 1
            print(f'   {ddmm(p)}: голосувань {n}; голосували: ' + ', '.join(f'{short(k)} ({v})' for k, v in sorted(c.items())))
    print('# Головуючого набір не зазначає — брати з тексту ухвали чи рішення («у складі головуючого – …»).')


class RegKeys(set):
    """Номери реєстру + пари «номер без суфікса, дата» (з колонки date або назви файлу): описка в номері
    («15-19» замість «15-18») не дає хибного «немає в реєстрі»."""
    def __init__(self, R):
        super().__init__(norm_num(r['number']) for r in R)
        self.pd = set()
        for r in R:
            p = norm_num(r['number']).split('/')[0]
            for d in [r.get('date') or ''] + re.findall(r'\d\d\.\d\d\.\d{4}', r.get('title') or ''):
                if d:
                    self.pd.add((p, d))

    def has(self, num, day):
        return norm_num(num) in self or (norm_num(num).split('/')[0], ddmm(day) if '-' in day else day) in self.pd


def reg_rows(path):
    R = list(csv.DictReader(open(path, encoding='utf-8')))
    return R, RegKeys(R)


def cmd_zvirka(a):
    reg = a.reiestr or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Реєстр_рішень_ВРП.csv')
    R, RN = reg_rows(reg)
    d = load_json(a)
    last = max(x['day'] for x in d)
    first = min(x['day'] for x in d)
    out = []
    seen = set()
    for x in d:
        k = (x['decisionNum'], x['day'])
        if k in seen or not SUB.search(x.get('title') or ''):
            continue
        seen.add(k)
        if not RN.has(x['decisionNum'], x['day']):
            out.append({'dzherelo': 'rishennia JSON', 'data': ddmm(x['day']), 'nomer': x['decisionNum'], 'organ': x['org'],
                        'nazva': re.sub(r'\s+', ' ', x['title']), 'fail': '; '.join(x['texts']) or '(тексту немає)',
                        'posylannia': '', 'golosy': f'{x.get("votingFor")}/{x.get("votingAgainst")}'})
    # окремі думки (файли «_od» в архіві)
    try:
        tdir = texts_dir(a)
        reg_od = {re.split(r'[_.]', r['title'])[0] + '_' + (re.search(r'\d\d\.\d\d\.\d{4}', r['title']) or [''])[0]
                  for r in R if 'окрем' in r['folder'].lower() or '_od' in r['title'].lower()}
        for fp in sorted(glob.glob(os.path.join(tdir, '**', '*_od*'), recursive=True)):
            if fp.endswith('.txt'):
                continue
            n = os.path.basename(fp)
            m = re.match(r'(\d+)[_ ,-]+(\d\d\.\d\d\.\d{4})', n)
            if m and f'{m[1]}_{m[2]}' not in reg_od:
                out.append({'dzherelo': 'rishennia архів', 'data': m[2], 'nomer': m[1], 'organ': '',
                            'nazva': 'Окрема думка (у реєстрі немає)', 'fail': n, 'posylannia': '', 'golosy': ''})
    except SystemExit:
        raise
    except Exception as e:
        print(f'# ! Окремі думки не перевірено: {e}')
    if a.xlsx:
        out += zvirka_xlsx(a, RN, {o['nomer'] for o in out})
    # зворотна звірка: рядки реєстру в періоді JSON, яких немає в JSON
    J = {norm_num(x['decisionNum']) for x in d}
    back = []
    for r in R:
        m = re.match(r'(\d\d)\.(\d\d)\.(\d{4})', r['date'] or '')
        if not m:
            continue
        iso = f'{m[3]}-{m[2]}-{m[1]}'
        if first <= iso <= last and norm_num(r['number']) not in J:
            back.append(f'{r["date"]} {r["number"]} ({r["folder"]})')
    out.sort(key=lambda o: (o['data'][-4:], o['data'][3:5], o['data'][:2]))
    print(f'# Реєстр: {reg} ({len(R)} рядків). Немає в реєстрі: {len(out)}.')
    print(f'# Рядки реєстру за {ddmm(first)}–{ddmm(last)}, яких немає в JSON за номером ({len(back)}; часто — описка '
          f'в номері): ' + '; '.join(back[:40]))
    write(out, ['dzherelo', 'data', 'nomer', 'organ', 'nazva', 'fail', 'posylannia', 'golosy'], a)


def xlsx_lists(a, single):
    try:
        import openpyxl
    except ImportError:
        subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', '--break-system-packages', 'openpyxl'], check=True)
        import openpyxl
    m = meta('akty', a.onovyty)
    rows = []
    for r in m['resources']:
        if (r.get('format') or '').upper() not in ('XLSX', 'XLS'):
            continue
        name = r['url'].rsplit('/', 1)[-1]
        if ('single' in name) != single or 'onovleno' in name:
            continue
        p = os.path.join(CACHE, name)
        if a.onovyty or not os.path.exists(p):
            curl(r['url'], p, 300)
        ws = openpyxl.load_workbook(p, read_only=True).worksheets[0]
        it = ws.iter_rows(values_only=True)
        hdr = [str(c).strip().lower() if c else '' for c in next(it)]
        next(it, None)          # другий рядок — українські назви стовпців
        for v in it:
            if not any(v):
                continue
            x = dict(zip(hdr, v))
            x['_res'] = name
            rows.append(x)
    return rows


def zvirka_xlsx(a, RN, have):
    out = []
    for x in xlsx_lists(a, single=False):
        t = re.sub(r'\s+', ' ', str(x.get('title') or ''))
        n = re.sub(r'\s|№', '', str(x.get('number') or ''))
        dd = str(x.get('date') or '')[:10]
        if SUB.search(t) and not RN.has(n, dd) and n not in have:
            out.append({'dzherelo': x['_res'], 'data': ddmm(dd), 'nomer': n, 'organ': '', 'nazva': t, 'fail': '',
                        'posylannia': x.get('link') or x.get('url') or '',
                        'golosy': f'{x.get("vote_for") or ""}/{x.get("vote_agains") or ""}'})
    return out


def cmd_poshuk(a):
    d = load_json(a)
    rx = re.compile(a.vyraz, re.I)
    org = PAL.get((a.organ or '').lower())
    docs = {}
    for x in d:
        if org and x['authority'] != org:
            continue
        if a.vid and x['day'] < a.vid or a.do and x['day'] > a.do:
            continue
        for n in x['texts']:
            docs.setdefault(n, x)
    tdir = texts_dir(a)
    print(f'# Переглядаю {len(docs)} файлів (DOC — через LibreOffice, перший раз повільно; звужуйте --organ/--vid/--do)')
    hits = 0
    for n, x in sorted(docs.items(), key=lambda kv: kv[1]['day']):
        fp = file_path(tdir, n)
        if not fp:
            continue
        t = re.sub(r'\s+', ' ', to_text(fp))
        ms = list(rx.finditer(t))
        if not ms:
            continue
        hits += 1
        print(f'\n== {ddmm(x["day"])} {x["org"]} № {x["decisionNum"]} — {n}\n   {re.sub(r"\s+", " ", x["title"])[:160]}')
        for m in ms[:a.max]:
            s = max(0, m.start() - a.context)
            print('   …' + t[s:m.end() + a.context] + '…')
        if a.limit and hits >= a.limit:
            print(f'# Зупинено на {hits} документах (--limit)')
            break
    print(f'\n# Документів зі збігом: {hits}')


def cmd_tekst(a):
    d = load_json(a)
    n = norm_num(a.nomer)
    xs = [x for x in d if norm_num(x['decisionNum']) == n and (not a.data or x['day'] == a.data)]
    if not xs:
        sys.exit(f'Акта № {a.nomer} у наборі немає (охоплення — див. перший рядок).')
    tdir = texts_dir(a)
    os.makedirs(a.dir, exist_ok=True)
    names = sorted({t for x in xs for t in x['texts']})
    # додатки й окремі думки до того самого акта
    day = xs[0]['day']
    base = f'{n.split("/")[0]}_{ddmm(day)}'
    names += [os.path.basename(p) for p in glob.glob(os.path.join(tdir, '**', base + '*'), recursive=True)
              if not p.endswith('.txt') and os.path.basename(p) not in names]
    for t in names:
        fp = file_path(tdir, t)
        if fp:
            shutil.copy(fp, os.path.join(a.dir, t))
            print(os.path.join(a.dir, t))
        else:
            print(f'! {t} — в архіві немає')
    print(f'# {xs[0]["org"]} № {xs[0]["decisionNum"]} від {ddmm(day)}: {re.sub(chr(10), " ", xs[0]["title"])[:200]}; '
          f'голосувань {len(xs)}; «за» / «проти»: ' + ', '.join(f'{x.get("votingFor")}/{x.get("votingAgainst")}' for x in xs[:10]))


def cmd_odnoosibni(a):
    rows = xlsx_lists(a, single=True)
    out = []
    rx = re.compile(a.nazva, re.I) if a.nazva else None
    khto = name_rx(a.khto) if a.khto else None
    for x in rows:
        dd = x.get('date')
        dd = dd.strftime('%Y-%m-%d') if hasattr(dd, 'strftime') else str(dd or '')[:10]
        if a.rik and not dd.startswith(str(a.rik)):
            continue
        member = re.sub(r'\s+', ' ', str(x.get('member') or '')).strip()
        if khto and not khto.search(re.sub(r'^([А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ]\.)\s?(.+)$', r'\2 \1', member)):
            continue
        t = re.sub(r'\s+', ' ', str(x.get('title') or ''))
        if rx and not rx.search(t):
            continue
        out.append({'data': ddmm(dd), 'nomer': x.get('number'), 'khto': member, 'nazva': t,
                    'posylannia': x.get('link') or x.get('url') or '', 'resurs': x['_res']})
    c = collections.Counter(re.sub(r'^\s*Про\s+', '', o['nazva'])[:45] for o in out)
    print(f'# Одноосібні акти (data.gov.ua {PKG["akty"]}): {len(out)}. Найчастіші: ' +
          '; '.join(f'{k}… — {v}' for k, v in c.most_common(5)))
    print('# З 2025 р. стовпець «Член ВРП» містить дисциплінарного інспектора. Дати з помилками (напр. 2027-03-27) — як у джерелі.')
    write(out, ['data', 'nomer', 'khto', 'nazva', 'posylannia', 'resurs'], a)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest='cmd', required=True)
    com = argparse.ArgumentParser(add_help=False)
    com.add_argument('--out')
    com.add_argument('--onovyty', action='store_true')
    sp.add_parser('stan', parents=[com])
    p = sp.add_parser('skarzhnyk', parents=[com]); p.add_argument('imia'); p.add_argument('--suddia'); p.add_argument('--teksty', action='store_true')
    p = sp.add_parser('suddia', parents=[com]); p.add_argument('imia'); p.add_argument('--sud'); p.add_argument('--teksty', action='store_true')
    p = sp.add_parser('sklad', parents=[com]); p.add_argument('--data', required=True); p.add_argument('--palata')
    p = sp.add_parser('zvirka', parents=[com]); p.add_argument('--reiestr'); p.add_argument('--xlsx', action='store_true')
    p = sp.add_parser('poshuk', parents=[com]); p.add_argument('vyraz'); p.add_argument('--organ'); p.add_argument('--vid'); p.add_argument('--do')
    p.add_argument('--context', type=int, default=300); p.add_argument('--max', type=int, default=3); p.add_argument('--limit', type=int, default=0)
    p = sp.add_parser('tekst', parents=[com]); p.add_argument('nomer'); p.add_argument('--data'); p.add_argument('--dir', default='vrp_teksty')
    p = sp.add_parser('odnoosibni', parents=[com]); p.add_argument('--khto'); p.add_argument('--rik'); p.add_argument('--nazva')
    a = ap.parse_args()
    {'stan': cmd_stan, 'skarzhnyk': cmd_skarzhnyk, 'suddia': cmd_suddia, 'sklad': cmd_sklad, 'zvirka': cmd_zvirka,
     'poshuk': cmd_poshuk, 'tekst': cmd_tekst, 'odnoosibni': cmd_odnoosibni}[a.cmd](a)


if __name__ == '__main__':
    main()
