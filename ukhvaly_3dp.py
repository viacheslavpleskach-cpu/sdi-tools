# -*- coding: utf-8 -*-
"""
ukhvaly_3dp.py — підготовка проєктів ухвал Третьої Дисциплінарної палати ВРП:
  «Про залишення без розгляду та повернення дисциплінарних скарг» (розділ 3 порядку денного)
  «Про відмову у відкритті дисциплінарних справ» (розділ 2 порядку денного)

Автор: В'ячеслав Плескач

Команди:
  extract --agenda ПОРЯДОК.docx --in ФАЙЛИ_АБО_ТЕКИ_АБО_АРХІВИ... --out ДАНІ.json [--report ЗВІТ.md]
      Конвертує висновки (DOCX/DOC/RTF, також у .zip/.7z), розбирає розділи 2 і 3 порядку денного,
      витягує з висновків реквізити (номери і дати скарг, інспектор, пропозиція, підстави),
      зіставляє висновки з пунктами порядку денного і пише звіт про прогалини й розбіжності.
  build --data ПРОЄКТ.json --out УХВАЛА.docx [--template Шаблон_ухвали_3ДП.docx | …_docx.b64]
      Збирає DOCX ухвали з тексту пунктів, складеного за промтом.

Залежності: pandoc, LibreOffice (для DOC/RTF), python-docx, py7zr (для .7z).
"""
import argparse, base64, copy, glob, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile

MONTHS = ['січня', 'лютого', 'березня', 'квітня', 'травня', 'червня', 'липня', 'серпня',
          'вересня', 'жовтня', 'листопада', 'грудня']
NUM_RE = r'(?:[А-ЯІЇЄҐA-Z]\s?-\s?)?\d{1,5}/\d{1,3}/\d{1,2}-\d{2}'
DATE_W = r'«?\s*(\d{1,2})\s*»?\s+(' + '|'.join(MONTHS) + r')\s+(\d{4})'
DATE_N = r'(\d{1,2})\.(\d{2})\.(\d{4})'

# Відомі інспектори: прізвище -> (як у ухвалі, дієслово)
INSPECTORS = {
    'ДАВИДОВИЧ': ('Давидович І.І.', 'склала'),
    'ЗОЛОТАР': ('Золотар М.М.', 'склав'),
    'САДОВИЙ': ('Садовий Р.М.', 'склав'),
    'ЧЕРЕП': ('Череп В.Ю.', 'склав'),
    'ПЛЕСКАЧ': ('Плескач В.Ю.', 'склав'),
    'ОСЄТРОВ': ('Осєтров В.М.', 'склав'),
    'НЕДИБАЛЮК': ('Недибалюк В.Д.', 'склала'),
}


def norm(s):
    s = re.sub('[\u200b\u200c\u200d\u2060\ufeff\u00ad]', '', s)
    return (s.replace('‑', '-').replace('–', '-').replace('‐', '-')
             .replace('\xa0', ' ').replace(' ', ' '))


LAT2CYR = str.maketrans('ABCEHIKMOPTXY', 'АВСЕНІКМОРТХУ')


def norm_num(n):
    return re.sub(r'\s', '', norm(n)).translate(LAT2CYR)


def fmt_date(d, m, y):
    return f'{int(d)} {m} {y} року'


def date_num_to_words(d, m, y):
    return fmt_date(d, MONTHS[int(m) - 1], y)


# ---------------------------------------------------------------- конвертація
def to_text(path, workdir):
    ext = os.path.splitext(path)[1].lower()
    src = path
    if ext in ('.doc', '.rtf', '.odt'):
        subprocess.run(['soffice', '--headless', '--convert-to', 'docx', '--outdir', workdir, path],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
        src = os.path.join(workdir, os.path.splitext(os.path.basename(path))[0] + '.docx')
        if not os.path.exists(src):
            return None
    r = subprocess.run(['pandoc', src, '-t', 'plain', '--wrap=none'], capture_output=True, timeout=180)
    if r.returncode != 0:
        return None
    return norm(r.stdout.decode('utf-8', 'replace'))


def unpack(path, dest):
    os.makedirs(dest, exist_ok=True)
    if path.lower().endswith('.zip'):
        z = zipfile.ZipFile(path)
        for i in z.infolist():
            n = i.filename
            if not (i.flag_bits & 0x800):
                for enc in ('utf-8', 'cp866'):
                    try:
                        n = n.encode('cp437').decode(enc); break
                    except Exception:
                        pass
            if n.endswith('/'):
                continue
            p = os.path.join(dest, os.path.basename(n))
            with open(p, 'wb') as f:
                f.write(z.read(i))
    elif path.lower().endswith('.7z'):
        import py7zr
        with py7zr.SevenZipFile(path) as z:
            z.extractall(dest)


def collect_inputs(paths, workdir):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, fs in os.walk(p):
                for f in fs:
                    files.append(os.path.join(root, f))
        else:
            files.append(p)
    out = []
    for f in files:
        low = f.lower()
        if low.endswith(('.zip', '.7z')):
            d = tempfile.mkdtemp(dir=workdir)
            unpack(f, d)
            out += collect_inputs([d], workdir)
        elif low.endswith(('.docx', '.doc', '.rtf', '.odt')) and not os.path.basename(f).startswith('~$'):
            out.append(f)
    return out


# ---------------------------------------------------------------- порядок денний
def parse_agenda(text):
    t = text
    i2 = t.find('2.  Розгляд висновків')
    if i2 < 0:
        i2 = t.find('Розгляд висновків дисциплінарних інспекторів Вищої ради правосуддя про відкриття')
    i3 = t.find('складених за результатами попередньої перевірки')
    i3 = t.rfind('\n', 0, i3)
    i4 = t.find('Процедурні і організаційні питання')
    sec2 = t[i2:i3]
    sec3 = t[i3:i4 if i4 > 0 else len(t)]
    m = re.search(r'засідання Третьої Дисциплінарної палати Вищої ради правосуддя\s+' + DATE_W, t)
    date = fmt_date(*m.groups()) if m else None
    mnum = re.search(r'Порядок денний\s*№\s*(\d+)', t)
    return {'date': date, 'agenda_no': mnum.group(1) if mnum else None,
            'section2': parse_section(sec2), 'section3': parse_section(sec3)}


def parse_section(sec):
    items = []
    chunks = re.split(r'\n(?=\s*\d+\)\s)', '\n' + sec)
    pending = []
    for ch in chunks:
        m = re.match(r'\s*(\d+)\)\s+(.*)', ch, re.S)
        if not m:
            continue
        body = m.group(2)
        mi = re.search(r'\(дисциплінарний інспектор[^)]*доповідач\s+([^)]+?)\s*\)', body)
        text = re.sub(r'\(дисциплінарний інспектор[^)]*\)', '', body)
        text = re.sub(r'\s+', ' ', text).strip()
        it = {'n': int(m.group(1)), 'text': text,
              'numbers': [norm_num(x) for x in re.findall(NUM_RE, text)], 'inspector': None}
        items.append(it)
        pending.append(it)
        if mi:
            ins = re.sub(r'\s+', ' ', mi.group(1)).strip().rstrip('.') + '.'
            for p in pending:
                p['inspector'] = ins
            pending = []
    return items


# ---------------------------------------------------------------- висновки
def find_dates_near(t, num_pos, num_len):
    """Дата скарги біля номера: спершу «від <дата>» одразу після номера, далі найближча дата у вікні."""
    after = t[num_pos + num_len: num_pos + num_len + 60]
    m = re.match(r'[\s,]*(?:від|від\s+)\s*' + DATE_W, after)
    if m:
        return fmt_date(*m.groups())
    m = re.match(r'[\s,]*(?:від)\s*' + DATE_N, after)
    if m:
        return date_num_to_words(*m.groups())
    # Осєтров: «[дата] до Вищої ради правосуддя надійшла …, якій присвоєно реєстраційний номер [номер]»
    if re.search(r'присвоєно\s+(?:реєстраційний\s+)?номер\s*$', t[max(0, num_pos - 60): num_pos]):
        seg = t[max(0, num_pos - 800): num_pos]
        k = seg.rfind('надійшл')
        if k > 0:
            mm = list(re.finditer(DATE_W, seg[max(0, k - 60):k]))
            if mm:
                return fmt_date(*mm[-1].groups())
    before = t[max(0, num_pos - 160): num_pos]
    cands = [(mm.end(), fmt_date(*mm.groups())) for mm in re.finditer(DATE_W, before)]
    cands += [(mm.end(), date_num_to_words(*mm.groups())) for mm in re.finditer(DATE_N, before)]
    if cands:
        return sorted(cands)[-1][1]
    return None


def parse_conclusion(path, text):
    t = re.sub(r'[ \t]+', ' ', text)
    flat = re.sub(r'\s+', ' ', t)
    info = {'file': os.path.basename(path)}
    # інспектор
    m = re.search(r'ДИСЦИПЛІНАРНИЙ ІНСПЕКТОР\s+([А-ЯІЇЄҐ’\'-]+)\s+([А-ЯІЇЄҐ’\'-]+)\s+([А-ЯІЇЄҐ’\'-]+)', t)
    surname = m.group(1) if m else None
    if surname in INSPECTORS:
        info['inspector'], info['verb'] = INSPECTORS[surname]
    elif m:
        info['inspector'] = f'{m.group(1).capitalize()} {m.group(2)[0]}.{m.group(3)[0]}.'
        info['verb'] = 'склала' if m.group(3).upper().endswith('ІВНА') or m.group(3).upper().endswith('ЇВНА') else 'склав'
    else:
        info['inspector'], info['verb'] = None, None
    # дата висновку
    md = re.search(DATE_W, flat[:1500]) or re.search(DATE_N, flat[:1500])
    info['conclusion_date'] = (fmt_date(*md.groups()) if md and len(md.groups()) == 3 and not md.group(2).isdigit()
                               else (date_num_to_words(*md.groups()) if md else None))
    # номери скарг і дати
    nums = []
    seen = set()
    for mm in re.finditer(NUM_RE, flat):
        n = norm_num(mm.group(0))
        if n in seen:
            continue
        ctx = flat[max(0, mm.start() - 120): mm.end() + 20]
        # відкидаємо номери ухвал/рішень (/15-, /18- тощо), справ і постанов
        if re.search(r'/(15|16|18|17|8)-\d\d$', n) and not re.search(r'скарг', ctx):
            continue
        seen.add(n)
        nums.append({'number': n, 'date': find_dates_near(flat, mm.start(), len(mm.group(0)))})
    info['complaints'] = nums
    # ухвали інспектора (п. 5 ч. 1 ст. 44 тощо)
    info['inspector_rulings'] = [re.sub(r'\s+', ' ', x)[:400] for x in
                                 re.findall(r'[^.]{0,200}ухвал\w* дисциплінарного інспектора[^.]{0,300}\.', flat)]
    # перший абзац «надійшла»
    m = re.search(r'надійшл[аи]', flat)
    info['received_par'] = flat[max(0, m.start() - 250): m.start() + 650].strip() if m else None
    # пропозиція
    prop = None
    ps = [mm.end() for mm in re.finditer(r'пропоную\s*:?', flat, re.I)]
    if ps:
        tail = flat[ps[-1]:]
        tail = re.split(r'Дисциплінарний інспектор\s', tail)[0]
        prop = tail.strip(' |-+=')
    mp = re.search(r'Пропозиція\s*:?', t)
    if mp:
        seg = t[mp.end(): mp.end() + 4000]
        seg = re.sub(r'[|+]', ' ', seg)
        seg = re.sub(r'[-=]{3,}', ' ', seg)
        seg = re.sub(r'\s+', ' ', seg).strip()
        seg = re.split(r'\s(?:І|I)\.\s|\sВступ\s|\s1\.\s', seg)[0]
        if len(seg) > 20 and (not prop or len(prop) < 20):
            prop = seg
    if not prop:
        # Плескач: останній пункт «З наявних матеріалів можна зробити висновок…»
        mm = list(re.finditer(r'(З наявних матеріалів[^.]{0,600}\.)', flat))
        prop = mm[-1].group(1) if mm else None
    info['proposal'] = re.sub(r'\s+', ' ', prop).strip() if prop else None
    pl = (info['proposal'] or '').lower()
    kinds = []
    if re.search(r'відкрити дисциплінарну справу', pl): kinds.append('відкриття')
    if re.search(r'відмовити у відкритті', pl): kinds.append('відмова')
    if re.search(r'залишити без розгляду|залишення без розгляду', pl): kinds.append('залишення')
    if not kinds:
        if 'повернення скарги' in pl or 'ознак дисциплінарного проступку' in pl or 'судом вищої інстанції' in pl or 'фактичні дані' in pl:
            kinds.append('залишення')
    info['kinds'] = kinds
    # підстави: згадки пунктів ст. 44/45 у фінальній третині
    tail = flat[int(len(flat) * 0.55):] + ' ' + (info['proposal'] or '')
    grounds = []
    for mm in re.finditer(r'пункт\w*\s+((?:\d+(?:\s*,\s*|\s+(?:та|і)\s+)?)+)\s*частини першої статті (4[45])', tail):
        grounds.append(f"ст.{mm.group(2)} п.{re.sub(r'\s+', '', mm.group(1)).replace('та', ',').replace('і', ',')}")
    info['grounds_mentions'] = sorted(set(grounds))
    flags = []
    for key, pat in [('п2_немає_ознак', r'не містить відомостей про (?:наявність )?ознак'),
                     ('п3_немає_фактичних_даних', r'(?:не містить|відсутні|не міститься) (?:\S+ ){0,3}посилан\w* на фактичні дані'),
                     ('п6_суд_вищої_інстанції', r'перевірені виключно судом вищої інстанції'),
                     ('ст45п4_незгода', r'незгоди із судов'),
                     ('ст45п1_вже_розглядались', r'вже були предметом'),
                     ('ст45п2_строк', r'закінчився (?:встановлений законом )?строк')]:
        if re.search(pat, tail, re.I):
            flags.append(key)
    info['grounds_flags'] = flags
    return info


# ---------------------------------------------------------------- зіставлення
def match(agenda, concls):
    rep = {'section2': [], 'section3': [], 'unmatched_conclusions': [], 'date_mismatch': []}
    used = set()
    for sec in ('section2', 'section3'):
        for it in agenda[sec]:
            hits = [c for c in concls if set(it['numbers']) & {x['number'] for x in c['complaints']}]
            for c in hits:
                used.add(c['file'])
            row = {'n': it['n'], 'agenda_text': it['text'], 'agenda_inspector': it['inspector'],
                   'numbers': it['numbers'], 'conclusions': [c['file'] for c in hits]}
            if hits:
                cn = set().union(*[{x['number'] for x in c['complaints']} for c in hits])
                row['numbers_not_in_conclusion'] = [n for n in it['numbers'] if n not in cn]
                kinds = sorted(set(sum([c['kinds'] for c in hits], [])))
                row['kinds'] = kinds
            rep[sec].append(row)
    rep['unmatched_conclusions'] = [c['file'] for c in concls if c['file'] not in used]
    # запасне зіставлення за прізвищами суддів (номери в порядку денному і висновку розходяться)
    texts = {c['file']: open(c['text_path']).read().upper() if c.get('text_path') else '' for c in concls}
    for sec in ('section2', 'section3'):
        for row in rep[sec]:
            if row['conclusions']:
                continue
            caps = re.findall(r"\b([А-ЯІЇЄҐ’'-]{3,})\s+[А-ЯІЇЄҐ][а-яіїєґ’']+", row['agenda_text'])
            if not caps:
                continue
            def hit(c):
                tt = texts.get(c['file'], '')
                return all(w[:max(4, len(w) - 3)] in tt for w in caps[:3])
            insp = (row.get('agenda_inspector') or '').split(' ')[0]
            same = lambda c: not insp or (c.get('inspector') or '').startswith(insp)
            pool = [c for c in concls if c['file'] in rep['unmatched_conclusions'] and same(c)]
            cands = [c['file'] for c in pool if hit(c)]
            if not cands:
                cands = [c['file'] for c in concls if hit(c) and same(c)]
            if cands:
                row['possible'] = cands
    return rep


def nums_or_text(r):
    return ', '.join(r['numbers']) if r['numbers'] else r['agenda_text'][:120]


def report_md(agenda, concls, rep):
    L = [f"# Звіт зіставлення — засідання {agenda['date']} (порядок денний № {agenda['agenda_no']})", '']
    L.append('## Розділ 2 (відмова у відкритті)')
    for r in rep['section2']:
        if not r['conclusions']:
            if r.get('possible'):
                L.append(f"- **п. {r['n']} ({nums_or_text(r)}; {r['agenda_inspector']}) — за номером висновку не знайдено; "
                         f"можливо, це {', '.join(r['possible'])} (номери розходяться — перевірити).**")
            else:
                L.append(f"- п. {r['n']} ({nums_or_text(r)}; {r['agenda_inspector']}) — висновку немає → "
                         f"ймовірно, пропозиція відкрити дисциплінарну справу; в ухвалі не відображається.")
        else:
            extra = ''
            if r.get('numbers_not_in_conclusion'):
                extra += f"; номерів з порядку денного немає у висновку: {', '.join(r['numbers_not_in_conclusion'])}"
            if 'відкриття' in r.get('kinds', []) or 'залишення' in r.get('kinds', []):
                extra += f"; вид пропозиції: {', '.join(r['kinds'])} — перевірити, чи це змішаний висновок"
            L.append(f"- п. {r['n']} → {', '.join(r['conclusions'])}{extra}")
    L += ['', '## Розділ 3 (залишення без розгляду та повернення)']
    for r in rep['section3']:
        if not r['conclusions']:
            if r.get('possible'):
                L.append(f"- **п. {r['n']} ({nums_or_text(r)}; {r['agenda_inspector']}) — за номером висновку не знайдено; "
                         f"можливо, це {', '.join(r['possible'])} (номери в порядку денному і висновку розходяться — перевірити).**")
            else:
                L.append(f"- **п. {r['n']} ({nums_or_text(r)}; {r['agenda_inspector']}) — ВИСНОВКУ НЕ НАДІСЛАНО.**")
        else:
            extra = ''
            if r.get('numbers_not_in_conclusion'):
                extra += f"; номерів з порядку денного немає у висновку: {', '.join(r['numbers_not_in_conclusion'])}"
            if r.get('kinds') and r['kinds'] != ['залишення']:
                extra += f"; вид пропозиції: {', '.join(r['kinds'])}"
            L.append(f"- п. {r['n']} → {', '.join(r['conclusions'])}{extra}")
    L += ['', '## Висновки, яких немає в порядку денному']
    L += [f'- {f}' for f in rep['unmatched_conclusions']] or ['- немає']
    return '\n'.join(L)


def cmd_extract(a):
    work = tempfile.mkdtemp()
    agenda = parse_agenda(to_text(a.agenda, work))
    files = collect_inputs(a.inputs, work)
    concls, failed, seen_hash = [], [], set()
    for f in files:
        if os.path.abspath(f) == os.path.abspath(a.agenda):
            continue
        t = to_text(f, work)
        if not t:
            failed.append(os.path.basename(f)); continue
        if 'ВИСНОВОК' not in t[:3000] and 'висновок' not in t[:3000].lower():
            continue
        h = hash(t)
        if h in seen_hash:
            continue
        seen_hash.add(h)
        concls.append(parse_conclusion(f, t))
        concls[-1]['text_path'] = os.path.join(work, f'c{len(concls):03d}.txt')
        open(concls[-1]['text_path'], 'w').write(t)
    rep = match(agenda, concls)
    data = {'agenda': agenda, 'conclusions': concls, 'match': rep, 'failed': failed}
    json.dump(data, open(a.out, 'w'), ensure_ascii=False, indent=1)
    md = report_md(agenda, concls, rep)
    if failed:
        md += '\n\n## Не вдалося прочитати\n' + '\n'.join(f'- {x}' for x in failed)
    if a.report:
        open(a.report, 'w').write(md)
    print(md)
    print(f'\nТексти висновків: {work}')


# ---------------------------------------------------------------- збирання DOCX
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'

TITLES = {'zalyshennia': ['Про залишення без розгляду та повернення дисциплінарних скарг'],
          'vidmova': ['Про відмову у відкритті ', 'дисциплінарних справ']}
DEFAULT_COMPOSITION = {
    'chair': {'gen': 'Сасевича О.М.', 'sign': 'Олександр САСЕВИЧ'},
    'members': [{'gen': 'Кандзюби О.В.', 'sign': 'Олег КАНДЗЮБА'},
                {'gen': 'Круподері Д.О.', 'sign': 'Дмитро КРУПОДЕРЯ'},
                {'gen': 'Попікової О.В.', 'sign': 'Ольга ПОПІКОВА'}]}


def boilerplate(kind):
    if kind == 'zalyshennia':
        dijshla = ('Третя Дисциплінарна палата Вищої ради правосуддя за результатами розгляду дисциплінарних скарг, '
                   'матеріалів їх попередньої перевірки та висновків дисциплінарних інспекторів Вищої ради правосуддя '
                   'дійшла висновку про залишення без розгляду та повернення зазначених дисциплінарних скарг.')
        keruy = ('Керуючись статтею 44 Закону України «Про Вищу раду правосуддя», пунктом 13.13 Регламенту '
                 'Вищої ради правосуддя, Третя Дисциплінарна палата Вищої ради правосуддя')
    else:
        dijshla = ('Третя Дисциплінарна палата Вищої ради правосуддя, за результатами розгляду дисциплінарних скарг, '
                   'матеріалів їх попередньої перевірки та висновків дисциплінарних інспекторів Вищої ради правосуддя, '
                   'дійшла висновку про відмову у відкритті дисциплінарних справ стосовно зазначених у скаргах суддів.')
        keruy = ('Керуючись статтею 45 Закону України «Про Вищу раду правосуддя», пунктом 13.13 Регламенту '
                 'Вищої ради правосуддя, Третя Дисциплінарна палата Вищої ради правосуддя')
    return dijshla, keruy


def composition_par(comp, single):
    members = ', '.join(m['gen'] for m in comp['members'])
    tail = ('розглянувши висновок дисциплінарного інспектора Вищої ради правосуддя – доповідача за результатами '
            'попередньої перевірки дисциплінарних скарг,') if single else (
        'розглянувши висновки дисциплінарних інспекторів Вищої ради правосуддя – доповідачів за результатами '
        'попередньої перевірки дисциплінарних скарг,')
    return (f"Третя Дисциплінарна палата Вищої ради правосуддя у складі головуючого – {comp['chair']['gen']}, "
            f"членів Третьої Дисциплінарної палати Вищої ради правосуддя {members}, {tail}")


def set_text(p, s):
    runs = p.findall(W + 'r')
    first = next((r for r in runs if r.find(W + 't') is not None), runs[0] if runs else None)
    if first is None:
        first = p.makeelement(W + 'r', {}); p.append(first)
    for r in runs:
        if r is not first:
            p.remove(r)
    for c in list(first):
        if c.tag != W + 'rPr':
            first.remove(c)
    t = first.makeelement(W + 't', {}); t.text = s; t.set(XMLSPACE, 'preserve'); first.append(t)


def load_template(path):
    import docx
    if path.endswith('.b64'):
        raw = base64.b64decode(open(path).read())
    else:
        raw = open(path, 'rb').read()
    return docx.Document(io.BytesIO(raw))


def cmd_build(a):
    import docx
    data = json.load(open(a.data))
    kind = data['kind']
    single = bool(data.get('single'))
    comp = data.get('composition') or DEFAULT_COMPOSITION
    d = load_template(a.template)
    body = d.element.body
    ph = {}
    for p in body.iter(W + 'p'):
        t = ''.join(x.text or '' for x in p.iter(W + 't'))
        m = re.fullmatch(r'\{\{(\w+)\}\}', t.strip())
        if m:
            ph[m.group(1)] = p
    # шапка
    set_text(ph['DATE'], '  ' + data['date'])
    set_text(ph['NUMBER'], data.get('number') or '№ ____/3дп/15-26')
    title_p = ph['TITLE']
    lines = TITLES[kind]
    set_text(title_p, lines[0])
    prev = title_p
    for ln in lines[1:]:
        np_ = copy.deepcopy(title_p); set_text(np_, ln); prev.addnext(np_); prev = np_
    set_text(ph['COMP'], data.get('intro') or composition_par(comp, single))
    # встановила
    proto = ph['ITEM']
    prev = proto
    first_item = True
    for par in data['established']:
        # порожній рядок між пунктами «встановила» (як в ухвалах палати)
        if re.match(r'\d+\)\s', par):
            if not first_item:
                gap = copy.deepcopy(proto); set_text(gap, ''); prev.addnext(gap); prev = gap
            first_item = False
        np_ = copy.deepcopy(proto); set_text(np_, par); prev.addnext(np_); prev = np_
    body.remove(proto) if proto.getparent() is not None else None
    dijshla, keruy = boilerplate(kind)
    if single and not data.get('conclusion_par'):
        sys.exit('Для окремої ухвали (single) потрібне поле conclusion_par')
    # додаткові абзаци між встановила і дійшла (напр. норма закону для окремої ухвали)
    prev = ph['DIJSHLA']
    set_text(ph['DIJSHLA'], data.get('conclusion_par') or dijshla)
    for par in data.get('before_conclusion', [])[::-1]:
        np_ = copy.deepcopy(ph['DIJSHLA']); set_text(np_, par); ph['DIJSHLA'].addprevious(np_)
    set_text(ph['KERUY'], keruy)
    set_text(ph['UKHV'], 'ухвалила:')
    for where in ('addprevious', 'addnext'):
        gap = copy.deepcopy(ph['KERUY']); set_text(gap, ''); getattr(ph['UKHV'], where)(gap)
    set_text(ph['VST'], 'встановила:')
    # ухвалила
    proto = ph['NUM']
    prev = proto
    for par in data['operative']:
        np_ = copy.deepcopy(proto)
        if single or len(data['operative']) == 1:
            ppr = np_.find(W + 'pPr')
            if ppr is not None and ppr.find(W + 'numPr') is not None:
                ppr.remove(ppr.find(W + 'numPr'))
        set_text(np_, par); prev.addnext(np_); prev = np_
    body.remove(proto)
    set_text(ph['FINAL'], 'Ухвала оскарженню не підлягає.')
    for _ in range(2):
        gap = copy.deepcopy(ph['FINAL']); set_text(gap, ''); ph['FINAL'].addnext(gap)
    # підписи
    tbl = d.tables[-1]
    tblpr = tbl._tbl.find(W + 'tblPr')
    if tblpr is not None and tblpr.find(W + 'tblpPr') is not None:
        tblpr.remove(tblpr.find(W + 'tblpPr'))
    rows = tbl.rows
    chair_cell = rows[0].cells[1]
    set_text(chair_cell.paragraphs[-1]._p, comp['chair']['sign'])
    proto_label = copy.deepcopy(rows[2]._tr)
    proto_blank = copy.deepcopy(rows[3]._tr)
    for r in list(rows)[2:]:
        tbl._tbl.remove(r._tr)
    def add_row(left, right):
        tr = copy.deepcopy(proto_label)
        cells = tr.findall(W + 'tc')
        for c, s in zip(cells, (left, right)):
            ps = c.findall(W + 'p')
            for extra in ps[1:]:
                c.remove(extra)
            set_text(ps[0], s)
        tbl._tbl.append(tr)
    labels = ['Члени Третьої Дисциплінарної', 'палати Вищої ради правосуддя']
    seq = []
    for i, m in enumerate(comp['members']):
        if i == 0:
            seq.append((labels[0], m['sign'])); seq.append((labels[1], ''))
        else:
            seq.append(('', m['sign']))
            if i < len(comp['members']) - 1:
                seq.append(('', ''))
    for l, r in seq:
        add_row(l, r)
    d.core_properties.author = "В'ячеслав Плескач"
    d.core_properties.last_modified_by = "В'ячеслав Плескач"
    d.core_properties.title = TITLES[kind][0].strip() + (' ' + TITLES[kind][1] if len(TITLES[kind]) > 1 else '')
    d.save(a.out)
    print('Збережено:', a.out)


def main():
    ap = argparse.ArgumentParser()
    sp = ap.add_subparsers(dest='cmd', required=True)
    e = sp.add_parser('extract')
    e.add_argument('--agenda', required=True)
    e.add_argument('--in', dest='inputs', nargs='+', required=True)
    e.add_argument('--out', required=True)
    e.add_argument('--report')
    b = sp.add_parser('build')
    b.add_argument('--data', required=True)
    b.add_argument('--out', required=True)
    here = os.path.dirname(os.path.abspath(__file__))
    tpl = next((os.path.join(here, n) for n in ('Шаблон_ухвали_3ДП.docx', 'Шаблон_ухвали_3ДП_docx.b64')
                if os.path.exists(os.path.join(here, n))), os.path.join(here, 'Шаблон_ухвали_3ДП.docx'))
    b.add_argument('--template', default=tpl)
    a = ap.parse_args()
    {'extract': cmd_extract, 'build': cmd_build}[a.cmd](a)


if __name__ == '__main__':
    main()
