# -*- coding: utf-8 -*-
"""Інструменти для перенесення олівцевих правок літературного редактора в DOCX
у режимі рецензування (tracked changes). Версія 3 (01.11.2026).

Робоча тека — поточна (cwd). Очікувані файли: orig.docx, scan.pdf, edits.py.

  python3 pravky_tools.py scan              # scan.pdf → pg/ (200 dpi), half/ (напівсторінки 1550 px), grid/ (огляд 8 стор.)
  python3 pravky_tools.py zoom P y0 y1 [x0 x1] [rot]   # фрагмент сторінки P → z.png (частки 0..1; rot = 90/-90/180)
  python3 pravky_tools.py index             # orig.docx → u/ + paras.txt («[індекс] текст», {курсив})
  python3 pravky_tools.py build             # правки з edits.py → unpacked/word/document.xml
  python3 pravky_tools.py comments          # коментарі з edits.COMMENTS
  python3 pravky_tools.py pack out.docx
  python3 pravky_tools.py check out.docx    # відхилити все = оригінал, лічильники, змінені абзаци
  python3 pravky_tools.py review A B        # текст «прийняти все» змінених абзаців A..B-1 (NBSP показано «°»)
  python3 pravky_tools.py all out.docx      # build + comments + pack + validate + check

Великий документ (800+ абзаців, сотні правок) збирається 2–4 хв — запускати у фоні.
"""
import html, os, re, shutil, subprocess, sys, zipfile

AUTHOR = "Літературний редактор"
INITIALS = "ЛР"
DATE = "2026-11-01T00:00:00Z"
NB = ' '
COMMENT_SCRIPT = '/mnt/skills/public/docx/scripts/comment.py'
VALIDATE_SCRIPT = '/mnt/skills/public/docx/scripts/office/validate.py'

RUN_RE = re.compile(r'<w:r(?:\s[^>]*)?>(?:(?!</?w:r[ >]).)*?</w:r>', re.S)
T_RE = re.compile(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', re.S)
RPR_RE = re.compile(r'<w:rPr>.*?</w:rPr>', re.S)
PARA_RE = re.compile(r'<w:p[ >].*?</w:p>', re.S)
TRACKED_RE = re.compile(r'<w:(?:ins|del)(?:\s[^>]*)?>.*?</w:(?:ins|del)>', re.S)
SPLIT = re.compile(r'<w:r((?:\s[^>]*)?)>(<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?((?:<w:(?:br|tab|noBreakHyphen)/>)+)(<w:t(?:\s[^>]*)?>[^<]*</w:t>)</w:r>', re.S)
COMPLEX = re.compile(r'<w:(?:noBreakHyphen|softHyphen|tab|br|cr|sym|fldChar|instrText|footnoteReference|endnoteReference|drawing|object)\b')


# ───────────────────────── підготовка XML ─────────────────────────
def safe_merge(xml):
    """Зливає суміжні runs з одним <w:t> і однаковим <w:rPr>, лише якщо між ними нічого немає."""
    def render(it):
        return (f'<w:r>{it["rpr"]}<w:t xml:space="preserve">{it["text"]}</w:t></w:r>'
                if it['only_t'] else it['raw'])
    out, last, pos, merged = [], None, 0, 0
    for m in RUN_RE.finditer(xml):
        gap, pos, run = xml[pos:m.start()], m.end(), m.group(0)
        tm, rm = T_RE.search(run), RPR_RE.search(run)
        rpr = rm.group(0) if rm else ''
        only_t = tm is not None and run.count('<w:t') == 1 and RPR_RE.sub('', run).count('<') == 4
        if last is not None and gap == '' and only_t and last['only_t'] and last['rpr'] == rpr:
            last['text'] += tm.group(1); merged += 1; continue
        if last is not None: out.append(render(last))
        out.append(gap)
        last = {'rpr': rpr, 'text': tm.group(1) if tm else '', 'raw': run, 'only_t': only_t}
    if last is not None: out.append(render(last))
    out.append(xml[pos:])
    return ''.join(out), merged


def prepare(xml):
    xml = xml.replace('<w:lastRenderedPageBreak/>', '')
    xml = re.sub(r'</w:t><w:t(?:\s[^>]*)?>', '', xml)
    before = [''.join(T_RE.findall(p)) for p in PARA_RE.findall(xml)]
    xml, nm = safe_merge(xml)
    xml = SPLIT.sub(lambda m: f'<w:r{m[1]}>{m[2] or ""}{m[3]}</w:r><w:r{m[1]}>{m[2] or ""}{m[4]}</w:r>', xml)
    assert before == [''.join(T_RE.findall(p)) for p in PARA_RE.findall(xml)], 'підготовка змінила текст'
    return xml, nm


# ───────────────────────── tracked-заміна ─────────────────────────
class Doc:
    def __init__(self, xml):
        self.xml, self._id = xml, 9000

    def nid(self):
        self._id += 1
        return self._id

    def paras(self):
        return list(PARA_RE.finditer(self.xml))

    @staticmethod
    def _plain(rpr, text):
        return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'

    def _mk(self, rpr, old, new):
        out = []
        if old:
            out.append(f'<w:del w:id="{self.nid()}" w:author="{AUTHOR}" w:date="{DATE}">'
                       f'<w:r>{rpr}<w:delText xml:space="preserve">{old}</w:delText></w:r></w:del>')
        if new:
            out.append(f'<w:ins w:id="{self.nid()}" w:author="{AUTHOR}" w:date="{DATE}">'
                       f'<w:r>{rpr}<w:t xml:space="preserve">{new}</w:t></w:r></w:ins>')
        return ''.join(out)

    def replace(self, old, new, expect=None, para=None, skip_italic=True):
        """Замінює old → new (del+ins) у вказаному абзаці. Runs у <w:ins>/<w:del> пропускаються;
        runs із полями, виносками, br/tab/noBreakHyphen — бар'єри; курсивні runs — лише з skip_italic=False."""
        applied = 0
        while True:
            paras = self.paras()
            rng = range(len(paras)) if para is None else ([para] if isinstance(para, int) else para)
            found = None
            for pi in rng:
                pm = paras[pi]; ptxt = pm.group(0)
                marked = [(mm.start(), mm.end()) for mm in TRACKED_RE.finditer(ptxt)]
                runs = []
                for rm in RUN_RE.finditer(ptxt):
                    if any(a <= rm.start() < b for a, b in marked): continue
                    r = rm.group(0); tm = T_RE.search(r)
                    cx = bool(COMPLEX.search(r)) or len(T_RE.findall(r)) > 1
                    runs.append([rm.start(), rm.end(), r, tm.group(1) if tm else '', '<w:i/>' in r, tm is not None and not cx])
                if not runs: continue
                concat = ''.join(r[3] for r in runs)
                cmap = [(ri, off) for ri, r in enumerate(runs) for off in range(len(r[3]))]
                idx = concat.find(old)
                while idx >= 0:
                    span = cmap[idx: idx + len(old)]
                    inv = sorted({c[0] for c in span})
                    if (skip_italic and any(runs[i][4] for i in inv)) or any(not runs[i][5] for i in range(inv[0], inv[-1] + 1)):
                        idx = concat.find(old, idx + 1); continue
                    found = (pm.start(), runs, span, inv); break
                if found: break
            if not found: break
            pstart, runs, span, inv = found
            first, last = inv[0], inv[-1]
            rf = RPR_RE.search(runs[first][2]); rf = rf.group(0) if rf else ''
            rl = RPR_RE.search(runs[last][2]); rl = rl.group(0) if rl else ''
            pre, post = runs[first][3][:span[0][1]], runs[last][3][span[-1][1] + 1:]
            parts = ([self._plain(rf, pre)] if pre else []) + [self._mk(rf, old, new)] + ([self._plain(rl, post)] if post else [])
            self.xml = self.xml[:pstart + runs[first][0]] + ''.join(parts) + self.xml[pstart + runs[last][1]:]
            applied += 1
            if expect is not None and applied >= expect: break
        if expect is not None and applied != expect:
            raise AssertionError(f"{old!r}: застосовано {applied}, очікувалось {expect}")
        if applied == 0:
            raise AssertionError(f"{old!r}: НЕ ЗНАЙДЕНО")
        return applied


# ───────────────────────── допоміжне ─────────────────────────
def _cwd_import_edits():
    sys.path.insert(0, os.getcwd())
    import importlib, edits
    importlib.reload(edits)
    return edits


def cmd_scan():
    from PIL import Image
    import glob
    for dname in ('pg', 'half', 'grid'): os.makedirs(dname, exist_ok=True)
    subprocess.run(['pdftoppm', '-r', '200', '-png', 'scan.pdf', 'pg/p'], check=True)
    fs = sorted(glob.glob('pg/p-*.png'))
    for f in fs:
        n = re.search(r'-(\d+)\.png$', f).group(1).zfill(2)
        if not f.endswith(f'p-{n}.png'): os.rename(f, f'pg/p-{n}.png'); f = f'pg/p-{n}.png'
        im = Image.open(f); W, H = im.size
        for k, (a, b) in enumerate([(0, 0.52), (0.48, 1)]):
            c = im.crop((0, int(H * a), W, int(H * b))); s = 1550 / W
            c.resize((1550, int(c.height * s))).convert('L').save(f'half/{n}{"ab"[k]}.png')
    fs = sorted(glob.glob('pg/p-*.png'))
    for g in range(0, len(fs), 8):
        ims = [Image.open(f).convert('L').resize((413, 584)) for f in fs[g:g + 8]]
        G = Image.new('L', (413 * 4, 584 * 2), 255)
        for i, im in enumerate(ims): G.paste(im, ((i % 4) * 413, (i // 4) * 584))
        G.save(f'grid/g{g // 8:02d}.png')
    print('сторінок:', len(fs))


def cmd_zoom(pg, y0, y1, x0='0', x1='1', rot='0'):
    from PIL import Image
    im = Image.open(f'pg/p-{int(pg):02d}.png'); W, H = im.size
    c = im.crop((int(W * float(x0)), int(H * float(y0)), int(W * float(x1)), int(H * float(y1))))
    if int(rot): c = c.rotate(int(rot), expand=True)
    s = 1550 / c.width
    c.resize((int(c.width * s), int(c.height * s))).save('z.png')
    print('z.png')


def cmd_index():
    shutil.rmtree('u', ignore_errors=True)
    subprocess.run(['unzip', '-q', 'orig.docx', '-d', 'u/'], check=True)
    x = open('u/word/document.xml', encoding='utf-8').read()
    ps = PARA_RE.findall(x)
    with open('paras.txt', 'w', encoding='utf-8') as out:
        for i, p in enumerate(ps):
            t = ''
            for rm in RUN_RE.finditer(p):
                r = rm.group(0); tt = ''.join(T_RE.findall(r))
                if '<w:noBreakHyphen/>' in r: tt = '‑' + tt
                t += ('{' + tt + '}') if '<w:i/>' in r and tt else tt
            out.write(f'[{i}] {t}\n')
    print('абзаців:', len(ps), '→ paras.txt')


# ───────────────────────── збирання ─────────────────────────
def cmd_build():
    ed = _cwd_import_edits()
    shutil.rmtree('unpacked', ignore_errors=True)
    subprocess.run(['unzip', '-q', 'orig.docx', '-d', 'unpacked/'], check=True)
    P = 'unpacked/word/document.xml'
    xml, nm = prepare(open(P, encoding='utf-8').read())
    d = Doc(xml)
    st = {'n': 0, 'fails': []}

    def ptext(pi):   # текст без tracked-фрагментів (ні ins, ні del)
        return ''.join(T_RE.findall(TRACKED_RE.sub('', d.paras()[pi].group(0))))

    def atext(pi):   # текст «прийняти все»
        x = re.sub(r'<w:del(?:\s[^>]*)?>.*?</w:del>', '', d.paras()[pi].group(0), flags=re.S)
        return ''.join(T_RE.findall(x))

    def adapt(pi, old):  # дозволяє NBSP на місці пробілу в old
        b = ptext(pi)
        if old in b: return old
        m = re.search('[  ]+'.join(re.escape(p) for p in old.split(' ')), b)
        return m.group(0) if m else None

    def rep(pi, old, new, exp=1, it=False):
        a = adapt(pi, old)
        if a is None:
            st['fails'].append(f'[{pi}] НЕ ЗНАЙДЕНО {old[:60]!r}'); return
        try:
            st['n'] += d.replace(a, new, expect=exp, para=pi, skip_italic=not it)
        except AssertionError as e:
            st['fails'].append(f'[{pi}] {e}')

    def gen(paras, old, new, it=True):
        """Системна заміна в кожному з абзаців paras (усі входження)."""
        for pi in paras:
            for _ in range(60):
                a = adapt(pi, old)
                if a is None: break
                try:
                    st['n'] += d.replace(a, new, expect=1, para=pi, skip_italic=not it)
                except AssertionError:
                    st['fails'].append(f'[{pi}] системна заміна не застосована {old!r}'); break

    def tail(pi, fn):
        """Зміна кінця абзацу: fn(хвіст) → новий хвіст (напр. додати «;»)."""
        t = ptext(pi).rstrip(); k = 6
        while t.count(t[-k:]) > 1 and k < len(t): k += 2
        old = t[-k:]; new = fn(old)
        if new != old:
            try: st['n'] += d.replace(old, new, expect=1, para=pi, skip_italic=False)
            except AssertionError as e: st['fails'].append(f'[{pi}] кінцівка {e}')

    # 1) точкові правки
    for item in ed.E:
        pi, old, new, exp = item[:4]
        rep(pi, old, new, exp, item[4] if len(item) > 4 else False)
    # 2) системні заміни: (абзаци, old, new[, курсив=True])
    for g in getattr(ed, 'GENERIC', []):
        gen(g[0], g[1], g[2], g[3] if len(g) > 3 else True)
    # 3) вкладені лапки «…» → “…” в абзацах-фрагментах цитат (зовнішніх лапок у них немає)
    for pi in getattr(ed, 'NESTED', []):
        for q, r in (('«', '“'), ('»', '”')):
            for _ in range(30):
                if q not in ptext(pi): break
                try: st['n'] += d.replace(q, r, expect=1, para=pi, skip_italic=False)
                except AssertionError: st['fails'].append(f'[{pi}] лапки {q}'); break
    # 4) абзаци з крапками («…..», «<...>») → «<…>»
    for pi in getattr(ed, 'ELLIPSIS', []):
        t = ptext(pi).strip()
        try: st['n'] += d.replace(t, '&lt;…&gt;', expect=1, para=pi, skip_italic=False)
        except AssertionError as e: st['fails'].append(f'[{pi}] <…> {e}')
    # 5) довільний гачок для специфічних системних правил документа
    if hasattr(ed, 'HOOK'):
        ed.HOOK(dict(d=d, rep=rep, gen=gen, tail=tail, ptext=ptext, atext=atext, st=st, re=re))

    # ── технічні зміни без рецензування ──
    def set_chars(pi, fn):
        spans = list(PARA_RE.finditer(d.xml)); m = spans[pi]; seg = m.group(0)
        cells = []
        for tm in re.finditer(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', seg):
            cells += [tm.start(1) + j for j in range(len(tm.group(1)))]
        ch = fn(''.join(seg[c] for c in cells))
        if not ch: return 0
        lst = list(seg)
        for off, c in ch.items(): lst[cells[off]] = c
        d.xml = d.xml[:m.start()] + ''.join(lst) + d.xml[m.end():]
        return len(ch)

    auto = 0
    for pi, frag in getattr(ed, 'NBSP', []):
        def f(c, frag=frag, pi=pi):
            parts = frag.split('~')
            pat = '([  ])'.join(re.escape(p).replace('\\ ', '[  ]') for p in parts)
            out, found = {}, False
            for mm in re.finditer(pat, c):
                found = True
                for g in range(1, len(parts)):
                    if c[mm.start(g)] == ' ': out[mm.start(g)] = NB
            if not found: st['fails'].append(f'[{pi}] NBSP не знайдено {frag!r}')
            return out
        auto += set_chars(pi, f)

    A, B = ed.RANGE   # межі частини (абзаци A..B включно) для автоматичних правил
    RULES = [r'(?<![А-Яа-яІЇЄҐіїєґ])а\. (?=с\.)', r'(?<![А-Яа-яІЇЄҐіїєґ])а\.[  ]с\. (?=\d)',
             r'(?<![А-Яа-яІЇЄҐіїєґ])т\. (?=\d)']
    for pi in range(A, B + 1):
        def f(c):
            out = {}
            for r in RULES:
                for mm in re.finditer(r, c): out[mm.end() - 1] = NB
            for mm in re.finditer(r'(?<=[А-Яа-яІЇЄҐіїєґ])["\'`ʼ](?=[а-яіїєґ])', c): out[mm.start()] = '’'
            return out
        auto += set_chars(pi, f)
    cnt = [0]
    def dbl(m):
        t = m.group(2); new = re.sub(r'(?<=\S)  +(?=\S)', ' ', t)
        if new != t: cnt[0] += 1
        return m.group(1) + new + m.group(3)
    spans = list(PARA_RE.finditer(d.xml))
    a, b = spans[A].start(), spans[B].end()
    d.xml = d.xml[:a] + re.sub(r'(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)', dbl, d.xml[a:b]) + d.xml[b:]
    open(P, 'w', encoding='utf-8').write(d.xml)
    print('злито runs', nm, '| правок', st['n'], '| NBSP/апостроф', auto, '| подвійні пробіли', cnt[0], '| невдачі', len(st['fails']))
    for f_ in st['fails']: print('  ', f_)
    return len(st['fails'])


def cmd_comments():
    ed = _cwd_import_edits()
    P = 'unpacked/word/document.xml'
    PPR = re.compile(r'^<w:p[^>]*>\s*<w:pPr>.*?</w:pPr>', re.S)
    CR = '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{i}"/></w:r>'
    s = open(P, encoding='utf-8').read()
    cid = 0
    for pi, text in sorted(ed.COMMENTS, key=lambda x: -x[0]):   # з кінця документа
        spans = list(PARA_RE.finditer(s)); p = spans[pi].group(0)
        subprocess.run(['python3', COMMENT_SCRIPT, 'unpacked/', text, '--id', str(cid), '--author', AUTHOR, '--initials', INITIALS], check=True, capture_output=True)
        m = PPR.search(p); pos = m.end() if m else p.index('>') + 1
        new = p[:pos] + f'<w:commentRangeStart w:id="{cid}"/>' + p[pos:-len('</w:p>')] + f'<w:commentRangeEnd w:id="{cid}"/>' + CR.format(i=cid) + '</w:p>'
        s = s[:spans[pi].start()] + new + s[spans[pi].end():]
        cid += 1
    open(P, 'w', encoding='utf-8').write(s)
    print('коментарів:', cid)


def cmd_pack(out):
    src = 'unpacked'
    names = sorted((os.path.relpath(os.path.join(r, f), src) for r, _, fs in os.walk(src) for f in fs),
                   key=lambda n: (n != '[Content_Types].xml', n))
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names: z.write(os.path.join(src, n), n)
    print('→', out)


def cmd_check(out):
    orig = zipfile.ZipFile('orig.docx').read('word/document.xml').decode()
    new = zipfile.ZipFile(out).read('word/document.xml').decode()
    INS = re.compile(r'<w:ins\s[^>]*>(.*?)</w:ins>', re.S); DEL = re.compile(r'<w:del\s[^>]*>(.*?)</w:del>', re.S)
    reject = DEL.sub(lambda m: m[1], INS.sub('', new.replace('<w:delText', '<w:t').replace('</w:delText>', '</w:t>')))
    accept = INS.sub(lambda m: m[1], DEL.sub('', new))
    TOK = re.compile(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>')
    norm = lambda s: re.sub(r'\s+', '', html.unescape(s).replace('`', '’').replace("'", '’').replace('"', '’').replace('ʼ', '’'))
    po, pr, pa = [PARA_RE.findall(x) for x in (orig, reject, accept)]
    bad = [i for i in range(len(po)) if norm(''.join(TOK.findall(po[i]))) != norm(''.join(TOK.findall(pr[i])))]
    print('абзаців (оригінал / відхилити / прийняти):', len(po), len(pr), len(pa))
    print('відхилити все ≠ оригінал:', bad or 'немає')
    for tag in ['<w:br/>', '<w:noBreakHyphen/>', '<w:tab/>', '<w:fldChar ', '<w:footnoteReference ']:
        c = (orig.count(tag), reject.count(tag), accept.count(tag))
        print(f'{tag:24} {c}', '' if len(set(c)) == 1 else '  ← РОЗБІЖНІСТЬ')
    ch = [i for i in range(len(po)) if ''.join(TOK.findall(po[i])) != ''.join(TOK.findall(pa[i]))]
    print('змінено абзаців:', len(ch), 'від', ch[:1], 'до', ch[-1:])
    return 0 if not bad and len(po) == len(pr) == len(pa) else 1


def cmd_review(a, b):
    s = open('unpacked/word/document.xml', encoding='utf-8').read()
    acc = re.sub(r'<w:del\s[^>]*>.*?</w:del>', '', s, flags=re.S)
    TOK = re.compile(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>')
    raw, ps = PARA_RE.findall(s), PARA_RE.findall(acc)
    for i in range(int(a), min(int(b), len(raw))):
        if '<w:ins ' in raw[i] or '<w:del ' in raw[i]:
            print(f'[{i}]', html.unescape(''.join(TOK.findall(ps[i]))).replace(NB, '°'))


def cmd_all(out):
    if cmd_build(): sys.exit('є невдачі — виправте edits.py')
    cmd_comments(); cmd_pack(out)
    subprocess.run(['python3', VALIDATE_SCRIPT, out], check=True)
    sys.exit(cmd_check(out))


if __name__ == '__main__':
    c, args = sys.argv[1], sys.argv[2:]
    {'scan': cmd_scan, 'zoom': cmd_zoom, 'index': cmd_index, 'build': cmd_build, 'comments': cmd_comments,
     'pack': cmd_pack, 'check': cmd_check, 'review': cmd_review, 'all': cmd_all}[c](*args)
