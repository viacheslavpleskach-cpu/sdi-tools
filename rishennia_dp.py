# -*- coding: utf-8 -*-
"""Проєкт рішення Дисциплінарної палати ВРП на основі висновку дисциплінарного інспектора
за результатами підготовки дисциплінарної справи до розгляду. Версія 1.0 (16.11.2026).
Автор: В'ячеслав Плескач. Інструкція — «Інструкція_рішення_ДП.md» у проєкті «СДІ».

  python3 rishennia_dp.py extract --in ВИСНОВОК.docx --out vysnovok.md
      текст висновку з індексами абзаців [pN] (після технічних виправлень) і підказками
  python3 rishennia_dp.py build --in ВИСНОВОК.docx --data data.json --out РІШЕННЯ.docx
      збирає проєкт рішення: шапка і кінець — напряму, зміни тексту — в режимі правок
  python3 rishennia_dp.py check --in РІШЕННЯ.docx [--orig ВИСНОВОК.docx]
      лічильники правок і текст «прийняти все» змінених абзаців
  python3 rishennia_dp.py template --from РІШЕННЯ_ПАЛАТИ.docx --out Шаблон_рішення_ДП.docx
      перезбирає шаблон шапки й кінця з готового рішення палати

Формат data.json — розділ 10 інструкції. Коротко:
  {"kind": "prytiahnennia" | "vidmova", "date": "…", "title": "…", "intro": "…",
   "operative": ["…", "…"], "composition": {"chair": "…", "members": ["…"]},
   "keruiuchys": "…" (необов'язково), "edits": [{"p": 12, "old": "…", "new": "…"},
   {"p": 40, "delete": true}, {"after": 55, "text": "…", "like": 54}]}
"""
import argparse, copy, datetime, difflib, io, json, os, re, sys, zipfile
from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
W = '{%s}' % W_NS
AUTHOR = "В'ячеслав Плескач"
DATE = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(HERE, 'Шаблон_рішення_ДП.docx')

HEADING_III = 'ІІІ. ВИСНОВКИ ТРЕТЬОЇ ДИСЦИПЛІНАРНОЇ ПАЛАТИ ВИЩОЇ РАДИ ПРАВОСУДДЯ'
PALATA = 'Третя Дисциплінарна палата Вищої ради правосуддя'
DEFAULT_COMPOSITION = {'chair': 'Олександр САСЕВИЧ',
                       'members': ['Олег КАНДЗЮБА', 'Дмитро КРУПОДЕРЯ', 'Ольга ПОПІКОВА']}
APPEAL = ('Рішення Третьої Дисциплінарної палати Вищої ради правосуддя може бути оскаржене '
          'в порядку і в строки, встановлені статтею 51 Закону України «Про Вищу раду правосуддя».')
OSNOVA = ('на підставі статті 106 Закону України «Про судоустрій і статус суддів», статей 34, 49, 50 '
          'Закону України «Про Вищу раду правосуддя» Третя Дисциплінарна палата Вищої ради правосуддя')
KERUIUCHYS = {'prytiahnennia': 'Керуючись викладеним, ' + OSNOVA,
              'vidmova': 'На підставі' + OSNOVA[len('на підставі'):]}


def w(tag):
    return W + tag


# ───────────────────────── технічні виправлення (вносяться напряму) ─────────────────────────
CYR = 'А-ЯҐЄІЇа-яґєії'
FIXES = [
    (re.compile(r"(?<=[%s])['ʼ`](?=[%s])" % (CYR, CYR)), '’'),                       # апостроф
    (re.compile(r'(?<=\S) [-—] (?=\S)'), ' – '),                                       # дефіс/довге тире між словами
    (re.compile(r'\b((?:19|20)\d{2})-((?:19|20)\d{2})\b'), r'\1–\2'),                  # 2023-2025
    (re.compile(r'(?<=\S)  +(?=\S)'), ' '),                                            # подвійні пробіли
]
RANGE = re.compile(r'\b(пункт(?:и|ах|ів|ами)|підпункт(?:и|ах|ів)|абзац(?:и|ах|ів)|стат(?:тях|ей|тями)'
                   r'|частин(?:ах|ами)|сторінк(?:ах|и)|розділ(?:и|ах|ів))\s+(\d+)-(\d+)\b')


def fix_text(s, first=False):
    for rx, rep in FIXES:
        s = rx.sub(rep, s)
    s = RANGE.sub(lambda m: f'{m[1]} {m[2]}–{m[3]}' if int(m[2]) < int(m[3]) else m[0], s)
    if first:
        s = re.sub(r'^[-—]\s*(?=\S)', '– ', s)                                         # «- ВП № …» → «– ВП № …»
    return s


# ───────────────────────── робота з абзацами ─────────────────────────
SIMPLE_SKIP = {w('rPr')}


def is_simple(r):
    kids = [k for k in r if k.tag not in SIMPLE_SKIP]
    return r.tag == w('r') and len(kids) == 1 and kids[0].tag == w('t')


def bar_text(el):
    out = []
    for x in el.iter():
        if x.tag == w('t'):
            out.append(x.text or '')
        elif x.tag == w('tab'):
            out.append('\t')
        elif x.tag in (w('noBreakHyphen'),):
            out.append('-')
        elif x.tag in (w('footnoteReference'), w('endnoteReference')):
            out.append('†')
    return ''.join(out)


def clean_para(p):
    for x in list(p.iter(w('proofErr'), w('lastRenderedPageBreak'))):
        x.getparent().remove(x)
    # злиття сусідніх простих runs з однаковим rPr
    prev = None
    for ch in list(p):
        if is_simple(ch):
            rpr = ch.find(w('rPr'))
            key = etree.tostring(rpr) if rpr is not None else b''
            if prev is not None and prev[1] == key and prev[0].getnext() is ch:
                t0, t1 = prev[0].find(w('t')), ch.find(w('t'))
                t0.text = (t0.text or '') + (t1.text or '')
                t0.set(XML_SPACE, 'preserve')
                p.remove(ch)
                continue
            prev = (ch, key)
        else:
            prev = None


def segments(p):
    """[(kind, el, text)] прямих нащадків абзацу: 'run' — простий run, 'bar' — усе інше з текстом."""
    segs = []
    for ch in p:
        if ch.tag == w('pPr'):
            continue
        if is_simple(ch):
            segs.append(('run', ch, ch.find(w('t')).text or ''))
        elif ch.tag in (w('r'), w('hyperlink'), w('smartTag'), w('ins'), w('fldSimple'), w('sdt')):
            segs.append(('bar', ch, bar_text(ch)))
        elif ch.tag == w('del'):
            segs.append(('bar', ch, ''))
    return segs


def ptext(p):
    return ''.join(s[2] for s in segments(p))


def mk_t(text, tag='t'):
    t = etree.Element(w(tag))
    t.text = text
    t.set(XML_SPACE, 'preserve')
    return t


class Tracker:
    def __init__(self, start=90000):
        self.id = start

    def mark(self, tag):
        self.id += 1
        e = etree.Element(w(tag))
        e.set(w('id'), str(self.id)); e.set(w('author'), AUTHOR); e.set(w('date'), DATE)
        return e


def split_run(r, off):
    """Розрізає простий run на два за зміщенням off; повертає другий."""
    t = r.find(w('t'))
    txt = t.text or ''
    r2 = copy.deepcopy(r)
    t.text = txt[:off]; t.set(XML_SPACE, 'preserve')
    t2 = r2.find(w('t')); t2.text = txt[off:]; t2.set(XML_SPACE, 'preserve')
    r.addnext(r2)
    return r2


def split_at(p, pos):
    acc = 0
    for kind, el, txt in segments(p):
        if acc < pos < acc + len(txt):
            if kind != 'run':
                raise ValueError('межа правки всередині поля / виноски / іншої правки')
            split_run(el, pos - acc)
            return
        acc += len(txt)


def rpr_of(r):
    rp = r.find(w('rPr')) if r is not None else None
    return copy.deepcopy(rp) if rp is not None else None


def apply_hunk(p, start, end, ins_text, tr):
    split_at(p, start); split_at(p, end)
    acc, dels, before, after = 0, [], None, None
    for kind, el, txt in segments(p):
        a, b = acc, acc + len(txt)
        acc = b
        if start < end and a >= start and b <= end and txt:
            if kind != 'run':
                raise ValueError(f'фрагмент для вилучення зачіпає поле/виноску/правку: {txt!r}')
            dels.append(el)
        elif start < end and a < end and b > start:
            raise ValueError('межа правки не збігається з межею run')
        elif start < end and kind == 'bar' and a == b and start < a < end:
            raise ValueError('у фрагменті для вилучення є виноска або поле')
        if b <= start and kind == 'run':
            before = el
        if a >= end and after is None and kind == 'run':
            after = el
    proto = dels[0] if dels else (before if before is not None else after)
    anchor = dels[-1] if dels else None
    for el in dels:
        d = tr.mark('del')
        el.addprevious(d)
        d.append(el)
        t = el.find(w('t')); t.tag = w('delText')
    if ins_text:
        ins = tr.mark('ins')
        r = etree.SubElement(ins, w('r'))
        rp = rpr_of(proto)
        if rp is not None:
            r.append(rp)
        r.append(mk_t(ins_text))
        if anchor is not None:
            anchor.getparent().addnext(ins)
        else:
            # вставка в позицію start: після останнього сегмента, що закінчується в start
            acc, prev_el = 0, None
            for kind, el, txt in segments(p):
                if acc + len(txt) <= start:
                    prev_el = el
                    acc += len(txt)
                else:
                    break
            if prev_el is not None:
                prev_el.addnext(ins)
            else:
                ppr = p.find(w('pPr'))
                if ppr is not None:
                    ppr.addnext(ins)
                else:
                    p.insert(0, ins)


TOK = re.compile(r'\w+|\s+|[^\w\s]', re.U)


def tracked_replace(p, old, new, tr, n=1, strict=True, whole=False):
    txt = ptext(p)
    idx, k = -1, 0
    while k < n:
        idx = txt.find(old, idx + 1)
        if idx < 0:
            raise ValueError(f'не знайдено (входження {n}): {old[:80]!r}')
        k += 1
    if strict and txt.find(old, idx + 1) >= 0 and n == 1:
        cnt = txt.count(old)
        if cnt > 1:
            raise ValueError(f'фрагмент неоднозначний ({cnt} входжень), розширте old або задайте "n": {old[:60]!r}')
    a, b = TOK.findall(old), TOK.findall(new)
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    hunks, pos = [], [0]
    for t in a:
        pos.append(pos[-1] + len(t))
    if whole:
        hunks = [(idx, idx + len(old), new)]
    else:
        raw = [(i1, i2, ''.join(b[j1:j2])) for op, i1, i2, j1, j2 in sm.get_opcodes() if op != 'equal']
        merged = []
        for h in raw:                       # зливаємо правки, між якими 1–2 слова: читається як одна заміна
            if merged and len(re.findall(r'\w+', ''.join(a[merged[-1][1]:h[0]]))) <= 2:
                ps, pe, pins = merged[-1]
                merged[-1] = (ps, h[1], pins + ''.join(a[pe:h[0]]) + h[2])
            else:
                merged.append(h)
        hunks = [(idx + pos[i1], idx + pos[i2], ins) for i1, i2, ins in merged]
    for s, e, ins in reversed(hunks):
        apply_hunk(p, s, e, ins, tr)
    return len(hunks)


def tracked_delete_para(p, tr):
    for ch in list(p):
        if ch.tag == w('r'):
            d = tr.mark('del')
            ch.addprevious(d); d.append(ch)
            for t in ch.iter(w('t')):
                t.tag = w('delText')
            for t in ch.iter(w('instrText')):
                t.tag = w('delInstrText')
    ppr = p.find(w('pPr'))
    if ppr is None:
        ppr = etree.Element(w('pPr')); p.insert(0, ppr)
    rpr = ppr.find(w('rPr'))
    if rpr is None:
        rpr = etree.SubElement(ppr, w('rPr'))
    rpr.insert(0, tr.mark('del'))


def first_run(p):
    for r in p.iter(w('r')):
        if r.find(w('t')) is not None:
            return r
    return None


def tracked_insert_para(after_p, text, like_p, tr):
    np_ = etree.Element(w('p'))
    ppr = like_p.find(w('pPr'))
    ppr = copy.deepcopy(ppr) if ppr is not None else etree.Element(w('pPr'))
    for x in ppr.findall(w('sectPr')):
        ppr.remove(x)
    rpr = ppr.find(w('rPr'))
    if rpr is None:
        rpr = etree.SubElement(ppr, w('rPr'))
    rpr.insert(0, tr.mark('ins'))
    np_.append(ppr)
    ins = tr.mark('ins')
    r = etree.SubElement(ins, w('r'))
    rp = rpr_of(first_run(like_p))
    if rp is not None:
        for b in rp.findall(w('b')) + rp.findall(w('i')):
            rp.remove(b)
        r.append(rp)
    r.append(mk_t(text))
    np_.append(ins)
    after_p.addnext(np_)
    return np_


def set_para_text(p, text):
    """Пряма заміна тексту абзацу (форматування першого run)."""
    r0 = first_run(p)
    rp = rpr_of(r0)
    for ch in list(p):
        if ch.tag != w('pPr'):
            p.remove(ch)
    r = etree.SubElement(p, w('r'))
    if rp is not None:
        r.append(rp)
    r.append(mk_t(text))


# ───────────────────────── пакет DOCX ─────────────────────────
class Pkg:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.files = {n: z.read(n) for n in z.namelist() if not n.endswith('/')}
        self.doc = etree.fromstring(self.files['word/document.xml'])
        self.body = self.doc.find(w('body'))

    def save(self, path):
        self.files['word/document.xml'] = etree.tostring(self.doc, xml_declaration=True, encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            for n, b in self.files.items():
                z.writestr(n, b)

    def paras(self):
        return list(self.body.iter(w('p')))

    def set_author(self):
        n = 'docProps/core.xml'
        if n in self.files:
            s = self.files[n].decode('utf8')
            s = re.sub(r'<dc:creator>.*?</dc:creator>|<dc:creator/>', f'<dc:creator>{AUTHOR}</dc:creator>', s)
            s = re.sub(r'<cp:lastModifiedBy>.*?</cp:lastModifiedBy>|<cp:lastModifiedBy/>',
                       f'<cp:lastModifiedBy>{AUTHOR}</cp:lastModifiedBy>', s)
            s = re.sub(r'<dc:title>.*?</dc:title>', '<dc:title></dc:title>', s)
            self.files[n] = s.encode('utf8')

    def numbering(self):
        """numId → abstractNumId"""
        n = 'word/numbering.xml'
        if n not in self.files:
            return {}
        x = etree.fromstring(self.files[n])
        return {e.get(w('numId')): e.find(w('abstractNumId')).get(w('val')) for e in x.findall(w('num'))}


def para_numbers(pkg):
    """Орієнтовна автонумерація: {id(p): 'N.'} для абзаців з numPr рівня 0 основного списку."""
    absmap = pkg.numbering()
    counts, out = {}, {}
    for p in pkg.paras():
        np_ = p.find(f'{w("pPr")}/{w("numPr")}')
        if np_ is None or np_.find(w('numId')) is None:
            continue
        nid = np_.find(w('numId')).get(w('val'))
        lvl = np_.find(w('ilvl')).get(w('val')) if np_.find(w('ilvl')) is not None else '0'
        if nid == '0':
            continue
        key = absmap.get(nid, nid)
        c = counts.setdefault(key, {})
        c[lvl] = c.get(lvl, 0) + 1
        for l in list(c):
            if int(l) > int(lvl):
                c[l] = 0
        out[id(p)] = (key, lvl, c[lvl])
    # основний список — той, у якому найбільше абзаців рівня 0
    tot = {}
    for key, lvl, _ in out.values():
        if lvl == '0':
            tot[key] = tot.get(key, 0) + 1
    main = max(tot, key=tot.get) if tot else None
    return {k: (f'{v[2]}.' if v[0] == main and v[1] == '0' else '•') for k, v in out.items()}


# ───────────────────────── розмітка висновку ─────────────────────────
def locate(pkg):
    P = pkg.paras()
    T = [ptext(p).strip() for p in P]
    top = list(pkg.body)
    loc = {'P': P, 'T': T}
    loc['proc'] = next((i for i, t in enumerate(T) if re.match(r'^(І|I)\.\s*ПРОЦЕДУРНІ', t)), None)
    loc['ocinka'] = next((i for i, t in enumerate(T) if re.match(r'^(ІІІ|III)\.\s*ОЦІНКА СКАРГ', t)), None)
    sig = [i for i, t in enumerate(T) if re.match(r'^Дисциплінарний інспектор\s*$', t)]
    loc['sig'] = sig[-1] if sig else None
    end = loc['sig'] if loc['sig'] is not None else len(P)
    prop = [i for i, t in enumerate(T[:end]) if re.match(r'^Пропоную\b', t)]
    loc['propose'] = prop[-1] if prop else None
    dod = [i for i, t in enumerate(T[:end]) if re.match(r'^Додатки?\s*:', t)]
    loc['dodatky'] = dod[-1] if dod and (loc['propose'] is None or dod[-1] > loc['propose']) else None
    loc['appendix'] = [i for i, t in enumerate(T) if re.match(r'^Додаток\s*№?\s*\d+\s+до\s+Висновку', t)]
    # останній абзац мотивувальної частини
    stop = min(x for x in (loc['dodatky'], loc['sig'], len(P)) if x is not None)
    last = stop - 1
    while last > 0 and not T[last]:
        last -= 1
    loc['last'] = last
    loc['propose_items'] = list(range(loc['propose'] + 1, last + 1)) if loc['propose'] is not None else []
    return loc


HINTS = [
    ('Самопосилання «Висновок» (крім «цього Висновку», «до Висновку» — їх скрипт замінює сам)',
     re.compile(r'(?<!цього )(?<!до )\bВисновк\w*')),
    ('«Ознаки» порушення / проступку (перевести у ствердну форму, крім цитат ухвал і фактичних описів)',
     re.compile(r'ознак\w*', re.I)),
    ('Перша особа інспектора', re.compile(r'\b(вважаю|пропоную|доходжу|на мою думку|вбачаю|переконаний|зазначу|наголошу|'
                                          r'звертаю увагу|констатую|уважаю|вважав би|мною)\b', re.I)),
    ('Безособові оцінкові звороти (кандидати на голос палати)',
     re.compile(r'^(Очевидно|Варто|Слід|Потрібно|Необхідно|Незайвим|Безумовно|Важливо|Належить|Варто|Під час '
                r'(підготовки|попередньої перевірки)[^,]*встановлено|Встановлено|З урахуванням (цього|викладеного) слід)', re.I)),
    ('Кандидати на зміни по суті — лише для звіту, НЕ змінювати',
     re.compile(r'11-125сап20|/[12]дп/|Посібник|URL:|https?://')),
]


def cmd_extract(a):
    pkg = Pkg(a.inp)
    for p in pkg.paras():
        clean_para(p)
    technical_fixes(pkg)
    loc = locate(pkg)
    nums = para_numbers(pkg)
    P, T = loc['P'], loc['T']
    out = [f'# Висновок: {os.path.basename(a.inp)}', '',
           'Текст подано після технічних виправлень (апостроф, тире, діапазони, подвійні пробіли) — '
           'саме з ним звіряються "old" у data.json. Індекс [pN] — для "p", "after", "like". † — виноска.', '']
    out.append('## Розмітка')
    for k in ('proc', 'ocinka', 'propose', 'last', 'dodatky', 'sig'):
        out.append(f'- {k}: {"p" + str(loc[k]) if loc[k] is not None else "—"}')
    out.append(f'- propose_items: {", ".join("p%d" % i for i in loc["propose_items"]) or "—"}')
    out.append(f'- appendix: {", ".join("p%d" % i for i in loc["appendix"]) or "—"}')
    out.append('')
    out.append('## Абзаци до розділу I (шапка висновку; у рішенні замінюється)')
    end_head = loc['proc'] if loc['proc'] is not None else 0
    for i in range(end_head):
        if T[i]:
            out.append(f'[p{i}] {T[i]}')
    out.append('')
    out.append('## Текст')
    stop = loc['sig'] if loc['sig'] is not None else len(P)
    pidx = {id(p): i for i, p in enumerate(P)}
    for i in range(end_head, stop):
        if not T[i]:
            continue
        tbl = next((a for a in P[i].iterancestors(w('tbl'))), None)
        if tbl is not None:                                     # таблиця — стисло, рядками
            first = pidx[id(next(tbl.iter(w('p'))))]
            if i != first:
                continue
            rows = []
            for tr_ in tbl.iter(w('tr')):
                cells = [' / '.join(ptext(q).strip() for q in tc.iter(w('p')) if ptext(q).strip()) for tc in tr_.iter(w('tc'))]
                rows.append(' | '.join(cells))
            last_p = pidx[id(list(tbl.iter(w('p')))[-1])]
            out.append(f'[таблиця p{first}–p{last_p}, рядків {len(rows)}]')
            for r_ in rows[:6]:
                out.append('    ' + r_[:300])
            if len(rows) > 6:
                out.append(f'    … ще {len(rows) - 6} рядків')
            continue
        n = nums.get(id(P[i]), '')
        out.append(f'[p{i}]{" (" + n + ")" if n else ""} {T[i]}')
    if loc['appendix']:
        out.append('')
        out.append('## Додатки (після підписів; заголовки «до Висновку» → «до рішення» скрипт змінює сам)')
        for i in loc['appendix']:
            out.append(f'[p{i}] {T[i]}')
    out.append('')
    out.append('## Підказки')
    for title, rx in HINTS:
        hits = [i for i in range(end_head, stop) if T[i] and rx.search(T[i])
                and next(P[i].iterancestors(w('tbl')), None) is None]
        out.append(f'### {title}: {len(hits)}')
        for i in hits:
            m = rx.search(T[i])
            s = max(0, m.start() - 80)
            out.append(f'- [p{i}] …{T[i][s:m.end() + 120]}…')
    with open(a.out, 'w', encoding='utf8') as f:
        f.write('\n'.join(out) + '\n')
    print(f'{a.out}: {sum(1 for t in T if t)} непорожніх абзаців; розмітка:',
          {k: loc[k] for k in ('proc', 'ocinka', 'propose', 'last', 'dodatky', 'sig')})


def technical_fixes(pkg):
    n = 0
    for p in pkg.paras():
        segs = segments(p)
        first, prev_space = True, False
        for kind, el, txt in segs:
            if kind == 'run':
                t = el.find(w('t'))
                new = fix_text(t.text or '', first=first)
                if prev_space and new.startswith(' ') and new.strip():
                    new = new.lstrip(' ')                              # подвійний пробіл на межі runs
                if new != (t.text or ''):
                    t.text = new; t.set(XML_SPACE, 'preserve'); n += 1
                txt = new
            if txt:
                prev_space = txt.endswith(' ')
            if txt.strip():
                first = False
    return n


# ───────────────────────── шаблон ─────────────────────────
def load_template(path):
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    doc = etree.fromstring(files['word/document.xml'])
    body = doc.find(w('body'))
    kids = [k for k in body if k.tag != w('sectPr')]
    i = next(i for i, k in enumerate(kids) if k.tag == w('p') and ptext(k).strip() == '{{BODY}}')
    rels = etree.fromstring(files['word/_rels/document.xml.rels'])
    media = {}
    for rel in rels:
        if rel.get('Type', '').endswith('/image'):
            media[rel.get('Id')] = ('word/' + rel.get('Target').lstrip('/').replace('word/', ''), files['word/' + rel.get('Target').lstrip('/').replace('word/', '')])
    for k in kids:                                  # закладки шаблону можуть дублювати id закладок висновку
        for b in list(k.iter(w('bookmarkStart'), w('bookmarkEnd'))):
            b.getparent().remove(b)
    return kids[:i], kids[i + 1:], media


def fill(el, mapping):
    """Замінює {{KEY}} у тексті всіх w:t елемента (плейсхолдер має бути в одному w:t)."""
    for t in el.iter(w('t')):
        if t.text and '{{' in t.text:
            for k, v in mapping.items():
                t.text = t.text.replace('{{%s}}' % k, v)
            t.set(XML_SPACE, 'preserve')


def has_ph(el, key):
    return any(t.text and '{{%s}}' % key in t.text for t in el.iter(w('t')))


def add_media(pkg, media):
    """Копіює зображення шаблону (герб) у пакет; повертає {старий rId: новий rId}."""
    rels_name = 'word/_rels/document.xml.rels'
    rels = etree.fromstring(pkg.files[rels_name])
    ct = etree.fromstring(pkg.files['[Content_Types].xml'])
    exts = {d.get('Extension').lower() for d in ct.findall('{%s}Default' % CT_NS)}
    mp = {}
    for k, (name, data) in media.items():
        ext = name.rsplit('.', 1)[-1].lower()
        new_name = f'word/media/rishennia_herb_{k}.{ext}'
        pkg.files[new_name] = data
        nid = f'rIdRish{k}'
        e = etree.SubElement(rels, '{%s}Relationship' % PR_NS)
        e.set('Id', nid); e.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        e.set('Target', new_name[len('word/'):])
        if ext not in exts:
            d = etree.SubElement(ct, '{%s}Default' % CT_NS)
            d.set('Extension', ext); d.set('ContentType', {'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpeg'}.get(ext, 'image/' + ext))
            exts.add(ext)
        mp[k] = nid
    pkg.files[rels_name] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    pkg.files['[Content_Types].xml'] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)
    return mp


# ───────────────────────── збирання ─────────────────────────
def cmd_build(a):
    data = json.load(open(a.data, encoding='utf8'))
    kind = data.get('kind', 'prytiahnennia')
    pkg = Pkg(a.inp)
    for p in pkg.paras():
        clean_para(p)
    nfix = technical_fixes(pkg)
    loc = locate(pkg)
    P, T = loc['P'], loc['T']
    if loc['proc'] is None:
        sys.exit('Не знайдено заголовок «І. ПРОЦЕДУРНІ ПИТАННЯ»')
    tr = Tracker()
    report, errors = [], []
    body_range = range(loc['proc'], loc['last'] + 1)

    # 1. Автоматичні правки: самопосилання «Висновок» → «рішення» (у режимі правок)
    auto = [('цього Висновку', 'цього рішення'), ('цьому Висновку', 'цьому рішенні'),
            ('до Висновку', 'до рішення'), ('цей Висновок', 'це рішення')]
    nauto = 0
    for i in body_range:
        for old, new in auto:
            k = 1
            while old in ptext(P[i]):
                try:
                    tracked_replace(P[i], old, new, tr, n=1, strict=False)
                    nauto += 1
                except ValueError as e:
                    errors.append(f'auto p{i}: {e}')
                    break
                k += 1
                if k > 50:
                    break

    # 2. Правки з data.json (у режимі правок). Вставки абзаців — наприкінці, з кінця.
    edits = data.get('edits', [])
    inserts = []
    for e in edits:
        try:
            if 'after' in e:
                inserts.append(e); continue
            i = e['p']
            if not (loc['proc'] <= i <= loc['last']):
                raise ValueError(f'p{i} поза мотивувальною частиною (p{loc["proc"]}–p{loc["last"]})')
            if e.get('delete'):
                tracked_delete_para(P[i], tr)
            elif 'text' in e:                                    # новий текст усього абзацу
                tracked_replace(P[i], ptext(P[i]), e['text'], tr, whole=e.get('whole', False))
            else:
                tracked_replace(P[i], e['old'], e['new'], tr, n=e.get('n', 1), whole=e.get('whole', False))
        except (ValueError, KeyError, IndexError) as ex:
            errors.append(f'edit {json.dumps(e, ensure_ascii=False)[:160]}: {ex}')

    # 3. Пропозиція інспектора → «Керуючись…» (у режимі правок)
    ker = data.get('keruiuchys') or KERUIUCHYS[kind]
    if loc['propose'] is not None:
        pp = P[loc['propose']]
        tracked_replace(pp, ptext(pp), ker, tr, whole=True)
        for i in loc['propose_items']:
            if T[i]:
                tracked_delete_para(P[i], tr)
        report.append(f'абзац «Пропоную…» p{loc["propose"]} замінено на «{ker[:30]}…», '
                      f'пункти пропозиції ({len(loc["propose_items"])}) вилучено правками')
        last_p = P[loc['last']]
    else:
        last_p = tracked_insert_para(P[loc['last']], ker, P[loc['last']], tr)
        report.append(f'абзац «{ker[:30]}…» додано правкою після p{loc["last"]}')

    for e in sorted(inserts, key=lambda e: -e['after']):
        try:
            i = e['after']
            if not (loc['proc'] <= i <= loc['last']):
                raise ValueError(f'p{i} поза мотивувальною частиною')
            tracked_insert_para(P[i], e['text'], P[e.get('like', i)], tr)
        except (ValueError, KeyError, IndexError) as ex:
            errors.append(f'insert {json.dumps(e, ensure_ascii=False)[:160]}: {ex}')

    # 4. Заголовок розділу III — напряму
    if loc['ocinka'] is not None:
        set_para_text(P[loc['ocinka']], HEADING_III)
    else:
        errors.append('не знайдено заголовок «ІІІ. ОЦІНКА СКАРГИ» — назву розділу III не змінено')

    # 5. Кінець (напряму): від «Додатки:» / підпису інспектора до першого додатка або розриву розділу
    head_t, tail_t, media = load_template(a.template)
    top = list(pkg.body)

    def top_of(p):
        while p.getparent() is not pkg.body:
            p = p.getparent()
        return p
    last_top = top_of(last_p)
    stop_el = None
    if loc['appendix']:
        stop_el = top_of(P[loc['appendix'][0]])
    k = top.index(last_top) + 1
    removed = 0
    while k < len(top):
        el = top[k]
        if el is stop_el or el.tag == w('sectPr'):
            break
        if el.find(f'.//{w("sectPr")}') is not None:          # абзац із розривом розділу — зберегти, очистити текст
            for ch in list(el):
                if ch.tag != w('pPr'):
                    el.remove(ch)
            break
        pkg.body.remove(el); removed += 1; k += 1
    comp = data.get('composition') or DEFAULT_COMPOSITION
    tail = [copy.deepcopy(x) for x in tail_t]
    new_tail = []
    for el in tail:
        if has_ph(el, 'OPERATIVE'):
            for txt in data['operative']:
                c = copy.deepcopy(el); fill(c, {'OPERATIVE': txt}); new_tail.append(c)
        elif has_ph(el, 'MEMBER'):
            grp = new_tail[-2:]                                   # два порожні абзаци перед прототипом
            for j, m in enumerate(comp['members'][1:]):
                if j > 0:
                    new_tail += [copy.deepcopy(g) for g in grp]
                c = copy.deepcopy(el); fill(c, {'MEMBER': m}); new_tail.append(c)
            if len(comp['members']) < 2:
                del new_tail[-2:]
        else:
            fill(el, {'APPEAL': data.get('appeal', APPEAL), 'CHAIR': comp['chair'],
                      'MEMBER1': comp['members'][0] if comp['members'] else ''})
            new_tail.append(el)
    style_ids = set(re.findall(r'w:styleId="([^"]+)"', pkg.files.get('word/styles.xml', b'').decode('utf8')))

    def normalize(el):
        # стилю шаблону може не бути у висновку іншої верстки; дублікати w14:paraId прибираємо
        for ps in list(el.iter(w('pStyle'))):
            if ps.get(w('val')) not in style_ids:
                ps.getparent().remove(ps)
        for x in el.iter():
            for k in [k for k in x.attrib if k.endswith('}paraId') or k.endswith('}textId')]:
                del x.attrib[k]
    for el in new_tail:
        normalize(el)
    anchor = last_top
    for el in new_tail:
        anchor.addnext(el); anchor = el
    for i in loc['appendix']:
        p = P[i]
        set_para_text(p, re.sub(r'до\s+Висновку', 'до рішення', ptext(p)))

    # 6. Шапка (напряму): усе до «І. ПРОЦЕДУРНІ ПИТАННЯ» замінюється шапкою шаблону
    proc_top = top_of(P[loc['proc']])
    for el in list(pkg.body):
        if el is proc_top:
            break
        pkg.body.remove(el)
    rid = add_media(pkg, media)
    head = [copy.deepcopy(x) for x in head_t]
    for el in head:
        fill(el, {'DATE': data.get('date', '___ ________ 2026 року'),
                  'NUMBER': data.get('number', '№ _______/3дп/15-26'),
                  'TITLE': data['title'], 'INTRO': data['intro']})
        for blip in el.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
            old = blip.get('{%s}embed' % R_NS)
            if old in rid:
                blip.set('{%s}embed' % R_NS, rid[old])
        for dp in el.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'):
            dp.set('id', '9001')
        normalize(el)
        proc_top.addprevious(el)

    pkg.set_author()
    pkg.save(a.out)
    # звіт
    print(f'Збережено: {a.out}')
    print(f'- технічних виправлень (напряму): {nfix}')
    print(f'- автоматичних замін «Висновок → рішення» (правками): {nauto}')
    print(f'- правок з data.json: {len(edits) - len([x for x in errors if x.startswith(("edit", "insert"))])} з {len(edits)}')
    print(f'- елементів кінця висновку вилучено: {removed}; додатків: {len(loc["appendix"])}')
    for r in report:
        print('- ' + r)
    if errors:
        print('ПОМИЛКИ:')
        for e in errors:
            print('  ! ' + e)
        sys.exit(2)


# ───────────────────────── перевірка ─────────────────────────
def view(p, mode):
    out = []
    for x in p.iter():
        anc = [a.tag for a in x.iterancestors()]
        if x.tag == w('t'):
            if mode == 'reject' and w('ins') in anc:
                continue
            out.append(x.text or '')
        elif x.tag == w('delText') and mode == 'reject':
            out.append(x.text or '')
    return ''.join(out)


def cmd_check(a):
    pkg = Pkg(a.inp)
    ins = pkg.doc.findall(f'.//{w("ins")}'); dels = pkg.doc.findall(f'.//{w("del")}')
    authors = {e.get(w('author')) for e in ins + dels}
    print(f'правок: вставок {len(ins)}, вилучень {len(dels)}; автори: {authors}')
    for p in pkg.paras():
        if p.find(f'.//{w("ins")}') is not None or p.find(f'.//{w("del")}') is not None:
            acc, rej = view(p, 'accept'), view(p, 'reject')
            print('—' * 40)
            print('БУЛО :', rej[:600])
            print('СТАЛО:', acc[:600])
    if a.orig:
        o = Pkg(a.orig)
        ot = [ptext(p) for p in o.paras() if ptext(p).strip()]
        rt = [view(p, 'reject') for p in pkg.paras() if view(p, 'reject').strip()]
        common = set(fix_text(t) for t in ot) & set(rt)
        print(f'\n«відхилити все»: {len(common)} з {len(ot)} абзаців висновку збігаються дослівно (з урахуванням технічних виправлень)')


# ───────────────────────── побудова шаблону ─────────────────────────
def cmd_template(a):
    """Шаблон з готового рішення палати: шапка (до «встановила:» включно + порожній абзац),
    {{BODY}}, кінець (від порожнього абзацу перед «вирішила:»)."""
    pkg = Pkg(getattr(a, 'from'))
    kids = [k for k in pkg.body if k.tag != w('sectPr')]
    tx = [ptext(k).strip() if k.tag == w('p') else '' for k in kids]
    i_vst = tx.index('встановила:')
    i_vyr = tx.index('вирішила:')
    head = kids[:i_vst + 2]
    tail = kids[i_vyr - 1:]
    for el in head:
        for p in el.iter(w('p')):
            clean_para(p)
            t = ptext(p).strip()
            if re.match(r'^_*\s*\S*\s*\d{4} року$', t) or re.match(r'^\d{1,2} \S+ \d{4} року$', t) or re.match(r'^_+ \S+ \d{4} року$', t):
                set_para_text(p, '{{DATE}}')
            elif t.startswith('№'):
                set_para_text(p, '{{NUMBER}}')
            elif t.startswith('Про '):
                set_para_text(p, '{{TITLE}}')
            elif t.startswith('Третя Дисциплінарна палата'):
                set_para_text(p, '{{INTRO}}')
    # кінець: прототипи
    out_tail, sig_seen, state = [], 0, 'before'
    for el in tail:
        if el.tag == w('p'):
            clean_para(el)
        t = ptext(el).strip() if el.tag == w('p') else ''
        if state == 'before' and t.startswith(('притягнути', 'відмовити', 'Притягнути', '1.')):
            set_para_text(el, '{{OPERATIVE}}'); out_tail.append(el); state = 'operative'
        elif state == 'operative' and not t.startswith('Рішення Третьої'):
            continue                                              # інші пункти резолютивної частини
        elif t.startswith('Рішення Третьої'):
            set_para_text(el, '{{APPEAL}}'); out_tail.append(el); state = 'after'
        elif re.search(r'[А-ЯҐЄІЇ’]{3,}$', t) and not t.startswith(('Головуючий', 'Члени')):
            sig_seen += 1
            key = {1: 'CHAIR', 2: 'MEMBER1'}.get(sig_seen, 'MEMBER')
            tabs = [x for x in el.iter(w('tab')) if x.getparent().tag == w('r')]
            after_tab = False if tabs else True
            first_done = False
            for r in list(el.iter(w('r'))):
                if tabs and r is tabs[-1].getparent():
                    after_tab = True; continue
                tt = r.find(w('t'))
                if after_tab and tt is not None:
                    if not first_done:
                        tt.text = '{{%s}}' % key; tt.set(XML_SPACE, 'preserve'); first_done = True
                    else:
                        r.getparent().remove(r)
            out_tail.append(el)
            if key == 'MEMBER':
                break
        else:
            out_tail.append(el)
    for el in list(pkg.body):
        if el.tag != w('sectPr'):
            pkg.body.remove(el)
    sect = pkg.body.find(w('sectPr'))
    for el in head:
        sect.addprevious(el)
    bp = etree.Element(w('p')); r = etree.SubElement(bp, w('r')); r.append(mk_t('{{BODY}}'))
    sect.addprevious(bp)
    for el in out_tail:
        sect.addprevious(el)
    # виноски шаблону не потрібні
    pkg.set_author()
    pkg.save(a.out)
    print('Шаблон:', a.out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest='cmd', required=True)
    e = sp.add_parser('extract'); e.add_argument('--in', dest='inp', required=True); e.add_argument('--out', required=True)
    b = sp.add_parser('build'); b.add_argument('--in', dest='inp', required=True); b.add_argument('--data', required=True)
    b.add_argument('--out', required=True); b.add_argument('--template', default=DEFAULT_TEMPLATE)
    c = sp.add_parser('check'); c.add_argument('--in', dest='inp', required=True); c.add_argument('--orig')
    t = sp.add_parser('template'); t.add_argument('--from', required=True); t.add_argument('--out', required=True)
    a = ap.parse_args()
    {'extract': cmd_extract, 'build': cmd_build, 'check': cmd_check, 'template': cmd_template}[a.cmd](a)


if __name__ == '__main__':
    main()
