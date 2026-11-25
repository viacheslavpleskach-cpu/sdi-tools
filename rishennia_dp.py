# -*- coding: utf-8 -*-
"""Проєкт рішення Дисциплінарної палати ВРП на основі висновку дисциплінарного інспектора
за результатами підготовки дисциплінарної справи до розгляду, а також проєкт ухвали про відкриття
дисциплінарної справи на основі висновку за результатами попередньої перевірки (kind "vidkryttia").
Версія 1.5 (25.11.2026): автозаміна «цього Висновку / до Висновку» з великої літери має ті самі винятки, що й з малої
  (цитати, Висновки КРЄС та інших органів: «до Висновку № 3 (2002)…» не чіпається), кожну заміну й кожен пропуск видно у звіті;
  вилучення абзацу охоплює гіперпосилання і знімає вставки автозаміни; data.json перевіряється до початку роботи;
  заголовки «І. Процедурні питання» розпізнаються в будь-якому регістрі; перша сторінка без номера, якщо так у шаблоні;
  час правок — київський; «check --orig» рахує лише мотивувальну частину.
Версія 1.3 (20.11.2026): дата й номер за замовчуванням — з поточним роком, а не жорстко 2026.
Версія 1.4 (20.11.2026): самопосилання з малої літери («у пункті 131 цього висновку», «у цьому висновку», «додаток 1 до висновку»)
  замінюються правками, крім цитат і висновків інших органів; run з нерозривним дефісом / табуляцією перед текстом
  розрізається, тож правка в такому абзаці більше не падає з «межа правки всередині поля».
Автор: В'ячеслав Плескач. Правила — скіл «proiekt-rishennia-dp»; довідка зі зразками — «Інструкція_рішення_ДП.md» у проєкті «СДІ».

  python3 rishennia_dp.py extract --in ВИСНОВОК.docx --out vysnovok.md
      текст висновку з індексами абзаців [pN] (після технічних виправлень) і підказками
  python3 rishennia_dp.py build --in ВИСНОВОК.docx --data data.json --out РІШЕННЯ.docx
      збирає проєкт рішення: шапка і кінець — напряму, зміни тексту — в режимі правок
  python3 rishennia_dp.py check --in РІШЕННЯ.docx [--orig ВИСНОВОК.docx]
      лічильники правок і текст «прийняти все» змінених абзаців
  python3 rishennia_dp.py template --from РІШЕННЯ_ПАЛАТИ.docx --out Шаблон_рішення_ДП.docx
      перезбирає шаблон шапки й кінця з готового рішення палати

Формат data.json — у скілі. Коротко:
  {"kind": "prytiahnennia" | "vidmova" | "vidkryttia", "date": "…", "title": "…", "intro": "…",
   "pidsumok": "Розглянувши висновок … (лише vidkryttia)", "zakl": N | null (необов'язково),
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
try:                                        # Word показує w:date як місцевий час — пишемо київський
    from zoneinfo import ZoneInfo
    DATE = datetime.datetime.now(ZoneInfo('Europe/Kyiv')).strftime('%Y-%m-%dT%H:%M:%SZ')
except Exception:
    DATE = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
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
              'vidmova': 'На підставі' + OSNOVA[len('на підставі'):],
              'vidkryttia': ('Керуючись статтею 46 Закону України «Про Вищу раду правосуддя», статтею 106 Закону '
                             'України «Про судоустрій і статус суддів», пунктом 13.21 Регламенту Вищої ради '
                             'правосуддя, Третя Дисциплінарна палата Вищої ради правосуддя')}
TEMPLATES = {'prytiahnennia': 'Шаблон_рішення_ДП.docx', 'vidmova': 'Шаблон_рішення_ДП.docx',
             'vidkryttia': 'Шаблон_ухвали_відкриття_ДП.docx'}
APPEALS = {'vidkryttia': 'Ухвала оскарженню не підлягає.'}
PIDSUMOK_START = ('Розглянувши висновок дисциплінарного інспектора Вищої ради правосуддя Плескача В.Ю. за результатами '
                  'попередньої перевірки дисциплінарної скарги, Третя Дисциплінарна палата Вищої ради правосуддя '
                  'вважає, що наведені у скарзі та встановлені фактичні обставини можуть свідчити про наявність у діях ')
PEREVIRKA = ('Перевірка та дослідження наведених обставин будуть здійснюватися дисциплінарним інспектором '
             'під час підготовки дисциплінарної справи до розгляду.')
# самопосилання «Висновок» → назва документа палати
AUTO = {'rishennia': [('цього Висновку', 'цього рішення'), ('цьому Висновку', 'цьому рішенні'),
                      ('до Висновку', 'до рішення'), ('цей Висновок', 'це рішення')],
        'ukhvala': [('цього Висновку', 'цієї ухвали'), ('цьому Висновку', 'цій ухвалі'),
                    ('до Висновку', 'до ухвали'), ('цей Висновок', 'ця ухвала')]}
# те саме з малої літери (1.4): не в цитатах «…» і не про висновки інших органів («цього висновку органу опіки»)
LC_NOT = (r'(?!\s+(?:орган\w*|ВККС|Консультативн\w*|Венеційськ\w*|експерт\w*|КРЄС|ЄСПЛ|№|N\s?\d|\(?\d|Ради|ради|[Кк]омісі\w*|[Кк]омітет\w*|'
          r'служб\w*|суду|Суду|інспектор\w*|прокурор\w*|Бюро|ГРЕКО|GRECO|ОБСЄ|щодо))')
REF = r'\b(?:пункт|підпункт|розділ|абзац|таблиц|додат|виноск)\w*\s+[^.;«»]{0,40}?'
LC_RX = {'цього': re.compile(REF + r'\bцього висновку\b' + LC_NOT),
         'цьому': re.compile(r'(?<!\w)[УуВв] цьому висновку\b' + LC_NOT),
         'додаток': re.compile(r'\b[Дд]одат\w*\s+(?:№\s*)?\d+(?:\s*[,–-]\s*\d+)*\s+до висновку\b')}
# «цей висновок» з малої не замінюється: зазвичай це висновок суду / ЄСПЛ («Зробивши цей висновок, Суд…»)
# з великої літери (1.5): ті самі винятки — «до Висновку № 3 (2002) Консультативної ради…» не є самопосиланням
AUTO_UC = {k: [(re.compile(re.escape(old) + r'\b' + LC_NOT), old, new) for old, new in v] for k, v in AUTO.items()}
AUTO_LC = {'rishennia': [(LC_RX['цього'], 'цього висновку', 'цього рішення'),
                         (LC_RX['цьому'], 'цьому висновку', 'цьому рішенні'),
                         (LC_RX['додаток'], 'до висновку', 'до рішення')],
           'ukhvala': [(LC_RX['цього'], 'цього висновку', 'цієї ухвали'),
                       (LC_RX['цьому'], 'цьому висновку', 'цій ухвалі'),
                       (LC_RX['додаток'], 'до висновку', 'до ухвали')]}


def in_quote(txt, pos):
    """Позиція всередині цитати «…» (незакрита «  перед pos)."""
    head = txt[:pos]
    return head.count('«') > head.count('»')


# відомості про призначення / обрання судді наприкінці висновку попередньої перевірки
APPOINT = re.compile(r'(Указом Президента України|Постановою Верховної Ради України|[Рр]ішенням Вищої ради правосуддя)'
                     r'.{0,250}?\b(призначен|обран|переведен)', re.S)
ZAKL = re.compile(r'Пропоную\s+відкрити\s+дисциплінарн')


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


GLYPHS = {w('noBreakHyphen'), w('softHyphen'), w('tab')}


def is_glyph_run(r):
    """run лише з нерозривного дефіса / м'якого переносу / табуляції — його можна вилучати правкою."""
    kids = [k for k in r if k.tag not in SIMPLE_SKIP]
    return r.tag == w('r') and bool(kids) and all(k.tag in GLYPHS for k in kids)


def split_mixed_runs(p):
    """run на кшталт <noBreakHyphen/><t>текст</t> розрізає на окремі runs (по одному елементу),
    щоб текстова частина стала простим run і правку можна було поставити всередині неї (1.4)."""
    for r in list(p.iter(w('r'))):
        kids = [k for k in r if k.tag not in SIMPLE_SKIP]
        if len(kids) < 2 or not all(k.tag in GLYPHS or k.tag == w('t') for k in kids):
            continue
        rpr = r.find(w('rPr'))
        prev = r
        for k in kids:
            nr = etree.Element(w('r'))
            for a, v in r.attrib.items():
                nr.set(a, v)
            if rpr is not None:
                nr.append(copy.deepcopy(rpr))
            nr.append(k)
            prev.addnext(nr); prev = nr
        r.getparent().remove(r)


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
    split_mixed_runs(p)
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
            if kind != 'run' and not is_glyph_run(el):
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
        t = el.find(w('t'))
        if t is not None:
            t.tag = w('delText')
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
    # нерозривний пробіл у висновку = звичайний у "old" (у vysnovok.md вони не розрізняються);
    # довжина та сама, тож позиції збігаються, а незмінені нерозривні пробіли зберігаються
    txt = ptext(p).replace('\xa0', ' ')
    old = old.replace('\xa0', ' ')
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
    # вставки, зроблені цим запуском (автозаміни, попередні правки), знімаємо: абзац вилучається повністю
    for ins in list(p.iter(w('ins'))):
        if ins.getparent().tag != w('rPr') and ins.get(w('author')) == AUTHOR and ins.get(w('date')) == DATE:
            ins.getparent().remove(ins)
    for ch in list(p.iter(w('r'))):                              # і runs у гіперпосиланнях, полях, чужих вставках
        if next(ch.iterancestors(w('p'))) is not p or next(ch.iterancestors(w('del')), None) is not None:
            continue
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


def unnumber(ppr, tr=None):
    """Прибирає автонумерацію з pPr і ставить звичайний абзацний відступ 1,25 см.
    Для наявного абзацу (tr задано) зміна фіксується як правка форматування (w:pPrChange)."""
    if ppr.find(w('numPr')) is None:
        return
    if tr is not None:
        old = etree.Element(w('pPr'))
        for c in ppr:
            if c.tag not in (w('rPr'), w('sectPr'), w('pPrChange')):
                old.append(copy.deepcopy(c))
    ppr.remove(ppr.find(w('numPr')))
    ind = ppr.find(w('ind'))
    if ind is None:
        ind = etree.Element(w('ind'))
        after_ind = [w(t) for t in ('contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
                                    'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle',
                                    'rPr', 'sectPr', 'pPrChange')]
        anchor = next((c for c in ppr if c.tag in after_ind), None)
        if anchor is not None:
            anchor.addprevious(ind)
        else:
            ppr.append(ind)
    for k in list(ind.attrib):
        del ind.attrib[k]
    ind.set(w('left'), '0'); ind.set(w('firstLine'), '709')
    if tr is not None:
        ch = tr.mark('pPrChange')
        ch.append(old)
        ppr.append(ch)                      # pPrChange — останній дочірній елемент pPr


def tracked_insert_para(after_p, text, like_p, tr, plain=False):
    np_ = etree.Element(w('p'))
    ppr = like_p.find(w('pPr'))
    ppr = copy.deepcopy(ppr) if ppr is not None else etree.Element(w('pPr'))
    for x in ppr.findall(w('sectPr')):
        ppr.remove(x)
    if plain:                               # «Керуючись…» — без номера пункту
        unnumber(ppr)
    rpr = ppr.find(w('rPr'))
    if rpr is None:
        rpr = etree.SubElement(ppr, w('rPr'))
    for x in rpr.findall(w('ins')) + rpr.findall(w('del')):   # позначки правок абзацу-зразка не копіювати
        rpr.remove(x)
    for x in ppr.findall(w('pPrChange')):
        ppr.remove(x)
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
            for tag in ('dc:creator', 'cp:lastModifiedBy'):
                if f'<{tag}>' not in s:
                    s = s.replace('</cp:coreProperties>', f'<{tag}>{AUTHOR}</{tag}></cp:coreProperties>')
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
    loc['proc'] = next((i for i, t in enumerate(T) if re.match(r'^[ІI]\.\s*ПРОЦЕДУРНІ', t.upper())), None)
    loc['ocinka'] = next((i for i, t in enumerate(T)
                          if re.match(r'^[ІI]{3}\.\s*ОЦІНКА\s+(?:ДИСЦИПЛІНАРН\w+\s+)?СКАРГ', t.upper())), None)
    sig = [i for i, t in enumerate(T) if re.match(r'^Дисциплінарний інспектор\s*$', t)]
    loc['sig'] = sig[-1] if sig else None
    end = loc['sig'] if loc['sig'] is not None else len(P)
    prop = [i for i, t in enumerate(T[:end]) if re.match(r'^Пропоную\b', t)]
    loc['propose'] = prop[-1] if prop else None
    dod = [i for i, t in enumerate(T[:end]) if re.match(r'^Додатки?\s*:', t)]
    loc['dodatky'] = dod[-1] if dod and (loc['propose'] is None or dod[-1] > loc['propose']) else None
    loc['appendix'] = [i for i, t in enumerate(T) if re.match(r'^Додаток\s*№?\s*\d+\s+до\s+[Вв]исновку', t)]
    # останній абзац мотивувальної частини
    stop = min(x for x in (loc['dodatky'], loc['sig'], len(P)) if x is not None)
    last = stop - 1
    while last > 0 and not T[last]:
        last -= 1
    loc['last'] = last
    loc['propose_items'] = list(range(loc['propose'] + 1, last + 1)) if loc['propose'] is not None else []
    # тип висновку: попередня перевірка (→ ухвала про відкриття) чи підготовка справи до розгляду (→ рішення)
    head = ' '.join(T[:loc['proc'] or 0])
    loc['tip'] = 'poperednia' if 'попередньої перевірки' in head else 'pidhotovka'
    # підсумковий абзац з пропозицією відкрити справу і відомості про призначення після нього
    lo = loc['ocinka'] if loc['ocinka'] is not None else (loc['proc'] or 0)
    zk = [i for i in range(lo, last + 1) if ZAKL.search(T[i])]
    loc['zakl'] = loc['propose'] if loc['propose'] is not None else (zk[-1] if zk else None)
    tail_from = loc['zakl'] + 1 if loc['zakl'] is not None else max(lo, last - 4)
    loc['appoint'] = [i for i in range(tail_from, last + 1) if T[i] and APPOINT.search(T[i])]
    return loc


HINTS = [
    ('Самопосилання «Висновок» (крім «цього Висновку», «до Висновку» — їх скрипт замінює сам)',
     re.compile(r'(?<!цього )(?<!до )\bВисновк\w*')),
    ('Самопосилання з малої «висновок» (скрипт сам замінює «пункт N цього висновку», «у цьому висновку», «додаток N до висновку», '
     'крім цитат «…» і висновків інших органів; решту — вручну)',
     re.compile(r'\b(?:цього|цьому|цей|даного|даному|цим)\s+висновк\w*|[Дд]одат\w*\s+(?:№\s*)?\d+\s+до\s+висновку')),
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
# для висновку попередньої перевірки (ухвала про відкриття) — інші підказки
HINTS_VIDKRYTTIA = [
    ('Самопосилання «Висновок» (крім «цього Висновку», «до Висновку» — їх скрипт замінює сам)',
     re.compile(r'(?<!цього )(?<!до )\bВисновк\w*')),
    ('Самопосилання з малої «висновок» (скрипт сам замінює «пункт N цього висновку», «у цьому висновку», «додаток N до висновку», '
     'крім цитат «…» і висновків інших органів; решту — вручну)',
     re.compile(r'\b(?:цього|цьому|цей|даного|даному|цим)\s+висновк\w*|[Дд]одат\w*\s+(?:№\s*)?\d+\s+до\s+висновку')),
    ('Ствердні висновки про дії судді (кандидати на гіпотезу: «могла не навести», «міг бути», «може не відповідати»)',
     re.compile(r'\b(вочевидь|очевидно|фактично|безумовно|свідчить|підтверджується|не навів|не навела|не врахував|'
                r'не врахувала|не надав|не надала|порушив|порушила|допустив|допустила|стало|призвело)\b', re.I)),
    ('Безособові оцінкові звороти і «встановлено» (кандидати на голос палати)',
     re.compile(r'^(Очевидно|Варто|Слід|Потрібно|Необхідно|Незайвим|Безумовно|Важливо|Належить|Додатково|Аналіз|'
                r'При цьому|Під час попередньої перевірки[^,]*встановлено|Встановлено|Оцінюючи)|належить урахувати|'
                r'варто зауважити|доходжу висновку', re.I)),
    ('Перша особа інспектора', re.compile(r'\b(вважаю|пропоную|доходжу|на мою думку|вбачаю|переконаний|зазначу|наголошу|'
                                          r'звертаю увагу|констатую|уважаю|мною)\b', re.I)),
    ('Робота інспектора в розділі III (запити, додаткові пояснення, межі попередньої перевірки, пошук у ЄДРСР)',
     re.compile(r'дисциплінарн\w* інспектор|попередньої перевірки|додатков\w* поясненн|запит\w*|ЄДРСР|не виявлено|'
                r'нез’ясован', re.I)),
    ('«Ознаки» (у тексті ухвали НЕ переводити у ствердну форму)', re.compile(r'ознак\w*', re.I)),
    ('Кандидати на зміни по суті — лише для звіту, НЕ змінювати',
     re.compile(r'11-125сап20|/[12]дп/|ухвалою від .{0,40}відкрила дисциплінарну справу|Конституційний Суд|'
                r'Суду справедливості|Посібник|URL:|https?://')),
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
    if loc['proc'] is None:
        out.append('- УВАГА: заголовок «І. ПРОЦЕДУРНІ ПИТАННЯ» не знайдено — тип висновку і межі розділів нижче ненадійні')
    vid = loc['tip'] == 'poperednia'
    out.append('- тип висновку: ' + ('попередня перевірка → проєкт УХВАЛИ про відкриття (kind "vidkryttia")' if vid
                                    else 'підготовка справи до розгляду → проєкт РІШЕННЯ'))
    for k in ('proc', 'ocinka', 'propose', 'zakl', 'last', 'dodatky', 'sig'):
        out.append(f'- {k}: {"p" + str(loc[k]) if loc[k] is not None else "—"}')
    if vid:
        out.append(f'- appoint (відомості про призначення — скрипт вилучає правкою): '
                   f'{", ".join("p%d" % i for i in loc["appoint"]) or "—"}')
        if loc['zakl'] is not None:
            rest = [i for i in range(loc['zakl'] + 1, loc['last'] + 1) if T[i] and i not in loc['appoint']]
            out.append(f'- після zakl залишаються: {", ".join("p%d" % i for i in rest) or "—"}')
    else:
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
    out.append('## Підказки' + (' (ухвала про відкриття; розділи I–II — без змін, підказки лише з розділу III)' if vid else ''))
    hint_from = (loc['ocinka'] or end_head) if vid else end_head
    for title, rx in (HINTS_VIDKRYTTIA if vid else HINTS):
        hits = [i for i in range(hint_from, stop) if T[i] and rx.search(T[i])
                and next(P[i].iterancestors(w('tbl')), None) is None]
        out.append(f'### {title}: {len(hits)}')
        for i in hits:
            m = rx.search(T[i])
            s = max(0, m.start() - 80)
            out.append(f'- [p{i}] …{T[i][s:m.end() + 120]}…')
    with open(a.out, 'w', encoding='utf8') as f:
        f.write('\n'.join(out) + '\n')
    print(f'{a.out}: {sum(1 for t in T if t)} непорожніх абзаців; тип: {loc["tip"]}; розмітка:',
          {k: loc[k] for k in ('proc', 'ocinka', 'propose', 'zakl', 'last', 'dodatky', 'sig')})


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
    sect = body.find(w('sectPr'))
    title_pg = sect is not None and sect.find(w('titlePg')) is not None
    return kids[:i], kids[i + 1:], media, title_pg


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
def validate(data):
    """Перевірка data.json до початку роботи: зрозуміле повідомлення замість traceback наприкінці."""
    bad = []

    def num(e, k):
        v = e.get(k)
        if isinstance(v, str) and v.strip().isdigit():
            e[k] = v = int(v.strip())
        if not isinstance(v, int) or isinstance(v, bool):
            bad.append(f'"{k}" має бути номером абзацу (числом): {json.dumps(e, ensure_ascii=False)[:100]}')
    for k in ('title', 'intro'):
        if not isinstance(data.get(k), str) or not data[k].strip():
            bad.append(f'немає поля "{k}" (рядок)')
    op = data.get('operative')
    if not isinstance(op, list) or not op or not all(isinstance(x, str) and x.strip() for x in op):
        bad.append('поле "operative" має бути непорожнім списком рядків')
    for k in ('date', 'number', 'appeal', 'keruiuchys', 'pidsumok'):
        if k in data and data[k] is not None and not isinstance(data[k], str):
            bad.append(f'поле "{k}" має бути рядком')
    comp = data.get('composition')
    if comp and not (isinstance(comp.get('chair'), str) and isinstance(comp.get('members'), list)
                     and all(isinstance(m, str) for m in comp['members'])):
        bad.append('"composition": {"chair": "Ім’я ПРІЗВИЩЕ", "members": ["…", "…"]}')
    if 'zakl' in data and data['zakl'] is not None:
        num(data, 'zakl')
    if not isinstance(data.get('edits', []), list):
        bad.append('"edits" має бути списком')
    else:
        for e in data.get('edits', []):
            if not isinstance(e, dict):
                bad.append(f'правка має бути об’єктом: {str(e)[:80]}'); continue
            if 'after' in e:
                num(e, 'after')
                if 'like' in e:
                    num(e, 'like')
                if not isinstance(e.get('text'), str):
                    bad.append(f'вставка без "text": {json.dumps(e, ensure_ascii=False)[:100]}')
            elif 'p' in e:
                num(e, 'p')
                if not (e.get('delete') or isinstance(e.get('text'), str)
                        or (isinstance(e.get('old'), str) and isinstance(e.get('new'), str))):
                    bad.append(f'правка без "delete", "text" або пари "old" / "new": {json.dumps(e, ensure_ascii=False)[:100]}')
                if 'n' in e:
                    num(e, 'n')
            else:
                bad.append(f'правка без "p" чи "after": {json.dumps(e, ensure_ascii=False)[:100]}')
    if bad:
        sys.exit('Помилки в data.json (нічого не зібрано):\n  - ' + '\n  - '.join(bad))


def cmd_build(a):
    data = json.load(open(a.data, encoding='utf-8-sig'))
    kind = data.get('kind', 'prytiahnennia')
    if kind not in KERUIUCHYS:
        sys.exit(f'Невідомий kind: {kind} (prytiahnennia / vidmova / vidkryttia)')
    validate(data)
    vid = kind == 'vidkryttia'
    template = a.template or os.path.join(HERE, TEMPLATES[kind])
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
    if vid and loc['tip'] != 'poperednia':
        report.append('УВАГА: kind "vidkryttia", але шапка висновку не містить «попередньої перевірки»')
    if not vid and loc['tip'] == 'poperednia':
        report.append('УВАГА: висновок за результатами попередньої перевірки, а kind — не "vidkryttia"')
    if vid and not str(data.get('pidsumok', '')).startswith('Розглянувши висновок'):
        sys.exit('Для kind "vidkryttia" потрібне поле "pidsumok", що починається з «Розглянувши висновок…»')
    body_range = range(loc['proc'], loc['last'] + 1)

    # 1. Автоматичні правки: самопосилання «Висновок» → «рішення» / «ухвала» (у режимі правок)
    key = 'ukhvala' if vid else 'rishennia'
    nauto, auto_lc, auto_uc, auto_skip = 0, [], [], []
    for i in body_range:
        for rules, log in ((AUTO_UC[key], auto_uc), (AUTO_LC[key], auto_lc)):
            for rx, old, new in rules:
                for _ in range(50):
                    txt = ptext(P[i]).replace('\xa0', ' ')
                    m = next((m for m in rx.finditer(txt) if not in_quote(txt, m.start())), None)
                    if m is None:
                        break
                    pos = m.start() + m.group(0).rfind(old)
                    n = txt[:pos].count(old) + 1
                    try:
                        tracked_replace(P[i], old, new, tr, n=n, strict=False)
                        nauto += 1
                        log.append(f'p{i}: «…{txt[max(0, pos - 25):pos]}[{old} → {new}]{txt[pos + len(old):pos + len(old) + 25]}…»')
                    except ValueError as e:
                        errors.append(f'auto p{i}: {e}')
                        break
        txt = ptext(P[i]).replace('\xa0', ' ')                        # що лишилося без заміни — у звіт
        for old, _ in AUTO[key]:
            for m in re.finditer(re.escape(old), txt):
                auto_skip.append(f'p{i}: «…{txt[max(0, m.start() - 25):m.end() + 35]}…»')

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
    if vid:
        # ухвала про відкриття: підсумковий абзац висновку («… наявні ознаки … Пропоную відкрити…») і відомості
        # про призначення вилучаються; наприкінці — «Розглянувши висновок…», «Перевірка…», «Керуючись…»
        zakl = data['zakl'] if 'zakl' in data else loc['zakl']
        if zakl is not None and not (loc['proc'] <= zakl <= loc['last']):
            errors.append(f'zakl p{zakl} поза мотивувальною частиною'); zakl = None
        numbered = [i for i in body_range if T[i] and P[i].find(f'{w("pPr")}/{w("numPr")}') is not None
                    and next(P[i].iterancestors(w('tbl')), None) is None]
        like = P[zakl] if zakl is not None and zakl in numbered else (P[numbered[-1]] if numbered else P[loc['last']])
        like = copy.deepcopy(like)                                 # зразок форматування до вилучень
        deleted = []
        if zakl is not None:
            tracked_delete_para(P[zakl], tr); deleted.append(zakl)
            report.append(f'підсумковий абзац висновку p{zakl} вилучено правкою')
        else:
            report.append('УВАГА: абзац «Пропоную відкрити…» не знайдено — підсумок висновку не вилучено')
        for i in loc['appoint']:
            if i not in deleted:
                tracked_delete_para(P[i], tr); deleted.append(i)
        if loc['appoint']:
            report.append('відомості про призначення вилучено правками: ' + ', '.join('p%d' % i for i in loc['appoint']))
        if zakl is not None:
            rest = [i for i in range(zakl + 1, loc['last'] + 1) if T[i] and i not in deleted]
            if rest:
                report.append('після підсумкового абзацу залишилися абзаци ' + ', '.join('p%d' % i for i in rest)
                              + ' — перевірити (кандидати для звіту)')
        p1 = tracked_insert_para(P[loc['last']], data['pidsumok'], like, tr)
        p2 = tracked_insert_para(p1, PEREVIRKA, like, tr)
        last_p = tracked_insert_para(p2, ker, like, tr, plain=True)
        report.append('додано правками: «Розглянувши висновок…», «Перевірка та дослідження…», «Керуючись…»')
    elif loc['propose'] is not None:
        pp = P[loc['propose']]
        tracked_replace(pp, ptext(pp), ker, tr, whole=True)
        if pp.find(w('pPr')) is not None:
            unnumber(pp.find(w('pPr')), tr)
        for i in loc['propose_items']:
            if T[i]:
                tracked_delete_para(P[i], tr)
        report.append(f'абзац «Пропоную…» p{loc["propose"]} замінено на «{ker[:30]}…», '
                      f'пункти пропозиції ({len(loc["propose_items"])}) вилучено правками')
        last_p = P[loc['last']]
    else:
        last_p = tracked_insert_para(P[loc['last']], ker, P[loc['last']], tr, plain=True)
        report.append(f'абзац «{ker[:30]}…» додано правкою після p{loc["last"]}')

    # з кінця; кілька вставок після одного абзацу — у порядку data.json
    for e in [x for _, x in sorted(enumerate(inserts), key=lambda t: (-t[1]['after'], -t[0]))]:
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
    head_t, tail_t, media, title_pg = load_template(template)
    sect = pkg.body.find(w('sectPr'))
    if title_pg and sect is not None and sect.find(w('titlePg')) is None:      # перша сторінка без номера, як у шаблоні
        tp = etree.Element(w('titlePg'))
        later = [w(t) for t in ('textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange')]
        anchor = next((c for c in sect if c.tag in later), None)
        anchor.addprevious(tp) if anchor is not None else sect.append(tp)
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
            fill(el, {'APPEAL': data.get('appeal', APPEALS.get(kind, APPEAL)), 'CHAIR': comp['chair'],
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
        set_para_text(p, re.sub(r'до\s+[Вв]исновку', 'до ухвали' if vid else 'до рішення', ptext(p)))

    # 6. Шапка (напряму): усе до «І. ПРОЦЕДУРНІ ПИТАННЯ» замінюється шапкою шаблону
    proc_top = top_of(P[loc['proc']])
    for el in list(pkg.body):
        if el is proc_top:
            break
        pkg.body.remove(el)
    rid = add_media(pkg, media)
    head = [copy.deepcopy(x) for x in head_t]
    for el in head:
        _y = datetime.date.today().year
        fill(el, {'DATE': data.get('date', f'___ ________ {_y} року'),
                  'NUMBER': data.get('number', f'№ _______/3дп/15-{_y % 100:02d}'),
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
    print(f'- вид документа: {"ухвала про відкриття дисциплінарної справи" if vid else "рішення"}; шаблон: {os.path.basename(template)}')
    print(f'- автоматичних замін «Висновок → {"ухвала" if vid else "рішення"}» (правками): {nauto}')
    for x in auto_uc:
        print('  ' + x)
    if auto_lc:
        print('  з них з малої літери (перевірити, що це самопосилання): ' + '; '.join(auto_lc))
    if auto_skip:
        print('- «Висновок» з великої літери НЕ замінено (цитата або висновок іншого органу — перевірити): ' + '; '.join(auto_skip))
    print(f'- правок з data.json: {len(edits) - len([x for x in errors if x.startswith(("edit", "insert"))])} з {len(edits)}')
    print(f'- елементів кінця висновку вилучено: {removed}; додатків: {len(loc["appendix"])}')
    for r in report:
        print('- ' + r)
    if errors:
        print('ПОМИЛКИ (файл збережено, але частину правок не внесено — не видавати без виправлення):')
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
        for p in o.paras():
            clean_para(p)
        technical_fixes(o)
        lo = locate(o)
        body = lo['P'][(lo['proc'] or 0):lo['last'] + 1]             # лише мотивувальна частина: шапку й кінець замінено
        strip = lambda t: re.sub(r'[†\t\xa0 ]+', ' ', t).strip()   # view() не бачить знаків виносок і табуляцій
        ot = [strip(ptext(p)) for p in body if strip(ptext(p))]
        rt = {strip(view(p, 'reject')) for p in pkg.paras()}
        common = [t for t in ot if t in rt]
        print(f'\n«відхилити все»: {len(common)} з {len(ot)} абзаців мотивувальної частини висновку збігаються дослівно '
              '(з урахуванням технічних виправлень; заголовок розділу III змінено напряму)')


# ───────────────────────── побудова шаблону ─────────────────────────
def cmd_template(a):
    """Шаблон з готового рішення (ухвали) палати: шапка (до «встановила:» включно + порожній абзац),
    {{BODY}}, кінець (від порожнього абзацу перед «вирішила:» / «ухвалила:»)."""
    pkg = Pkg(getattr(a, 'from'))
    kids = [k for k in pkg.body if k.tag != w('sectPr')]
    tx = [ptext(k).strip() if k.tag == w('p') else '' for k in kids]
    i_vst = tx.index('встановила:')
    i_vyr = next(i for i, t in enumerate(tx) if t in ('вирішила:', 'ухвалила:'))
    head = kids[:i_vst + 2]
    tail = kids[i_vyr - 1:]
    for el in head:
        for p in el.iter(w('p')):
            clean_para(p)
            if p.find(f'.//{w("drawing")}') is not None:      # абзац з гербом: лише зображення,
                for r in list(p.iter(w('r'))):                    # залишки тексту бланка прибрати
                    if r.find(f'.//{w("drawing")}') is None and r.find(w('t')) is not None:
                        r.getparent().remove(r)
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
        if state == 'before' and t.startswith(('притягнути', 'відмовити', 'Притягнути', 'відкрити', '1.')):
            set_para_text(el, '{{OPERATIVE}}'); out_tail.append(el); state = 'operative'
        elif state == 'operative' and not t.startswith(('Рішення Третьої', 'Ухвала ')):
            continue                                              # інші пункти резолютивної частини
        elif t.startswith(('Рішення Третьої', 'Ухвала ')):
            lead_tab = ptext(el).startswith('\t')
            set_para_text(el, '{{APPEAL}}'); out_tail.append(el); state = 'after'
            if lead_tab:                                          # «\tУхвала оскарженню не підлягає.»
                r0 = el.find(w('r')); rt = etree.Element(w('r'))
                if r0.find(w('rPr')) is not None:
                    rt.append(copy.deepcopy(r0.find(w('rPr'))))
                etree.SubElement(rt, w('tab')); r0.addprevious(rt)
        elif re.search(r'[А-ЯҐЄІЇ’]{3,}$', t) and not t.startswith(('Головуючий', 'Члени')):
            sig_seen += 1
            key = {1: 'CHAIR', 2: 'MEMBER1'}.get(sig_seen, 'MEMBER')
            tabs = [x for x in el.iter(w('tab')) if x.getparent().tag == w('r')]
            after_tab = False if tabs else True
            first_done = False
            for r in list(el.iter(w('r'))):
                if tabs and r is tabs[-1].getparent():
                    after_tab = True
                    seen, found = False, False                    # ім'я в тому самому run після табуляції
                    for ch in list(r):
                        if ch is tabs[-1]:
                            seen = True
                        elif seen and ch.tag == w('t'):
                            r.remove(ch); found = True
                    if found:
                        tabs[-1].addnext(mk_t('{{%s}}' % key)); first_done = True
                    continue
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
    b.add_argument('--out', required=True); b.add_argument('--template', default=None,
                                                           help='типово — за kind: ' + ', '.join(f'{k}: {v}' for k, v in TEMPLATES.items()))
    c = sp.add_parser('check'); c.add_argument('--in', dest='inp', required=True); c.add_argument('--orig')
    t = sp.add_parser('template'); t.add_argument('--from', required=True); t.add_argument('--out', required=True)
    a = ap.parse_args()
    {'extract': cmd_extract, 'build': cmd_build, 'check': cmd_check, 'template': cmd_template}[a.cmd](a)


if __name__ == '__main__':
    main()
