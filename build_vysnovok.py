#!/usr/bin/env python3
"""Збирає DOCX висновку за результатами попередньої перевірки дисциплінарної скарги.

Версія 1.4 (21.11.2026): заголовки розділів («# ») і підзаголовки («## ») не відриваються від наступного
абзацу внизу сторінки (w:keepNext).
Версія 1.3 (20.11.2026): вкладений перелік «1) …; 2) …» (рядки блоку, що починаються з «- ») — для розділу I,
коли інспектор того самого дня надіслав кілька запитів (правило від 20.11.2026, зразок «Смик витяг»):
нумерований пункт «[дата] дисциплінарний інспектор … надіслав наступні запити:», під ним — підпункти
без наскрізного номера, нумерація «1)» починається заново в кожному переліку.
Версія 1.2 (18.11.2026): кожен фрагмент «[ПОТРЕБУЄ ДОЗАПОВНЕННЯ…]» (і в тексті, і в таблиці, і у виносках)
виділяється жовтим маркером (w:highlight yellow), щоб інспектор одразу бачив прогалини.
1.1 (14.11.2026): автор у властивостях документа — «В'ячеслав Плескач» (dc:creator і cp:lastModifiedBy),
незалежно від шаблону.

Використання:
    python3 build_vysnovok.py ШАБЛОН.docx ЗМІСТ.md РЕЗУЛЬТАТ.docx

ШАБЛОН — будь-який висновок Плескача В.Ю. тієї самої верстки (за замовчуванням «Шаблон_висновку.docx» =
«143132 Боровков.docx»; для висновку з пропозицією відкрити справу — «Шаблон_висновку_відкриття.docx»):
з нього беруться шапка, таблиці, стилі, нумерація, підпис.

Формат ЗМІСТ.md
---------------
Шапка (рядки на початку файлу):
    @дата: 22 вересня 2026
    @скаржник: Яремків Ольга Михайлівна
    @скарга: Я-3997/1/7-26 від 18.09.2026          (можна кілька рядків @скарга:)
    @суддя: Боровков Дмитро Олександрович - Залізничний районний суд міста Львова   (кілька — кілька рядків)
    @пропозиція: Відмовити у відкритті ...          (кілька абзаців — кілька рядків)
    @акт: Указ Президента України від ... № ...      (лише якщо в шаблоні є рядок «Акт про призначення»)

Тіло (абзаци розділяються порожнім рядком):
    # І. Процедурні питання              -> заголовок розділу (жирний)
    ## Загальні принципи                 -> підзаголовок (жирний курсив, підкреслений)
    > «Текст цитати...                   -> цитата: курсив, відступ, без номера
    звичайний абзац                      -> нумерований пункт (наскрізна автонумерація Word)
    =текст                               -> абзац без номера (звичайний відступ)
    - перший запит;                      -> блок, усі рядки якого починаються з «- », — вкладений перелік
    - другий запит.                         «1) …», «2) …» (нумерація заново в кожному блоці)
Усередині абзацу:
    *курсив*, **жирний**, [^текст виноски] -> виноска внизу сторінки
"""
import re
import sys
import zipfile
import html

AUTHOR = "В'ячеслав Плескач"
W_T = re.compile(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>')


def esc(s):
    return html.escape(s, quote=False)


def top_elements(body):
    return re.findall(r'(<w:p[ >].*?</w:p>|<w:tbl>.*?</w:tbl>|<w:sectPr.*?</w:sectPr>)', body, re.S)


def text_of(e):
    return ''.join(W_T.findall(e))


def ppr_of(p):
    m = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S)
    return m.group(0) if m else ''


def first_rpr(p):
    m = re.search(r'<w:r(?:\s[^>]*)?>\s*(<w:rPr>.*?</w:rPr>)', p, re.S)
    return m.group(1) if m else '<w:rPr></w:rPr>'


def rpr_with(rpr, italic=None, bold=None):
    inner = rpr[len('<w:rPr>'):-len('</w:rPr>')]
    if italic is not None:
        inner = re.sub(r'<w:i/>|<w:iCs/>', '', inner)
    if bold is not None:
        inner = re.sub(r'<w:b/>|<w:bCs/>', '', inner)
    add = ''
    if bold:
        add += '<w:b/><w:bCs/>'
    if italic:
        add += '<w:i/><w:iCs/>'
    # rFonts має йти першим
    m = re.match(r'(<w:rFonts[^>]*/>)(.*)', inner, re.S)
    if m:
        inner = m.group(1) + add + m.group(2)
    else:
        inner = add + inner
    return '<w:rPr>' + inner + '</w:rPr>'


GAP = re.compile(r'\[ПОТРЕБУЄ ДОЗАПОВНЕННЯ[^\]]*\]')
RUN = re.compile(r'<w:r>(<w:rPr>.*?</w:rPr>)?<w:t(?: xml:space="preserve")?>([^<]*)</w:t></w:r>', re.S)


def highlight_gaps(xml):
    """Розбиває прогони з «[ПОТРЕБУЄ ДОЗАПОВНЕННЯ…]» і виділяє ці фрагменти жовтим маркером."""
    def one(m):
        rpr, text = m.group(1) or '<w:rPr></w:rPr>', m.group(2)
        if 'ПОТРЕБУЄ ДОЗАПОВНЕННЯ' not in text:
            return m.group(0)
        hl = re.sub(r'<w:highlight [^>]*/>', '', rpr)
        # порядок у w:rPr за схемою: highlight — перед u, effect, bdr, shd, fitText, vertAlign, rtl, cs, em, lang …
        m2 = re.search(r'<w:(?:u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath)[ >/]', hl)
        cut = m2.start() if m2 else hl.rindex('</w:rPr>')
        hl = hl[:cut] + '<w:highlight w:val="yellow"/>' + hl[cut:]
        out, pos = [], 0
        for g in GAP.finditer(text):
            if g.start() > pos:
                out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{text[pos:g.start()]}</w:t></w:r>')
            out.append(f'<w:r>{hl}<w:t xml:space="preserve">{g.group(0)}</w:t></w:r>')
            pos = g.end()
        if pos < len(text):
            out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{text[pos:]}</w:t></w:r>')
        return ''.join(out)
    return RUN.sub(one, xml)


def set_author(core):
    a = esc(AUTHOR)
    for tag in ('dc:creator', 'cp:lastModifiedBy'):
        if re.search(rf'<{tag}[ >/]', core):
            core = re.sub(rf'<{tag}(?:\s[^>]*)?(?:/>|>.*?</{tag}>)', f'<{tag}>{a}</{tag}>', core, flags=re.S)
        else:
            core = core.replace('</cp:coreProperties>', f'<{tag}>{a}</{tag}></cp:coreProperties>')
    return core


def keep_next(p):
    """Заголовок не лишається внизу сторінки окремо від тексту (w:keepNext)."""
    if '<w:keepNext/>' in p:
        return p
    if '<w:pPr>' in p:
        m = re.search(r'<w:pPr>(<w:pStyle [^>]*/>)?', p)
        return p[:m.end()] + '<w:keepNext/>' + p[m.end():]
    return p.replace('<w:p>', '<w:p><w:pPr><w:keepNext/></w:pPr>', 1)


class Doc:
    def __init__(self, tpl):
        self.z = zipfile.ZipFile(tpl)
        self.files = {n: self.z.read(n) for n in self.z.namelist()}
        d = self.files['word/document.xml'].decode('utf8')
        b0 = d.find('<w:body>') + len('<w:body>')
        b1 = d.rfind('</w:body>')
        self.head, self.tail = d[:b0], d[b1:]
        self.els = top_elements(d[b0:b1])
        self.fn_xml = self.files['word/footnotes.xml'].decode('utf8')
        self.fn_next = 1
        self._find_protos()

    def _find_protos(self):
        els = self.els
        idx = {}
        tables = [i for i, e in enumerate(els) if e.startswith('<w:tbl>')]
        self.i_date, self.i_info = tables[0], tables[1]
        for i, e in enumerate(els):
            t = text_of(e).strip()
            if 'h1' not in idx and 'ПРОЦЕДУРНІ ПИТАННЯ' in t.upper():
                idx['h1'] = i
            if 'h2' not in idx and t.lower() in ('загальні принципи', 'зміст скарги',
                                                 'зміст дисциплінарної скарги',
                                                 'застосування загальних принципів у цій справі'):
                idx['h2'] = i
            if t.startswith('Дисциплінарний інспектор'):
                idx['sign'] = i
        self.i_h1 = idx['h1']
        self.p_head = els[idx['h1']]
        self.p_sub = els[idx['h2']] if 'h2' in idx else None
        self.i_sign = idx['sign']
        nums = [e for e in els if '<w:numPr>' in e and e.startswith('<w:p')]
        self.p_num = nums[0]
        quotes = [e for e in els if e.startswith('<w:p') and '<w:numPr>' not in e
                  and 'w:left="397"' in e]
        self.p_quote = quotes[0] if quotes else None
        # шапка до першого заголовка; хвіст: порожній абзац перед підписом + підпис + sectPr
        self.i_post = idx['sign'] - 1

    # --- побудова абзаців ---
    def runs(self, text, base_rpr, italic=False):
        out = []
        parts = re.split(r'(\[\^[^\]]+\])', text)
        for part in parts:
            if not part:
                continue
            m = re.match(r'\[\^([^\]]+)\]', part)
            if m:
                out.append(self.footnote(m.group(1), base_rpr))
                continue
            for seg in re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', part):
                if not seg:
                    continue
                b = seg.startswith('**') and seg.endswith('**')
                i = (not b) and seg.startswith('*') and seg.endswith('*')
                s = seg[2:-2] if b else seg[1:-1] if i else seg
                rp = rpr_with(base_rpr, italic=True, bold=b) if (i or italic) else (
                    rpr_with(base_rpr, bold=True) if b else base_rpr)
                out.append(f'<w:r>{rp}<w:t xml:space="preserve">{esc(s)}</w:t></w:r>')
        return ''.join(out)

    def footnote(self, text, base_rpr):
        n = self.fn_next
        self.fn_next += 1
        sup = rpr_with(base_rpr).replace('</w:rPr>', '<w:vertAlign w:val="superscript"/></w:rPr>')
        small = ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
                 'w:cs="Times New Roman"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="uk-UA"/></w:rPr>')
        small_sup = small.replace('</w:rPr>', '<w:vertAlign w:val="superscript"/></w:rPr>')
        fn = (f'<w:footnote w:id="{n}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
              f'</w:pPr><w:r>{small_sup}<w:footnoteRef/></w:r><w:r>{small}<w:t xml:space="preserve"> '
              f'{esc(text)}</w:t></w:r></w:p></w:footnote>')
        self.fn_xml = self.fn_xml.replace('</w:footnotes>', fn + '</w:footnotes>')
        return f'<w:r>{sup}<w:footnoteReference w:id="{n}"/></w:r>'

    def para(self, proto, text, italic=False):
        ppr = ppr_of(proto)
        rpr = first_rpr(proto)
        return f'<w:p>{ppr}{self.runs(text, rpr, italic)}</w:p>'

    def quote(self, text):
        proto = self.p_quote or self.p_num
        ppr = ppr_of(proto)
        if not self.p_quote:
            ppr = re.sub(r'<w:numPr>.*?</w:numPr>', '', ppr, flags=re.S)
            ppr = ppr.replace('<w:spacing', '<w:ind w:left="397" w:firstLine="454"/><w:spacing', 1)
        rpr = rpr_with(first_rpr(self.p_num), italic=False, bold=False)
        lead = ''
        if text.startswith('«'):
            lead = f'<w:r>{rpr}<w:t>«</w:t></w:r>'
            text = text[1:]
        return f'<w:p>{ppr}{lead}{self.runs(text, rpr, italic=True)}</w:p>'

    def plain(self, text):
        ppr = re.sub(r'<w:numPr>.*?</w:numPr>', '', ppr_of(self.p_num), flags=re.S)
        ppr = ppr.replace('<w:spacing', '<w:ind w:firstLine="709"/><w:spacing', 1)
        return f'<w:p>{ppr}{self.runs(text, first_rpr(self.p_num))}</w:p>'

    def sublist(self, items):
        """Вкладений перелік «1)» з окремою нумерацією для кожного блоку (запити одного дня)."""
        n = self.files['word/numbering.xml'].decode('utf8')
        if not hasattr(self, 'abs_sub'):
            ids = [int(x) for x in re.findall(r'w:abstractNumId="(\d+)"', n)]
            self.abs_sub = max(ids) + 1 if ids else 0
            rpr = re.sub(r'<w:b/>|<w:bCs/>|<w:i/>|<w:iCs/>', '', first_rpr(self.p_num))
            absn = (f'<w:abstractNum w:abstractNumId="{self.abs_sub}"><w:multiLevelType w:val="hybridMultilevel"/>'
                    f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/>'
                    f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="709" w:firstLine="284"/></w:pPr>{rpr}</w:lvl>'
                    f'</w:abstractNum>')
            # abstractNum мають стояти перед першим w:num
            k = n.find('<w:num ')
            n = n[:k] + absn + n[k:] if k > 0 else n.replace('</w:numbering>', absn + '</w:numbering>')
        nums = [int(x) for x in re.findall(r'<w:num w:numId="(\d+)"', n)]
        nid = max(nums) + 1 if nums else 1
        n = n.replace('</w:numbering>', f'<w:num w:numId="{nid}"><w:abstractNumId w:val="{self.abs_sub}"/>'
                      f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>')
        self.files['word/numbering.xml'] = n.encode('utf8')
        ppr = ppr_of(self.p_num)
        ppr = re.sub(r'<w:numId w:val="\d+"/>', f'<w:numId w:val="{nid}"/>', ppr)
        ppr = re.sub(r'<w:ind [^>]*/>', '<w:ind w:left="709" w:firstLine="284"/>', ppr)
        ppr = re.sub(r'<w:spacing w:before="\d+"', '<w:spacing w:before="60"', ppr)
        rpr = first_rpr(self.p_num)
        return ''.join(f'<w:p>{ppr}{self.runs(t, rpr)}</w:p>' for t in items)

    # --- таблиці ---
    @staticmethod
    def fill_cell(tc, lines):
        ps = re.findall(r'<w:p[ >].*?</w:p>', tc, re.S)
        proto = ps[0]
        ppr = ppr_of(proto)
        rpr = first_rpr(proto)
        new = ''.join(f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{esc(l)}</w:t></w:r></w:p>'
                      for l in lines) or f'<w:p>{ppr}</w:p>'
        start = tc.find(ps[0])
        end = tc.rfind(ps[-1]) + len(ps[-1])
        return tc[:start] + new + tc[end:]

    def fill_info(self, meta):
        t = self.els[self.i_info]
        rows = re.findall(r'<w:tr[ >].*?</w:tr>', t, re.S)
        labels = [text_of(r).strip() for r in rows]

        def row_after(label_start):
            for k, lab in enumerate(labels):
                if lab.startswith(label_start):
                    return k + 1
            raise ValueError('Не знайдено рядок ' + label_start)

        def set_row(k, cols):
            cells = re.findall(r'<w:tc>.*?</w:tc>', rows[k], re.S)
            r = rows[k]
            for c, lines in zip(cells, cols):
                if lines is None:
                    continue
                r = r.replace(c, self.fill_cell(c, lines), 1)
            return r

        new_rows = list(rows)
        k = row_after('Скаржник')
        new_rows[k] = set_row(k, [meta['скаржник'], None])
        k = row_after('Дисциплінарна скарга')
        new_rows[k] = set_row(k, [meta['скарга'], meta['суддя']])
        k = row_after('Пропозиція')
        new_rows[k] = set_row(k, [meta['пропозиція']])
        if any(l.startswith('Акт про призначення') for l in labels):
            k = row_after('Акт про призначення')
            cells_k = re.findall(r'<w:tc>.*?</w:tc>', rows[k], re.S)
            cols = [None] * (len(cells_k) - 1) + [meta.get('акт') or ['']]
            new_rows[k] = set_row(k, cols)
        for a, b in zip(rows, new_rows):
            t = t.replace(a, b, 1)
        self.els[self.i_info] = t

    def fill_date(self, date):
        day, month, year = date.split()[:3]
        t = self.els[self.i_date]
        cells = re.findall(r'<w:tc>.*?</w:tc>', t, re.S)
        vals = {1: day, 3: month, 4: year}
        for k, v in vals.items():
            t = t.replace(cells[k], self.fill_cell(cells[k], [v]), 1)
        self.els[self.i_date] = t

    def build(self, md):
        meta = {'скаржник': [], 'скарга': [], 'суддя': [], 'пропозиція': [], 'дата': [], 'акт': []}
        body_lines = []
        for line in md.splitlines():
            m = re.match(r'@(\w+):\s*(.*)$', line)
            if m and m.group(1) in meta:
                meta[m.group(1)].append(m.group(2).strip())
            else:
                body_lines.append(line)
        if meta['дата']:
            self.fill_date(meta['дата'][0])
        self.fill_info(meta)
        blocks = [b.strip() for b in re.split(r'\n\s*\n', '\n'.join(body_lines)) if b.strip()]
        out = []
        for b in blocks:
            lines = [x.strip() for x in b.splitlines() if x.strip()]
            if lines and all(x.startswith('- ') for x in lines):
                out.append(self.sublist([x[2:].strip() for x in lines]))
                continue
            b = ' '.join(lines)
            if b.startswith('## '):
                if self.p_sub:
                    out.append(keep_next(self.para(self.p_sub, b[3:])))
                else:  # у шаблоні немає підзаголовка — жирний курсив без номера
                    ppr = re.sub(r'<w:numPr>.*?</w:numPr>', '', ppr_of(self.p_num), flags=re.S)
                    rp = rpr_with(first_rpr(self.p_num), italic=True, bold=True)
                    out.append(keep_next(f'<w:p>{ppr}<w:r>{rp}<w:t xml:space="preserve">{esc(b[3:])}</w:t></w:r></w:p>'))
            elif b.startswith('# '):
                out.append(keep_next(self.para(self.p_head, b[2:])))
            elif b.startswith('> '):
                out.append(self.quote(b[2:]))
            elif b.startswith('='):
                out.append(self.plain(b[1:].strip()))
            else:
                out.append(self.para(self.p_num, b))
        pre, post = self.els[:self.i_h1], self.els[self.i_post:]
        doc = self.head + ''.join(pre) + ''.join(out) + ''.join(post) + self.tail
        self.files['word/document.xml'] = doc.encode('utf8')
        self.files['word/footnotes.xml'] = self.fn_xml.encode('utf8')
        if 'docProps/core.xml' in self.files:
            self.files['docProps/core.xml'] = set_author(self.files['docProps/core.xml'].decode('utf8')).encode('utf8')

    def save(self, path):
        for n in ('word/document.xml', 'word/footnotes.xml'):
            if n in self.files:
                self.files[n] = highlight_gaps(self.files[n].decode('utf8')).encode('utf8')
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            for n in self.z.namelist():
                z.writestr(n, self.files[n])


def main():
    tpl, md, out = sys.argv[1:4]
    d = Doc(tpl)
    d.build(open(md, encoding='utf8').read())
    d.save(out)
    import xml.dom.minidom
    with zipfile.ZipFile(out) as z:
        for n in ('word/document.xml', 'word/footnotes.xml', 'docProps/core.xml'):
            if n in z.namelist():
                xml.dom.minidom.parseString(z.read(n))
    print('OK', out)


if __name__ == '__main__':
    main()
