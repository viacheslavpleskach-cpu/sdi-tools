#!/usr/bin/env python3
"""Збирає DOCX висновку на скаргу на рішення Дисциплінарної палати ВРП і файлу рекомендацій до розділу IV.

Версія 1.1 (20.11.2026): дата без --date — з поточним роком. Автор у властивостях документа — «В'ячеслав Плескач».

Використання:
    python3 build_vysnovok_skarga.py ШАБЛОН.docx ЗМІСТ.txt РЕЗУЛЬТАТ.docx full|plain [--date "18 листопада 2026"]
    python3 build_vysnovok_skarga.py check ЗМІСТ.txt          # перевірка розмітки до збирання

full  — висновок: бланк (герб, «ВИЩА РАДА ПРАВОСУДДЯ», інспектор, дата, «ВИСНОВОК», назва з рядка T) + текст + підпис;
plain — рекомендації: лише текст + підпис (без бланка).
Дата без --date — ««___» ____________ <поточний рік> року».

ШАБЛОН — «Шаблон_висновку_скарга.docx» (перші 10 абзаців — бланк; стилі, нумерація numId=3 для наскрізних пунктів,
abstractNumId=5 для маркованих переліків).

Формат ЗМІСТ.txt (кожен блок — один рядок; порожні рядки ігноруються):
    T за результатами опрацювання скарги ...   -> назва під «ВИСНОВОК» (лише full)
    # І. ПРОЦЕДУРНІ ПИТАННЯ                    -> заголовок розділу (жирний)
    ## Підзаголовок                            -> жирний курсив, підкреслений
    ### Підпідзаголовок                        -> курсив, підкреслений
    P текст                                    -> наскрізно нумерований пункт (автонумерація Word)
    P                                          -> порожній нумерований пункт (розділ IV для інспектора)
    - текст                                    -> пункт переліку (маркер «–», нумерація переліку з 1 не потрібна)
    [[текст]]                                  -> виділення жовтим жирним (звірити / дозаповнити)
Апостроф ' -> ’, нерозривні пробіли після «№», «від» перед датою, між прізвищем та ініціалами — автоматично.
"""
import os, re, shutil, sys, tempfile, zipfile
from xml.sax.saxutils import escape

NB = ' '
AUTHOR = 'В’ячеслав Плескач'


def typo(t):
    t = t.replace("'", '’')
    t = re.sub(r'№ ', '№' + NB, t)
    t = re.sub(r'(?<![А-Яа-яІіЇїЄєҐґ])від (?=\d)', 'від' + NB, t)
    t = re.sub(r'([А-ЯІЇЄҐ][а-яіїєґ’\-]+) ([А-ЯІЇЄҐ]\.[А-ЯІЇЄҐ]\.)', r'\1' + NB + r'\2', t)
    return t


def rpr(pre='', post='', hl=False):
    return ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            + pre + '<w:sz w:val="28"/><w:szCs w:val="28"/>'
            + ('<w:highlight w:val="yellow"/>' if hl else '') + post
            + '<w:lang w:val="uk-UA"/></w:rPr>')


def runs(text, pre='', post=''):
    out = []
    for i, part in enumerate(re.split(r'\[\[(.*?)\]\]', typo(text))):
        if not part:
            continue
        hl = (i % 2 == 1)
        b = '<w:b/><w:bCs/>' if hl and '<w:b/>' not in pre else ''
        out.append('<w:r>' + rpr(pre + b, post, hl) + '<w:t xml:space="preserve">' + escape(part) + '</w:t></w:r>')
    return ''.join(out)


SP = '<w:spacing w:before="160" w:line="240" w:lineRule="auto"/>'


def para(text, kind, num_id=None):
    if kind == 'H':
        pre = '<w:b/><w:bCs/>'
        return '<w:p><w:pPr><w:keepNext/>' + SP + '<w:jc w:val="both"/>' + rpr(pre) + '</w:pPr>' + runs(text, pre) + '</w:p>'
    if kind in ('S', 'SS'):
        pre = '<w:b/><w:bCs/><w:i/><w:iCs/>' if kind == 'S' else '<w:i/><w:iCs/>'
        post = '<w:u w:val="single"/>'
        return ('<w:p><w:pPr><w:pStyle w:val="a3"/><w:keepNext/>' + SP +
                '<w:ind w:left="567"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/>' + rpr(pre, post) + '</w:pPr>'
                + runs(text, pre, post) + '</w:p>')
    if kind == 'P':
        return ('<w:p><w:pPr><w:pStyle w:val="a3"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>' + SP +
                '<w:ind w:left="567" w:hanging="567"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/>'
                + rpr('<w:bCs/>') + '</w:pPr>' + runs(text, '<w:bCs/>') + '</w:p>')
    if kind == 'L':
        return ('<w:p><w:pPr><w:pStyle w:val="a3"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % num_id + SP +
                '<w:ind w:left="927" w:hanging="360"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/>'
                + rpr('<w:bCs/>') + '</w:pPr>' + runs(text, '<w:bCs/>') + '</w:p>')
    raise ValueError(kind)


def parse(src):
    blocks = []
    for n, line in enumerate(open(src, encoding='utf8').read().split('\n'), 1):
        line = line.rstrip()
        if not line:
            continue
        if line == 'P':
            blocks.append(('P', ''))
            continue
        for pref, k in (('### ', 'SS'), ('## ', 'S'), ('# ', 'H'), ('P ', 'P'), ('- ', 'L'), ('T ', 'T')):
            if line.startswith(pref):
                blocks.append((k, line[len(pref):]))
                break
        else:
            raise SystemExit('рядок %d: невідома розмітка: %s' % (n, line[:70]))
    return blocks


def check(src):
    blocks = parse(src)
    text = '\n'.join(t for _, t in blocks)
    bad = text.count('[[') - text.count(']]')
    print('блоків:', len(blocks), '| пунктів P:', sum(k == 'P' for k, _ in blocks),
          '| виділень [[…]]:', len(re.findall(r'\[\[', text)), '| незакритих [[:', bad)
    first = [i for i, (k, _) in enumerate(blocks) if k == 'L']
    orphan = [i for i in first if i == 0 or blocks[i - 1][0] not in ('P', 'L')]
    if orphan:
        print('перелік без вступного пункту P перед ним (блоки):', orphan[:10])
    for ph in ('цього Висновку', 'суддя суддя', '  '):
        if ph in text:
            print('увага: трапляється «%s»' % ph)
    return 0 if bad == 0 else 2


def build(template, src, out, mode, date=None):
    blocks = parse(src)
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(template) as z:
        z.extractall(tmp)
    dp = os.path.join(tmp, 'word/document.xml')
    doc = open(dp, encoding='utf8').read()
    head = doc[:doc.index('<w:body>') + len('<w:body>')]
    sect = doc[doc.index('<w:sectPr'):doc.index('</w:body>')]
    ps = re.findall(r'<w:p [^>]*>.*?</w:p>', doc[len(head):], re.S)[:10]

    body = []
    if mode == 'full':
        def retext(p, text):
            ppr = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S).group(0)
            r = re.search(r'<w:r(?: [^>]*)?>(<w:rPr>.*?</w:rPr>)', p, re.S).group(1)
            return '<w:p>' + ppr + '<w:r>' + r + '<w:t xml:space="preserve">' + escape(text) + '</w:t></w:r></w:p>'
        body += ps[0:6]
        d = ps[6]
        r = re.search(r'<w:r(?: [^>]*)?>(<w:rPr>.*?</w:rPr>)', d, re.S).group(1)
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', d, re.S).group(0)
        if date:
            dd, mm, yy = date.split()[:3]
            left = '«%s» %s %s року' % (dd, mm, yy)
        else:
            left = '«___» ____________ %d року' % __import__('datetime').date.today().year
        tabs = ''.join(re.findall(r'<w:tab/>', d)) or '<w:tab/>'
        body.append('<w:p>' + ppr + '<w:r>' + r + '<w:t xml:space="preserve">' + escape(left) + '</w:t></w:r>'
                    + '<w:r>' + r + tabs + '<w:t>місто Київ</w:t></w:r></w:p>')
        body += [ps[7], ps[8]]
        titles = [t for k, t in blocks if k == 'T']
        if not titles:
            raise SystemExit('для full потрібен рядок «T …» з назвою висновку')
        body.append(retext(ps[9], typo(titles[0])))

    new_nums, next_id, prev, cur = [], 100, None, None
    for kind, text in blocks:
        if kind == 'T':
            continue
        if kind == 'L':
            if prev != 'L':
                cur = next_id
                next_id += 1
                new_nums.append('<w:num w:numId="%d"><w:abstractNumId w:val="5"/><w:lvlOverride w:ilvl="0">'
                                '<w:startOverride w:val="1"/></w:lvlOverride></w:num>' % cur)
            body.append(para(text, 'L', cur))
        else:
            body.append(para(text, kind))
        prev = kind

    sig = ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="Calibri" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
           '<w:b/><w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="uk-UA"/></w:rPr>')
    sppr = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="567"/><w:jc w:val="both"/>' + sig + '</w:pPr>'
    body.append('<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>')
    body.append('<w:p>' + sppr + '<w:r>' + sig + '<w:t xml:space="preserve">Дисциплінарний інспектор </w:t></w:r></w:p>')
    body.append('<w:p>' + sppr + '<w:r>' + sig + '<w:t>Вищої ради правосуддя</w:t></w:r>'
                + ('<w:r>' + sig + '<w:tab/></w:r>') * 4 + '<w:r>' + sig + '<w:tab/><w:t>В’ячеслав ПЛЕСКАЧ</w:t></w:r></w:p>')

    open(dp, 'w', encoding='utf8').write(head + ''.join(body) + sect + '</w:body></w:document>')
    np_ = os.path.join(tmp, 'word/numbering.xml')
    nx = open(np_, encoding='utf8').read().replace('</w:numbering>', ''.join(new_nums) + '</w:numbering>')
    open(np_, 'w', encoding='utf8').write(nx)
    fp = os.path.join(tmp, 'word/footnotes.xml')
    if os.path.exists(fp):
        fn = open(fp, encoding='utf8').read()
        fn = re.sub(r'<w:footnote (?![^>]*w:type=)[^>]*>.*?</w:footnote>', '', fn, flags=re.S)
        open(fp, 'w', encoding='utf8').write(fn)
    cp = os.path.join(tmp, 'docProps/core.xml')
    if os.path.exists(cp):
        c = open(cp, encoding='utf8').read()
        c = re.sub(r'(<dc:creator>)[^<]*(</dc:creator>)', r'\g<1>' + AUTHOR + r'\2', c)
        c = re.sub(r'(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)', r'\g<1>' + AUTHOR + r'\2', c)
        open(cp, 'w', encoding='utf8').write(c)
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        order = ['[Content_Types].xml']
        for root, _, files in os.walk(tmp):
            for f in files:
                p = os.path.relpath(os.path.join(root, f), tmp)
                if p not in order:
                    order.append(p)
        for p in order:
            z.write(os.path.join(tmp, p), p)
    shutil.rmtree(tmp)
    print('блоків:', len(blocks), '| пунктів:', sum(1 for k, _ in blocks if k == 'P'), '->', out)


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'check':
        sys.exit(check(a[1]))
    date = None
    if '--date' in a:
        i = a.index('--date')
        date = a[i + 1]
        del a[i:i + 2]
    if len(a) != 4 or a[3] not in ('full', 'plain'):
        raise SystemExit(__doc__)
    build(a[0], a[1], a[2], a[3], date)
