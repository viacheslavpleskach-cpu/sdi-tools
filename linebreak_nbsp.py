#!/usr/bin/env python3
"""linebreak_nbsp.py — нерозривні пробіли лише в місцях невдалих розривів рядка (Довідник редагування ВРП, розд. 12).

Автор: В'ячеслав Плескач.

  python3 linebreak_nbsp.py <вхідний.docx> <вихідний.docx> [--max-iter 6] [--dry-run]

Порядок (рішення користувача від 21.10.2026):
  * завжди — нерозривний пробіл між «а. с.» і номером аркуша, між «т.» і номером тому;
  * лише в місці невдалого розриву рядка:
      1) однолітерне слово (а, в, з, і, й, о, у) у кінці рядка;
      2) ініціали перенесено окремо від прізвища («Петров / І.І.»);
      3) дата словами розірвана як «23 / жовтня 2026 року» або «23 жовтня 2026 / року»;
      4) «№» у кінці рядка окремо від номера.
Документ рендериться (LibreOffice → PDF) з прийнятими правками; у кожному знайденому місці
замінюється ОДИН пробіл (а не всі схожі фрагменти), документ рендериться знову — до --max-iter разів.
Наявні нерозривні пробіли не змінюються; правки вносяться без рецензування (розд. 12).
Верстка Word може трохи відрізнятися від LibreOffice, тому поодинокі розриви можливі.
"""
import argparse, copy, io, os, re, shutil, subprocess, sys, tempfile, zipfile
from lxml import etree as ET   # lxml зберігає оголошення просторів імен (mc:Ignorable)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = '{%s}' % W
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
MON = 'січня|лютого|березня|квітня|травня|червня|липня|серпня|вересня|жовтня|листопада|грудня'
ONE = set('аАвВзЗіІйЙоОуУ')


# ---------------------------------------------------------------- текст абзаців і вузли w:t
def para_nodes(root):
    """Для кожного абзацу: список (вузол w:t, початок у тексті абзацу) без видалених фрагментів."""
    res = []
    def walk(el, acc, pos):
        for ch in el:
            if not isinstance(ch.tag, str):
                continue
            if ch.tag in (NS + 'del', NS + 'moveFrom'):
                continue
            if ch.tag == NS + 'p':
                continue
            if ch.tag == NS + 't':
                acc.append((ch, pos[0]))
                pos[0] += len(ch.text or '')
            elif ch.tag == NS + 'tab':
                acc.append((None, pos[0])); pos[0] += 1
            elif ch.tag in (NS + 'br', NS + 'cr'):
                acc.append((None, pos[0])); pos[0] += 1
            elif ch.tag == NS + 'noBreakHyphen':
                acc.append((None, pos[0])); pos[0] += 1
            else:
                walk(ch, acc, pos)
    for p in root.iter(NS + 'p'):
        acc, pos = [], [0]
        walk(p, acc, pos)
        text = ''
        for node, start in acc:
            if node is None:
                text += ' '
            else:
                text += node.text or ''
        res.append((p, acc, text))
    return res

def set_char(acc, idx, ch):
    for node, start in acc:
        if node is None:
            continue
        t = node.text or ''
        if start <= idx < start + len(t):
            node.text = t[:idx - start] + ch + t[idx - start + 1:]
            node.set(XML_SPACE, 'preserve')
            return True
    return False

# ---------------------------------------------------------------- рендер
def accepted_copy(src, dst):
    """Копія DOCX із прийнятими правками (для рендеру)."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                s = data.decode('utf-8')
                s = re.sub(r'<w:del\b[^>]*/>', '', s)
                s = re.sub(r'<w:del\b[^>]*>.*?</w:del>', '', s, flags=re.S)
                s = re.sub(r'<w:moveFrom\b[^>]*>.*?</w:moveFrom>', '', s, flags=re.S)
                s = re.sub(r'</?w:(?:ins|moveTo)\b[^>]*>', '', s)
                data = s.encode('utf-8')
            zout.writestr(it, data)

def render_lines(docx):
    import pdfplumber
    d = tempfile.mkdtemp()
    acc = os.path.join(d, 'r.docx')
    accepted_copy(docx, acc)
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', d, acc],
                   check=True, capture_output=True, timeout=300)
    lines = []
    with pdfplumber.open(os.path.join(d, 'r.pdf')) as pdf:
        for pg in pdf.pages:
            for ln in pg.extract_text_lines(layout=False, strip=True):
                t = ln['text'].strip()
                if t and not re.fullmatch(r'\d{1,3}', t):   # номери сторінок
                    lines.append(t)
    shutil.rmtree(d, ignore_errors=True)
    return lines

# ---------------------------------------------------------------- пошук невдалих розривів
def norm(s):
    return re.sub(r'\s+', ' ', s.replace('\u00a0', ' '))

def normmap(text):
    """Нормалізований текст (послідовність пробілів → один пробіл) і відповідність індексів оригіналу."""
    out, imap, prev_sp = [], [], False
    for i, ch in enumerate(text):
        sp = ch in ' \u00a0\t'
        if sp and prev_sp:
            continue
        out.append(' ' if sp else ch); imap.append(i); prev_sp = sp
    return ''.join(out), imap

def bad_breaks(lines):
    out = []   # (опис, хвіст рядка, початок наступного, номер входження в потоці тексту)
    flow = ''
    for idx, (a, b) in enumerate(zip(lines, lines[1:])):
        flow += norm(a) + ' '

        wa, wb = a.split(), b.split()
        if not wa or not wb:
            continue
        last, first = wa[-1], wb[0]
        why = None
        if len(last) == 1 and last in ONE:
            why = '1: однолітерне слово в кінці рядка'
        elif re.fullmatch(r'[А-ЯІЇЄҐ]\.[А-ЯІЇЄҐ]\.[,;:.)»]?', first) and re.fullmatch(r'[А-ЯІЇЄҐ][а-яіїєґ’\'-]+', last):
            why = '2: ініціали окремо від прізвища'
        elif re.fullmatch(r'\d{1,2}', last) and re.match(rf'(?:{MON})\b', first):
            why = '3а: дата словами — число окремо від місяця'
        elif re.fullmatch(r'\d{4}', last) and re.match(r'року\b', first) and len(wa) >= 2 and re.fullmatch(MON, wa[-2]):
            why = '3б: дата словами — «року» окремо від року'
        elif last == '№':
            why = '4: «№» у кінці рядка'
        if why:
            tail = ' '.join(wa[-3:])
            head = ' '.join(wb[:2])
            key = norm(tail) + ' ' + norm(head)
            nth = (flow + norm(b)).count(key) - 1 if key in (flow + norm(b)) else 0
            out.append((why, tail, head, max(nth, 0)))
    return out

def norm(s):
    return re.sub(r'\s+', ' ', s.replace(' ', ' '))

def apply(docx_in, docx_out, breaks, always=True):
    with zipfile.ZipFile(docx_in) as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    paras = para_nodes(root)
    done, log = 0, []
    # а) завжди: «а. с.» + номер, «т.» + номер тому
    if always:
        for p, acc, text in paras:
            for m in re.finditer(r'а\. с\.( )(?=\d)|(?:(?<=\()|(?<=,\s)|(?<=;\s))т\.( )(?=\d)', text):
                i = m.start(1) if m.group(1) else m.start(2)
                if set_char(acc, i, ' '):
                    done += 1; log.append(('завжди: «а. с.» / «т.» + номер', text[max(0, i-12):i+6]))
        paras = para_nodes(root)
    for why, tail, head, nth in breaks:
        key = norm(tail) + ' ' + norm(head)
        occ = []   # (абзац, індекс пробілу)
        for pi, (p, acc, text) in enumerate(paras):
            nt, imap = normmap(text)
            j = nt.find(key)
            while j >= 0:
                k = j + len(norm(tail))          # пробіл між хвостом і наступним рядком (у нормалізованому тексті)
                a0 = imap[k]
                b0 = a0
                while b0 + 1 < len(text) and text[b0 + 1] in ' \u00a0':
                    b0 += 1
                occ.append((pi, a0, b0))
                j = nt.find(key, j + 1)
        if not occ:
            continue
        order = ([occ[nth]] if nth < len(occ) else []) + occ
        for pi, a0, b0 in order:
            p, acc, text = paras[pi]
            run = text[a0:b0 + 1]
            if ' ' in run:                       # є звичайний пробіл, на якому рядок рветься
                for i in range(b0, a0, -1):      # зайві пробіли вилучити (подвійний пробіл — технічна описка)
                    set_char(acc, i, '')
                set_char(acc, a0, '\u00a0')
                done += 1; log.append((why, (tail + ' ° ' + head)))
                break
        paras = para_nodes(root)
    with zipfile.ZipFile(docx_in) as zin, zipfile.ZipFile(docx_out + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                data = ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            zout.writestr(it, data)
    os.replace(docx_out + '.tmp', docx_out)
    return done, log

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('src'); ap.add_argument('dst')
    ap.add_argument('--max-iter', type=int, default=6)
    ap.add_argument('--dry-run', action='store_true', help='лише показати невдалі розриви')
    a = ap.parse_args()
    if a.dry_run:
        for why, tail, head, _n in bad_breaks(render_lines(a.src)):
            print(f'{why}: …{tail} / {head}…')
        return
    if os.path.abspath(a.src) != os.path.abspath(a.dst):   # дозволено запуск «на місці»
        shutil.copy(a.src, a.dst)
    total, first = 0, True
    for it in range(1, a.max_iter + 1):
        br = bad_breaks(render_lines(a.dst))
        if not br and not first:
            print(f'ітерація {it}: невдалих розривів немає'); break
        n, log = apply(a.dst, a.dst, br, always=first)
        first = False
        total += n
        print(f'ітерація {it}: знайдено розривів {len(br)}, замінено пробілів {n}')
        for why, frag in log:
            print(f'   {why}: …{frag}…')
        if n == 0:
            break
    left = bad_breaks(render_lines(a.dst))
    print(f'усього нерозривних пробілів: {total}; невдалих розривів лишилося: {len(left)}')
    for why, tail, head, _n in left:
        print(f'   не виправлено ({why}): …{tail} / {head}…')

if __name__ == '__main__':
    main()
