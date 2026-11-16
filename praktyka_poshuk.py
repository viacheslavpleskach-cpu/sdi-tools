#!/usr/bin/env python3
"""Пошук позицій у рішеннях ВРП і Дисциплінарних палат та в постановах ВП ВС.

Автор: В'ячеслав Плескач. Версія 1.2 (16.11.2026): практика — у публічній теці Диска
https://drive.google.com/drive/folders/1u3pg4mfNgqvmp7h_WagnwXUeZWomCBzG («Рішення ВРП в DOC», «постанови_ВП_ВС в DOCX»
з підтеками «До 16.05.2026» і «Після 16.05.2026»); постанови ВП ВС з 05.2026 тепер теж мають файли на Диску;
колонка edrsr реєстру ВП ВС — номер документа в ЄДРСР (для посилання reyestr.court.gov.ua/Review/<номер>).
Версія 1.1 (14.11.2026): постанови без файлу на Диску, але з позначкою «ЄДРСР <doc_id>»
у колонці title (нові постанови з дайджесту ВП ВС), завантажуються з ЄДРСР через edrsr_vybirka.py (має лежати поруч);
шляхи до реєстрів — параметрами --reg / --reg-vpvs.

Використання:
    python3 praktyka_poshuk.py "фраза або регулярний вираз" [--vpvs] [--folder 3дп|2дп|1дп|залишення|зміна|скасування]
                                [--qual "п. 2"] [--from 2025] [--topic T02] [--max 15] [--context 1]
Приклади:
    python3 praktyka_poshuk.py "суддівськ\\w* помилк" --from 2024
    python3 praktyka_poshuk.py "навантаженн" --qual "п. 2" --folder залишення
    python3 praktyka_poshuk.py "навантаженн" --vpvs --topic T02 --reg-vpvs ../postanovi_vp-vs-monitoring/indeks/Реєстр_постанов_ВП_ВС.csv

Реєстр рішень ВРП — «Реєстр_рішень_ВРП.csv» (sdi-tools, поруч зі скриптом). Реєстр постанов ВП ВС —
«Реєстр_постанов_ВП_ВС.csv» з репозиторію postanovi_VP-VS-monitoring (тека indeks; його щодня оновлює дайджест).
Скрипт завантажує лише ті тексти, які проходять фільтри, кешує їх у ./_cache і друкує
нумеровані пункти, де трапляється фраза, з реквізитами для цитування.
"""
import argparse, csv, datetime as dt, html, os, re, subprocess, sys, zipfile
import concurrent.futures as cf

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, '_cache')

# --- постанови ВП ВС: DOCX, конвертовані з PDF (рядок = run без пробілу) ---
L = r'[A-Za-zА-ЯІЇЄҐа-яіїєґ’\'`ʼ0-9]'


def convert(path):
    x = zipfile.ZipFile(path).read('word/document.xml').decode('utf8')
    out = []
    for p in re.findall(r'<w:p[ >].*?</w:p>', x, re.S):
        runs = []
        for r in re.findall(r'<w:r[ >].*?</w:r>', p, re.S):
            t = ''.join(re.findall(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', r))
            if '<w:tab/>' in r:
                t = ' ' + t
            if t:
                runs.append(html.unescape(t))
        s = ''
        for t in runs:
            if s:
                if re.search(r'[А-ЯІЇЄҐа-яіїєґ]-$', s) and re.match(r'[а-яіїєґ]', t):
                    s = s[:-1]                      # перенос слова
                elif re.search(L + r'$', s) and re.match(L, t):
                    s += ' '
                elif re.search(r'[,.;:)»”]$', s) and re.match(L + r'|[«(“]', t):
                    s += ' '
            s += t
        s = re.sub(r'\s+', ' ', s).strip()
        if s:
            out.append(s)
    return '\n\n'.join(out)


def edrsr_id(row):
    m = re.match(r'ЄДРСР\s+(\d+)', row.get('title') or '')
    return m.group(1) if m else (row.get('edrsr') or None)


def prefetch_edrsr(rows):
    """Завантажує тексти постанов з ЄДРСР (за doc_id) одним запуском edrsr_vybirka.py; кешує _cache/edrsr_<id>.txt."""
    need = [r for r in rows if not r.get('id') and edrsr_id(r) and
            not os.path.exists(os.path.join(CACHE, f'edrsr_{edrsr_id(r)}.txt'))]
    if not need:
        return
    ev = os.path.join(HERE, 'edrsr_vybirka.py')
    if not os.path.exists(ev):
        print('Немає edrsr_vybirka.py поруч — постанови з позначкою «ЄДРСР» пропущено (запустіть get.sh).')
        return
    years = sorted({int(r['date'][-4:]) for r in need})
    tmp = os.path.join(CACHE, '_edrsr')
    os.makedirs(tmp, exist_ok=True)
    ids = ','.join(edrsr_id(r) for r in need)
    tsv = os.path.join(tmp, 'docs.tsv')
    subprocess.run([sys.executable, ev, 'doc', '--id', ids, '--years', f'{years[0]}-{min(years[-1] + 1, dt.date.today().year)}', '--out', tsv],
                   check=True, stderr=subprocess.DEVNULL)
    subprocess.run([sys.executable, ev, 'texts', '--in', tsv, '--dir', tmp], check=True, stderr=subprocess.DEVNULL)
    for r in need:
        i = edrsr_id(r)
        src = os.path.join(tmp, i + '.txt')
        if os.path.exists(src):
            lines = [re.sub(r'\s+', ' ', l).strip() for l in open(src, encoding='utf-8', errors='replace')]
            open(os.path.join(CACHE, f'edrsr_{i}.txt'), 'w', encoding='utf-8').write('\n\n'.join(l for l in lines if l))


def text_of(row):
    os.makedirs(CACHE, exist_ok=True)
    if not row.get('id'):
        i = edrsr_id(row)
        p = os.path.join(CACHE, f'edrsr_{i}.txt') if i else ''
        return open(p, encoding='utf-8').read() if p and os.path.exists(p) else ''
    t = os.path.join(CACHE, row['id'] + '.txt')
    if os.path.exists(t):
        return open(t, encoding='utf-8', errors='replace').read()
    ext = os.path.splitext(row['title'])[1].lower()
    f = os.path.join(CACHE, row['id'] + ext)
    subprocess.run(['curl', '-sS', '-L', '-m', '120', '-o', f,
                    f"https://drive.usercontent.google.com/download?id={row['id']}&export=download&confirm=t"], check=True)
    if row.get('_vpvs'):
        s = convert(f)
    elif ext == '.docx':
        s = subprocess.run(['pandoc', f, '-t', 'plain', '--wrap=none'], capture_output=True).stdout.decode('utf-8', 'replace')
    elif ext == '.pdf':
        s = subprocess.run(['pdftotext', '-layout', f, '-'], capture_output=True).stdout.decode('utf-8', 'replace')
    else:
        subprocess.run(['soffice', '--headless', '--convert-to', 'txt:Text (encoded):UTF8', '--outdir', CACHE, f], capture_output=True)
        p = os.path.splitext(f)[0] + '.txt'
        s = open(p, encoding='utf-8', errors='replace').read() if os.path.exists(p) else ''
    open(t, 'w', encoding='utf-8').write(s)
    return s


def paragraphs(s):
    """[(номер пункту або '', текст)] — пункти нумерованих списків pandoc: '12.  Текст'."""
    out, cur, num = [], [], ''
    for line in s.splitlines():
        m = re.match(r'^\s{0,4}(\d{1,3})\.\s+(\S.*)$', line)
        if m:
            if cur:
                out.append((num, ' '.join(cur)))
            num, cur = m.group(1), [m.group(2)]
        elif line.strip():
            if cur and line.startswith('    '):
                cur.append(line.strip())
            else:
                if cur:
                    out.append((num, ' '.join(cur)))
                num, cur = '', [line.strip()]
        else:
            if cur:
                out.append((num, ' '.join(cur))); num, cur = '', []
    if cur:
        out.append((num, ' '.join(cur)))
    return out


def paragraphs_vpvs(s):
    out = []
    for p in s.split('\n\n'):
        m = re.match(r'^(\d{1,3})\.\s+(.*)', p, re.S)
        out.append((m.group(1), m.group(2)) if m else ('', p))
    return out


def cite(row):
    if row.get('_vpvs'):
        pr = f" (провадження № {row['provadzhennia']})" if row.get('provadzhennia') else ''
        i = edrsr_id(row)
        return f"ВП ВС, постанова від {row['date']} у справі № {row['case']}{pr}" + (f' [ЄДРСР {i}]' if i else '')
    org = {'ВРП': 'ВРП, рішення', '1дп': '1ДП, рішення', '2дп': '2ДП, рішення', '3дп': '3ДП, рішення'}.get(row['org'], row['org'])
    return f"{org} від {row['date']} № {row.get('number', '')}".strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pattern'); ap.add_argument('--folder'); ap.add_argument('--qual')
    ap.add_argument('--from', dest='yfrom', type=int); ap.add_argument('--max', type=int, default=15)
    ap.add_argument('--context', type=int, default=0)
    ap.add_argument('--vpvs', action='store_true', help='шукати в постановах ВП ВС')
    ap.add_argument('--topic', help='код теми: T01…T14 (див. Індекс), фільтр за колонкою topics')
    ap.add_argument('--limit', type=int, default=120, help='скільки найсвіжіших рішень переглянути')
    ap.add_argument('--reg', default=os.path.join(HERE, 'Реєстр_рішень_ВРП.csv'))
    ap.add_argument('--reg-vpvs', default=os.path.join(HERE, 'Реєстр_постанов_ВП_ВС.csv'))
    a = ap.parse_args()
    if a.vpvs:
        if not os.path.exists(a.reg_vpvs):
            sys.exit(f'Немає {a.reg_vpvs}. Підключіть postanovi_VP-VS-monitoring і передайте --reg-vpvs <клон>/indeks/Реєстр_постанов_ВП_ВС.csv')
        rows = list(csv.DictReader(open(a.reg_vpvs, encoding='utf-8')))
        for r in rows:
            r['_vpvs'] = 1; r.setdefault('folder', ''); r['qual_106'] = r.get('qual_106') or ''
            r['subject'] = r.get('subject', ''); r['title'] = r.get('title') or 'x.docx'
    else:
        rows = list(csv.DictReader(open(a.reg, encoding='utf-8')))
    if a.folder:
        rows = [r for r in rows if a.folder in r['folder']]
    if a.qual:
        rows = [r for r in rows if a.qual in r.get('qual_106', '')]
    if a.yfrom:
        rows = [r for r in rows if r['date'][-4:].isdigit() and int(r['date'][-4:]) >= a.yfrom]
    if a.topic:
        rows = [r for r in rows if a.topic in (r.get('topics') or '').split()]
    rows = [r for r in rows if 'окремі' not in r.get('folder', '').lower() and 'прокурор' not in r.get('folder', '').lower()]
    no_src = [r for r in rows if not r.get('id') and not edrsr_id(r)]
    if no_src:
        print(f"Пропущено {len(no_src)} рядків без джерела тексту: "
              + ', '.join(r.get('title') or r.get('case', '') for r in no_src[:5]) + (' …' if len(no_src) > 5 else ''))
    rows = [r for r in rows if r.get('id') or edrsr_id(r)]
    rows.sort(key=lambda r: r['date'][-4:] + r['date'][3:5] + r['date'][:2], reverse=True)
    rows = rows[:a.limit]
    os.makedirs(CACHE, exist_ok=True)
    prefetch_edrsr(rows)
    rx = re.compile(a.pattern, re.I)
    hits = 0
    with cf.ThreadPoolExecutor(8) as ex:   # паралельне завантаження з Диска
        list(ex.map(text_of, [r for r in rows[:40] if r.get('id')]))
    for r in rows:
        ps = paragraphs_vpvs(text_of(r)) if r.get('_vpvs') else paragraphs(text_of(r))
        found = [i for i, (_, p) in enumerate(ps) if rx.search(p)]
        if not found:
            continue
        print(f"\n=== {cite(r)} — {r.get('about') or r.get('subject', '')} [{r['result']}; {r.get('qual_106', '')}] (id {r.get('id') or edrsr_id(r)})")
        for i in found[:3]:
            for j in range(max(0, i - a.context), min(len(ps), i + a.context + 1)):
                n, p = ps[j]
                print(f"  п. {n or '—'}: {p[:1200]}")
        hits += 1
        if hits >= a.max:
            break
    print(f'\nЗнайдено в {hits} рішеннях (переглянуто до {len(rows)}).')


if __name__ == '__main__':
    main()
