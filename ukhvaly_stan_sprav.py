#!/usr/bin/env python3
"""Аналіз руху справ судді за всіма документами ЄДРСР (рішення + ухвали).

Автор: В'ячеслав Плескач. Версія 1.1 (14.11.2026): дата оцінки END — параметр (раніше була жорстко 15.12.2022,
день ліквідації ОАСК).

1) edrsr_vybirka.py meta --court <код> --judge "<Суддя І. Б.>" (усі форми) → TSV; texts → TXT.
2) python3 -I ukhvaly_stan_sprav.py uhv   <tsv> <txtdir> <kinds.pkl>
   — вид кожної ухвали за резолютивною частиною (OPEN, CLOSE, LEAVE, RETURN, SUSPEND, EVIDENCE …).
3) python3 -I ukhvaly_stan_sprav.py cases <kinds.pkl> <усі_документи_суду.pkl> "<Суддя>" <cases.pkl> [END РРРР-ММ-ДД]
   — стан кожної справи на END: завершена / передана іншому судді / незавершена; днів без документів.
   END — дата, станом на яку оцінюється стан (типово — сьогодні; для ОАСК — 2022-12-15).
   <усі_документи_суду.pkl> — pandas DataFrame з колонками judge, cause_num, adjudication_date (meta по суду).
Незавершена справа: суддя відкрив/прийняв провадження, а до END немає підсумкового акта
(рішення, закриття провадження, залишення позову без розгляду, повернення, відмова у відкритті,
передача за підсудністю) або після нього провадження продовжено.
Перевірка: такі справи мають з'являтися в документах інших судів у наступних роках.
"""
import sys, os, re, html
import pandas as pd


RES = re.compile(r'У\s?Х\s?В\s?А\s?Л\s?И\s?В|П\s?О\s?С\s?Т\s?А\s?Н\s?О\s?В\s?И\s?В|В\s?И\s?Р\s?І\s?Ш\s?И\s?В|ухвалив\s*:|постановив\s*:|вирішив\s*:')


def clean(t):
    if re.search(r'<(html|meta|body|!--)', t[:5000], re.I):
        t = re.sub(r'(?is)<(style|script|head)[^>]*>.*?</\1>', ' ', t)
        t = re.sub(r'(?s)<!--.*?-->', ' ', t)
        t = re.sub(r'<[^>]+>', ' ', t)
        t = html.unescape(t)
    t = t.replace('`', "'").replace('’', "'").replace('ʼ', "'")
    return re.sub(r'\s+', ' ', t)


def resol(t):
    ms = list(RES.finditer(t))
    if ms:
        return t[ms[-1].end():ms[-1].end() + 1500]
    return t[-1200:]


def kind(r):
    s = r.lower()
    head = s[:600]
    if re.search(r'відмовити\s+(у|в)\s+відкритті|(у|в)\s+відкритті\s+провадження[^.]{0,300}відмовити', head): return 'REFUSE_OPEN'
    if re.search(r'передати[^.]{0,160}(за\s+підсудністю|за\s+територіальною)|направити[^.]{0,120}за\s+(територіальною\s+)?підсудністю', head): return 'TRANSFER'
    if re.search(r'закрити\s+провадження', head) and 'підготовче' not in head[:80]: return 'CLOSE'
    if re.search(r'(позов\w*\s+заяв\w*|позов|адміністративн\w+\s+справ\w*)[^.]{0,150}залишити\s+без\s+розгляду|залишити\s+без\s+розгляду[^.]{0,60}(позов|адміністративн\w+\s+справ)', head) and not re.search(r'забезпеч', head[:250]):
        return 'LEAVE'
    if re.search(r'(позовну\s+заяву|позов)[^.]{0,80}повернути|повернути[^.]{0,80}позовн', head): return 'RETURN'
    if re.search(r'відкрити\s+(\w+\s+){0,3}провадження|відкрити\s+провадження|прийняти\s+позовну\s+заяву\s+до\s+розгляду', head) and not re.search(r'апеляц', head[:200]): return 'OPEN'
    if re.search(r'прийняти\s+(адміністративну\s+)?справу[^.]{0,60}до\s+(свого\s+)?провадження|прийняти\s+до\s+(свого\s+)?провадження', head): return 'ACCEPT'
    if 'без руху' in head: return 'NO_MOTION'
    if re.search(r'зупинити\s+провадження', head): return 'SUSPEND'
    if re.search(r'поновити\s+провадження', head): return 'RESUME'
    if re.search(r"об'єднати", head): return 'JOIN'
    if re.search(r'відвід|самовідвід', head): return 'RECUSAL'
    if re.search(r'витребувати|зобов\'язати[^.]{0,80}(надати|подати)', head): return 'EVIDENCE'
    if re.search(r'загального\s+позовного\s+провадження|спрощеного\s+позовного', head): return 'PROC_FORM'
    if re.search(r'призначити[^.]{0,80}(засідання|розгляд)|закрити\s+підготовче', head): return 'HEARING'
    if re.search(r'описк|арифметичн|роз\'ясн', head): return 'POST'
    if re.search(r'судов\w+\s+контрол|звіт|виконавч\w+\s+(лист|документ)|замін\w+\s+(сторон|стягувач|боржник)', head): return 'ENFORCE'
    if re.search(r'забезпеч', head): return 'SECURE'
    if re.search(r'апеляційн', head): return 'APPEAL'
    if re.search(r'залучити|третьої\s+особи|третьою\s+особою', head): return 'PARTY'
    if re.search(r'клопотанн|заяв', head): return 'MOTION'
    return 'OTHER'


def run_uhv(argv):
    tsv, d, out = argv[0:3]
    df = pd.read_csv(tsv, sep='\t', dtype=str)
    kinds, snips = [], []
    for _, r in df.iterrows():
        p = os.path.join(d, r.doc_id + '.txt')
        if r.judgment_form == 'Рішення':
            k = 'DECISION'; sn = ''
            if os.path.exists(p):
                sn = resol(clean(open(p, encoding='utf-8', errors='ignore').read()))[:300]
        elif not os.path.exists(p):
            k = 'NOTEXT'; sn = ''
        else:
            ct = clean(open(p, encoding='utf-8', errors='ignore').read())
            if len(ct) < 300:
                k = 'NOTEXT'; sn = ''
            else:
                rr = resol(ct); k = kind(rr); sn = rr[:300]
        kinds.append(k); snips.append(sn)
    df['kind'] = kinds; df['res'] = snips
    df.to_pickle(out)
    print(df.kind.value_counts())


def run_cases(argv):
    FINAL = {'DECISION', 'CLOSE', 'LEAVE', 'TRANSFER', 'REFUSE_OPEN', 'RETURN'}
    START = {'OPEN', 'ACCEPT'}
    POSTK = {'POST', 'ENFORCE', 'APPEAL', 'NOTEXT'}
    END = pd.Timestamp(argv[4]) if len(argv) > 4 else pd.Timestamp.today().normalize()
    print('Стан на', END.date())

    k = pd.read_pickle(argv[0]); allc = pd.read_pickle(argv[1]); judge = argv[2]
    k['d'] = pd.to_datetime(k.adjudication_date)
    k = k[k.d <= END]
    allc['d'] = pd.to_datetime(allc.adjudication_date, errors='coerce')
    oth = allc[(allc.judge != judge) & allc.cause_num.isin(k.cause_num.unique())]
    othlast = oth.groupby('cause_num').d.max()
    rows = []
    for c, g in k.sort_values(['d', 'doc_id']).groupby('cause_num'):
        st = g[g.kind.isin(START)]
        start = st.d.min() if len(st) else pd.NaT
        fin = g[g.kind.isin(FINAL)]
        last_fin = fin.d.max() if len(fin) else pd.NaT
        # активні дії після останнього підсумкового акта
        if pd.notna(last_fin):
            after = g[(g.d > last_fin) & ~g.kind.isin(POSTK | FINAL)]
            finished = len(after) == 0
        else:
            finished = False
        first_fin = fin.d.min() if len(fin) else pd.NaT
        lastdoc = g.d.max()
        lastact = lastdoc
        ol = othlast.get(c, pd.NaT)
        left = (not finished) and pd.notna(ol) and ol > lastdoc
        ds = g.d.sort_values().drop_duplicates()
        gaps = ds.diff().dt.days
        maxgap = gaps.max() if len(ds) > 1 else 0
        rows.append(dict(case=c, start=start, first_doc=g.d.min(), first_fin=first_fin, last_fin=last_fin,
                         finished=finished, left=left, lastdoc=lastdoc, lastact=lastact, ndocs=len(g),
                         nuhv=(g.judgment_form == 'Ухвала').sum(), nrish=(g.kind == 'DECISION').sum(),
                         kinds=' '.join(g.kind), cat=g.category_name.dropna().iloc[0] if g.category_name.notna().any() else '',
                         maxgap=maxgap, has_start=pd.notna(start)))
    cs = pd.DataFrame(rows)
    cs['pending'] = cs.has_start & ~cs.finished & ~cs.left
    cs['idle'] = (END - cs.lastact).dt.days
    cs['age'] = (END - cs.start).dt.days
    cs['dur'] = (cs.last_fin - cs.start).dt.days
    cs.to_pickle(argv[3])
    print('справ', len(cs), 'з відкриттям', cs.has_start.sum(), 'завершено', cs.finished.sum(), 'передано іншим', cs.left.sum(), 'незавершено', cs.pending.sum())


if __name__ == '__main__':
    {'uhv': run_uhv, 'cases': run_cases}[sys.argv[1]](sys.argv[2:])
