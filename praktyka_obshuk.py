#!/usr/bin/env python3
"""praktyka_obshuk.py 1.0 — аналіз практики слідчого судді у справах про дозвіл на обшук
і порівняння з іншими слідчими суддями того самого суду (В'ячеслав Плескач).

Для висновків, де треба показати, що недоліки ухвал судді не поодинокі (тенденції практики),
і відокремити їх від особливостей суду. Результат — реєстр Excel, який читач може перевірити
сам: кожен рядок має посилання на текст в ЄДРСР і фрагмент тексту, де знайдено ключові слова,
а аркуш «Зведення» рахується формулами з реєстру.

Підготовка даних (edrsr_vybirka.py з edrsr-tool; категорія 40467 — «проведення обшуку житла
чи іншого володіння особи»):
    python3 edrsr_vybirka.py meta --years 2024-2026 --court 2606 --kind 2 --form 5 --cat 40467 \
        --from 2024-01-01 --to 2025-12-31 --out obshuk.tsv
    python3 edrsr_vybirka.py texts --in obshuk.tsv --dir texts

Аналіз:
    python3 praktyka_obshuk.py --in obshuk.tsv --texts texts --suddia "Білоцерківець О. А." \
        --roky-suddi 2024-2025 --roky-inshi 2025 \
        --zrazok-sprava 757/35459/25-к [--zrazok-text motiv.txt] --out praktyka
    → praktyka.json (усі ухвали з показниками і зведення), praktyka.xlsx (реєстр), зведення в консолі.

--in і --texts приймають кілька значень через кому (напр. окрема вибірка судді за 2024 рік
і вибірка суду за 2025 рік); дублікати doc_id відкидаються.
--zrazok-text — файл з мотивувальною частиною ухвали-зразка, якщо її текст в ЄДРСР закрито
(від «Відповідно до ч. 1 ст. 234» до «Керуючись»); інакше вона береться з вибірки за --zrazok-sprava.
--kody — позначити інших суддів умовними номерами замість прізвищ.

Показники (усі — автоматично за словами, тому в реєстрі є фрагмент тексту і посилання):
  результат — за резолютивною частиною (задоволено повністю / частково / відмовлено по суті /
    через неявку / повернуто; «лише вступна і резолютивна частини», «виправлення описки»);
  власна оцінка — частина мотивувальної від першої фрази «Слідчий суддя, заслухавши / дослідивши…»
    (після 20 % мотивувальної частини) до резолютивної; у ній шукаються:
    пропорційність («пропорційн», «співмірн», «розумний баланс», «необхідн… в демократичному»,
    «єдиний доцільний спосіб»), ЄСПЛ / Конвенція, конкретні матеріали (протоколи, рапорти тощо);
  ч. 3 ст. 233 — посилання в тексті; норма про перевірку — «чи дійсно були наявні підстав»;
    ОП ККС — «466/525/22» (будь-де в тексті);
  типовий текст — найбільший збіг власної оцінки з іншою ухвалою того самого судді (Жаккар за
    послідовностями з 5 слів без цифр і знеособлених даних), ≥ 80 % — типовий;
  збіг зі зразком — difflib-подібність мотивувальної частини до ухвали-зразка;
  три ключові речення — стандартні речення ухвал Білоцерківця О. А. (див. KEY); для іншого
    судді задати свої через --key "речення1|речення2|…".
Обмеження: показники свідчать про те, що в ухвалі написано, а не про правильність рішення.
"""
import argparse, csv, difflib, json, os, re, statistics as st, sys
from collections import Counter, defaultdict
from datetime import date

URL = 'https://reyestr.court.gov.ua/Review/'
EVAL = re.compile(r'(Слідчий суддя,? (?:заслухавши|дослідивши|вивчивши|оглянувши|перевіривши|розглянувши|оцінивши|проаналізувавши)|Дослідивши|Вивчивши|Оцінивши|Заслухавши|Перевіривши|Проаналізувавши|Розглянувши клопотання (?:та|і)|слідчий суддя (?:дійшов|приходить|вважає|зазначає)|(?:Відповідно до|Згідно (?:з )?|Згідно із )(?:ч\.|частин\w*)\s*\S+\s*ст(?:\.|атті)\s*234)')
IND = {
    'prop': r'пропорційн|співмірн|співрозмірн|розумн\w* баланс|необхідн\w* в демократичному|єдин\w* (?:доцільн|можлив|ефективн)\w* спос',
    'espl': r'Європейськ\w* суд\w* з прав людини|ЄСПЛ|Конвенці\w* про захист прав',
    'konkr': r'протокол\w* (?:допиту|огляду|обшуку|затримання|за результатами|проведення)|рапорт\w*|висновк\w* (?:експерт|судов)|показан\w* (?:свідк|потерпіл|підозрюв)|витяг\w* з (?:Державного|ЄДР|Єдиного)|відповід\w* на запит|матеріал\w* (?:НСРД|негласн)',
}
CHECK = re.compile(r'чи дійсно (?:були|існували) наявн\w* підстав')
S233 = re.compile(r'(?:ч\.|частин\w*)\s*(?:3|третьо\w*)\s*ст\S*\s*233')
SECR = 'не наводить зміст документів'
KEY = [r'не наводить зміст документів,? доданих до клопотання \(доказів\),? оскільки він обмежений таємницею досудового розслідування',
       r'дані дають достатні підстави вважати, що відомості, які містяться у відшукуваних речах і документах, мають значення для досудового розслідування',
       r'Непроведення обшуку могло б перешкодити меті кримінального провадження']


def norm(t):
    t = re.sub(r'\s+', ' ', t)
    for w in ('УХВАЛИВ', 'ПОСТАНОВИВ', 'ВСТАНОВИВ'):
        t = re.sub(r'\s?'.join(w), w, t)
    return t


def parts(t):
    """(вступ, мотивувальна від ВСТАНОВИВ до «Керуючись», резолютивна)"""
    m = list(re.finditer(r'(УХВАЛИВ|ПОСТАНОВИВ)\s*:?', t))
    end = m[-1].start() if m else len(t)
    k = None
    for mm in re.finditer(r'[Кк]еруючись|На підставі викладеного|Враховуючи викладене та керуючись', t[:end]):
        k = mm.start()
    be = k if k else end
    s = t.find('ВСТАНОВИВ'); s = s if 0 <= s < be else 0
    return t[:s], t[s:be], t[end:]


def motive(t):
    a = t.find('Відповідно до ч. 1 ст. 234'); b = t.find('Керуючись', a if a > 0 else 0)
    if a < 0: a = max(0, t.find('ВСТАНОВИВ'))
    return t[a:b] if b > a else t[a:a + 1500]


def shingles(s, n=5):
    w = re.sub(r'ОСОБА_\d+|АДРЕСА_\d+|НОМЕР_\d+|ІНФОРМАЦІЯ_\d+|\d+', '#', s).lower().split()
    return set(tuple(w[i:i + n]) for i in range(max(0, len(w) - n + 1)))


def snip(t, m, a=160, b=200):
    return ('…' if m.start() > a else '') + t[max(0, m.start() - a):m.end() + b].strip() + '…'


def years(s):
    a, _, b = s.partition('-'); return set(range(int(a), int(b or a) + 1))


def classify(raw):
    if 'Інформація заборонена для оприлюднення' in raw and len(raw) < 1500:
        return None
    t = norm(raw); a, b, c = parts(t); cl = c[:900].lower(); head = t[:900].lower()
    if not (re.search(r'обшук', head) and 'про арешт майна' not in head):
        return None
    typ = 'повна'
    if 'вступна та резолютивна' in head or 'вступна і резолютивна' in head: typ = 'вступна і резолютивна'
    if re.search(r'описк', cl) or re.search(r'описк', head[:600]): typ = 'виправлення описки'
    if re.search(r'задов\w*ити частково|частково задов\w*', cl): res = 'частково'
    elif re.search(r'повернути', cl[:400]): res = 'повернуто'
    elif re.search(r'відмовити|залишити без задоволення', cl[:700]): res = 'відмовлено'
    elif re.search(r'залишити без розгляду', cl): res = 'без розгляду'
    elif re.search(r'задов\w*ити|надати (?:дозвіл|слідч|прокурор|детектив|старш)', cl): res = 'задоволено'
    else: res = '?'
    why = ''
    if res == 'відмовлено':
        why = 'неявка' if re.search(r"не з.яв(?:ил|ив|ля)|неявк|не прибул", b) or re.search(r"не з.яв(?:ил|ив|ля)", cl) else 'по суті'
    d = dict(typ=typ, res=res, why=why)
    start = None
    for m in EVAL.finditer(b):
        if m.start() > 0.2 * len(b): start = m.start(); break
    if start is None: start = int(len(b) * 0.6)
    ev = b[start:]
    d['eval_len'] = len(ev)
    for k, p in IND.items():
        m = re.search(p, ev); d[k] = bool(m); d[k + '_snip'] = snip(ev, m) if m else ''
    d['secr'] = SECR in t
    d['s233'] = bool(S233.search(t))
    m = CHECK.search(t); d['check'] = bool(m); d['check_snip'] = snip(t, m) if m else ''
    d['op466'] = '466/525/22' in t
    d['key'] = all(re.search(k, t) for k in KEY)
    d['merits'] = typ == 'повна' and res in ('задоволено', 'частково')
    d['_ev'] = ev; d['_mot'] = motive(t)
    return d


def label(d):
    if d['typ'] == 'вступна і резолютивна': return 'задоволено (оприлюднено лише вступну і резолютивну частини)'
    if d['typ'] == 'виправлення описки': return 'виправлення описки'
    if d['res'] == 'відмовлено':
        return 'відмовлено по суті' if d['why'] == 'по суті' else 'відмовлено через неявку сторони обвинувачення'
    return {'задоволено': 'задоволено повністю', 'частково': 'задоволено частково', 'повернуто': 'повернуто',
            'без розгляду': 'залишено без розгляду'}.get(d['res'], 'не визначено')


def summary(L):
    M = [d for d in L if d['merits']]
    ref = sum(1 for d in L if d['why'] == 'по суті' and d['typ'] == 'повна')
    s233 = [d for d in L if d['s233'] and d['typ'] == 'повна' and (d['merits'] or d['why'] == 'по суті')]
    f = lambda k: sum(1 for d in M if d[k])
    return dict(usi=len(L), po_suti=len(M) + ref, chastkovo=sum(1 for d in M if d['res'] == 'частково'), vidmova_po_suti=ref,
                zadov_povnyi_tekst=len(M), kliuchovi=f('key'), taiemnytsia=f('secr'), proportsiinist=f('prop'), yespl=f('espl'),
                typovyi=sum(1 for d in M if (d.get('maxsim') or 0) >= 80),
                mediana_otsinky=int(st.median([d['eval_len'] for d in M])) if M else 0,
                nevidkladni=len(s233), z_normoiu=sum(d['check'] for d in s233), op_kks=sum(d['op466'] for d in s233),
                povernuto=sum(1 for d in L if d['res'] == 'повернуто'), neiavka=sum(1 for d in L if d['why'] == 'неявка'))


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--in', dest='inp', required=True)
    ap.add_argument('--texts', required=True)
    ap.add_argument('--suddia', required=True, help='як у колонці judge ЄДРСР, напр. «Білоцерківець О. А.»')
    ap.add_argument('--roky-suddi', default='2024-2025')
    ap.add_argument('--roky-inshi', default='2025')
    ap.add_argument('--zrazok-sprava', default='')
    ap.add_argument('--zrazok-text', default='')
    ap.add_argument('--key', default='', help='ключові речення через «|» замість типових')
    ap.add_argument('--kody', action='store_true', help='інших суддів позначити умовними номерами')
    ap.add_argument('--top', type=int, default=7, help='скільки інших суддів показувати окремо')
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    if a.key:
        KEY[:] = [re.escape(x.strip()) for x in a.key.split('|') if x.strip()]
    tdirs = a.texts.split(',')
    ys, yo = years(a.roky_suddi), years(a.roky_inshi)
    nm = lambda j: re.sub(r'\s+', ' ', j).strip()
    J = nm(a.suddia)
    rows, seen = [], set()
    for f in a.inp.split(','):
        for x in csv.DictReader(open(f, encoding='utf-8'), delimiter='\t'):
            if x['doc_id'] in seen or x.get('status', '1') == '0': continue
            if x.get('category_code') and x['category_code'] != '40467': continue
            y = int(x['adjudication_date'][:4]); mine = nm(x['judge']) == J
            if (mine and y not in ys) or (not mine and y not in yo): continue
            p = next((os.path.join(t, x['doc_id'] + '.txt') for t in tdirs if os.path.exists(os.path.join(t, x['doc_id'] + '.txt'))), None)
            if not p: print('немає тексту:', x['doc_id'], x['cause_num'], file=sys.stderr); continue
            d = classify(open(p, encoding='utf-8', errors='ignore').read())
            if d is None: continue
            seen.add(x['doc_id'])
            d.update(doc_id=x['doc_id'], case=x['cause_num'], judge=nm(x['judge']), date=x['adjudication_date'], mine=mine)
            rows.append(d)
    # схожість у межах судді
    G = defaultdict(list)
    for d in rows:
        if d['merits']: d['_sh'] = shingles(d['_ev']); G[d['judge']].append(d)
    for L in G.values():
        for d in L:
            best, bc = 0, ''
            for e in L:
                if e is d or not d['_sh'] or not e['_sh']: continue
                j = len(d['_sh'] & e['_sh']) / len(d['_sh'] | e['_sh'])
                if j > best: best, bc = j, e['case']
            d['maxsim'], d['maxsim_case'] = round(best * 100), bc
    base = open(a.zrazok_text, encoding='utf-8').read().strip() if a.zrazok_text else \
        next((d['_mot'] for d in rows if d['case'] == a.zrazok_sprava), '')
    for d in rows:
        d['sim_zrazok'] = round(difflib.SequenceMatcher(None, base, d['_mot']).ratio() * 100) if base and d['mine'] and d['merits'] else None
        for k in ('_sh', '_ev', '_mot'): d.pop(k, None)
    # зведення
    mine = [d for d in rows if d['mine']]; oth = [d for d in rows if not d['mine']]
    po = Counter()
    for d in oth: po[d['judge']] += 1 if (d['merits'] or (d['why'] == 'по суті' and d['typ'] == 'повна')) else 0
    order = sorted({d['judge'] for d in oth}, key=lambda j: (-po[j], -sum(1 for d in oth if d['judge'] == j), j))
    code = {j: f'Слідчий суддя {i + 1}' for i, j in enumerate(order)}
    short = lambda j: j.replace('. ', '.')
    disp = (lambda j: code[j]) if a.kody else short
    top = order[:a.top]
    S = {}
    for y in sorted(ys): S[f'{short(J)} – {y}'] = summary([d for d in mine if d['date'][:4] == str(y)])
    S[f'{short(J)} – разом'] = summary(mine)
    S['Інші слідчі судді разом'] = summary(oth)
    for j in top: S[disp(j)] = summary([d for d in oth if d['judge'] == j])
    if len(order) > len(top): S[f'Інші {len(order) - len(top)} слідчих суддів'] = summary([d for d in oth if d['judge'] not in top])
    json.dump(dict(rows=rows, zvedennia=S, poriadok=[disp(j) for j in order]), open(a.out + '.json', 'w', encoding='utf-8'), ensure_ascii=False)
    cols = ['usi', 'po_suti', 'chastkovo', 'vidmova_po_suti', 'zadov_povnyi_tekst', 'kliuchovi', 'taiemnytsia', 'proportsiinist', 'yespl', 'typovyi', 'nevidkladni', 'z_normoiu', 'op_kks']
    print('група\t' + '\t'.join(cols))
    for k, v in S.items(): print(k + '\t' + '\t'.join(str(v[c]) for c in cols))
    xlsx(a, rows, J, disp, top, order, short)


def xlsx(a, rows, J, disp, top, order, short):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter as L
    TNR = 'Times New Roman'
    F = Font(name=TNR, size=10); FB = Font(name=TNR, size=10, bold=True); FL = Font(name=TNR, size=10, color='0563C1', underline='single')
    W = Alignment(wrap_text=True, vertical='top'); HF = PatternFill('solid', fgColor='DDEBF7'); th = Side(style='thin', color='BFBFBF')
    BR = Border(left=th, right=th, top=th, bottom=th)
    yn = lambda b: 'так' if b else 'ні'
    HEAD = ['№', 'Слідчий суддя', 'Рядок зведення', 'Справа №', 'Дата ухвали', 'Рік', 'Результат розгляду клопотання',
            'Розглянуто по суті (повний текст)', 'Дозвіл звужено (задоволено частково)', 'Є всі ключові речення стандартного тексту',
            'Є фраза «не наводить зміст документів»', 'Є оцінка пропорційності втручання', 'Фрагмент тексту про пропорційність',
            'Є посилання на ЄСПЛ або Конвенцію', 'Фрагмент тексту про ЄСПЛ / Конвенцію', 'Клопотання про невідкладний обшук (ч. 3 ст. 233 КПК)',
            'Наведено норму про перевірку підстав невідкладного обшуку', 'Є посилання на постанову ОП ККС у справі № 466/525/22',
            'Найбільший збіг власної оцінки з іншою ухвалою цього судді, %', 'З якою ухвалою найбільший збіг', 'Текст ухвали в ЄДРСР',
            f'Збіг мотивувальної частини з ухвалою у справі № {a.zrazok_sprava}, %' if a.zrazok_sprava else 'Збіг зі зразком, %']
    rest = f'Інші {len(order) - len(top)} слідчих суддів'
    wb = Workbook(); ws = wb.active; ws.title = 'Як перевірити'
    txt = [('Реєстр ухвал про дозвіл на обшук, використаних для аналізу практики', True), ("Укладач: В'ячеслав Плескач", False), ('', False),
           ('Як перевірити показник: увімкніть фільтр у рядку заголовків аркуша реєстру, відберіть потрібне значення, відкрийте ухвалу за посиланням в останній колонці і знайдіть у ній (Ctrl+F) слова, за якими визначено показник; порівняйте кількість рядків із числом на аркуші «Зведення» (його рахують формули з реєстру).', False),
           ('Результат – за резолютивною частиною. «Розглянуто по суті» – задоволено повністю, частково або відмовлено по суті (повний текст); решту показників визначено лише для ухвал про задоволення з повним текстом (в інших – «—»).', False),
           ('Оцінка пропорційності – у частині власної оцінки слідчого судді (з фрази «Слідчий суддя, заслухавши / дослідивши…» після викладу клопотання до резолютивної частини) є слова «пропорційн…», «співмірн…», «співрозмірн…», «розумний баланс», «необхідн… в демократичному суспільстві» або «єдиний доцільний / можливий / ефективний спосіб». ЄСПЛ або Конвенція – там само згадано Європейський суд з прав людини, ЄСПЛ чи Конвенцію про захист прав людини.', False),
           ('Невідкладний обшук – в ухвалі є посилання на частину третю статті 233 КПК України; норма про перевірку – слова «чи дійсно були наявні підстави»; ОП ККС – номер справи 466/525/22.', False),
           ('Збіг власної оцінки – частка однакових послідовностей із п’яти слів (без цифр і знеособлених даних) з найсхожою іншою ухвалою того самого слідчого судді; 80 % і більше – типовий текст.', False),
           ('Обмеження: показники визначено автоматично за словами в тексті; вони свідчать про те, що в ухвалі написано, а не про правильність рішення. Ухвали, тексти яких в ЄДРСР закрито чи не оприлюднено, до реєстру не входять.', False)]
    ws.column_dimensions['A'].width = 130
    for i, (s, b) in enumerate(txt, 1):
        c = ws.cell(i, 1, s); c.font = Font(name=TNR, size=12 if i == 1 else 11, bold=b); c.alignment = W
    WID = [6, 20, 18, 16, 11, 6, 30, 11, 11, 12, 12, 12, 45, 12, 45, 12, 13, 12, 12, 16, 40, 13]
    def row(i, d):
        g = d['merits']; m = g or (d['why'] == 'по суті' and d['typ'] == 'повна'); s233 = d['s233'] and d['typ'] == 'повна' and m
        dash = lambda v: v if g else '—'
        j = short(d['judge']) if d['mine'] else disp(d['judge'])
        grp = j if d['mine'] or d['judge'] in top else rest
        return [i, j, grp, d['case'], date.fromisoformat(d['date']), int(d['date'][:4]), label(d), yn(m),
                ('так' if d['res'] == 'частково' else 'ні') if m else '—', dash(yn(d['key'])), dash(yn(d['secr'])), dash(yn(d['prop'])),
                d['prop_snip'] if g else '', dash(yn(d['espl'])), d['espl_snip'] if g else '', yn(s233),
                yn(d['check']) if s233 else '—', yn(d['op466']) if s233 else '—', d.get('maxsim') if g else None,
                d.get('maxsim_case', '') if g else '', URL + d['doc_id'], d.get('sim_zrazok')]
    def sheet(title, data):
        s = wb.create_sheet(title)
        for j, v in enumerate(HEAD, 1):
            c = s.cell(1, j, v); c.font = FB; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); c.border = BR
        s.row_dimensions[1].height = 80
        for i, d in enumerate(data, 1):
            for j, v in enumerate(row(i, d), 1):
                c = s.cell(i + 1, j, v); c.font = F; c.alignment = W; c.border = BR
                if j == 5: c.number_format = 'DD.MM.YYYY'
            lc = s.cell(i + 1, 21); lc.hyperlink = lc.value; lc.font = FL
        for j, w in enumerate(WID, 1): s.column_dimensions[L(j)].width = w
        s.freeze_panes = 'E2'; s.auto_filter.ref = f'A1:{L(len(HEAD))}{len(data) + 1}'
        return len(data) + 1
    mine = sorted([d for d in rows if d['mine']], key=lambda d: (d['date'], d['case']))
    rank = {j: i for i, j in enumerate(order)}
    oth = sorted([d for d in rows if not d['mine']], key=lambda d: (rank[d['judge']], d['date'], d['case']))
    t1 = f'{short(J)} {a.roky_suddi}'[:31]; t2 = f'Інші судді {a.roky_inshi}'[:31]
    n1 = sheet(t1, mine); n2 = sheet(t2, oth)
    z = wb.create_sheet('Зведення', 1)
    cols = [('Усього ухвал', []), ('Розглянуто по суті (повний текст)', [('H', 'так')]), ('з них задоволено частково', [('I', 'так')]), ('частка', None),
            ('Відмовлено по суті', [('G', 'відмовлено по суті')]), ('Ухвал про задоволення з повним текстом', [('J', '<>—'), ('H', 'так')]),
            ('Є всі ключові речення', [('J', 'так')]), ('Є фраза «не наводить зміст документів»', [('K', 'так')]),
            ('Є оцінка пропорційності', [('L', 'так')]), ('частка', None), ('Є посилання на ЄСПЛ / Конвенцію', [('N', 'так')]), ('частка', None),
            ('Типовий текст (збіг ≥ 80 %)', [('S', '>=80')]), ('частка', None), ('Клопотань про невідкладний обшук', [('P', 'так')]),
            ('з них з нормою про перевірку', [('Q', 'так')]), ('з них з посиланням на ОП ККС № 466/525/22', [('R', 'так')])]
    grp = [(f'{short(J)} – {y}', 1, [('F', y)]) for y in sorted(years(a.roky_suddi))] + [(f'{short(J)} – разом', 1, []), ('Інші слідчі судді разом', 2, [])] + \
          [(disp(j), 2, [('C', disp(j))]) for j in top] + ([(rest + ' разом', 2, [('C', rest)])] if len(order) > len(top) else [])
    z.cell(1, 1, 'Зведення (формули рахують значення з аркушів реєстру)').font = Font(name=TNR, size=12, bold=True)
    z.cell(2, 1, 'Частки: «задоволено частково» – від розглянутих по суті; решта – від ухвал про задоволення з повним текстом.').font = Font(name=TNR, size=10, italic=True)
    hr = 4; z.cell(hr, 1, 'Слідчий суддя / період')
    for j, (n, _) in enumerate(cols, 2): z.cell(hr, j, n)
    for j in range(1, len(cols) + 2):
        c = z.cell(hr, j); c.font = FB; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); c.border = BR
    z.row_dimensions[hr].height = 75
    for i, (name, sh, crit) in enumerate(grp, hr + 1):
        z.cell(i, 1, name)
        sn, n = (f"'{t1}'", n1) if sh == 1 else (f"'{t2}'", n2)
        rng = lambda col: f'{sn}!${col}$2:${col}${n}'
        for j, (cn, cc) in enumerate(cols, 2):
            if cc is None:
                den = L(3) if j == 5 else L(7)
                c = z.cell(i, j, f'=IF({den}{i}=0,"—",{L(j - 1)}{i}/{den}{i})'); c.number_format = '0%'
            else:
                cr = crit + cc
                f = f"=COUNTA({rng('A')})" if not cr else '=COUNTIFS(' + ','.join(f'{rng(col)},{v if isinstance(v, int) else chr(34) + v + chr(34)}' for col, v in cr) + ')'
                z.cell(i, j, f)
        for j in range(1, len(cols) + 2):
            c = z.cell(i, j); c.font = F; c.border = BR; c.alignment = W
    z.column_dimensions['A'].width = 34
    for j in range(2, len(cols) + 2): z.column_dimensions[L(j)].width = 12
    z.freeze_panes = 'B5'
    wb.save(a.out + '.xlsx'); print('→', a.out + '.xlsx', '(перерахуйте формули: soffice або recalc.py скіла xlsx)')


if __name__ == '__main__':
    main()
