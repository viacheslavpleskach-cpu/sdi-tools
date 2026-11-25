#!/usr/bin/env python3
"""praktyka_obshuk.py 1.2 — аналіз практики слідчого судді у справах про дозвіл на обшук
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
Скорочені тексти (1.1): ухвали, де оприлюднено лише вступну і резолютивну або лише резолютивну
  частину («…вважаю можливим обмежитися складанням і оголошенням лише резолютивної частини ухвали»,
  «вступна та резолютивна частини»), з аналізу повністю виключаються — суддя не зобов’язаний наводити
  в них мотиви; аналізуються лише повні тексти. У реєстрі вони не наводяться, а їх кількість
  показано окремо (аркуш «Як перевірити»).
Зведення — окремо за кожен рік (--roky-suddi, --roky-inshi) для судді, інших суддів разом і кожного
  з --top інших суддів, а також динаміка між першим і останнім роком (аркуш «Динаміка»).
Обмеження: показники свідчать про те, що в ухвалі написано, а не про правильність рішення.
Зміни 1.2 (25.11.2026): скороченим текстом вважається також ухвала без мотивувальної частини
  (після «встановив» одразу резолютивна частина), навіть без застереження про скорочений текст;
  повторна публікація того самого тексту в тій самій справі з іншою датою у вступній частині
  враховується один раз (лишається примірник, рік якого збігається з роком у номері справи);
  різні ухвали в одній справі за один день більше не відкидаються як повтор (повтором є лише текст,
  схожий на вже залишений щонайменше на 90 %); згадки про пропорційність поділено за характером
  (цитата норми / цитата практики ЄСПЛ / власний висновок) — колонка W реєстру; колонки X–Y —
  ухвали, з яких видно, що обшук уже проведено, але посилання на ч. 3 ст. 233 КПК немає (кандидати,
  кожну слід прочитати); нові колонки додано після V, тож літери колонок у формулах не змінилися;
  рядок «Інші N слідчих суддів» названо «Решта слідчих суддів»; «УСТАНОВИВ» розпізнається нарівні
  з «ВСТАНОВИВ».
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
SHORT = re.compile(r'обмежитися (?:складанням (?:і|та) )?(?:оголошенням|проголошенням) лише|лише (?:вступн\w* (?:та|і) )?резолютивн\w* частин|(?:складається|оголошується|проголошується)(?: (?:та|і) (?:оголошується|проголошується))? (?:лише )?(?:вступна (?:та|і) )?резолютивна частина|ч\.\s*2\s*ст\.\s*376 КПК|повний текст ухвали (?:слідчого судді )?буде (?:оголошено|проголошено|складено|виготовлено)', re.I)
KEY = [r'не наводить зміст документів,? доданих до клопотання \(доказів\),? оскільки він обмежений таємницею досудового розслідування',
       r'дані дають достатні підстави вважати, що відомості, які містяться у відшукуваних речах і документах, мають значення для досудового розслідування',
       r'Непроведення обшуку могло б перешкодити меті кримінального провадження']


def norm(t):
    t = re.sub(r'\s+', ' ', t)
    for w in ('УХВАЛИВ', 'ПОСТАНОВИВ', 'ВСТАНОВИВ', 'УСТАНОВИВ'):
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


EXPOST = re.compile(r'здійснено проникнення|обшук\w* до постановлення ухвали|невідкладн\w+ обшук|(?:в ході|під час) якого (?:було )?виявлено та вилучено', re.I)
P_NORMA = re.compile(r'пропорційним втручанню в особисте|найбільш доцільним та ефективним способом відшукання та вилучення речей і документів, які мають значення для досудового розслідування, а також встановлення|ч(?:\.|астин\w+)\s*(?:5|п.ято\w+)\s*ст(?:\.|атті)\s*234')
P_ESPL = re.compile(r'ЄСПЛ|Європейськ\w* [Сс]уд|у справі [«"“]|проти [А-ЯІЇЄ]|див\.|§|Конвенці|заява №|суду необхідно оцінювати|у кожному конкретному випадку')


def prop_kind(ev):
    kinds, own, first = [], None, None
    for m in re.finditer(IND['prop'], ev):
        a = ev.rfind('. ', 0, m.start()) + 1; b = ev.find('. ', m.end()); b = b if b > 0 else len(ev)
        s = ev[a:b + 1]
        if P_NORMA.search(s): k = 'норма'
        elif P_ESPL.search(s): k = 'єспл'
        else: k = 'власна'; own = own or m
        kinds.append(k); first = first or m
    return kinds, own, first


def classify(raw):
    if 'Інформація заборонена для оприлюднення' in raw and len(raw) < 1500:
        return None
    t = norm(raw); a, b, c = parts(t); cl = c[:900].lower(); head = t[:900].lower()
    if not (re.search(r'обшук', head) and 'про арешт майна' not in head):
        return None
    typ = 'повна'
    if ('вступна та резолютивна' in head or 'вступна і резолютивна' in head) and len(b) < 2500: typ = 'вступна і резолютивна'
    if SHORT.search(t) and len(b) < 2500: typ = 'вступна і резолютивна'  # мотивувальної частини немає
    hasust = re.search(r'[ВУ]СТАНОВИ[ВЛ]', t[:len(a) + len(b) + 12]) is not None
    bb = re.sub(r'^[ВУ]СТАНОВИ[ВЛ]А?\s*:?', '', b.strip()).strip(' :-,')
    if (hasust and len(bb) < 300) or (not hasust and len(b) < 2500): typ = 'вступна і резолютивна'  # 1.2: порожня мотивувальна частина
    if re.search(r'описк', cl) or re.search(r'описк', head[:600]): typ = 'виправлення описки'
    cl = c[:2500].lower()
    fv = re.search(r'задов\w*ти частково|частково задов\w*ти|залишити без задоволення|в задоволенні[^.]{0,400}?відмовити|відмовити|залишити без розгляду|повернути|задов\w*ти|задовільнити|надати (?:дозвіл|слідч|прокурор|детектив|старш)', cl)
    v = fv.group(0) if fv else ''
    if re.search(r'повернути', cl[:700]) and not re.search(r'задов\w*ти частково|частково задов', cl[:700]): res = 'повернуто'
    elif 'частков' in v or (v.startswith(('задов', 'надати')) and re.search(r'в (?:іншій|решті) частин\w*[^.]{0,200}?відмов|\bчастково\b', cl)): res = 'частково'
    elif v == 'повернути': res = 'повернуто'
    elif v.startswith(('відмов', 'в задоволенні', 'залишити без задоволення')): res = 'відмовлено'
    elif v == 'залишити без розгляду': res = 'без розгляду'
    elif v: res = 'задоволено'
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
    kinds, own, first = prop_kind(ev)
    d['prop_kind'] = ('власний висновок' if 'власна' in kinds else 'цитата практики ЄСПЛ' if 'єспл' in kinds else 'цитата норми') if kinds else ''
    if own: d['prop_snip'] = snip(ev, own)
    d['expost_bez'] = (not S233.search(t)) and bool(EXPOST.search(t))
    mm = EXPOST.search(t); d['expost_snip'] = snip(t, mm) if (mm and d['expost_bez']) else ''
    import hashlib
    d['_nt'] = re.sub(r'[\d\W_]+', '', t.lower())[:12000]
    d['_fp'] = hashlib.md5(re.sub(r'[\d\W_]+', '', t.lower()).encode()).hexdigest()
    d['secr'] = SECR in t
    d['s233'] = bool(S233.search(t))
    m = CHECK.search(t); d['check'] = bool(m); d['check_snip'] = snip(t, m) if m else ''
    d['op466'] = '466/525/22' in t
    d['key'] = all(re.search(k, t) for k in KEY)
    d['merits'] = typ == 'повна' and res in ('задоволено', 'частково')
    d['_ev'] = ev; d['_mot'] = motive(t)
    m = re.search(r'1-кс[-/ ]*\d+/\d+', t[:600]); d['pr'] = m.group(0).replace(' ', '') if m else ''
    d['_len'] = len(t)
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
                povernuto=sum(1 for d in L if d['res'] == 'повернуто'), neiavka=sum(1 for d in L if d['why'] == 'неявка'),
                prop_vlasna=sum(1 for d in M if d.get('prop_kind') == 'власний висновок'), prop_norma=sum(1 for d in M if d.get('prop_kind') == 'цитата норми'), prop_espl=sum(1 for d in M if d.get('prop_kind') == 'цитата практики ЄСПЛ'), expost_bez=sum(1 for d in M if d.get('expost_bez')), bez_rozghliadu=sum(1 for d in L if d['res'] == 'без розгляду'))


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
    ap.add_argument('--min-po-suti', type=int, default=0, help='окремо показувати всіх суддів, які хоча б за один рік розглянули по суті стільки клопотань (замість --top)')
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    if a.key:
        KEY[:] = [re.escape(x.strip()) for x in a.key.split('|') if x.strip()]
    tdirs = a.texts.split(',')
    ys, yo = years(a.roky_suddi), years(a.roky_inshi)
    nm = lambda j: re.sub(r'\s+', ' ', j).strip()
    J = nm(a.suddia)
    rows, seen, skipped = [], set(), Counter()
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
            if d['typ'] != 'повна':  # скорочені тексти й виправлення описок не аналізуються
                skipped[('суддя' if mine else 'інші', y, d['typ'])] += 1; continue
            d.update(doc_id=x['doc_id'], case=x['cause_num'], judge=nm(x['judge']), date=x['adjudication_date'], mine=mine)
            rows.append(d)
    # 1.2: той самий текст у тій самій справі під іншою датою (різниця лише в цифрах) — лишається примірник, рік якого збігається з роком у номері справи
    fpg = defaultdict(list)
    for d in rows: fpg[(d['judge'], d['case'], d['_fp'])].append(d)
    drop = set()
    for L in fpg.values():
        if len(L) > 1:
            L.sort(key=lambda d: (d['date'], d['doc_id']))
            keep = next((d for d in L if re.search(r'/(\d\d)-', d['case']) and re.search(r'/(\d\d)-', d['case']).group(1) == d['date'][2:4]), L[0])
            for d in L:
                if d is not keep:
                    drop.add(d['doc_id']); skipped[('суддя' if d['mine'] else 'інші', int(d['date'][:4]), 'повторна публікація')] += 1
                    print('повторна публікація під іншою датою:', d['judge'], d['case'], d['date'], d['doc_id'], '→ залишено', keep['date'])
    rows = [d for d in rows if d['doc_id'] not in drop]
    # повторні публікації того самого повного тексту (той самий суддя, справа, № провадження, дата, результат) — лишається довший
    KG = defaultdict(list)
    for d in rows: KG[(d['judge'], d['case'], d['pr'] or d['doc_id'], d['date'], d['res'])].append(d)
    drop2 = set()
    for L in KG.values():
        if len(L) < 2: continue
        L.sort(key=lambda d: -d['_len']); reps = []
        for d in L:  # 1.2: повторною публікацією вважається лише текст, близький до вже залишеного (різні ухвали в одній справі за один день лишаються)
            if any(difflib.SequenceMatcher(None, r['_nt'], d['_nt'], autojunk=False).ratio() >= 0.9 for r in reps):
                drop2.add(d['doc_id']); skipped[('суддя' if d['mine'] else 'інші', int(d['date'][:4]), 'повторна публікація')] += 1
            else:
                if reps: print('різні ухвали в одній справі за один день залишено:', d['judge'], d['case'], d['date'], d['doc_id'], 'і', reps[0]['doc_id'])
                reps.append(d)
    if drop2: print('повторні публікації відкинуто:', len(drop2))
    rows = [d for d in rows if d['doc_id'] not in drop2]
    # схожість у межах судді
    G = defaultdict(list)
    for d in rows:
        if d['merits']: d['_sh'] = shingles(d['_ev']); G[(d['judge'], d['date'][:4])].append(d)  # у межах судді й року
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
        for k in ('_sh', '_ev', '_mot', '_len', '_fp', '_nt'): d.pop(k, None)
    # зведення
    mine = [d for d in rows if d['mine']]; oth = [d for d in rows if not d['mine']]
    po = Counter()
    for d in oth: po[d['judge']] += 1 if (d['merits'] or (d['why'] == 'по суті' and d['typ'] == 'повна')) else 0
    order = sorted({d['judge'] for d in oth}, key=lambda j: (-po[j], -sum(1 for d in oth if d['judge'] == j), j))
    code = {j: f'Слідчий суддя {i + 1}' for i, j in enumerate(order)}
    short = lambda j: j.replace('. ', '.')
    disp = (lambda j: code[j]) if a.kody else short
    top = order[:a.top]
    if a.min_po_suti:
        py = Counter((d['judge'], d['date'][:4]) for d in oth if d['merits'] or (d['why'] == 'по суті' and d['typ'] == 'повна'))
        top = [j for j in order if any(py[(j, y)] >= a.min_po_suti for y in {d['date'][:4] for d in oth})]
    S = {}
    yy = sorted(ys | yo)
    grp = [(short(J), lambda d: d['mine'])] + [('Інші слідчі судді разом', lambda d: not d['mine'])] + \
          [(disp(j), (lambda j: lambda d: d['judge'] == j)(j)) for j in top]
    if len(order) > len(top): grp.append(('Решта слідчих суддів', lambda d: not d['mine'] and d['judge'] not in top))
    for name, f in grp:
        for y in yy:
            L = [d for d in rows if f(d) and d['date'][:4] == str(y)]
            if L: S[f'{name} – {y}'] = summary(L)
    json.dump(dict(rows=rows, zvedennia=S, poriadok=[disp(j) for j in order], vidkynuto={f'{k[0]} {k[1]} {k[2]}': v for k, v in skipped.items()}), open(a.out + '.json', 'w', encoding='utf-8'), ensure_ascii=False)
    cols = ['usi', 'po_suti', 'chastkovo', 'vidmova_po_suti', 'zadov_povnyi_tekst', 'kliuchovi', 'taiemnytsia', 'proportsiinist', 'yespl', 'typovyi', 'nevidkladni', 'z_normoiu', 'op_kks']
    print('група\t' + '\t'.join(cols))
    for k, v in S.items(): print(k + '\t' + '\t'.join(str(v[c]) for c in cols))
    for k, v in sorted(skipped.items()): print('відкинуто (не повний текст):', *k, v)
    xlsx(a, rows, J, disp, top, order, short, grp, yy, skipped)


def xlsx(a, rows, J, disp, top, order, short, groups, yy, skipped):
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
            'Є фраза «не наводить зміст документів»', 'Згадано пропорційність втручання', 'Фрагмент тексту про пропорційність',
            'Є посилання на ЄСПЛ або Конвенцію', 'Фрагмент тексту про ЄСПЛ / Конвенцію', 'Клопотання про невідкладний обшук (ч. 3 ст. 233 КПК)',
            'Наведено норму про перевірку підстав невідкладного обшуку', 'Є посилання на постанову ОП ККС у справі № 466/525/22',
            'Найбільший збіг власної оцінки з іншою ухвалою цього судді, %', 'З якою ухвалою найбільший збіг', 'Текст ухвали в ЄДРСР',
            f'Збіг мотивувальної частини з ухвалою у справі № {a.zrazok_sprava}, %' if a.zrazok_sprava else 'Збіг зі зразком, %',
            'Характер згадки про пропорційність', 'Обшук уже проведено, але посилання на частину третю статті 233 КПК немає', 'Фрагмент тексту про вже проведений обшук']
    rest = 'Решта слідчих суддів'
    wb = Workbook(); ws = wb.active; ws.title = 'Як перевірити'
    txt = [('Реєстр ухвал про дозвіл на обшук, використаних для аналізу практики', True), ("Укладач: В'ячеслав Плескач", False), ('', False),
           ('Як перевірити показник: увімкніть фільтр у рядку заголовків аркуша реєстру, відберіть потрібне значення, відкрийте ухвалу за посиланням в останній колонці і знайдіть у ній (Ctrl+F) слова, за якими визначено показник; порівняйте кількість рядків із числом на аркуші «Зведення» (його рахують формули з реєстру).', False),
           ('Результат – за резолютивною частиною. «Розглянуто по суті» – задоволено повністю, частково або відмовлено по суті (повний текст); решту показників визначено лише для ухвал про задоволення з повним текстом (в інших – «—»).', False),
           ('Оцінка пропорційності – у частині власної оцінки слідчого судді (з фрази «Слідчий суддя, заслухавши / дослідивши…» після викладу клопотання до резолютивної частини) є слова «пропорційн…», «співмірн…», «співрозмірн…», «розумний баланс», «необхідн… в демократичному суспільстві» або «єдиний доцільний / можливий / ефективний спосіб». ЄСПЛ або Конвенція – там само згадано Європейський суд з прав людини, ЄСПЛ чи Конвенцію про захист прав людини.', False),
           ('Невідкладний обшук – в ухвалі є посилання на частину третю статті 233 КПК України; норма про перевірку – слова «чи дійсно були наявні підстави»; ОП ККС – номер справи 466/525/22.', False),
           ('Збіг власної оцінки – частка однакових послідовностей із п’яти слів (без цифр і знеособлених даних) з найсхожою іншою ухвалою того самого слідчого судді; 80 % і більше – типовий текст.', False),
           ('Обмеження: показники визначено автоматично за словами в тексті; вони свідчать про те, що в ухвалі написано, а не про правильність рішення. Ухвали, тексти яких в ЄДРСР закрито чи не оприлюднено, до реєстру не входять.', False)]
    sk = Counter()
    for (who, y, typ), v in skipped.items(): sk[(who, y)] += v
    txt.append(('Не аналізувалися (скорочені тексти – лише вступна і резолютивна або лише резолютивна частина, виправлення описок, повторні публікації того самого тексту): ' +
                '; '.join(f"{'суддя ' + short(J) if w == 'суддя' else 'інші судді'}, {y} рік – {v}" for (w, y), v in sorted(sk.items())) + '.', False))
    ws.column_dimensions['A'].width = 130
    for i, (s, b) in enumerate(txt, 1):
        c = ws.cell(i, 1, s); c.font = Font(name=TNR, size=12 if i == 1 else 11, bold=b); c.alignment = W
    WID = [6, 20, 18, 16, 11, 6, 30, 11, 11, 12, 12, 12, 45, 12, 45, 12, 13, 12, 12, 16, 40, 13, 18, 14, 45]
    def row(i, d):
        g = d['merits']; m = g or (d['why'] == 'по суті' and d['typ'] == 'повна'); s233 = d['s233'] and d['typ'] == 'повна' and m
        dash = lambda v: v if g else '—'
        j = short(d['judge']) if d['mine'] else disp(d['judge'])
        grp = j if d['mine'] or d['judge'] in top else rest
        return [i, j, grp, d['case'], date.fromisoformat(d['date']), int(d['date'][:4]), label(d), yn(m),
                ('так' if d['res'] == 'частково' else 'ні') if m else '—', dash(yn(d['key'])), dash(yn(d['secr'])), dash(yn(d['prop'])),
                d['prop_snip'] if g else '', dash(yn(d['espl'])), d['espl_snip'] if g else '', yn(s233),
                yn(d['check']) if s233 else '—', yn(d['op466']) if s233 else '—', d.get('maxsim') if g else None,
                d.get('maxsim_case', '') if g else '', URL + d['doc_id'], d.get('sim_zrazok'),
                (d.get('prop_kind') or '—') if g else '—', (yn(d.get('expost_bez'))) if g else '—', d.get('expost_snip', '') if g else '']
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
            ('Згадано пропорційність', [('L', 'так')]), ('частка', None), ('Є посилання на ЄСПЛ / Конвенцію', [('N', 'так')]), ('частка', None),
            ('Типовий текст (збіг ≥ 80 %)', [('S', '>=80')]), ('частка', None), ('Клопотань про невідкладний обшук', [('P', 'так')]),
            ('з них з нормою про перевірку', [('Q', 'так')]), ('з них з посиланням на ОП ККС № 466/525/22', [('R', 'так')]),
            ('Власний висновок про пропорційність конкретного обшуку', [('W', 'власний висновок')]), ('Обшук уже проведено, посилання на частину третю статті 233 КПК немає', [('X', 'так')])]
    gdef = [(short(J), 1, []), ('Інші слідчі судді разом', 2, [])] + [(disp(j), 2, [('C', disp(j))]) for j in top] + \
           ([(rest + ' разом', 2, [('C', rest)])] if len(order) > len(top) else [])
    def ingrp(d, sh, cr):
        if sh == 1: return d['mine']
        if d['mine']: return False
        if not cr: return True
        return (disp(d['judge']) if d['judge'] in top else rest) == cr[0][1]
    grp = [(f'{n} – {y}', sh, [('F', y)] + cr) for n, sh, cr in gdef for y in yy
           if any(ingrp(d, sh, cr) and d['date'][:4] == str(y) for d in rows)]
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
    rowidx = {name: i for i, (name, _, _) in enumerate(grp, hr + 1)}
    if len(yy) > 1:
        dz = wb.create_sheet('Динаміка', 2)
        y0, y1 = yy[0], yy[-1]
        dz.cell(1, 1, f'Динаміка показників: {y0} → {y1} рік (частки рахуються з аркуша «Зведення»)').font = Font(name=TNR, size=12, bold=True)
        dz.cell(2, 1, 'Зміна – різниця часток у відсоткових пунктах; «—» – немає даних за рік.').font = Font(name=TNR, size=10, italic=True)
        pok = [('Задоволено частково (від розглянутих по суті)', 'D', 'C'), ('Є всі ключові речення', 'H', 'G'),
               ('Є фраза «не наводить зміст документів»', 'I', 'G'), ('Згадано пропорційність', 'J', 'G'),
               ('Є посилання на ЄСПЛ / Конвенцію', 'L', 'G'), ('Типовий текст (збіг ≥ 80 %)', 'N', 'G'),
               ('Невідкладні обшуки: наведено норму про перевірку', 'Q', 'P')]
        hd = ['Слідчий суддя', 'Показник', f'{y0} рік', f'{y1} рік', 'Зміна, в. п.']
        for j, v in enumerate(hd, 1):
            c = dz.cell(4, j, v); c.font = FB; c.fill = HF; c.border = BR; c.alignment = Alignment(wrap_text=True, horizontal='center')
        r = 5
        for n, _, _ in gdef:
            a0, a1 = rowidx.get(f'{n} – {y0}'), rowidx.get(f'{n} – {y1}')
            for k, (pn, num, den) in enumerate(pok):
                dz.cell(r, 1, n if k == 0 else ''); dz.cell(r, 2, pn)
                for j, ri in ((3, a0), (4, a1)):
                    dz.cell(r, j, f"=IF(Зведення!{den}{ri}=0,\"—\",Зведення!{num}{ri}/Зведення!{den}{ri})" if ri else '—').number_format = '0%'
                dz.cell(r, 5, f'=IF(OR(ISTEXT(C{r}),ISTEXT(D{r})),"—",ROUND((D{r}-C{r})*100,0))')
                for j in range(1, 6):
                    c = dz.cell(r, j); c.font = FB if (j == 1 and k == 0) else F; c.border = BR; c.alignment = W
                r += 1
        dz.column_dimensions['A'].width = 30; dz.column_dimensions['B'].width = 46
        for col in 'CDE': dz.column_dimensions[col].width = 12
        dz.freeze_panes = 'A5'
    wb.save(a.out + '.xlsx'); print('→', a.out + '.xlsx', '(перерахуйте формули: soffice або recalc.py скіла xlsx)')


if __name__ == '__main__':
    main()
