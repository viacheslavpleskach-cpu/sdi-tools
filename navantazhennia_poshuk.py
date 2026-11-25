#!/usr/bin/env python3
"""Навантаження суду за даними ДСА України (довідник навантаження проєкту «СДІ»).

Використання:
    python3 navantazhennia_poshuk.py "<суд>" [--data ДД.ММ.РРРР] [--period 2025_4]
Приклади:
    python3 navantazhennia_poshuk.py "Київський районний суд м. Одеси" --data 14.07.2026
    python3 navantazhennia_poshuk.py "Барський" --period 2024_4

Що робить:
  * знаходить суд (без урахування регістру, «м.» = «міста», апостроф будь-який; достатньо частини назви);
  * друкує таблицю за всіма періодами: надійшло, судді (за рішенням ВРП / з повноваженнями / за
    нормативами), справ на суддю з повноваженнями, середньомісячне надходження на суддю,
    кількість робочих днів, потрібних одному судді для розгляду всього, що надійшло за період,
    місце суду в області та в Україні (серед судів того самого виду, за справами на суддю);
  * з --data (дата подання скарги або подій) — бере останній період, що закінчився до цієї дати,
    і останній повний рік; друкує готовий абзац для висновку (Посібник, розд. 9) і дані ВККСУ про посади.

Довідник — файли «Навантаження_суди.csv», «Навантаження_періоди.csv», «Навантаження_дані.csv»
і «Навантаження_посади.csv» (посади суддів за даними ВККСУ) поруч зі скриптом (або --dir <тека>).
Новий період ДСА дописує nav_build.py --append.

Версія 1.1 (25.11.2026): дата — ДД.ММ.РРРР або РРРР-ММ-ДД; повідомлення, якщо до дати немає жодного періоду або дані
суду починаються пізніше / обриваються (перейменування, зміна підсудності); суд без суддів з повноваженнями — окреме
формулювання замість «близько — справ»; суди з однаковим показником ділять одне місце.
1.2 (26.11.2026): перейменовані суди — одна історія під поточною назвою (колонка «попередні_назви» довідника, nav_build.py
--zvesty); суд шукається і за попередньою назвою; в абзаці за період до перейменування — назва того часу з приміткою
«(нині — …)».
"""
import argparse, csv, datetime as dt, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
KIND_GEN = {'Місцеві загальні суди': 'місцевих загальних судів',
            'Апеляційні загальні суди': 'апеляційних загальних судів',
            'Окружні адміністративні суди': 'окружних адміністративних судів',
            'Апеляційні адміністративні суди': 'апеляційних адміністративних судів',
            'Місцеві господарські суди': 'місцевих господарських судів',
            'Апеляційні господарські суди': 'апеляційних господарських судів'}


def norm(s):
    s = s.lower().replace('’', "'").replace('`', "'").replace('ʼ', "'")
    s = re.sub(r'\bм\.\s*', 'міста ', s)
    s = s.replace('«', '').replace('»', '').replace('"', '')
    return re.sub(r'\s+', ' ', s).strip()


def f(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def d(s):
    for f in ('%d.%m.%Y', '%Y-%m-%d'):
        try:
            return dt.datetime.strptime(s.strip(), f).date()
        except ValueError:
            pass
    sys.exit(f'Не розпізнано дату «{s}»: потрібно ДД.ММ.РРРР або РРРР-ММ-ДД.')


def region_gen(r):
    if not r:
        return ''
    if r == 'м. Київ':
        return 'міста Києва'
    if 'Крим' in r:
        return 'Автономної Республіки Крим'
    if 'Севастопол' in r:
        return 'міста Севастополя'
    return r.replace('ська область', 'ської області')


def fmt(x, nd=0):
    if x is None:
        return '—'
    s = f'{x:,.{nd}f}'.replace(',', ' ')
    return s.replace('.', ',')


KINDS = {'1': 'Місцеві загальні суди', '2': 'Апеляційні загальні суди', '3': 'Окружні адміністративні суди',
         '4': 'Апеляційні адміністративні суди', '5': 'Місцеві господарські суди',
         '6': 'Апеляційні господарські суди'}
POSADY_SRC = {'01.05.2023': 'лист ВККСУ від 17.05.2023 № 21-725/23',
              '01.08.2023': 'лист ВККСУ від 08.08.2023 № 21-1360/23',
              '24.01.2024': 'інформація ВККСУ про кількість посад суддів станом на 24.01.2024'}


def rd(d, name):
    p = os.path.join(d, name)
    if not os.path.exists(p):
        return []
    with open(p, encoding='utf-8') as fh:
        return list(csv.DictReader(fh))


def load(d):
    """Збирає повні рядки з трьох файлів довідника; показники на місяць і дні рахує за методикою ДСА:
    на місяць = надійшло / (місяців · 11/12) / суддів з повноваженнями;
    днів = нормативний час / (суддів з повноваженнями · 8)."""
    courts = {c['id']: c for c in rd(d, 'Навантаження_суди.csv')}
    periods = {p['код']: p for p in rd(d, 'Навантаження_періоди.csv')}
    if not courts:
        sys.exit(f'Немає файлів довідника у теці {d}: потрібні Навантаження_суди.csv, '
                 'Навантаження_періоди.csv, Навантаження_дані.csv')
    rows = []
    for x in rd(d, 'Навантаження_дані.csv'):
        c, p = courts[x['id']], periods[x['період']]
        inc, nh, act = f(x['надійшло']), f(x['норм_год']), f(x['повн'])
        m = f(p['місяців'])
        r = {'період_код': x['період'], 'період': p['період'], 'стан_на': p['стан_на'],
             'судді_станом_на': p['судді_станом_на'], 'вид_суду': KINDS[c['вид']], 'регіон': c['регіон'],
             'суд': c['суд'], '_id': c['id'], 'надійшло': x['надійшло'], 'норм_год': x['норм_год'],
             'суддів_ВРП': x['ВРП'], 'суддів_з_повноваженнями': x['повн'], 'суддів_за_нормативами': x['норм'],
             'на_суддю_на_місяць': f'{inc / (m * 11 / 12) / act:.1f}' if inc is not None and act else '',
             'днів_на_суддю': f'{nh / (act * 8):.1f}' if nh is not None and act else '',
             'джерело': p['джерело'], 'примітка': p['примітка']}
        r['_na_suddiu'] = inc / act if inc is not None and act else None
        r['_n'] = norm(r['суд'])
        r['_old'] = [x.rsplit('@', 1) for x in (c.get('попередні_назви') or '').split(' | ') if '@' in x]
        r['суд_тоді'] = next((nm for nm, last in sorted(r['_old'], key=lambda y: y[1]) if x['період'] <= last), r['суд'])
        rows.append(r)
    # місця: у межах періоду і виду суду (Україна) та області (лише місцеві загальні)
    groups = {}
    for r in rows:
        groups.setdefault((r['період_код'], r['вид_суду']), []).append(r)
    for (_, kind), rs in groups.items():
        ok = sorted([r for r in rs if r['_na_suddiu']], key=lambda r: -r['_na_suddiu'])

        def rank(rr, field):                 # однаковий показник — однакове місце (1, 2, 2, 4)
            k = 0
            for i, r in enumerate(rr, 1):
                if i == 1 or abs(r['_na_suddiu'] - rr[i - 2]['_na_suddiu']) > 1e-9:
                    k = i
                r[field] = (k, len(rr))
        rank(ok, '_rank_ua')
        if kind == 'Місцеві загальні суди':
            byreg = {}
            for r in ok:
                byreg.setdefault(r['регіон'], []).append(r)
            for reg, rr in byreg.items():
                rank(rr, '_rank_reg')
    posady = rd(d, 'Навантаження_посади.csv')
    return rows, posady


def find(rows, q):
    toks = norm(q).split()
    def hit(n):
        words = n.split()
        return all(any(w.startswith(t) for w in words) for t in toks)
    names = sorted({r['_n'] for r in rows if hit(r['_n']) or any(hit(norm(nm)) for nm, _ in r['_old'])})
    exact = sorted({r['_n'] for r in rows if r['_n'] in names and (r['_n'] == norm(q) or
                                                                   any(norm(nm) == norm(q) for nm, _ in r['_old']))})
    return exact or names


def court_gen(name):
    """Родовий відмінок назви суду: слова до «суд» включно (прикметники на -ий/-ій → -ого/-ього)."""
    words = name.split()
    out, done = [], False
    for w in words:
        if done:
            out.append(w); continue
        if w.lower() == 'суд':
            out.append(w + 'у'); done = True; continue
        head, sep, last = w.rpartition('-')
        if last.endswith('ий'):
            last = last[:-2] + 'ого'
        elif last.endswith('ій'):
            last = last[:-2] + 'ього'
        out.append(head + sep + last)
    return ' '.join(out) if done else name


def source_ref(src):
    """«лист ДСА України від … № …» → «лист від … № …»; «інформація ДСА України (вх. …)» → «вх. …»."""
    if src.startswith('лист ДСА України'):
        return src.replace('лист ДСА України', 'лист', 1)
    m = re.search(r'\((вх\. ВРП[^)]*)\)', src)
    return m.group(1) if m else src


def paragraph(r, posady):
    reg = region_gen(r['регіон'])
    kind = r['вид_суду']
    act, vrp, nrm = f(r['суддів_з_повноваженнями']), f(r['суддів_ВРП']), f(r['суддів_за_нормативами'])
    vrp_word = 'визначених рішенням ВРП' if r['період_код'] >= '2023_3' else 'визначених наказами ДСА України'
    ref = source_ref(r['джерело'])
    lead = f'За даними Державної судової адміністрації України ({ref})'
    sud = court_gen(r['суд_тоді']) + (f' (нині — {r["суд"]})' if r['суд_тоді'] != r['суд'] else '')
    if not act:                              # суддів з повноваженнями немає: показники на суддю не рахуються
        return (f'{lead} за {r["період"]} до {sud} надійшло {fmt(f(r["надійшло"]))} справ і матеріалів; '
                f'станом на {r.get("судді_станом_на") or r["стан_на"]} у суді не було суддів, які здійснювали повноваження '
                f'(з {fmt(vrp)}, {vrp_word}; за нормативами потрібно {fmt(nrm)}), тому навантаження на одного суддю не розраховується.')
    t = (f'{lead} за {r["період"]} до {sud} надійшло {fmt(f(r["надійшло"]))} справ і матеріалів; '
         f'станом на {r.get("судді_станом_на") or r["стан_на"]} кількість суддів, які здійснювали повноваження, '
         f'становила {fmt(act)} з {fmt(vrp)}, {vrp_word}, тоді як за нормативами потрібно {fmt(nrm)}. Отже, на одного суддю з повноваженнями припадало '
         f'близько {fmt(r["_na_suddiu"])} справ (у середньому {fmt(f(r["на_суддю_на_місяць"]))} на місяць), а для '
         f'розгляду справ і матеріалів, що надійшли за цей період, одному судді за нормативним часом '
         f'потрібно було б близько {fmt(f(r["днів_на_суддю"]))} робочих днів.')
    if r.get('_rank_reg'):
        k, n = r['_rank_reg']
        t += f' За кількістю справ на одного суддю суд посідав {k} місце серед {n} {KIND_GEN[kind]} {reg}'
        if r.get('_rank_ua'):
            k2, n2 = r['_rank_ua']
            t += f' і {k2} місце серед {n2} {KIND_GEN[kind]} України.'
        else:
            t += '.'
    elif r.get('_rank_ua'):
        k2, n2 = r['_rank_ua']
        t += f' За кількістю справ на одного суддю суд посідав {k2} місце серед {n2} {KIND_GEN[kind]} України.'
    return t


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('sud')
    ap.add_argument('--data')
    ap.add_argument('--period')
    ap.add_argument('--dir', default=HERE, help='тека з файлами довідника (за замовчуванням — тека скрипта)')
    a = ap.parse_args()
    rows, posady = load(a.dir)
    names = find(rows, a.sud)
    if not names:
        sys.exit(f'Суд «{a.sud}» не знайдено. Спробуйте коротшу частину назви.')
    if len(names) > 6:
        print('Знайдено забагато судів, уточніть назву:'); [print(' -', n) for n in names[:30]]; return
    for n in names:
        rs = sorted([r for r in rows if r['_n'] == n], key=lambda r: r['період_код'])
        r0 = rs[-1]
        print(f'\n=== {r0["суд"]} ({r0["вид_суду"]}{", " + r0["регіон"] if r0["регіон"] else ""})')
        if r0['_old']:
            print('Попередні назви: ' + '; '.join(f'{nm} (дані до {last})' for nm, last in r0['_old']))
        allp = sorted({r['період_код'] for r in rows})
        if rs[0]['період_код'] != allp[0] or rs[-1]['період_код'] != allp[-1]:
            print(f'⚠ Дані цього суду є лише за {rs[0]["період"]} – {rs[-1]["період"]} (довідник: {allp[0]} – {allp[-1]}). '
                  'Можливі причини: суд перейменовано (шукайте й за попередньою / новою назвою), змінено підсудність, '
                  'суд не працює. Зазначити у звіті.')
        print('період | надійшло | ВРП / з повн. / норм. | справ на суддю | на місяць | днів на суддю | місце обл. | місце Укр. | джерело')
        for r in rs:
            reg = '{}/{}'.format(*r['_rank_reg']) if r.get('_rank_reg') else '—'
            ua = '{}/{}'.format(*r['_rank_ua']) if r.get('_rank_ua') else '—'
            src = 'РСУ' if 'rsu.gov.ua' in r['примітка'] else 'ДСА'
            print(f'{r["період"]} | {fmt(f(r["надійшло"]))} | {r["суддів_ВРП"]} / {r["суддів_з_повноваженнями"]} / '
                  f'{r["суддів_за_нормативами"]} | {fmt(r["_na_suddiu"])} | {r["на_суддю_на_місяць"]} | '
                  f'{fmt(f(r["днів_на_суддю"]))} | {reg} | {ua} | {src}')
        pick = []
        if a.period:
            pick = [r for r in rs if r['період_код'] == a.period]
        elif a.data:
            dd = d(a.data)
            before = [r for r in rs if d(r['стан_на']) < dd]
            if before:
                pick.append(before[-1])
                years = [r for r in before if r['період_код'].endswith('_4')]
                if years and years[-1] is not before[-1]:
                    pick.append(years[-1])
                elif not years:
                    print(f'\n⚠ Повного року до {a.data} для цього суду в довіднику немає.')
            else:
                print(f'\n⚠ До {a.data} у довіднику немає жодного періоду для цього суду '
                      f'(найраніший закінчується {rs[0]["стан_на"]}) — абзац не складено.')
        if a.period and not pick:
            print(f'\n⚠ Періоду {a.period} для цього суду в довіднику немає.')
        for r in pick:
            print(f'\n--- Абзац для висновку ({r["період"]}):')
            print(paragraph(r, posady))
            if r['примітка']:
                print(f'   [примітка: {r["примітка"]}]')
        if posady:
            ids = {r['_id'] for r in rs}
            pp = [p for p in posady if p['id'] in ids]
            lim = d(a.data) if a.data else (d(pick[0]['стан_на']) if pick else None)
            if lim:
                pp = [p for p in pp if d(p['стан_на']) <= lim] or pp
            if pp:
                p = pp[-1]
                extra = (f', з них щодо {p["триває_добір"]} триває процедура зайняття'
                         if p['триває_добір'] not in ('', '0') else '')
                print(f'\n--- ВККСУ ({POSADY_SRC.get(p["стан_на"], "")}): станом на {p["стан_на"]} гранична кількість посад суддів — '
                      f'{p["гранична"]}, призначено (обрано) — {p["призначено"]}, вакантних — {p["вакантні"]}{extra}.')


if __name__ == '__main__':
    main()
