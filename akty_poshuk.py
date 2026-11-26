#!/usr/bin/env python3
"""Пошук актів призначення / обрання / переведення судді у «Базі актів призначення суддів» (Excel на Google Диску).

Використання:
    python3 akty_poshuk.py "Прізвище [І.Б. | Ім'я По батькові]" [--sud "суд"] [--akt] [--onovyty] [--base файл.xlsx]
    python3 akty_poshuk.py 00161                 # за номером досьє ВККС
Приклади:
    python3 akty_poshuk.py "Євграфова Є.П."
    python3 akty_poshuk.py "Кифлюк Володимир Федорович" --sud "Тернопільський"
    python3 akty_poshuk.py "Прізвище І.Б." --akt   # рядок для таблиці висновку («@акт:»)

Джерело — «База_актів_призначення_суддів.xlsx» у головній теці
https://drive.google.com/drive/folders/1lK8gmOJoflDSMCIlFaSeX3OEB_As2Myh (файл шукається в теці за назвою «База_актів_призначення_суддів….xlsx»; запасний ID — 1F5b3YfxUIQCKQ9pvSRIiCnjoUsaqbTaJ; аркуш «База»).
Скрипт сам завантажує файл з Диска (доступ за посиланням, curl/urllib через проксі робочого середовища) і кешує його
в ~/.cache/akty_baza/ на 24 години; --onovyty — завантажити заново; --base — взяти локальний файл замість Диска.
Якщо Диск недоступний, береться кеш (з попередженням). Звідки взято базу і коли — скрипт друкує першим рядком;
це зазначати у звіті.

Друкує кар'єру судді в хронологічному порядку з посиланнями на акти — у формі, придатній для висновку
(«Указом Президента України від … № … призначена на посаду судді … строком на п'ять років»).
Акти з сайту ВРУ — пріоритетне джерело. Примітка «Можлива зміна прізвища» означає, що акт виданий щодо особи з іншим
прізвищем (ім'я, по батькові та суд збігаються) — перед використанням у висновку перевірити.
Потрібен openpyxl (pip install openpyxl --break-system-packages, якщо його немає).

Версія 1.2 (26.11.2026): `--sud` перевіряється і тоді, коли знайдено одного суддю: суд у базі інший — попередження
(можливе відрядження або тезка), раніше мовчки показувався суддя іншого суду; якщо файл бази в теці Диска за назвою
не знайдено і взято запасний ID, це зазначено в рядку «Джерело» (база може бути застарілою).
Версія 1.1 (25.11.2026): дата списку ВККС і дата бази беруться з самого файла (раніше «19.10.2026» було вшито);
попередження, якщо `--sud` не збігся з жодним судом; `--sud` розуміє «м.» = «міста» (лише в назві суду: у ПІБ «М.» —
ініціал) і суд у непрямому відмінку («Печерського районного суду» — порівняння із suddi_poshuk.py); із кількох файлів бази
без точної назви береться той, що має найновішу дату в назві; запуск без ПІБ — довідка.
"""
import datetime as dt
import html
import os
import re
import shutil
import sys
import time
import urllib.request

FILE_ID = '1F5b3YfxUIQCKQ9pvSRIiCnjoUsaqbTaJ'   # запасний: якщо файл у теці не знайдено за назвою
FOLDER_ID = '1lK8gmOJoflDSMCIlFaSeX3OEB_As2Myh'
FOLDER = f'https://drive.google.com/drive/folders/{FOLDER_ID}'
BASE_NAME = 'База_актів_призначення_суддів'
CACHE_DIR = os.path.expanduser('~/.cache/akty_baza')
CACHE = os.path.join(CACHE_DIR, 'База_актів_призначення_суддів.xlsx')
MAX_AGE = 24 * 3600

TYPE = {'Указ Президента України': 'Указом Президента України',
        'Постанова Верховної Ради України': 'Постановою Верховної Ради України',
        'Рішення Вищої ради правосуддя': 'Рішенням Вищої ради правосуддя',
        'Рішення ради (місцевої/обласної)': 'Рішенням ради народних депутатів'}
MON = ['січня', 'лютого', 'березня', 'квітня', 'травня', 'червня', 'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня']


def norm(s):
    s = (s or '').lower().replace('’', "'").replace('`', "'").replace('ʼ', "'")
    return re.sub(r'\s+', ' ', s).strip()


def norm_court(s):
    """Назва суду для --sud: «м.» = «міста». Лише для суду — у ПІБ «М.» є ініціалом («Мельник М.О.»)."""
    return re.sub(r'\s+', ' ', re.sub(r'\bм\.\s*', 'міста ', norm(s))).strip()


def same_court(sud, court):
    """Суд із --sud проти суду з бази: порівняння suddi_poshuk.py (відмінок «Печерського районного суду», «м.», описка);
    якщо того файла поруч немає — входження підрядка."""
    try:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from suddi_poshuk import same_court as sc
        if sc(sud, court):
            return True
    except ImportError:
        pass
    return norm_court(sud) in norm_court(court)


def title_date(title):
    """Дата в назві файла бази (ДД.ММ.РРРР або РРРР-ММ-ДД) для вибору найновішого; без дати — найраніший."""
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', title)
    if m:
        return f'{m[3]}-{m[2]}-{m[1]}'
    m = re.search(r'\d{4}-\d{2}-\d{2}', title)
    return m.group(0) if m else ''


def stamp(path):
    return dt.datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d.%m.%Y %H:%M')


FALLBACK_USED = False


def file_id():
    """ID файлу бази: шукає в теці Диска файл «База_актів_призначення_суддів….xlsx» за назвою
    (якщо базу завантажено заново, її ID змінюється); не знайдено чи тека недоступна — запасний FILE_ID."""
    try:
        req = urllib.request.Request(f'https://drive.google.com/embeddedfolderview?id={FOLDER_ID}',
                                     headers={'User-Agent': 'akty_poshuk'})
        with urllib.request.urlopen(req, timeout=60) as r:
            page = r.read().decode('utf-8', 'replace')
        found = []
        for m in re.finditer(r'id="entry-([^"]+)".*?class="flip-entry-title">([^<]+)', page, re.S):
            title = html.unescape(m.group(2)).strip()
            if title.startswith(BASE_NAME) and title.lower().endswith('.xlsx'):
                found.append((title != BASE_NAME + '.xlsx', title, m.group(1)))
        if found:                    # точна назва; інакше — з найновішою датою в назві
            exact = [x for x in found if not x[0]]
            return (exact or sorted(found, key=lambda x: (title_date(x[1]), x[1])))[-1][2]
    except Exception:
        pass
    global FALLBACK_USED
    FALLBACK_USED = True
    return FILE_ID


def download():
    os.makedirs(CACHE_DIR, exist_ok=True)
    tmp = CACHE + '.part'
    url = f'https://drive.usercontent.google.com/download?id={file_id()}&export=download&confirm=t'
    req = urllib.request.Request(url, headers={'User-Agent': 'akty_poshuk'})
    with urllib.request.urlopen(req, timeout=120) as r, open(tmp, 'wb') as f:
        shutil.copyfileobj(r, f)
    with open(tmp, 'rb') as f:
        if f.read(2) != b'PK':
            raise RuntimeError('Диск повернув не xlsx (можливо, змінено доступ за посиланням)')
    os.replace(tmp, CACHE)


def get_base(local, refresh):
    if local:
        return local, f'локальний файл {local} (змінено {stamp(local)})'
    fresh = os.path.exists(CACHE) and time.time() - os.path.getmtime(CACHE) < MAX_AGE
    if refresh or not fresh:
        try:
            download()
            fb = ('; файл бази в теці за назвою не знайдено — узято запасний ID, база може бути застарілою'
                  if FALLBACK_USED else '')
            return CACHE, f'Google Диск, {FOLDER} — завантажено {stamp(CACHE)}{fb}'
        except Exception as e:
            if not os.path.exists(CACHE):
                sys.exit(f'Не вдалося завантажити базу з Диска ({e}). Завантажте «База_актів_призначення_суддів.xlsx» '
                         f'з {FOLDER} і передайте --base шлях.')
            print(f'⚠ Диск недоступний ({e}); використано кеш від {stamp(CACHE)}.', file=sys.stderr)
            return CACHE, f'кеш копії з Google Диска від {stamp(CACHE)} (Диск був недоступний)'
    return CACHE, f'Google Диск, {FOLDER} — кеш від {stamp(CACHE)}'


def load(path):
    try:
        import openpyxl
    except ImportError:
        sys.exit('Немає openpyxl: pip install openpyxl --break-system-packages')
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['База']
    head = ' '.join(str(c or '') for c in next(ws.iter_rows(min_row=1, max_row=1, max_col=15, values_only=True)))
    m = re.search(r'список ВККС\w* на (\d\d\.\d\d\.\d{4})', head)
    INFO['spysok'] = m.group(1) if m else None
    INFO['zmineno'] = wb.properties.modified.strftime('%d.%m.%Y') if wb.properties.modified else None
    rows, cur = [], None
    for r in ws.iter_rows(min_row=2, max_col=15, values_only=True):
        r = list(r) + [None] * (15 - len(r))
        if r[1]:  # новий блок судді (колонки A–G об'єднані, значення лише в першому рядку)
            cur = {'dos': str(r[1]).strip().zfill(5), 'pib': r[2] or '', 'sex': r[3] or '', 'court': r[4] or '',
                   'inst': r[5] or '', 'acts': []}
            rows.append(cur)
        if cur is None or not r[7] or r[7] == 'відомостей не знайдено':
            continue
        cur['acts'].append({'kind': r[7], 'date': r[8], 'num': str(r[9] or '').strip(), 'verb': r[10] or '',
                            'term': r[11] or '', 'court': r[12] or '', 'url': r[13] or '', 'note': r[14] or ''})
    return rows


INFO = {}


def pokryttia():
    """«список суддів ВККС на …; базу змінено …» — з файла бази, а не з коду."""
    a = f"список суддів ВККС на {INFO['spysok']}" if INFO.get('spysok') else 'дата списку суддів ВККС у базі не зазначена'
    return a + (f"; базу змінено {INFO['zmineno']}" if INFO.get('zmineno') else '')


def human_date(d):
    if isinstance(d, (dt.date, dt.datetime)):
        return f'{d.day} {MON[d.month - 1]} {d.year} року'
    m = re.match(r'(\d\d)\.(\d\d)\.(\d{4})', str(d or ''))
    return f'{int(m[1])} {MON[int(m[2]) - 1]} {m[3]} року' if m else 'дата невідома'


def find(rows, q):
    q = q.strip()
    if re.fullmatch(r'\d{4,5}', q):
        return [r for r in rows if r['dos'] == q.zfill(5)]
    qq = norm(q).replace('.', '. ').split()
    sur, rest = qq[0], [x.strip('.') for x in qq[1:] if x.strip('.')]
    res = []
    for r in rows:
        p = norm(r['pib']).split()
        if not p or p[0] != sur:
            continue
        if all(i + 1 < len(p) and p[i + 1].startswith(part) for i, part in enumerate(rest[:2])):
            res.append(r)
    return res


def sentence(a, he):
    if he is None:                       # стать у базі не зазначена — обидві форми, щоб не вгадувати
        v = {'призначення': 'призначений (призначена)', 'обрання': 'обраний (обрана)',
             'переведення': 'переведений (переведена)'}.get(a['verb'], a['verb'])
    else:
        v = {'призначення': 'призначений' if he else 'призначена', 'обрання': 'обраний' if he else 'обрана',
             'переведення': 'переведений' if he else 'переведена'}.get(a['verb'], a['verb'])
    c = a['court']
    tr = re.match(r'^(.*) \(з (.*)\)$', c)
    pos = f'на посаду судді {tr.group(1)} (з посади судді {tr.group(2)})' if tr else f'на посаду судді {c}'
    t = f" {a['term']}" if a['term'] else ''
    num = f" № {a['num']}" if a['num'] else ''
    return f"{TYPE.get(a['kind'], a['kind'])} від {human_date(a['date'])}{num} {v} {pos}{t}"


def main():
    args = sys.argv[1:]
    if not args or args[0] in ('-h', '--help'):
        print(__doc__); return
    local = sud = None
    if '--base' in args:
        k = args.index('--base'); local = args[k + 1]; del args[k:k + 2]
    if '--sud' in args:
        k = args.index('--sud'); sud = args[k + 1]; del args[k:k + 2]
    akt = '--akt' in args
    refresh = '--onovyty' in args
    args = [a for a in args if a not in ('--akt', '--onovyty')]
    if not args or not args[0].strip():
        sys.exit('Не задано ПІБ або номер досьє. Приклад: python3 akty_poshuk.py "Прізвище І.Б." --akt')
    path, src = get_base(local, refresh)
    rows = load(path)
    print(f'Джерело: «База_актів_призначення_суддів.xlsx» ({src}); {pokryttia()}.')
    found = find(rows, args[0])
    if sud and found:
        here = [r for r in found if same_court(sud, r['court'])]
        if not here:
            print(f'⚠ --sud «{sud}» не збігся з судом ' + ('знайденого судді' if len(found) == 1 else 'жодного зі знайдених суддів')
                  + ' — можливе відрядження або тезка; показано ' + ('його' if len(found) == 1 else 'всіх') + '.')
        found = here or found
    if not found:
        print(f'«{args[0]}» у базі не знайдено (перевірте написання — потрібен називний відмінок; база охоплює діючих суддів: {pokryttia()}).')
        return
    if len(found) > 1:
        print(f'⚠ Знайдено {len(found)} суддів — уточніть ім\'я, по батькові або --sud.')
    for r in found:
        he = {'Ж': False, 'Ч': True}.get(str(r['sex']).strip().upper())
        print(f"\n{r['pib']} — {r['court']} (досьє № {r['dos']})")
        if not r['acts']:
            print('  Відомостей про акти не знайдено → [ПОТРЕБУЄ ДОЗАПОВНЕННЯ: акт про призначення]'); continue
        if akt:
            parts = [f"{a['kind']} від {human_date(a['date'])}" + (f" № {a['num']}" if a['num'] else '') for a in r['acts']]
            print('  @акт: ' + '; '.join(parts))
        for a in r['acts']:
            print(f"  • {sentence(a, he)}" + (f" ({a['url']})" if a['url'] else '') + '.')
            if a['note']:
                print(f"      ⚠ {a['note']}")


if __name__ == '__main__':
    main()
