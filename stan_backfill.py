#!/usr/bin/env python3
"""Ретроархів «Інформація щодо стану розгляду справ»: зміни стану за минулі місяці з щоденних архівів ДСА.

Версія 2.0 (10.11.2026). Зведено дві гілки 1.x і виправлено помилки запуску 01.2025–09.2026:
  * пошук у масиві «відомого» через np.uint64 (раніше — катастрофічне уповільнення опорного дня, до 2+ год);
  * стійкі завантаження: повтори, обрив завислих з'єднань (< 10 КБ/с 60 с), докачування; перевірка цілісності
    zip ще в потоці завантаження (битий архів завантажується заново, а не валить обробку);
  * день, який сервер ДСА не віддає, пропускається (SKIP, `skipped_days` у .json), а не зупиняє місяць;
  * опорний день, що не завантажується, замінюється попереднім доступним (до 3 спроб) замість ERROR;
  * формат до 19.11.2023: один CSV замість zip, відносні посилання, дні «0.00 KB» на сайті, 404 / HTML-заглушка;
  * prepare читає «відоме» і з індексу parquet (stan_istoriia_*.parquet), не лише з .tsv.xz; пам'ять — array('Q');
  * zvit — підсумкова таблиця за місяцями; run друкує, скільки місяців у діапазоні, оброблено й залишилося.
Правила відбору не змінювалися (ті самі, що для жовтня 2026 і для 01.2025–09.2026).

Джерело — сайт ДСА https://dsa.court.gov.ua/dsa/inshe/oddata/532/?page=N (щоденні архіви
«Стан розгляду справ від ДД.ММ.РРРР.zip|csv» з 11.2019; ті самі файли, що й на data.gov.ua).
Місяці обробляються у ЗВОРОТНОМУ порядку: «відомим» уважається все, що вже збережено пізніше
(вихідний знімок на 30.10.2026, ретроархів 01.2025–09.2026 / індекс, оброблені місяці, за наявності — репозиторій).

Правила відбору:
  * стан провадження (суд + справа + провадження + вид) = registration_date, judge, judges, stage_date,
    stage_name, cause_result, cause_dep; зміни лише participants / description не зберігаються;
  * версія запису зберігається один раз — у день, коли її вперше видно, — якщо такого стану в провадження
    не було ні на початок місяця, ні раніше в місяці, і він не збережений пізніше;
  * «нове» зі стадією й реєстрацією раніше за перше число попереднього місяця (застарілий запис) —
    лише якщо стан цього провадження далі змінився (у місяці або пізніше);
  * стан на останній доступний день попереднього місяця («змінене_було», повний вміст) — лише для проваджень,
    стан яких змінився в місяці.
Результат: <out>/stan_zminy_РРРР-ММ.tsv.xz (snapshot_date, change, row_hash + 13 полів ДСА) і .json зі статистикою.

Команди:
  stan_backfill.py list    --work W                                   # перелік файлів ДСА → W/files_all.json
  stan_backfill.py prepare --work W [--base BASE.tsv.xz] [--layers F.tsv.xz ...] [--parquet 'D/stan_istoriia_*.parquet']
  stan_backfill.py month   2024-12 --work W --out OUT                 # один місяць
  stan_backfill.py run     --from 2024-11 --to 2023-01 --work W --out OUT [--max-hours 1.8]
  stan_backfill.py check   2024-12 --out OUT [--examples 3 --court "Печерський"]
  stan_backfill.py zvit    --out OUT [--from 2023-01 --to 2024-12] [--porig 1000000]
Автор: В'ячеслав Плескач
"""
import argparse, array, collections, concurrent.futures as cf, csv, datetime, glob, hashlib, io, json, lzma, os, re
import subprocess, sys, threading, time, zipfile
import numpy as np

csv.field_size_limit(10 ** 9)
HEAD = ['court_name', 'case_number', 'case_proc', 'registration_date', 'judge', 'judges', 'participants',
        'stage_date', 'stage_name', 'cause_result', 'cause_dep', 'type', 'description']
OUT = ['snapshot_date', 'change', 'row_hash'] + HEAD
SITE = 'https://dsa.court.gov.ua/dsa/inshe/oddata/532/?page={}'
CURL = ['curl', '-sS', '-L', '-f', '--retry', '10', '--retry-all-errors', '--retry-delay', '3',
        '--speed-limit', '10000', '--speed-time', '60', '-m', '2400']


def say(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


# ---------- дані ----------
def src_rows(path):
    """Рядки архіву ДСА: zip із кількома CSV (з 21.11.2023) або один CSV (до 19.11.2023)."""
    def from_text(t, name):
        r = csv.reader(t, delimiter='\t')
        h = [c.strip('"').strip('﻿') for c in next(r, [])]
        if h[:13] != HEAD:
            sys.exit(f'Неочікуваний заголовок у {name}: {h[:13]}')
        for x in r:
            x = x + [''] * (13 - len(x))
            yield x[:13]
    if zipfile.is_zipfile(path):
        z = zipfile.ZipFile(path)
        for nm in sorted(z.namelist(), key=lambda s: [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', s)]):
            if nm.lower().endswith('.csv'):
                with z.open(nm) as f:
                    yield from from_text(io.TextIOWrapper(f, encoding='utf-8', errors='replace', newline=''), f'{path}/{nm}')
    else:
        with open(path, encoding='utf-8', errors='replace', newline='') as f:
            yield from from_text(f, path)


def valid(path):
    """'ok' — файл придатний; 'missing' — на сайті немає (HTML-заглушка); 'bad' — битий / неповний."""
    try:
        if not os.path.exists(path) or os.path.getsize(path) == 0:
            return 'bad'
        with open(path, 'rb') as t:
            head = t.read(64)
        if head.lstrip().lower().startswith((b'<!doctype', b'<html')):
            return 'missing'
        if head[:2] == b'PK':
            with zipfile.ZipFile(path) as z:
                return 'ok' if z.testzip() is None and any(n.lower().endswith('.csv') for n in z.namelist()) else 'bad'
        return 'ok' if head.lstrip(b'\xef\xbb\xbf').startswith(b'"court_name"') else 'bad'
    except Exception:
        return 'bad'


def arch_rows(path):
    with lzma.open(path, 'rt', encoding='utf-8', newline='') as f:
        r = csv.reader(f, delimiter='\t')
        h = next(r)
        if h[:3] != OUT[:3]:
            sys.exit(f'Неочікуваний заголовок у {path}')
        for x in r:
            yield x


def dig(x): return hashlib.blake2b('\t'.join(x).encode('utf-8'), digest_size=12).digest()
def kdig(x): return hashlib.blake2b('\x1f'.join((x[0], x[1], x[2], x[11])).encode('utf-8'), digest_size=8).digest()
def sdig(x): return hashlib.blake2b('\x1f'.join((x[3], x[4], x[5], x[7], x[8], x[9], x[10])).encode('utf-8'), digest_size=8).digest()
def kint(k): return int.from_bytes(k, 'little')
def pint(k, s): return int.from_bytes(hashlib.blake2b(k + s, digest_size=8).digest(), 'little')
def iso(s): return f'{s[6:10]}-{s[3:5]}-{s[0:2]}' if re.fullmatch(r'\d\d\.\d\d\.\d{4}', s or '') else ''


def addset(d, k, v):
    c = d.get(k)
    if c is None: d[k] = v
    elif isinstance(c, bytes):
        if c != v: d[k] = {c, v}
    else: c.add(v)
def inset(d, k, v):
    c = d.get(k)
    return c is not None and (c == v if isinstance(c, bytes) else v in c)
def size(d, k):
    c = d.get(k)
    return 0 if c is None else (1 if isinstance(c, bytes) else len(c))


# ---------- перелік файлів ДСА ----------
def cmd_list(a):
    def page(p):
        r = subprocess.run(['curl', '-sS', '-m', '90', '--retry', '3', SITE.format(p)], capture_output=True)
        s = r.stdout.decode('cp1251', 'replace')
        out = []
        for m in re.finditer(r'download="([^"]+)"\s+href="([^"]+)"', s):
            sz = re.search(r'([\d.]+\s*[KMG]?B)', re.sub(r'<[^>]+>', ' ', s[m.end():m.end() + 600]))
            out.append((p, m.group(1), m.group(2), sz.group(1) if sz else ''))
        return out
    first = subprocess.run(['curl', '-sS', '-m', '90', '--retry', '3', SITE.format(1)], capture_output=True).stdout.decode('cp1251', 'replace')
    n = max(int(x) for x in re.findall(r'page=(\d+)', first))
    rows = []
    with cf.ThreadPoolExecutor(8) as ex:
        for r in ex.map(page, range(1, n + 1)): rows += r
    files = {}
    for p, name, url, sz in rows:
        m = re.search(r'від (\d\d)\.(\d\d)\.(\d{4})', name)
        if m:
            if url.startswith('/'): url = 'https://dsa.court.gov.ua' + url   # до 11.2023 посилання відносні
            files[f'{m[3]}-{m[2]}-{m[1]}'] = (url, sz)
    json.dump(files, open(os.path.join(a.work, 'files_all.json'), 'w'), ensure_ascii=False, indent=0, sort_keys=True)
    say(f'сторінок {n}, щоденних архівів {len(files)}: {min(files)} … {max(files)}; '
        f'порожніх на сайті (0 КБ): {sum(1 for v in files.values() if missing(v[1]))}')


def files_for(work):
    f = os.path.join(work, 'files_all.json')
    if not os.path.exists(f):
        sys.exit('Немає files_all.json — спершу: stan_backfill.py list')
    return json.load(open(f))


def missing(sz): return sz.strip() in ('0.00 KB', '0 KB', '0.00 B', '')


# ---------- індекс «відомого» ----------
def idx_paths(work): return os.path.join(work, 'known_pairs.npy'), os.path.join(work, 'known_keys.npy')


def load_idx(work):
    p, k = idx_paths(work)
    if not os.path.exists(p): sys.exit('Немає індексу — спершу: stan_backfill.py prepare')
    return np.load(p), np.load(k)


def merge_idx(work, pairs, keys):
    p, k = idx_paths(work)
    P = np.unique(np.frombuffer(pairs, dtype=np.uint64) if isinstance(pairs, array.array) else np.array(pairs, dtype=np.uint64))
    K = np.unique(np.frombuffer(keys, dtype=np.uint64) if isinstance(keys, array.array) else np.array(keys, dtype=np.uint64))
    if os.path.exists(p):
        P = np.union1d(np.load(p), P); K = np.union1d(np.load(k), K)
    np.save(p + '.tmp.npy', P); np.save(k + '.tmp.npy', K)
    os.replace(p + '.tmp.npy', p); os.replace(k + '.tmp.npy', k)
    return len(P), len(K)


def cmd_prepare(a):
    pairs, keys = array.array('Q'), array.array('Q')
    def add(x):
        k = kdig(x); pairs.append(pint(k, sdig(x))); keys.append(kint(k))
    if a.base:
        for x in arch_rows(a.base): add(x[3:])
        say('вихідний знімок', a.base, len(pairs))
    for f in a.layers or []:
        n0 = len(pairs)
        for x in arch_rows(f):
            # «змінене_було» / «зникле» щоденних файлів після 30.10 мають лише ключові поля — їх пропускаємо
            if x[1] in ('нове', 'змінене', 'змінене_було', 'знімок') and x[3] and (x[6] or x[11]):
                add(x[3:])
        say('шар', os.path.basename(f), len(pairs) - n0)
    if a.parquet:
        try:
            import duckdb
        except ImportError:
            subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', '--break-system-packages', 'duckdb'], check=True)
            import duckdb
        fs = sorted(glob.glob(a.parquet))
        if not fs: sys.exit(f'Немає файлів {a.parquet}')
        con = duckdb.connect(); con.execute('SET threads=2')
        cols = ', '.join(f"coalesce({c},'')" for c in HEAD)
        for f in fs:
            n0 = len(pairs)
            cur = con.execute(f"select {cols} from read_parquet('{f}') where coalesce(court_name,'')<>'' and (coalesce(registration_date,'')<>'' or coalesce(stage_name,'')<>'')")
            while True:
                b = cur.fetchmany(200000)
                if not b: break
                for x in b: add(list(x))
            say('індекс', os.path.basename(f), len(pairs) - n0)
    nP, nK = merge_idx(a.work, pairs, keys)
    say(f'індекс: станів {nP}, проваджень {nK}')


# ---------- один місяць ----------
def month_days(files, ym):
    files = {d: v for d, v in files.items() if not missing(v[1])}
    y, m = map(int, ym.split('-'))
    d0 = datetime.date(y, m, 1)
    d1 = (d0 + datetime.timedelta(days=32)).replace(day=1)
    days = sorted(d for d in files if d0.isoformat() <= d < d1.isoformat())
    prev = sorted(d for d in files if d < d0.isoformat())
    fresh = (d0 - datetime.timedelta(days=1)).replace(day=1).isoformat()
    return prev[-3:][::-1], days, fresh


def fetch_file(url, f):
    """Повертає шлях до перевіреного файла, 'missing' (на сайті немає) або None (не завантажується)."""
    for att in range(4):
        r = subprocess.run(CURL + ['-C', '-', '-o', f, '-w', '%{http_code}', url], capture_output=True, text=True)
        code = r.stdout.strip()
        v = valid(f)
        if v == 'ok':
            return f
        if code == '404' or v == 'missing':
            try: os.remove(f)
            except OSError: pass
            return 'missing'
        if r.returncode == 0 or code in ('416', '200'):      # файл «завантажено», але він битий — качати з нуля
            try: os.remove(f)
            except OSError: pass
        time.sleep(30)
    try: os.remove(f)
    except OSError: pass
    return None


def cmd_month(a, ym=None):
    ym = ym or a.month
    allf = files_for(a.work)
    refs, days, fresh = month_days(allf, ym)
    if not refs or not days: sys.exit(f'{ym}: немає файлів ДСА')
    os.makedirs(a.out, exist_ok=True)
    dl = os.path.join(a.work, 'dl'); tmpd = os.path.join(a.work, f'tmp_{ym}')
    os.makedirs(dl, exist_ok=True); os.makedirs(tmpd, exist_ok=True)
    KP, KK = load_idx(a.work)
    def kp(v): v = np.uint64(v); i = np.searchsorted(KP, v); return i < len(KP) and KP[i] == v
    def kk(v): v = np.uint64(v); i = np.searchsorted(KK, v); return i < len(KK) and KK[i] == v
    t_start = time.time()
    gaps = sorted(d for d, v in allf.items() if missing(v[1]) and d[:7] == ym)
    say(f'{ym}: днів {len(days)}, «свіжі» з {fresh}; індекс {len(KP)} станів'
        + (f'; на сайті порожні (0 КБ): {", ".join(gaps)}' if gaps else ''))

    # опорний день — синхронно, з запасними
    ref = rf = None; ref_tried = []
    for d in refs:
        r = fetch_file(allf[d][0], os.path.join(dl, f'{d}.dat'))
        ref_tried.append(d)
        if r not in (None, 'missing'):
            ref, rf = d, r; break
        say(f'опорний день {d} не завантажується — пробую попередній')
    if not ref: sys.exit(f'ERROR опорний день: жоден із {ref_tried} не завантажується')

    got = {}; cond = threading.Condition(); sem = threading.Semaphore(2)
    def fetch(d):
        r = fetch_file(allf[d][0], os.path.join(dl, f'{d}.dat'))
        with cond: got[d] = r; cond.notify_all()
    def feeder():
        for d in days:
            sem.acquire(); threading.Thread(target=fetch, args=(d,), daemon=True).start()
    threading.Thread(target=feeder, daemon=True).start()

    seen_h = set(); seen_k = set(); stat = {}; mon_s = {}; ref_cand = {}
    stats = []; skipped = []; absent = []
    for d in [ref] + days:
        if d == ref:
            f = rf
        else:
            with cond:
                while d not in got: cond.wait()
                f = got.pop(d)
            sem.release()
        if f in (None, 'missing'):
            (absent if f == 'missing' else skipped).append(d)
            say(f'SKIP {d}: ' + ('файла на сайті немає (404)' if f == 'missing' else 'файл не завантажується') + ' — день пропущено')
            stats.append({'date': d, 'rows': 0, 'skipped': 'немає_на_сайті' if f == 'missing' else 'не_завантажено'})
            json.dump(stats, open(os.path.join(tmpd, 'stats.json'), 'w'), ensure_ascii=False)
            continue
        t0 = time.time(); st = collections.Counter(); day_k = set()
        w = None
        if d != ref:
            fo = open(os.path.join(tmpd, f'{d}.tsv'), 'w', encoding='utf-8', newline='')
            w = csv.writer(fo, delimiter='\t', lineterminator='\n')
        for x in src_rows(f):
            st['rows'] += 1
            h = dig(x)
            if h in seen_h: st['вже_було'] += 1; continue
            seen_h.add(h); k = kdig(x); s = sdig(x); day_k.add(k)
            if d == ref:
                addset(stat, k, s)
                if not kp(pint(k, s)): ref_cand[h] = x; st['кандидати_відліку'] += 1
                continue
            if inset(stat, k, s): st['без_зміни_стану'] += 1; continue
            if kp(pint(k, s)): st['збережено_пізніше'] += 1; continue
            addset(stat, k, s); addset(mon_s, k, s)
            if k in seen_k: w.writerow(['змінене', h.hex()] + x); st['змінене'] += 1
            elif iso(x[7]) >= fresh or iso(x[3]) >= fresh: w.writerow(['нове', h.hex()] + x); st['нове'] += 1
            else: w.writerow(['нове?', h.hex()] + x); st['нове_застаріле'] += 1
        if w: fo.close()
        seen_k |= day_k; os.remove(f)
        st['date'] = d; st['src_size'] = allf[d][1]; st['sec'] = round(time.time() - t0)
        stats.append(dict(st)); say(json.dumps(st, ensure_ascii=False))
        json.dump(stats, open(os.path.join(tmpd, 'stats.json'), 'w'), ensure_ascii=False)

    def changed(k, s):
        return kk(kint(k)) or size(mon_s, k) > (1 if inset(mon_s, k, s) else 0)
    fin = collections.Counter(); pairs = array.array('Q'); keys = array.array('Q')
    tsv = os.path.join(tmpd, 'all.tsv')
    with open(tsv, 'w', encoding='utf-8', newline='') as g:
        w = csv.writer(g, delimiter='\t', lineterminator='\n'); w.writerow(OUT)
        for h, x in ref_cand.items():
            k = kdig(x); s = sdig(x)
            if changed(k, s):
                w.writerow([ref, 'змінене_було', h.hex()] + x); fin['змінене_було'] += 1
                pairs.append(pint(k, s)); keys.append(kint(k))
        ref_cand.clear()
        for d in days:
            if d in skipped or d in absent: continue
            with open(os.path.join(tmpd, f'{d}.tsv'), encoding='utf-8', newline='') as r:
                for row in csv.reader(r, delimiter='\t'):
                    c, h, x = row[0], row[1], row[2:]
                    k = kdig(x); s = sdig(x)
                    if c == 'нове?':
                        if not changed(k, s): fin['відкинуто_застаріле'] += 1; continue
                        c = 'нове'
                    w.writerow([d, c, h] + x); fin[c] += 1
                    pairs.append(pint(k, s)); keys.append(kint(k))
    dst = os.path.join(a.out, f'stan_zminy_{ym}.tsv.xz')
    subprocess.run(['xz', '-T0', '-6', '-f', tsv], check=True)
    os.replace(tsv + '.xz', dst)
    nP, nK = merge_idx(a.work, pairs, keys)
    info = {'month': ym, 'reference': ref, 'reference_tried': ref_tried, 'days': days, 'skipped_days': skipped,
            'absent_days_404': absent, 'empty_on_site': gaps, 'fresh_from': fresh, 'result': fin,
            'file': os.path.basename(dst), 'file_bytes': os.path.getsize(dst),
            'minutes': round((time.time() - t_start) / 60, 1), 'index_after': [nP, nK], 'by_day': stats,
            'script': 'stan_backfill.py 2.0'}
    json.dump(info, open(os.path.join(a.out, f'stan_zminy_{ym}.json'), 'w'), ensure_ascii=False, indent=1)
    subprocess.run(['rm', '-rf', tmpd])
    say(f'DONE {ym}: ' + json.dumps(fin, ensure_ascii=False) + f'; {dst} ({os.path.getsize(dst)} байт), {info["minutes"]} хв'
        + (f'; SKIP {skipped}' if skipped else '') + (f'; 404 {absent}' if absent else '') + (f'; опорний {ref}' if ref != refs[0] else ''))


def months(frm, to):
    y, m = map(int, frm.split('-')); out = []
    while f'{y:04d}-{m:02d}' >= to:
        out.append(f'{y:04d}-{m:02d}'); y, m = (y, m - 1) if m > 1 else (y - 1, 12)
    return out


def cmd_run(a):
    ms = months(a.frm, a.to)
    done = [m for m in ms if os.path.exists(os.path.join(a.out, f'stan_zminy_{m}.json'))]
    say(f'місяців у діапазоні {a.frm} → {a.to}: {len(ms)}; оброблено {len(done)}; залишилося {len(ms) - len(done)}')
    t0 = time.time(); per = 60 * 60
    for ym in ms:
        if ym in done: continue
        if time.time() - t0 + per > a.max_hours * 3600:
            left = sum(1 for m in ms if not os.path.exists(os.path.join(a.out, f'stan_zminy_{m}.json')))
            say(f'STOP: часу на {ym} не вистачить — перезапустіть ту саму команду (залишилося {left})'); return
        cmd_month(a, ym)
    say('ALL DONE')


# ---------- перевірка і звіт ----------
def cmd_check(a):
    f = os.path.join(a.out, f'stan_zminy_{a.month}.tsv.xz')
    n = bad = dup = 0; seen = set(); c = collections.Counter(); ch = collections.defaultdict(list)
    for x in arch_rows(f):
        n += 1; h = bytes.fromhex(x[2])
        if h in seen: dup += 1
        seen.add(h)
        if dig(x[3:]) != h: bad += 1
        c[x[1]] += 1
        if a.court and a.court in x[3]:
            ch[(x[3], x[4], x[5], x[14])].append((x[0], x[1], x[7][:45], x[10], x[11], x[12][:40]))
    j = os.path.join(a.out, f'stan_zminy_{a.month}.json')
    exp = sum(json.load(open(j))['result'].get(k, 0) for k in ('нове', 'змінене', 'змінене_було')) if os.path.exists(j) else None
    print(f'рядків {n} (за .json {exp}), дублікатів {dup}, хибних хешів {bad}; ' + ', '.join(f'{k} {v}' for k, v in c.items()))
    ex = sorted((k for k, v in ch.items() if len(v) >= 3), key=lambda k: -len(ch[k]))[:a.examples]
    for k in ex:
        print('===', ' | '.join(k))
        for y in ch[k]: print('   ', ' | '.join(y))


def cmd_zvit(a):
    print('місяць | записів | МБ | хв | днів у файлі | дні < порога | не завантажено | 404 | 0 КБ на сайті | опорний')
    tot = tb = 0; allsk = []
    for f in sorted(glob.glob(os.path.join(a.out, 'stan_zminy_*.json')), reverse=True):
        i = json.load(open(f)); m = i['month']
        if a.frm and m < a.frm or a.to and m > a.to: continue
        r = sum(i['result'].get(k, 0) for k in ('нове', 'змінене', 'змінене_було'))
        low = [d['date'][8:] for d in i['by_day'] if 0 < d.get('rows', 0) < a.porig]
        sk = i.get('skipped_days', []); ab = i.get('absent_days_404', []); em = i.get('empty_on_site', [])
        tot += r; tb += i['file_bytes']; allsk += sk
        print(f"{m} | {r} | {i['file_bytes'] / 1e6:.1f} | {i.get('minutes')} | {len(i['days'])} | {len(low)} | "
              f"{','.join(d[5:] for d in sk) or '—'} | {','.join(d[5:] for d in ab) or '—'} | {len(em)} | {i['reference']}")
    print(f'РАЗОМ: {tot} записів, {tb / 1e9:.2f} ГБ; не завантажено днів: {len(allsk)} {sorted(allsk)}')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest='cmd', required=True)
    p = sp.add_parser('list'); p.add_argument('--work', required=True)
    p = sp.add_parser('prepare'); p.add_argument('--work', required=True); p.add_argument('--base')
    p.add_argument('--layers', nargs='*'); p.add_argument('--parquet', help="glob, напр. 'D/stan_istoriia_*.parquet'")
    p = sp.add_parser('month'); p.add_argument('month'); p.add_argument('--work', required=True); p.add_argument('--out', required=True)
    p = sp.add_parser('run'); p.add_argument('--from', dest='frm', required=True); p.add_argument('--to', required=True)
    p.add_argument('--work', required=True); p.add_argument('--out', required=True); p.add_argument('--max-hours', type=float, default=1.8)
    p = sp.add_parser('check'); p.add_argument('month'); p.add_argument('--out', required=True)
    p.add_argument('--examples', type=int, default=3); p.add_argument('--court', default='районний суд')
    p = sp.add_parser('zvit'); p.add_argument('--out', required=True); p.add_argument('--from', dest='frm')
    p.add_argument('--to'); p.add_argument('--porig', type=int, default=1000000)
    a = ap.parse_args()
    for d in (getattr(a, 'work', None), getattr(a, 'out', None)):
        if d: os.makedirs(d, exist_ok=True)
    {'list': cmd_list, 'prepare': cmd_prepare, 'month': cmd_month, 'run': cmd_run, 'check': cmd_check,
     'zvit': cmd_zvit}[a.cmd](a)


if __name__ == '__main__':
    main()
