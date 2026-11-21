#!/usr/bin/env python3
"""Пошук у поточних знімках ДСА: «Інформація щодо стану розгляду справ» і «Список справ, призначених до
розгляду». Обидва набори щодня перезаписуються, тому скрипт щоразу бере свіжий файл (кешує його до наступного
оновлення набору).

Автор: В'ячеслав Плескач. Версія 1.1 (21.11.2026): резервне джерело — сайт ДСА (court.gov.ua/opendata, ті самі
файли): stan — https://dsa.court.gov.ua/dsa/inshe/oddata/532/, spisok — …/oddata/513/. --dzherelo auto (типово:
data.gov.ua, а якщо не відповідає — сайт ДСА) | portal | dsa.

  stan   — https://data.gov.ua/dataset/7e659fca-90c7-4002-bd96-00f9dbbbd251 (архів ~900 МБ; Верховний Суд — уся
           історія з 2018 р., місцеві й апеляційні суди — лише записи з рухом за останні тижні)
  spisok — https://data.gov.ua/dataset/42eaff6e-45da-4426-b4a1-f30989bfd36f (CSV ~400 МБ; засідання на тижні вперед)

Приклади:
  python3 potochni_dani.py stan --case "757/30191/25-к"
  python3 potochni_dani.py stan --court "Печерський районний суд міста Києва" --judge "Вовк" --out stan.tsv
  python3 potochni_dani.py spisok --case "757/30191/25-к"
  python3 potochni_dani.py spisok --court "Печерський" --judge "Вовк" --from 2026-11-09 --to 2026-11-13
Фільтри: --case (точно; кілька — через кому), --court і --judge (частина рядка, без урахування регістру),
--from/--to (для stan — дата стадії, для spisok — дата засідання).
"""
import argparse, csv, io, json, os, re, subprocess, sys, time, zipfile, datetime as dt

csv.field_size_limit(10 ** 9)
PKG = {'stan': '7e659fca-90c7-4002-bd96-00f9dbbbd251', 'spisok': '42eaff6e-45da-4426-b4a1-f30989bfd36f'}
DSA = {'stan': 532, 'spisok': 513}


def meta_portal(kind):
    """→ (url, локальна назва, позначка версії, джерело) з data.gov.ua."""
    meta = json.loads(subprocess.run(['curl', '-sS', '-f', '-m', '60', '--retry', '1',
                                      f'https://data.gov.ua/api/3/action/package_show?id={PKG[kind]}'],
                                     check=True, capture_output=True).stdout)['result']
    rs = meta['resources']
    if kind == 'stan':
        rs = [r for r in rs if re.search(r'vid-\d\d-\d\d-\d{4}', r['name'])]
        r = sorted(rs, key=lambda r: re.search(r'vid-(\d\d)-(\d\d)-(\d{4})', r['name']).groups()[::-1])[-1]
    else:
        r = rs[0]
    return r['url'], r['name'], meta['metadata_modified'], 'data.gov.ua'


def meta_dsa(kind):
    """→ (url, локальна назва, позначка версії, джерело) з паспорта набору на сайті ДСА."""
    n = DSA[kind]
    pas = json.loads(subprocess.run(['curl', '-sS', '-f', '-m', '120', '--retry', '2',
                                     f'https://dsa.court.gov.ua/open_data_json.php?json={n}'],
                                    check=True, capture_output=True).stdout)
    fs = [(k, v) for f in pas.get('Файли', []) for k, v in f.items()]
    if kind == 'stan':
        c = []
        for k, v in fs:
            m = re.search(r'від (\d\d)\.(\d\d)\.(\d{4})\.zip$', k)
            if m:
                c.append(((m[3], m[2], m[1]), v, f'stan-rozglyadu-sprav-vid-{m[1]}-{m[2]}-{m[3]}.zip'))
        _, url, name = max(c)
    else:
        url = fs[0][1]; name = 'spisok-sprav-dsa.csv'
    hdr = subprocess.run(['curl', '-sS', '-f', '-I', '-m', '60', url], check=True, capture_output=True).stdout.decode('latin-1')
    lm = (re.search(r'(?im)^last-modified:\s*(.+?)\s*$', hdr) or [0, ''])[1]
    try:
        import email.utils
        lm = email.utils.parsedate_to_datetime(lm).astimezone(dt.timezone.utc).replace(tzinfo=None).isoformat()
    except Exception:
        pass
    return url, name, lm, f'сайт ДСА (oddata/{n})'


def fresh(kind, cache, src='auto'):
    os.makedirs(cache, exist_ok=True)
    got = None
    if src in ('auto', 'portal'):
        try:
            got = meta_portal(kind)
        except Exception as e:
            if src == 'portal':
                sys.exit(f'data.gov.ua не відповідає: {e}')
            print('data.gov.ua не відповідає — резервне джерело: сайт ДСА', file=sys.stderr)
    if got is None:
        got = meta_dsa(kind)
    url, name, stamp_v, where = got
    path = os.path.join(cache, name)
    stamp = path + '.stamp'
    if not (os.path.exists(path) and os.path.exists(stamp) and open(stamp).read() == stamp_v):
        for f in os.listdir(cache):
            if f.startswith('stan-rozglyadu' if kind == 'stan' else 'spisok'):
                os.remove(os.path.join(cache, f))
        t = time.time()
        subprocess.run(['curl', '-sS', '-L', '-f', '--retry', '3', '-m', '1800', '-o', path, url], check=True)
        open(stamp, 'w').write(stamp_v)
        print(f'[{kind}] завантажено {name} з {where} за {time.time() - t:.0f} с (набір оновлено {stamp_v[:16]})', file=sys.stderr)
    else:
        print(f'[{kind}] кеш {name} з {where} (набір оновлено {stamp_v[:16]})', file=sys.stderr)
    return path, name, stamp_v + (f' ({where})' if where != 'data.gov.ua' else '')


def rows(path):
    def rd(t):
        r = csv.reader(t, delimiter='\t')
        h = [c.strip('"').strip('﻿') for c in next(r)]
        for x in r:
            yield dict(zip(h, x))
    if path.endswith('.zip'):
        z = zipfile.ZipFile(path)
        for nm in z.namelist():
            if nm.lower().endswith('.csv'):
                with z.open(nm) as f:
                    yield from rd(io.TextIOWrapper(f, encoding='utf-8', errors='replace', newline=''))
    else:
        with open(path, encoding='utf-8', errors='replace', newline='') as f:
            yield from rd(f)


def iso(d):
    try:
        return dt.datetime.strptime(d[:10], '%d.%m.%Y').date().isoformat()
    except Exception:
        return ''


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('kind', choices=['stan', 'spisok'])
    ap.add_argument('--case'); ap.add_argument('--court'); ap.add_argument('--judge')
    ap.add_argument('--from', dest='dfrom'); ap.add_argument('--to', dest='dto')
    ap.add_argument('--cache', default='potochni_cache'); ap.add_argument('--out')
    ap.add_argument('--dzherelo', choices=['auto', 'portal', 'dsa'], default='auto',
                    help='auto — data.gov.ua, а якщо не відповідає — сайт ДСА; portal; dsa')
    a = ap.parse_args()
    if not (a.case or a.court or a.judge):
        sys.exit('Задайте --case, --court або --judge.')
    path, name, stamp = fresh(a.kind, a.cache, a.dzherelo)
    cases = set(a.case.split(',')) if a.case else None
    ck, jk, nk, dk = ('court_name', 'judge', 'case_number', 'stage_date') if a.kind == 'stan' else ('court_name', 'judges', 'case', 'date')
    out = []; t = time.time()
    for x in rows(path):
        if cases and x.get(nk, '') not in cases: continue
        if a.court and a.court.lower() not in x.get(ck, '').lower(): continue
        if a.judge and a.judge.lower() not in (x.get(jk, '') + ' ' + x.get('judges', '')).lower(): continue
        d = iso(x.get(dk, ''))
        if a.dfrom and d < a.dfrom: continue
        if a.dto and d > a.dto: continue
        out.append(x)
    out.sort(key=lambda x: (iso(x.get(dk, '')), x.get(dk, '')))
    print(f'Знайдено {len(out)} записів за {time.time() - t:.0f} с. Джерело: {name}, набір оновлено {stamp[:10]}{stamp[19:] if len(stamp) > 19 else ""}', file=sys.stderr)
    if not out:
        return
    cols = list(out[0].keys())
    if a.out:
        with open(a.out, 'w', encoding='utf-8', newline='') as f:
            w = csv.DictWriter(f, fieldnames=cols, delimiter='\t', lineterminator='\n'); w.writeheader(); w.writerows(out)
        print(f'  → {a.out}', file=sys.stderr)
    else:
        print('\t'.join(cols))
        for x in out[:200]:
            print('\t'.join(str(x.get(c, ''))[:250] for c in cols))


if __name__ == '__main__':
    main()
