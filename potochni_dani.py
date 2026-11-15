#!/usr/bin/env python3
"""Пошук у поточних знімках ДСА на data.gov.ua: «Інформація щодо стану розгляду справ» і «Список справ,
призначених до розгляду». Обидва набори щодня перезаписуються, тому скрипт щоразу бере свіжий файл
(кешує його до наступного оновлення набору).

Автор: В'ячеслав Плескач

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


def fresh(kind, cache):
    os.makedirs(cache, exist_ok=True)
    meta = json.loads(subprocess.run(['curl', '-sS', '-f', '-m', '60', f'https://data.gov.ua/api/3/action/package_show?id={PKG[kind]}'],
                                     check=True, capture_output=True).stdout)['result']
    rs = meta['resources']
    if kind == 'stan':
        rs = [r for r in rs if re.search(r'vid-\d\d-\d\d-\d{4}', r['name'])]
        r = sorted(rs, key=lambda r: re.search(r'vid-(\d\d)-(\d\d)-(\d{4})', r['name']).groups()[::-1])[-1]
    else:
        r = rs[0]
    path = os.path.join(cache, r['name'])
    stamp = path + '.stamp'
    if not (os.path.exists(path) and os.path.exists(stamp) and open(stamp).read() == meta['metadata_modified']):
        for f in os.listdir(cache):
            if f.startswith('stan-rozglyadu' if kind == 'stan' else 'spisok'):
                os.remove(os.path.join(cache, f))
        t = time.time()
        subprocess.run(['curl', '-sS', '-L', '-f', '--retry', '3', '-m', '1800', '-o', path, r['url']], check=True)
        open(stamp, 'w').write(meta['metadata_modified'])
        print(f'[{kind}] завантажено {r["name"]} за {time.time() - t:.0f} с (набір оновлено {meta["metadata_modified"][:16]})', file=sys.stderr)
    else:
        print(f'[{kind}] кеш {r["name"]} (набір оновлено {meta["metadata_modified"][:16]})', file=sys.stderr)
    return path, r['name'], meta['metadata_modified']


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
    a = ap.parse_args()
    if not (a.case or a.court or a.judge):
        sys.exit('Задайте --case, --court або --judge.')
    path, name, stamp = fresh(a.kind, a.cache)
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
    print(f'Знайдено {len(out)} записів за {time.time() - t:.0f} с. Джерело: {name}, набір оновлено {stamp[:10]}', file=sys.stderr)
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
