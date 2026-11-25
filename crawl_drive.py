#!/usr/bin/env python3
"""Рекурсивно обходить публічну папку Google Диска і завантажує всі файли.
python3 crawl.py <folder_id> <out_dir>   ->  out_dir/manifest.csv + файли у підпапках"""
import csv, html, os, re, subprocess, sys, concurrent.futures as cf

def get(url, out=None):
    cmd = ['curl', '-sS', '-L', '-m', '120', '--retry', '3', url]
    if out:
        cmd[1:1] = ['-o', out]
        subprocess.run(cmd, check=True); return None
    return subprocess.run(cmd, capture_output=True, check=True).stdout.decode('utf-8', 'replace')

def entries(fid):
    h = get(f'https://drive.google.com/embeddedfolderview?id={fid}')
    out = []
    for m in re.finditer(r'<div class="flip-entry" id="entry-([^"]+)".*?<a href="([^"]+)".*?class="flip-entry-title">([^<]+)', h, re.S):
        eid, href, title = m.group(1), html.unescape(m.group(2)), html.unescape(m.group(3)).strip()
        out.append((eid, title, '/folders/' in href))
    return out

def walk(fid, path, acc):
    for eid, title, is_dir in entries(fid):
        if is_dir:
            walk(eid, path + [title], acc)
        else:
            acc.append(('/'.join(path), title, eid))

root, outd = sys.argv[1], sys.argv[2]
os.makedirs(outd, exist_ok=True)
acc = []
walk(root, [], acc)
with open(os.path.join(outd, 'manifest.csv'), 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f); w.writerow(['folder', 'title', 'id']); w.writerows(acc)
print('files:', len(acc), flush=True)

def dl(row):
    folder, title, eid = row
    d = os.path.join(outd, folder); os.makedirs(d, exist_ok=True)
    safe = title.replace('/', '_')
    p = os.path.join(d, f'{eid}__{safe}')
    if os.path.exists(p) and os.path.getsize(p) > 0:
        return 0
    try:
        get(f'https://drive.usercontent.google.com/download?id={eid}&export=download&confirm=t', p)
        return 1
    except Exception as e:
        print('ERR', eid, title, e, flush=True); return 0

n = 0
with cf.ThreadPoolExecutor(8) as ex:
    for k, r in enumerate(ex.map(dl, acc), 1):
        n += r
        if k % 200 == 0: print('done', k, flush=True)
print('downloaded', n, 'of', len(acc), flush=True)
