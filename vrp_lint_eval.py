#!/usr/bin/env python3
"""vrp_lint_eval.py — оцінка повноти vrp_lint.py на парах «до / після редактора» (автор: В'ячеслав Плескач).
python3 vrp_lint_eval.py <тека з *_do.docx і *_pislia.docx> [--show N] [--json out.json]
"""
import sys, re, json, difflib, importlib.util, argparse, collections, os

spec = importlib.util.spec_from_file_location('vl', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vrp_lint.py'))
vl = importlib.util.module_from_spec(spec); spec.loader.exec_module(vl)

def norm(t):
    return t.replace(' ', ' ').replace('‑', '-').replace('\t', ' ')

TOK = re.compile(r"[\w’'ʼ]+|\s+|.", re.U)

def hits_for(t, ital):
    qs = vl.quote_spans(t, ital, 8)
    out = []
    for rid, grp, sec, desc, rx, flt in vl.R:
        for m in rx.finditer(t):
            if flt and not flt(t, m):
                continue
            out.append((m.start(), m.end(), rid))
    return out

def edits(a, b):
    ta, tb = TOK.findall(a), TOK.findall(b)
    sm = difflib.SequenceMatcher(None, ta, tb, autojunk=False)
    pos = [0]
    for x in ta:
        pos.append(pos[-1] + len(x))
    res = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            continue
        old, new = ''.join(ta[i1:i2]), ''.join(tb[j1:j2])
        if old.strip() == '' and new.strip() == '' and old.replace(' ', '') == new.replace(' ', ''):
            if old.count(' ') == new.count(' '):
                continue
        res.append((pos[i1], pos[i2], old, new))
    return res

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('dir'); ap.add_argument('--show', type=int, default=0); ap.add_argument('--json')
    a = ap.parse_args()
    tot = collections.Counter(); missed = []; byrule = collections.Counter(); fp = collections.Counter(); hitsall = collections.Counter()
    for f in sorted(os.listdir(a.dir)):
        if not f.endswith('_do.docx'):
            continue
        name = f[:-8]
        rawA = vl.paras_docx(os.path.join(a.dir, f))
        A = [(norm(x[0]), x[1]) for x in rawA]
        POS = {}
        for k0, s1, e1, rid, *_r in vl.pos_doc_rules([(norm(x[0]), x[1], x[2]) for x in rawA]):
            POS.setdefault(k0, []).append((s1, e1, rid))
        B = [norm(x[0]) for x in vl.paras_docx(os.path.join(a.dir, name + '_pislia.docx'))]
        At = [t for t, _ in A]
        sm = difflib.SequenceMatcher(None, At, B, autojunk=False)
        pairs = []
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            if op == 'equal':
                for k in range(i2 - i1):
                    pairs.append((i1 + k, j1 + k))
            elif op == 'replace':
                # зіставити за схожістю
                used = set()
                for i in range(i1, i2):
                    best, bj = 0, None
                    for j in range(j1, j2):
                        if j in used: continue
                        r = difflib.SequenceMatcher(None, At[i], B[j], autojunk=False).quick_ratio()
                        if r > best: best, bj = r, j
                    if bj is not None and best > 0.6:
                        used.add(bj); pairs.append((i, bj))
        n_ed = 0
        for i, j in pairs:
            t, it = A[i]
            if t == B[j]:
                hs = hits_for(t, it) + POS.get(i, [])
                for s, e, r in hs: fp[r] += 1; hitsall[r] += 1
                continue
            hs = hits_for(t, it) + POS.get(i, [])
            eds = edits(t, B[j])
            covered_h = set()
            for s0, e0, old, new in eds:
                small = len(old) <= 40 and len(new) <= 40
                kind = 'small' if small else 'big'
                tot[kind] += 1; n_ed += 1
                cov = [r for (s, e, r) in hs if s <= e0 + 2 and e >= s0 - 2]
                for (s, e, r) in hs:
                    if s <= e0 + 2 and e >= s0 - 2: covered_h.add((s, e, r))
                if cov:
                    tot[kind + '_cov'] += 1
                    for r in set(cov): byrule[r] += 1
                elif small:
                    ctx = t[max(0, s0 - 30):min(len(t), e0 + 30)]
                    missed.append({'doc': name, 'old': old, 'new': new, 'ctx': ctx})
            for h in hs:
                hitsall[h[2]] += 1
                if h not in covered_h: fp[h[2]] += 1
        print(f'{name}: абзаців {len(A)}, зіставлено {len(pairs)}, правок {n_ed}')
    print()
    for k in ('small', 'big'):
        if tot[k]:
            print(f'{k}: {tot[k]} правок, позначено скриптом {tot[k+"_cov"]} ({100*tot[k+"_cov"]/tot[k]:.1f} %)')
    # групи пропущених
    def key(m):
        o, n = m['old'].strip().lower(), m['new'].strip().lower()
        return f'{o!r} → {n!r}'
    grp = collections.Counter(key(m) for m in missed)
    print('\nНайчастіші пропущені правки:')
    for k, v in grp.most_common(a.show or 60):
        print(f'{v:4d}  {k}')
    print('\nПравила: спрацювань / збіглося з правкою / без правки')
    for r in sorted(hitsall, key=lambda r: -hitsall[r]):
        print(f'  {r:6s} {hitsall[r]:5d} {byrule[r]:5d} {fp[r]:5d}')
    if a.json:
        json.dump({'missed': missed, 'tot': tot}, open(a.json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

if __name__ == '__main__':
    main()
