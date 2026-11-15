#!/usr/bin/env python3
"""Перебудова індексу історії «Стан розгляду справ» для stan_poshuk.py (додавання нових місячних файлів).

Версія 1.1 (11.11.2026). Відтворює спосіб, яким 10.11.2026 зібрано індекс 01.2025–09.2026 + знімок 30.10.2026 (11.11.2026 — 01.2023–30.10.2026):
  ключ справи k = lower(номер без пробілів і кінцевого суфікса «-к», «-ц»…), шард = md5(k)[:8] % 32,
  у шарді записи відсортовано за k, судом, провадженням, видом, snapshot_date, stage_date.
Формули ключа й шарда збігаються з stan_poshuk.py (функції key, shard) — їх не змінювати.
Зміни 1.1 (результат не змінюється): шарди пишуться окремим проходом кожен (без тимчасового PARTITION_BY, що на 62 млн записів
вичерпував диск); зведення по суддях рахується через унікальні (суд, суддя, провадження, клас, місяць); --prodovzh пропускає готові шарди;
perevir рахує записи нових місяців за .json поруч із .tsv.xz (wc -l хибно враховує переноси рядків усередині полів).

Команди:
  stan_indeks.py zavantazh --folder <id теки Диска> --dir D [--mask stan_istoriia_]
        завантажити файли публічної теки Диска (доступ за посиланням) за назвами
  stan_indeks.py build --old 'D/stan_istoriia_*.parquet' --months 'O/stan_zminy_2023-*.tsv.xz' 'O/stan_zminy_2024-*.tsv.xz' \\
        --out NEW --work TMP [--prodovzh] [--vid 2023-01-01 --misiats-do 2026-09 --intervaly-do 2026-10-30]
        старі шарди + нові місячні файли → 32 нові шарди + stan_suddi_misiats.parquet + stan_suddi_intervaly.parquet
  stan_indeks.py perevir --dir NEW [--old 'D/stan_istoriia_*.parquet'] [--months ...] [--case "757/28368/24-к"]
        кількість записів (стара + нові = нова), дублікати row_hash, розміри шардів, приклад справи
Автор: В'ячеслав Плескач
"""
import argparse, glob, hashlib, os, re, shutil, subprocess, sys, time

NSHARD = 32
COLS = ['k', 'snapshot_date', 'change', 'court_name', 'case_number', 'case_proc', 'type', 'registration_date', 'judge',
        'judges', 'stage_date', 'stage_name', 'cause_result', 'cause_dep', 'description', 'participants', 'row_hash']
KSQL = r"lower(regexp_replace(regexp_replace(coalesce(case_number,''),'\s+','','g'),'-[^/0-9-]{1,4}$',''))"
SSQL = "(('0x'||md5(k)[1:8])::UBIGINT % 32)::INT"


def say(*a):
    print(time.strftime('%H:%M:%S'), *a, flush=True)


def duck():
    try:
        import duckdb
    except ImportError:
        subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', '--break-system-packages', 'duckdb'], check=True)
        import duckdb
    return duckdb


def con_for(work, db=None):
    duckdb = duck()
    os.makedirs(os.path.join(work, 'tmp_duck'), exist_ok=True)
    c = duckdb.connect(db) if db else duckdb.connect()
    c.execute(f"SET threads=2; SET memory_limit='5GB'; SET preserve_insertion_order=false; "
              f"SET temp_directory='{os.path.join(work, 'tmp_duck')}'")
    return c


def key(case):  # те саме, що в stan_poshuk.py
    s = re.sub(r'\s+', '', (case or '')).lower()
    return re.sub(r'-[^/0-9-]{1,4}$', '', s)


def shard(k):
    return int(hashlib.md5(k.encode('utf-8')).hexdigest()[:8], 16) % NSHARD


# ---------- завантаження з Диска ----------
def folder_files(folder):
    r = subprocess.run(['curl', '-sS', '-L', '-m', '60', f'https://drive.google.com/embeddedfolderview?id={folder}'],
                       capture_output=True)
    html = r.stdout.decode('utf-8', 'replace')
    return {m.group(2).strip(): m.group(1)
            for m in re.finditer(r'/file/d/([A-Za-z0-9_-]+).*?flip-entry-title">([^<]+)', html, re.S)}


def cmd_zavantazh(a):
    fs = {n: i for n, i in folder_files(a.folder).items() if n.startswith(a.mask)}
    if not fs: sys.exit(f'У теці {a.folder} не знайдено файлів «{a.mask}*» (немає доступу за посиланням?)')
    os.makedirs(a.dir, exist_ok=True)
    def one(item):
        n, i = item; p = os.path.join(a.dir, n)
        if os.path.exists(p) and os.path.getsize(p) > 0: return n, 'є'
        rc = subprocess.run(['curl', '-sS', '-L', '-f', '--retry', '3', '-m', '1800', '-o', p + '.part',
                             f'https://drive.usercontent.google.com/download?id={i}&export=download&confirm=t']).returncode
        if rc: return n, 'ПОМИЛКА'
        os.replace(p + '.part', p); return n, 'ok'
    import concurrent.futures as cf
    with cf.ThreadPoolExecutor(4) as ex:
        res = list(ex.map(one, sorted(fs.items())))
    bad = [n for n, s in res if s == 'ПОМИЛКА']
    say(f'файлів {len(res)}, помилок {len(bad)}' + (f': {bad}' if bad else ''))
    if bad: sys.exit(1)


# ---------- побудова ----------
def cmd_build(a):
    t = time.time()
    old = sorted(glob.glob(a.old)); months = sorted({f for g in a.months for f in glob.glob(g)})
    if len(old) != NSHARD: sys.exit(f'Старих шардів {len(old)}, очікувалося {NSHARD}: {a.old}')
    if not months: sys.exit('Не знайдено місячних файлів --months')
    pq = os.path.join(a.work, 'pq'); part = os.path.join(a.work, 'part')
    for d in (pq, a.out): os.makedirs(d, exist_ok=True)
    c = con_for(a.work)
    # 1. місячні .tsv.xz → parquet з ключем
    for f in months:
        m = re.search(r'(\d{4}-\d\d)', os.path.basename(f)).group(1); dst = os.path.join(pq, f'{m}.parquet')
        if os.path.exists(dst): continue
        tmp = os.path.join(a.work, f'tmp_{m}.tsv')
        subprocess.run(f'xz -dc -T2 "{f}" > "{tmp}"', shell=True, check=True)
        c.execute(f"""COPY (select {KSQL} k, * from read_csv('{tmp}', delim='\t', quote='"', escape='"', header=true,
                      all_varchar=true, strict_mode=false, max_line_size=20000000)) TO '{dst}' (FORMAT parquet, COMPRESSION zstd)""")
        os.remove(tmp)
        say('місяць', m, c.execute(f"select count(*) from '{dst}'").fetchall()[0][0])
    # 2. об'єднання зі старими шардами і розкладання за шардами
    cols = ', '.join(COLS)
    c.execute(f"""create or replace view u as
        select {cols} from read_parquet({old}) union all select {cols} from read_parquet('{pq}/*.parquet')""")
    # розкладання за шардами: кожен шард читається окремим проходом (без тимчасового PARTITION_BY, що вимагав десятків ГБ диска)
    for s in range(NSHARD):
        dst = os.path.join(a.out, f'stan_istoriia_{s:02d}.parquet')
        if a.prodovzh and os.path.exists(dst): continue
        c.execute(f"""COPY (select {cols} from u where {SSQL} = {s} order by k, court_name, case_proc, type, snapshot_date, stage_date)
                      TO '{dst}' (FORMAT parquet, COMPRESSION zstd, COMPRESSION_LEVEL 9, ROW_GROUP_SIZE 50000)""")
        say('шард', s, os.path.getsize(dst) // 2 ** 20, 'МБ')
    # 3. зведення по суддях і судах (як 10.11.2026, з ширшим періодом)
    db = os.path.join(a.work, 'agg.db')
    if os.path.exists(db): os.remove(db)
    g = con_for(a.work, db)
    J = r"""trim(regexp_extract(coalesce(judge,''), ':\s*([^,;]+)', 1))"""
    g.execute(rf"""create or replace table x as
     with r as (select court_name, hash(k, case_proc, type) pid, type,
      case when stage_name in ('Призначено склад суду','Передано судді','Передано в доповідь') then 'R'
           when stage_name='Призначено до судового розгляду' then 'P'
           when stage_name in ('Розглянуто','Розглянуто у апеляційній інстанції','Розглянуто у касаційній інстанції') then 'Z'
           when stage_name='Залишено без руху' then 'B'
           when stage_name like 'Повернуто%' then 'V' else 'O' end c,
      try_strptime(stage_date,'%d.%m.%Y')::date sd, try_strptime(registration_date,'%d.%m.%Y')::date rd, {J} j0
      from read_parquet('{a.out}/stan_istoriia_*.parquet'))
     select court_name, pid, type, c, sd, rd,
      case when regexp_matches(j0,'^\S+\s+\S+\s+\S+$') then regexp_replace(j0,'^(\S+)\s+(\S)\S*\s+(\S)\S*$','\1 \2.\3.')
           when regexp_matches(j0,'^\S+\s+\S\.\s*\S\.?$') then regexp_replace(j0,'^(\S+)\s+(\S)\.\s*(\S)\.?$','\1 \2.\3.')
           else j0 end j
     from r""")
    vid, mdo, ido = a.vid, a.misiats_do, a.intervaly_do
    g.execute(f"""create or replace table e as select distinct court_name, j, pid, c, sd, strftime(sd,'%Y-%m') ym
                  from x where sd between '{vid}' and last_day(('{mdo}-01')::date)""")
    g.execute("create or replace table e2 as select distinct court_name, j, pid, c, ym from e")
    g.execute("create or replace table e3 as select distinct court_name, j, pid, ym from e2")
    g.execute("create or replace table e4 as select distinct court_name, pid, c, ym from e2")
    g.execute("create or replace table e5 as select distinct court_name, pid, ym from e4")
    cnt = """count(*) filter (where c='R') rozpodileno, count(*) filter (where c='Z') rozghlianuto,
      count(*) filter (where c='B') bez_rukhu, count(*) filter (where c='V') povernuto"""
    g.execute(f"""create or replace table m1 as
      select a.court_name, a.j suddia, a.ym, a.provadzhen_z_rukhom, b.rozpodileno, coalesce(p.pryznachen,0) pryznachen, b.rozghlianuto, b.bez_rukhu, b.povernuto
      from (select court_name, j, ym, count(*) provadzhen_z_rukhom from e3 group by all) a
      left join (select court_name, j, ym, {cnt} from e2 group by all) b using (court_name, j, ym)
      left join (select court_name, j, ym, count(*) pryznachen from e where c='P' group by all) p using (court_name, j, ym)""")
    g.execute(f"""create or replace table m2 as
      select a.court_name, '*' suddia, a.ym, a.provadzhen_z_rukhom, b.rozpodileno, coalesce(p.pryznachen,0) pryznachen, b.rozghlianuto, b.bez_rukhu, b.povernuto
      from (select court_name, ym, count(*) provadzhen_z_rukhom from e5 group by all) a
      left join (select court_name, ym, {cnt} from e4 group by all) b using (court_name, ym)
      left join (select court_name, ym, count(*) pryznachen from (select distinct court_name, pid, c, sd, ym from e where c='P') group by all) p using (court_name, ym)""")
    g.execute("""create or replace table q as with p as (select distinct court_name, j, pid, type, sd from x where c='P' and sd is not null)
                 select *, sd - lag(sd) over (partition by court_name, pid order by sd) gap from p""")
    g.execute("create or replace table d as select court_name, j, pid, type, min(rd) rd, min(sd) filter (where c='Z') rz from x group by all")
    g.execute(f"""create or replace table gg as
     with a as (select court_name, j suddia, type, count(gap) intervaliv, median(gap) med_interval, quantile_cont(gap,0.9) p90_interval from q group by all
                union all select court_name, '*', type, count(gap), median(gap), quantile_cont(gap,0.9) from q group by all),
     b as (select court_name, j suddia, type, count(*) rozghl_n, median(rz-rd) med_dniv_reiestr_rozghl from d
             where rz between '{vid}' and '{ido}' and rz>=rd group by all
           union all select court_name, '*', type, count(*), median(rz-rd) from d where rz between '{vid}' and '{ido}' and rz>=rd group by all)
     select * from a full join b using (court_name, suddia, type)""")
    g.execute(f"COPY (select * from (select * from m1 union all select * from m2) order by court_name, suddia, ym) TO '{a.out}/stan_suddi_misiats.parquet' (FORMAT parquet, COMPRESSION zstd)")
    g.execute(f"COPY (select * from gg order by court_name, suddia, type) TO '{a.out}/stan_suddi_intervaly.parquet' (FORMAT parquet, COMPRESSION zstd)")
    for f in ('stan_suddi_misiats', 'stan_suddi_intervaly'):
        p = os.path.join(a.out, f + '.parquet'); say(f, g.execute(f"select count(*) from '{p}'").fetchall()[0][0], os.path.getsize(p) // 1024, 'КБ')
    g.close(); os.remove(db); shutil.rmtree(pq); shutil.rmtree(part, ignore_errors=True)
    shutil.rmtree(os.path.join(a.work, 'tmp_duck'), ignore_errors=True)
    say('BUILD DONE', round((time.time() - t) / 60, 1), 'хв')


def cmd_perevir(a):
    import tempfile
    c = con_for(tempfile.mkdtemp(prefix='stan_perevir_'))
    new = sorted(glob.glob(os.path.join(a.dir, 'stan_istoriia_*.parquet')))
    n_new = c.execute(f"select count(*) from read_parquet({new})").fetchall()[0][0]
    print(f'шардів {len(new)}, записів {n_new}, розмір {sum(os.path.getsize(f) for f in new) / 2 ** 30:.2f} ГБ, '
          f'найбільший шард {max(os.path.getsize(f) for f in new) // 2 ** 20} МБ')
    if a.old and a.months:
        n_old = c.execute(f"select count(*) from read_parquet({sorted(glob.glob(a.old))})").fetchall()[0][0]
        n_m = 0
        for g in a.months:
            for f in sorted(glob.glob(g)):
                j = f.replace('.tsv.xz', '.json')
                if os.path.exists(j):
                    import json
                    r = json.load(open(j))['result']; n_m += sum(r.get(k, 0) for k in ('нове', 'змінене', 'змінене_було'))
                else:
                    n_m += int(subprocess.run(f'xz -dc "{f}" | tail -n +2 | wc -l', shell=True, capture_output=True, text=True).stdout)
        print(f'старий індекс {n_old} + нові місяці {n_m} = {n_old + n_m}; збіг: {n_old + n_m == n_new}')
    bad = c.execute(f"select count(*) from read_parquet({new}, filename=true) where {SSQL} <> "
                    f"cast(regexp_extract(filename, 'stan_istoriia_(\\d+)', 1) as int)").fetchall()[0][0]
    dup = c.execute(f"select count(*) from (select row_hash from read_parquet({new}) group by 1 having count(*)>1)").fetchall()[0][0]
    rng = c.execute(f"select min(snapshot_date), max(snapshot_date) from read_parquet({new}) where change<>'знімок'").fetchall()[0]
    print(f'записів не у своєму шарді: {bad}; повторів row_hash: {dup}; snapshot_date ретроархіву: {rng[0]} … {rng[1]}')
    if a.case:
        k = key(a.case); f = os.path.join(a.dir, f'stan_istoriia_{shard(k):02d}.parquet')
        print(c.execute(f"""select snapshot_date, change, court_name, stage_date, stage_name, cause_result from '{f}'
                          where k=? order by snapshot_date, stage_date""", [k]).fetchdf().to_string(index=False, max_colwidth=40))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest='cmd', required=True)
    p = sp.add_parser('zavantazh'); p.add_argument('--folder', required=True); p.add_argument('--dir', required=True)
    p.add_argument('--mask', default='stan_istoriia_')
    p = sp.add_parser('build'); p.add_argument('--old', required=True); p.add_argument('--months', nargs='+', required=True)
    p.add_argument('--out', required=True); p.add_argument('--work', required=True)
    p.add_argument('--prodovzh', action='store_true'); p.add_argument('--vid', default='2023-01-01'); p.add_argument('--misiats-do', default='2026-09')
    p.add_argument('--intervaly-do', default='2026-10-30')
    p = sp.add_parser('perevir'); p.add_argument('--dir', required=True); p.add_argument('--old'); p.add_argument('--months', nargs='*')
    p.add_argument('--case')
    a = ap.parse_args()
    {'zavantazh': cmd_zavantazh, 'build': cmd_build, 'perevir': cmd_perevir}[a.cmd](a)


if __name__ == '__main__':
    main()
