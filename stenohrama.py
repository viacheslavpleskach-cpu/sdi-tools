#!/usr/bin/env python3
"""Стенограми аудіо- та відеозаписів судових засідань із таймкодами та учасниками. Автор: В'ячеслав Плескач.

Версія 1.0 (15.11.2026), перевірено на справі № 487/8888/25 (Вітовський районний суд Миколаївської обл.).

Команди:
  fetch  <file_id|url> ... [--out DIR]
        завантажує файли з Google Диска напряму (файл або тека мають бути доступні за посиланням);
        ID файлів у теці дає конектор Диска (search_files: parentId = '<id теки>').
  protokol <протокол.html> [--csv OUT]
        розбирає протокол технічної фіксації ВКЗ (vkz.court.gov.ua, HTML з «Електронного суду»):
        склад, назва файлу запису, час «Включено відеозапис», перелік дій із часом і учасником.
  transcribe <запис.mp4|wav> [--model large-v3|turbo|medium] [--names "Прізвище, НАЗК, ..."] [--out X.jsonl]
        розпізнає мовлення (faster-whisper, українська). Пише сегменти в JSONL по одному, тому
        перерваний запуск продовжується з місця зупинки (повторіть ту саму команду).
  draft  <сегменти.jsonl> [--protokol протокол.html] [--start 09:58:28] [--out чернетка.md]
        чернетка стенограми: «В файлі» (хв:с від початку запису) | «Час» (астрономічний) | «Особа» | «Текст».
        Особу підставляє за останньою дією учасника в протоколі (секретар ставить позначки із запізненням
        у кілька секунд — особу перевіряти за змістом); сегменти з низькою впевненістю позначаються [?].

Досвід (див. також скіл «court-hearing-transcripts»):
  * У середовищі Claude: 2 CPU без GPU. large-v3 ≈ 1,2× тривалості запису, turbo ≈ 0,45× (якість гірша:
    спотворює терміни й назви). Довгі записи — turbo, ключові фрагменти — повторно large-v3 через --from/--to.
  * PyAV 19 несумісний з faster-whisper 1.2 (av.open metadata_errors) — тому аудіо готується ffmpeg у WAV
    16 кГц і подається масивом numpy.
  * Записи ВКЗ часто починаються вже після «Засідання розпочато» — перші слова можуть бути втрачені.
"""
import argparse, html, json, os, re, subprocess, sys, time, wave

MODELS = {"large-v3": "large-v3", "turbo": "mobiuslabsgmbh/faster-whisper-large-v3-turbo",
          "medium": "medium", "small": "small"}
PROMPT = "Судове засідання. Головуючий суддя, секретар судового засідання, прокурор, захисник, свідок. КУпАП."


# ---------- fetch ----------
def drive_id(s):
    m = re.search(r"/d/([\w-]{20,})|[?&]id=([\w-]{20,})", s)
    return (m.group(1) or m.group(2)) if m else s


def cmd_fetch(a):
    os.makedirs(a.out, exist_ok=True)
    for x in a.ids:
        fid = drive_id(x)
        tmp = os.path.join(a.out, fid + ".part")
        url = f"https://drive.usercontent.google.com/download?id={fid}&export=download&confirm=t"
        r = subprocess.run(["curl", "-sS", "-L", "-D", "-", "-o", tmp, url], capture_output=True, text=True)
        name = re.findall(r'filename="([^"]+)"', r.stdout)
        dest = os.path.join(a.out, name[-1] if name else fid)
        os.replace(tmp, dest)
        head = open(dest, "rb").read(4000).lower()
        if b"accounts.google.com" in head or b"<title>google drive" in head:
            print(f"! {fid}: замість файлу отримано HTML — файл не відкритий за посиланням", file=sys.stderr)
        print(dest, os.path.getsize(dest))


# ---------- protokol ----------
def parse_protokol(path):
    t = open(path, encoding="utf-8", errors="replace").read()
    t = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", t, flags=re.S)
    t = re.sub(r"</tr>", "\n", t)
    t = html.unescape(re.sub(r"<[^>]+>", " ", t))
    t = re.sub(r"[ \t\r]+", " ", t).replace("Зі сховища", "").replace("Локально", "")
    p = {"sprava": (re.search(r"Справа № ([\d/\-]+)", t) or [None, ""])[1],
         "data": (re.search(r"(\d\d\.\d\d\.\d{4}) ", t) or [None, ""])[1],
         "fajl": sorted(set(re.findall(r"[0-9a-f]{32}_[0-9a-f-]{36}\.mp4", t))),
         "sklad": re.findall(r"(Головуючий суддя|Суддя|Прокурор|Захисник|Свідок|Секретар судового засідання|"
                             r"Особа, яка притягається до адміністративної відповідальності|Позивач|Відповідач|"
                             r"Представник[^ ]*|Обвинувачений|Потерпіл\w+|Експерт|Спеціаліст)\s+([А-ЯІЇЄҐ][^()\n]{2,60}?)\s*\(", t),
         "dii": []}
    for m in re.finditer(r"^\s*(\d+)\s+(\d\d:\d\d:\d\d)(?:\s+\d\d\.\d\d\.\d{4})?\s+(.*?)\s+-\s+(.*)$", t, flags=re.M):
        who = m.group(3).strip()
        role, name = (who.split(" : ", 1) + [""])[:2] if " : " in who else (who, "")
        p["dii"].append({"n": int(m.group(1)), "chas": m.group(2), "rol": role.strip(),
                         "osoba": name.strip(), "dia": m.group(4).strip()})
    st = [d["chas"] for d in p["dii"] if "Включено відеозапис" in d["dia"]]
    p["start"] = st[0] if st else None
    return p


def cmd_protokol(a):
    p = parse_protokol(a.html)
    print(f"Справа {p['sprava']}, {p['data']}; файл(и): {', '.join(p['fajl'])}; початок запису: {p['start']}")
    for r, n in p["sklad"]:
        print(f"  {r}: {n.strip()}")
    for d in p["dii"]:
        if d["osoba"]:
            print(f"  {d['chas']}  {d['rol']}: {d['osoba']} — {d['dia']}")
    if a.csv:
        import csv
        with open(a.csv, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["n", "chas", "rol", "osoba", "dia"])
            w.writeheader(); w.writerows(p["dii"])


# ---------- transcribe ----------
def to_wav(src):
    if src.lower().endswith(".wav"):
        return src
    dst = os.path.splitext(src)[0] + ".16k.wav"
    if not os.path.exists(dst):
        subprocess.run(["ffmpeg", "-v", "error", "-y", "-i", src, "-vn", "-ac", "1", "-ar", "16000", dst], check=True)
    return dst


def cmd_transcribe(a):
    import numpy as np
    from faster_whisper import WhisperModel
    wav = to_wav(a.audio)
    out = a.out or os.path.splitext(a.audio)[0] + f".{a.model}.jsonl"
    done = 0.0
    if os.path.exists(out):
        lines = [json.loads(l) for l in open(out, encoding="utf-8") if l.strip()]
        if lines:
            done = lines[-1]["e"]
    start = max(done, a.from_)
    w = wave.open(wav); sr = w.getframerate(); total = w.getnframes() / sr
    end = min(a.to or total, total)
    if start >= end - 1:
        print(f"Вже розпізнано: {out}"); return
    w.setpos(int(start * sr))
    audio = np.frombuffer(w.readframes(int((end - start) * sr)), dtype=np.int16).astype(np.float32) / 32768.0
    prompt = PROMPT + (" " + a.names if a.names else "")
    m = WhisperModel(MODELS.get(a.model, a.model), device="cpu", compute_type="int8", cpu_threads=os.cpu_count())
    t0 = time.time()
    segs, _ = m.transcribe(audio, language="uk", beam_size=5, vad_filter=True,
                           vad_parameters={"min_silence_duration_ms": 500},
                           initial_prompt=prompt, condition_on_previous_text=False)
    with open(out, "a", encoding="utf-8") as f:
        for s in segs:
            f.write(json.dumps({"s": round(s.start + start, 2), "e": round(s.end + start, 2), "t": s.text.strip(),
                                "lp": round(s.avg_logprob, 2), "ns": round(s.no_speech_prob, 2)},
                               ensure_ascii=False) + "\n")
            f.flush()
    el = time.time() - t0
    print(f"{out}: {end - start:.0f} с аудіо за {el:.0f} с ({el / max(end - start, 1):.2f}×)", file=sys.stderr)


# ---------- draft ----------
def hms(sec):
    sec = int(round(sec)); return f"{sec // 3600:02d}:{sec % 3600 // 60:02d}:{sec % 60:02d}"


def mmss(sec):
    sec = int(sec); return f"{sec // 3600}:{sec % 3600 // 60:02d}:{sec % 60:02d}" if sec >= 3600 else f"{sec // 60:02d}:{sec % 60:02d}"


def to_sec(c):
    h, m, s = map(int, c.split(":")); return h * 3600 + m * 60 + s


def short(rol, osoba):
    r = {"Головуючий суддя": "Суддя", "Особа, яка притягається до адміністративної відповідальності": "Особа, що притягається"}.get(rol, rol)
    return f"{r} {osoba}".strip()


def cmd_draft(a):
    segs = [json.loads(l) for l in open(a.jsonl, encoding="utf-8") if l.strip()]
    dii, start = [], a.start
    if a.protokol:
        p = parse_protokol(a.protokol)
        start = start or p["start"]
        dii = [(to_sec(d["chas"]), short(d["rol"], d["osoba"])) for d in p["dii"] if d["osoba"]]
    t0 = to_sec(start) if start else None
    rows = ["| В файлі | Час | Особа | Текст |", "| --- | --- | --- | --- |"]
    prev, row_start = None, 0.0
    for s in segs:
        ab = t0 + s["s"] if t0 is not None else None
        who = ""
        if ab is not None and dii:
            cand = [w for t, w in dii if t <= ab + a.lag]
            who = cand[-1] if cand else ""
        txt = s["t"] + (" [?]" if s["lp"] < a.min_lp else "")
        if a.merge and who == prev and len(rows) > 2 and s["s"] - row_start < a.merge:
            rows[-1] = rows[-1][:-2] + " " + txt + " |"
            continue
        row_start = s["s"]
        rows.append(f"| {mmss(s['s'])} | {hms(ab) if ab is not None else ''} | {who} | {txt} |")
        prev = who
    out = "\n".join(rows) + "\n"
    if a.out:
        open(a.out, "w", encoding="utf-8").write(out); print(a.out)
    else:
        print(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    f = sp.add_parser("fetch"); f.add_argument("ids", nargs="+"); f.add_argument("--out", default="media")
    p = sp.add_parser("protokol"); p.add_argument("html"); p.add_argument("--csv")
    t = sp.add_parser("transcribe"); t.add_argument("audio"); t.add_argument("--model", default="large-v3")
    t.add_argument("--names", default=""); t.add_argument("--out")
    t.add_argument("--from", dest="from_", type=float, default=0.0, help="секунда початку")
    t.add_argument("--to", type=float, default=None, help="секунда кінця")
    d = sp.add_parser("draft"); d.add_argument("jsonl"); d.add_argument("--protokol"); d.add_argument("--start")
    d.add_argument("--out"); d.add_argument("--lag", type=float, default=3.0, help="запізнення позначок секретаря, с")
    d.add_argument("--min-lp", dest="min_lp", type=float, default=-0.6, help="поріг avg_logprob для позначки [?]")
    d.add_argument("--merge", type=float, default=0, help="зливати поспіль репліки однієї особи в рядок до N секунд (напр. 30)")
    a = ap.parse_args()
    {"fetch": cmd_fetch, "protokol": cmd_protokol, "transcribe": cmd_transcribe, "draft": cmd_draft}[a.cmd](a)


if __name__ == "__main__":
    main()
