#!/usr/bin/env python3
"""Стенограми аудіо- та відеозаписів судових засідань із таймкодами та учасниками. Автор: В'ячеслав Плескач.

Версія 1.2 (16.11.2026): журнали-скани замість HTML (draft --dii CSV), пошук пропусків і обрізаних сегментів
у власному JSONL (import *.jsonl), --max-tokens, виправлено merge (не «з'їдає» сусідні сегменти). Перевірено на
справі № 757/3699/26-к (три записи: 8, 10 і 17 хв, лише хмара).
Версія 1.1 (15.11.2026): гібрид «ПК користувача + хмара» (import / fix / merge), фільтр фантомних фраз і відлуння
підказки. Перевірено на справі № 487/8888/25 (три записи: 4, 61 і 69 хв).

Команди:
  fetch  <file_id|url> ... [--out DIR]   — завантаження з Google Диска напряму (файл/тека доступні за посиланням).
  protokol <протокол.html> [--csv OUT]   — розбір протоколу технічної фіксації ВКЗ: склад, файл запису,
        час «Включено відеозапис» (нуль відліку), посекундні дії учасників.
  transcribe <запис> [--model large-v3|turbo] [--names "..."] [--from/--to с]   — розпізнавання в хмарі
        (faster-whisper, uk), JSONL пишеться по сегменту; перерваний запуск продовжується тією ж командою.
  import <xxl.json> [--duration с]   — JSON Faster-Whisper-XXL з ПК користувача → *.local.jsonl (з позначкою
        збійних сегментів) + *.spans.json (збійні місця й пропуски > 4 с) для перерозпізнавання.
  fix <запис> <spans.json> [--names "..."]   — перерозпізнає в хмарі лише ці відрізки (large-v3).
  merge <local.jsonl> <fix.jsonl>   — зводить якісні локальні сегменти з перерозпізнаним.
  draft <jsonl> [--protokol html | --dii журнал.csv] [--merge 30]   — чернетка «В файлі | Час | Особа | Текст».
        Журнал-скан (PDF) переписати в CSV chas,rol,osoba,dia; перший рядок — «Включено відеозапис» (нуль відліку).

Досвід (детальніше — скіл «court-hearing-transcripts»):
  * Хмара Claude: 2 CPU без GPU. large-v3 ≈ 1,2× тривалості запису; turbo ≈ 0,45× (гірше з термінами);
    fix лише збійних відрізків ≈ 0,8× їхньої тривалості.
  * ПК користувача (Windows, RTX 2060 6 ГБ, Faster-Whisper-XXL r245.4, -m large-v3 --compute_type int8_float16,
    решта за замовчуванням): 130 хв за кілька хвилин, але 35–40% сегментів збійні (temperature fallback до 1.0,
    англійські/корейські/китайські вставки), ≈ 15 хв мовлення пропущено. Тому: import → fix → merge.
    Для наступних запусків на ПК випробувати (спершу на 4-хв записі): --compute_type float16 --beam_size 5
    --condition_on_previous_text False --vad_filter True --initial_prompt "<учасники, терміни>".
  * PowerShell: цикл `Get-ChildItem *.mp4 | ForEach-Object { & "<шлях>\faster-whisper-xxl.exe" $_.FullName ... }`
    (синтаксис cmd `for %f in (...)` у PowerShell не працює).
  * PyAV 19 несумісний з faster-whisper 1.2 — аудіо готується ffmpeg у WAV 16 кГц і подається масивом.
  * initial_prompt з іменами модель на тиші повторює дослівно; «Дякую за перегляд», одиночне «Дякую» на тиші —
    фантоми. Фільтруються is_phantom().
  * Позначки секретаря в протоколі ВКЗ запізнюються на 2–10 с і є не для кожної репліки; у залі один мікрофон —
    особу перевіряти за змістом (звертання, питання/відповідь, самопредставлення).
  * large-v3 на віддаленому від мікрофона мовленні ставить no_speech_prob 0,6–0,8 і викидає цілі 30-с вікна
    (пропуски по 1–1,5 хв при гучному звуці). Шлях: import <свій>.jsonl → fix --model turbo (turbo такі місця
    розпізнає хоч фрагментарно) → merge. Вимкнення VAD / порогу тиші для large-v3 не допомагає.
  * max_new_tokens=220 обрізає щільні 30-с сегменти посеред слова; import позначає їх (≥ 20 с без кінцевого
    розділового знака) як збійні, fix --max-tokens 0 --model large-v3 їх перерозпізнає повністю.
  * Записи ВКЗ часто починаються після «Засідання розпочато» — перші слова можуть бути втрачені.
"""
import argparse, html, json, os, re, subprocess, sys, time, wave

MODELS = {"large-v3": "large-v3", "turbo": "mobiuslabsgmbh/faster-whisper-large-v3-turbo",
          "medium": "medium", "small": "small"}
PROMPT = "Судове засідання. Головуючий суддя, секретар судового засідання, прокурор, захисник, свідок. КУпАП."


# Параметри декодування, перевірені в хмарі (2 CPU, ліміт пам'яті ≈ 6,4 ГБ на процес):
# beam 3 замість 5, лише дві «температури», обмеження токенів і штраф за повтори — інакше на шумних
# фрагментах модель зациклюється (хвилини на фрагмент) або перевищує пам'ять (процес убивається, код 137).
DECODE = dict(language="uk", vad_filter=True, vad_parameters={"min_silence_duration_ms": 500},
              condition_on_previous_text=False, temperature=[0.0, 0.4], max_new_tokens=220, repetition_penalty=1.1)
PROMPT_MAX_WORDS = 30  # довша підказка (> ~224 токенів) змушує Whisper зависати на тиші


def decode(a):
    d = dict(DECODE)
    if getattr(a, "max_tokens", None) is not None:
        d["max_new_tokens"] = a.max_tokens or None
    return d


def make_prompt(names):
    words = (PROMPT + (" " + names if names else "")).split()
    if len(words) > PROMPT_MAX_WORDS:
        print(f"! підказку обрізано до {PROMPT_MAX_WORDS} слів (було {len(words)}): довга підказка зависає", file=sys.stderr)
    return " ".join(words[:PROMPT_MAX_WORDS])


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
    prompt = make_prompt(a.names)
    m = WhisperModel(MODELS.get(a.model, a.model), device="cpu", compute_type="int8", cpu_threads=os.cpu_count())
    t0 = time.time()
    segs, _ = m.transcribe(audio, beam_size=a.beam, initial_prompt=prompt, **decode(a))
    with open(out, "a", encoding="utf-8") as f:
        for s in segs:
            if is_phantom(s.text, prompt):
                continue
            f.write(json.dumps({"s": round(s.start + start, 2), "e": round(s.end + start, 2), "t": s.text.strip(),
                                "lp": round(s.avg_logprob, 2), "ns": round(s.no_speech_prob, 2)},
                               ensure_ascii=False) + "\n")
            f.flush()
    el = time.time() - t0
    print(f"{out}: {end - start:.0f} с аудіо за {el:.0f} с ({el / max(end - start, 1):.2f}×)", file=sys.stderr)


# ---------- import / fix / merge: гібрид «локальний ПК + хмара» ----------
FOREIGN = re.compile(r"[A-Za-z]{3,}|[぀-鿿가-힯]")
# Типові «фантомні» фрази Whisper на тиші/шумі
PHANTOM = re.compile(r"дякую за перегляд|субтитр|продовження (в|у) наступн|підписуйтесь|редактор субтитрів|^(дякую|дякуємо)[.!]?$",
                     re.I)


def is_phantom(text, prompt=PROMPT):
    """Фантом або відлуння initial_prompt (модель повторює підказку з іменами на тиші)."""
    t = text.strip().lower()
    if PHANTOM.search(t):
        return True
    words = re.findall(r"\w+", t)
    pw = " ".join(re.findall(r"\w+", prompt.lower()))
    if len(words) < 4:
        return False
    cov = set()
    for i in range(len(words) - 3):
        if " ".join(words[i:i + 4]) in pw:
            cov.update(range(i, i + 4))
    if len(cov) / len(words) >= 0.6:  # сегмент здебільшого повторює підказку
        return True
    # перекручене відлуння («Головуючий суддя Вовк, секретар суддя Ємець, …»): майже всі слова — з підказки
    pv = set(pw.split()) | {"суддя", "секретар", "головуючий", "слідчий"}
    return sum(w in pv for w in words) / len(words) >= 0.9  # 0,75 відкидало справжнє «… з'явилися»


def is_bad(s):
    """Збійний сегмент: іншомовні вставки, низька впевненість, «розігрів» (temperature fallback), повтори."""
    return bool(FOREIGN.search(s["t"])) or s["lp"] < -1.0 or s.get("temp", 0) > 0.4 or s.get("cr", 0) > 2.4


def cmd_import(a):
    """JSON Faster-Whisper-XXL (або openai-whisper) → наш JSONL + файл відрізків для перерозпізнавання."""
    if a.json.endswith(".jsonl"):  # власний JSONL хмари (transcribe) — шукати пропуски мовлення для turbo-доповнення
        segs = [json.loads(l) for l in open(a.json, encoding="utf-8") if l.strip()]
        segs = [s for s in segs if not is_phantom(s["t"], make_prompt(a.names))]
        for s in segs:  # обрізаний max_new_tokens сегмент: довгий і без кінцевого розділового знака
            if s["e"] - s["s"] >= 20 and not s["t"].rstrip().endswith((".", "?", "!", "…")):
                s["cr"] = 9.9
    else:
        S = json.load(open(a.json, encoding="utf-8"))["segments"]
        segs = [{"s": round(s["start"], 2), "e": round(s["end"], 2), "t": s["text"].strip(),
                 "lp": round(s.get("avg_logprob", 0), 2), "temp": s.get("temperature", 0),
                 "cr": round(s.get("compression_ratio", 0), 2)} for s in S]
    dur = a.duration or (segs[-1]["e"] if segs else 0)
    spans = [[s["s"] - 0.5, s["e"] + 0.5] for s in segs if is_bad(s)]
    pts = [(0, 0)] + [(s["s"], s["e"]) for s in segs] + [(dur, dur)]
    spans += [[pts[i][1] - 0.5, pts[i + 1][0] + 0.5] for i in range(len(pts) - 1) if pts[i + 1][0] - pts[i][1] > a.gap]
    spans.sort(); m = []
    for x, y in spans:
        x, y = max(0, x), min(dur, y)
        if m and x <= m[-1][1] + 3:
            m[-1][1] = max(m[-1][1], y)
        else:
            m.append([x, y])
    base = os.path.splitext(a.out or a.json)[0].replace(".large-v3", "")
    with open(base + ".local.jsonl", "w", encoding="utf-8") as f:
        for s in segs:
            s["bad"] = is_bad(s); f.write(json.dumps(s, ensure_ascii=False) + "\n")
    json.dump([[round(x, 2), round(y, 2)] for x, y in m], open(base + ".spans.json", "w"))
    nb = sum(s["bad"] for s in segs)
    print(f"{len(segs)} сегментів, збійних {nb} ({100 * nb // max(len(segs), 1)}%); "
          f"на перерозпізнавання {len(m)} відрізків, {sum(y - x for x, y in m) / 60:.1f} хв → {base}.spans.json")


def cmd_fix(a):
    """Перерозпізнає лише відрізки зі spans.json (large-v3, хмара). Можна перервати й повторити."""
    import numpy as np
    from faster_whisper import WhisperModel
    wav = to_wav(a.audio)
    spans = json.load(open(a.spans))
    out = a.out or os.path.splitext(a.spans)[0] + ".fix.jsonl"
    done = {json.loads(l)["span"] for l in open(out, encoding="utf-8")} if os.path.exists(out) else set()
    m = WhisperModel(MODELS.get(a.model, a.model), device="cpu", compute_type="int8", cpu_threads=os.cpu_count())
    w = wave.open(wav); sr = w.getframerate()
    prompt = make_prompt(a.names)
    t_all = time.time(); aud = 0; n = 0
    for i, (x, y) in enumerate(spans):
        if i in done:
            continue
        if a.limit and n >= a.limit:
            break
        n += 1
        w.setpos(int(x * sr))
        audio = np.frombuffer(w.readframes(int((y - x) * sr)), dtype=np.int16).astype(np.float32) / 32768.0
        segs, _ = m.transcribe(audio, beam_size=a.beam, initial_prompt=prompt, **decode(a))
        res = [{"s": round(s.start + x, 2), "e": round(s.end + x, 2), "t": s.text.strip(), "lp": round(s.avg_logprob, 2)}
               for s in segs if not is_phantom(s.text, prompt)]
        with open(out, "a", encoding="utf-8") as f:
            f.write(json.dumps({"span": i, "a": x, "b": y, "segs": res}, ensure_ascii=False) + "\n")
        aud += y - x
    print(f"{out}: {aud:.0f} с аудіо за {time.time() - t_all:.0f} с", file=sys.stderr)


def cmd_merge(a):
    """Зводить: якісні локальні сегменти поза відрізками + перерозпізнане в хмарі."""
    loc = [json.loads(l) for l in open(a.local, encoding="utf-8")]
    fix = [json.loads(l) for l in open(a.fix, encoding="utf-8")]
    spans = [(r["a"], r["b"]) for r in fix]
    # сегмент відкидається, лише якщо відрізок перекриває його більш ніж наполовину (запас ±0,5 с на межах
    # відрізків інакше «з'їдає» сусідні якісні сегменти)
    keep = [s for s in loc if not s.get("bad") and
            not any(min(s["e"], y) - max(s["s"], x) > max(0.3, 0.5 * (s["e"] - s["s"])) for x, y in spans)]
    allseg = [s for s in keep if not is_phantom(s["t"])] + [dict(s, src="fix") for r in fix for s in r["segs"] if not is_phantom(s["t"])]
    allseg.sort(key=lambda s: s["s"])
    out = a.out or a.local.replace(".local.jsonl", ".merged.jsonl")
    with open(out, "w", encoding="utf-8") as f:
        for s in allseg:
            f.write(json.dumps(s, ensure_ascii=False) + "\n")
    print(f"{out}: {len(keep)} локальних + {len(allseg) - len(keep)} з хмари")


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
    elif a.dii:  # журнал-скан, переписаний у CSV: chas,rol,osoba,dia
        import csv
        rows = list(csv.DictReader(open(a.dii, encoding="utf-8")))
        st = [r["chas"] for r in rows if "Включено відеозапис" in r.get("dia", "")]
        start = start or (st[0] if st else rows[0]["chas"])
        dii = [(to_sec(r["chas"]), short(r["rol"], r["osoba"])) for r in rows if r.get("osoba")]
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
    t.add_argument("--beam", type=int, default=3)
    t.add_argument("--max-tokens", dest="max_tokens", type=int, default=None, help="ліміт токенів сегмента; 0 — без ліміту")
    t.add_argument("--from", dest="from_", type=float, default=0.0, help="секунда початку")
    t.add_argument("--to", type=float, default=None, help="секунда кінця")
    i = sp.add_parser("import"); i.add_argument("json"); i.add_argument("--duration", type=float)
    i.add_argument("--gap", type=float, default=4.0, help="пропуск між сегментами, що вважається втраченим мовленням, с")
    i.add_argument("--out"); i.add_argument("--names", default="", help="для відсіву відлуння підказки (JSONL)")
    x = sp.add_parser("fix"); x.add_argument("audio"); x.add_argument("spans"); x.add_argument("--model", default="large-v3")
    x.add_argument("--names", default=""); x.add_argument("--out")
    x.add_argument("--beam", type=int, default=3, help="1 — для фрагмента, на якому процес убивається через пам'ять")
    x.add_argument("--max-tokens", dest="max_tokens", type=int, default=None, help="0 — без ліміту (обрізані сегменти)")
    x.add_argument("--limit", type=int, default=0, help="скільки відрізків за запуск (порції; 0 — усі)")
    g = sp.add_parser("merge"); g.add_argument("local"); g.add_argument("fix"); g.add_argument("--out")
    d = sp.add_parser("draft"); d.add_argument("jsonl"); d.add_argument("--protokol"); d.add_argument("--start"); d.add_argument("--dii", help="журнал-скан у CSV: chas,rol,osoba,dia")
    d.add_argument("--out"); d.add_argument("--lag", type=float, default=3.0, help="запізнення позначок секретаря, с")
    d.add_argument("--min-lp", dest="min_lp", type=float, default=-0.6, help="поріг avg_logprob для позначки [?]")
    d.add_argument("--merge", type=float, default=0, help="зливати поспіль репліки однієї особи в рядок до N секунд (напр. 30)")
    a = ap.parse_args()
    {"fetch": cmd_fetch, "protokol": cmd_protokol, "transcribe": cmd_transcribe, "import": cmd_import,
     "fix": cmd_fix, "merge": cmd_merge, "draft": cmd_draft}[a.cmd](a)


if __name__ == "__main__":
    main()
