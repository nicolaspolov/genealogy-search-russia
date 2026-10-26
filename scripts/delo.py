#!/usr/bin/env python3
"""Дело по поиску предков: таблицы, проверка, отчёт.

Только стандартная библиотека Python 3.8+.

    python3 scripts/delo.py novoe   ~/rodoslovnaya/smirnovy
    python3 scripts/delo.py proverka ~/rodoslovnaya/smirnovy
    python3 scripts/delo.py otchet  ~/rodoslovnaya/smirnovy [--polnyi] [--bez-pdf]

Таблицы — CSV с разделителем «;» в кодировке UTF-8 с BOM: открываются в Excel и
LibreOffice двойным щелчком. Описание колонок и статусов — delo-i-otchet.md.
"""
import argparse
import csv
import datetime as dt
import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
TEMPLATES = SKILL_DIR / "shablony"
TABLES = ["lyudi.csv", "dokumenty.csv", "svyazi.csv", "poisk.csv"]

STATUSES = {
    "доказано": ("#1b7f3b", "есть прямой документ о связи и независимое подтверждение, что это те же люди"),
    "вероятно": ("#2f6db5", "признаки сходятся, но прямого документа или независимого подтверждения пока нет"),
    "гипотеза": ("#8a6d1d", "рабочее предположение; нужна проверка"),
    "спорно": ("#b5532f", "есть противоречащий документ, вопрос не решён"),
    "опровергнуто": ("#777777", "документы показывают, что связи нет"),
}
RANK = {"доказано": 4, "вероятно": 3, "гипотеза": 2, "спорно": 1, "опровергнуто": 0}
WEAK_KINDS = {"индекс", "дерево", "рассказ"}
PRIMARY_KINDS = {"оригинал", "скан оригинала"}
PARENT_WORDS = {"отец", "мать"}


# ---------- чтение таблиц ----------

def read_table(path):
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8-sig")
    if not text.strip():
        return []
    delim = ";" if text.splitlines()[0].count(";") >= text.splitlines()[0].count(",") else ","
    rows = []
    for n, row in enumerate(csv.DictReader(text.splitlines(), delimiter=delim), 2):
        if None in row:
            sys.exit(f"{path.name}, строка {n}: полей больше, чем колонок. Списки внутри поля "
                     f"(например, документы D1, D2) пишите через запятую, не через «{delim}».")
        row = {(k or "").strip(): (v or "").strip() for k, v in row.items()}
        if any(row.values()) and not row.get("id", "").upper().startswith("ПРИМЕР"):
            rows.append(row)
    return rows


def ids(value):
    return [x for x in re.split(r"[\s,;]+", value or "") if x]


def load(case):
    case = Path(case).expanduser()
    if not case.is_dir():
        sys.exit(f"Нет папки дела: {case}")
    data = {name[:-4]: read_table(case / name) for name in TABLES}
    data["delo"] = parse_delo(case / "delo.md")
    data["path"] = case
    return data


def parse_delo(path):
    info = {"text": "", "glavnyi": "", "vopros": ""}
    if not path.exists():
        return info
    text = re.sub(r"<!--.*?-->", "", path.read_text(encoding="utf-8"), flags=re.S)
    info["text"] = text
    m = re.search(r"^Главный человек:\s*(\S+)", text, re.M)
    if m:
        info["glavnyi"] = m.group(1).strip()
    m = re.search(r"^Вопрос:\s*(.+)$", text, re.M)
    if m:
        info["vopros"] = m.group(1).strip()
    return info


def section(text, title):
    m = re.search(rf"^##\s*{re.escape(title)}\s*$(.*?)(?=^##\s|\Z)", text, re.M | re.S)
    if not m:
        return ""
    body = "\n".join(l for l in m.group(1).strip().splitlines() if not l.strip().startswith("<!--"))
    return body.strip()


# ---------- проверка ----------

def check(data):
    errors, warnings = [], []
    people = {p["id"]: p for p in data["lyudi"] if p.get("id")}
    docs = {d["id"]: d for d in data["dokumenty"] if d.get("id")}

    for name, key in (("lyudi", "id"), ("dokumenty", "id"), ("svyazi", "id")):
        seen = set()
        for row in data[name]:
            rid = row.get(key, "")
            if not rid:
                errors.append(f"{name}.csv: строка без id: {row}")
            elif rid in seen:
                errors.append(f"{name}.csv: id {rid} повторяется")
            seen.add(rid)

    for p in data["lyudi"]:
        for d in ids(p.get("документы")):
            if d not in docs:
                errors.append(f"человек {p['id']}: нет документа {d} в dokumenty.csv")

    for d in data["dokumenty"]:
        if not d.get("группа"):
            warnings.append(f"документ {d['id']}: не указана группа независимости (копии одной записи — одна группа)")
        kind = d.get("вид", "")
        if kind and kind not in PRIMARY_KINDS | WEAK_KINDS | {"копия", "выписка", "справка", "публикация"}:
            warnings.append(f"документ {d['id']}: необычный вид «{kind}»")

    for s in data["svyazi"]:
        sid = s.get("id", "?")
        for col in ("кто", "кому"):
            if s.get(col) not in people:
                errors.append(f"связь {sid}: в колонке «{col}» нет человека {s.get(col)!r} в lyudi.csv")
        status = s.get("статус", "")
        if status not in STATUSES:
            errors.append(f"связь {sid}: статус «{status}» не из списка: {', '.join(STATUSES)}")
        direct = s.get("прямой_документ", "")
        support = ids(s.get("подтверждения"))
        against = ids(s.get("против"))
        for d in ids(direct) + support + against:
            if d not in docs:
                errors.append(f"связь {sid}: нет документа {d} в dokumenty.csv")
        if direct and direct in docs and docs[direct].get("вид") in WEAK_KINDS:
            errors.append(f"связь {sid}: прямым документом указан «{docs[direct].get('вид')}» ({direct}) — это наводка, не документ")
        if status == "доказано":
            if not direct:
                errors.append(f"связь {sid}: «доказано» без прямого документа")
            elif direct in docs:
                if docs[direct].get("вид") not in PRIMARY_KINDS | WEAK_KINDS:
                    warnings.append(f"связь {sid}: прямой документ {direct} — «{docs[direct].get('вид')}», не оригинал и не скан оригинала")
                group = docs[direct].get("группа") or direct
                other = {docs[d].get("группа") or d for d in support if d in docs} - {group}
                if not other:
                    errors.append(f"связь {sid}: «доказано», но нет подтверждения из другой группы независимости (копия или индекс той же записи не считается)")
            if against:
                errors.append(f"связь {sid}: «доказано» при документах против ({', '.join(against)}) — разобрать противоречие или поставить «спорно»")
        if status in ("доказано", "вероятно") and not (direct or support):
            errors.append(f"связь {sid}: «{status}» без единого документа")

    for i, row in enumerate(data["poisk"], 1):
        res = row.get("результат", "").lower()
        if res.startswith(("не найден", "нет", "пусто", "не встрет")) and not row.get("охват"):
            warnings.append(f"поиск, строка {i}: отрицательный результат без охвата — это «не нашёл», а не «нет»")

    for p in people.values():
        if p.get("жив", "").lower() == "да":
            warnings.append(f"человек {p['id']} жив: в отчёте будут только имя и год рождения (полный вид — ключ --polnyi)")

    glav = data["delo"].get("glavnyi")
    if glav and glav not in people:
        errors.append(f"delo.md: главный человек {glav} не найден в lyudi.csv")
    return errors, warnings


# ---------- древо ----------

def parents_of(data):
    par = {}
    for s in data["svyazi"]:
        if s.get("кем_приходится") in PARENT_WORDS:
            par.setdefault(s["кому"], []).append(s)
    return par


def ancestor_rows(data, root, people):
    """Строки: (колено, человек, связь, статус цепочки = самое слабое звено)."""
    par = parents_of(data)
    out, seen = [], set()

    def walk(pid, depth, chain_status):
        for s in sorted(par.get(pid, []), key=lambda x: x.get("кем_приходится") != "отец"):
            anc = s["кто"]
            st = s.get("статус", "гипотеза")
            chain = st if chain_status is None or RANK.get(st, 0) < RANK.get(chain_status, 0) else chain_status
            out.append((depth, anc, s, chain))
            if (anc, depth) not in seen and depth < 30:
                seen.add((anc, depth))
                walk(anc, depth + 1, chain)

    if root in people:
        walk(root, 1, None)
    return sorted(out, key=lambda r: r[0])


# ---------- отчёт ----------

def esc(x):
    return html.escape(x or "")


def badge(status):
    color = STATUSES.get(status, ("#555", ""))[0]
    return f'<span class="b" style="background:{color}">{esc(status or "—")}</span>'


def person_label(p, full):
    name = p.get("имя_как_в_документе") or p.get("id")
    years = " – ".join(x for x in (year(p.get("рождение")), year(p.get("смерть"))) if x)
    if p.get("жив", "").lower() == "да" and not full:
        y = year(p.get("рождение"))
        return f"{esc(name)}{' (' + y + ')' if y else ''}"
    place = p.get("место_рождения", "")
    tail = ", ".join(x for x in (years, place) if x)
    return f"{esc(name)}{' <span class=m>(' + esc(tail) + ')</span>' if tail else ''}"


def year(value):
    m = re.search(r"(1[5-9]\d\d|20\d\d)", value or "")
    return m.group(1) if m else (value or "")


def generation_name(n):
    names = {1: "родители", 2: "деды и бабушки", 3: "прадеды", 4: "прапрадеды"}
    return names.get(n, f"{n}-е колено")


def build_html(data, full):
    people = {p["id"]: p for p in data["lyudi"] if p.get("id")}
    docs = {d["id"]: d for d in data["dokumenty"] if d.get("id")}
    delo = data["delo"]
    root = delo.get("glavnyi")
    today = dt.date.today().strftime("%d.%m.%Y")
    title = delo.get("vopros") or "Родословное исследование"

    counts = {k: 0 for k in STATUSES}
    for s in data["svyazi"]:
        if s.get("статус") in counts:
            counts[s["статус"]] += 1

    parts = []
    parts.append(f"<h1>{esc(title)}</h1><p class=m>Отчёт от {today}. Главный человек: "
                 f"{person_label(people[root], full) if root in people else '—'}</p>")

    itog = section(delo["text"], "Итог")
    if itog:
        parts.append("<h2>Главное</h2>" + md_paragraphs(itog))

    parts.append("<h2>Как читать статусы</h2><table class=st>" + "".join(
        f"<tr><td>{badge(k)}</td><td>{esc(v[1])}</td><td class=n>{counts[k]}</td></tr>" for k, v in STATUSES.items())
        + "</table><p class=m>Дальнее родство складывается из звеньев «родитель — ребёнок». "
          "Статус предка в схеме — по самому слабому звену цепочки до главного человека.</p>")

    rows = ancestor_rows(data, root, people)
    if rows:
        parts.append("<h2>Предки</h2>")
        current = None
        for depth, anc, s, chain in rows:
            if depth != current:
                if current is not None:
                    parts.append("</ul>")
                parts.append(f"<h3>{esc(generation_name(depth))}</h3><ul class=tree>")
                current = depth
            child = people.get(s["кому"], {}).get("имя_как_в_документе", s["кому"])
            parts.append(f"<li>{badge(chain)} <b>{person_label(people.get(anc, {'id': anc}), full)}</b>"
                         f" <span class=m>— {esc(s.get('кем_приходится'))} {esc(child)}</span></li>")
        parts.append("</ul>")

    parts.append("<h2>Связи по звеньям</h2><table><thead><tr><th>Кто</th><th>Кем приходится</th><th>Кому</th>"
                 "<th>Статус</th><th>Прямой документ</th><th>Подтверждения</th><th>Против</th><th>Что проверить</th></tr></thead>")
    for s in data["svyazi"]:
        a = people.get(s.get("кто"), {"id": s.get("кто")})
        b = people.get(s.get("кому"), {"id": s.get("кому")})
        parts.append(f"<tr><td>{person_label(a, full)}</td><td>{esc(s.get('кем_приходится'))}</td>"
                     f"<td>{person_label(b, full)}</td><td>{badge(s.get('статус'))}</td>"
                     f"<td>{esc(s.get('прямой_документ'))}</td><td>{esc(s.get('подтверждения'))}</td>"
                     f"<td>{esc(s.get('против'))}</td><td>{esc(s.get('что_проверить'))}</td></tr>")
    parts.append("</table>")

    parts.append("<h2>Люди</h2>")
    for p in data["lyudi"]:
        alive = p.get("жив", "").lower() == "да"
        lines = [f"<div class=card><b>{esc(p.get('id'))}. {person_label(p, full)}</b>"]
        if not alive or full:
            for label, col in (("Варианты написания", "варианты_написания"), ("Рождение", "рождение"),
                               ("Место рождения", "место_рождения"), ("Смерть", "смерть"),
                               ("Место смерти", "место_смерти"), ("Сословие, занятие", "сословие_занятие"),
                               ("Вероисповедание", "вероисповедание"), ("Документы", "документы"),
                               ("Примечание", "примечание")):
                if p.get(col):
                    lines.append(f"<div><span class=m>{label}:</span> {esc(p[col])}</div>")
        else:
            lines.append("<div class=m>Живой человек: подробности не показаны.</div>")
        lines.append("</div>")
        parts.append("".join(lines))

    parts.append("<h2>Документы</h2><table><thead><tr><th>ID</th><th>Что за документ</th><th>Шифр</th>"
                 "<th>Сайт, кадр</th><th>Вид</th><th>Группа</th><th>Что говорит</th></tr></thead>")
    for d in data["dokumenty"]:
        site = d.get("сайт_и_кадр", "")
        site_html = (f'<a href="{esc(site)}">{esc(site)}</a>' if site.startswith("http") else esc(site))
        parts.append(f"<tr><td>{esc(d.get('id'))}</td><td>{esc(d.get('что_за_документ'))}"
                     f"{' <span class=m>(' + esc(d.get('дата_документа')) + ')</span>' if d.get('дата_документа') else ''}</td>"
                     f"<td>{esc(d.get('архив_шифр'))}</td><td>{site_html}</td><td>{esc(d.get('вид'))}</td>"
                     f"<td>{esc(d.get('группа'))}</td><td>{esc(d.get('что_говорит'))}</td></tr>")
    parts.append("</table>")

    if data["poisk"]:
        parts.append("<h2>Где искали</h2><p class=m>Пустой результат здесь значит «не найдено в "
                     "указанном охвате», а не «записи не существует».</p><table><thead><tr><th>Дата</th><th>Где</th>"
                     "<th>Что искали</th><th>Охват</th><th>Написания</th><th>Результат</th><th>Дальше</th></tr></thead>")
        for r in data["poisk"]:
            parts.append(f"<tr><td>{esc(r.get('дата'))}</td><td>{esc(r.get('где'))}</td><td>{esc(r.get('что_искали'))}"
                         f"{'<br><span class=m>' + esc(r.get('запрос_или_шифр')) + '</span>' if r.get('запрос_или_шифр') else ''}</td>"
                         f"<td>{esc(r.get('охват'))}</td><td>{esc(r.get('написания'))}</td>"
                         f"<td>{esc(r.get('результат'))}</td><td>{esc(r.get('дальше'))}</td></tr>")
        parts.append("</table>")

    todo = [s for s in data["svyazi"] if s.get("статус") != "доказано" and s.get("что_проверить")]
    todo_text = section(delo["text"], "Открытые вопросы")
    if todo or todo_text:
        parts.append("<h2>Что дальше</h2>")
        if todo_text:
            parts.append(md_paragraphs(todo_text))
        if todo:
            parts.append("<ul>" + "".join(
                f"<li>{badge(s.get('статус'))} {esc(people.get(s['кто'], {}).get('имя_как_в_документе', s['кто']))} — "
                f"{esc(s.get('кем_приходится'))} {esc(people.get(s['кому'], {}).get('имя_как_в_документе', s['кому']))}: "
                f"{esc(s.get('что_проверить'))}</li>" for s in todo) + "</ul>")

    css = """
    @page { size: A4; margin: 16mm 14mm; }
    body { font-family: "PT Serif", Georgia, "Times New Roman", serif; color:#1d1d1d; font-size:10.5pt; line-height:1.45; }
    h1 { font-size:19pt; margin:0 0 4pt; } h2 { font-size:13.5pt; margin:18pt 0 6pt; border-bottom:1px solid #ccc; padding-bottom:2pt; page-break-after:avoid; break-after:avoid; }
    h3 { font-size:11pt; margin:10pt 0 3pt; color:#444; page-break-after:avoid; break-after:avoid; }
    thead { display:table-header-group; }
    .m { color:#666; font-size:9.5pt; }
    .b { color:#fff; border-radius:3pt; padding:0 5pt; font-size:8.5pt; font-family: Arial, sans-serif; white-space:nowrap; }
    table { border-collapse:collapse; width:100%; font-size:9pt; margin:4pt 0; }
    th, td { border:1px solid #d6d6d6; padding:3pt 4pt; vertical-align:top; text-align:left; }
    th { background:#f2f2f2; } tr { page-break-inside:avoid; }
    table.st td { border:none; padding:2pt 6pt 2pt 0; } td.n { color:#666; }
    ul.tree { margin:0 0 4pt 0; padding-left:14pt; } ul.tree li { margin:2pt 0; }
    .card { border:1px solid #e0e0e0; border-radius:4pt; padding:5pt 7pt; margin:5pt 0; page-break-inside:avoid; }
    a { color:#1f4f99; text-decoration:none; word-break:break-all; }
    """
    return ("<!doctype html><html lang=ru><head><meta charset=utf-8><title>" + esc(title) +
            "</title><style>" + css + "</style></head><body>" + "\n".join(parts) + "</body></html>")


def md_paragraphs(text):
    out = []
    for block in re.split(r"\n\s*\n", text.strip()):
        lines = [l.strip() for l in block.splitlines() if l.strip()]
        if lines and all(l.startswith(("- ", "* ")) for l in lines):
            out.append("<ul>" + "".join(f"<li>{esc(l[2:])}</li>" for l in lines) + "</ul>")
        else:
            out.append("<p>" + esc(" ".join(lines)) + "</p>")
    return "".join(out)


# ---------- PDF ----------

def find_browser():
    env = os.environ.get("CHROME_PATH")
    if env and Path(env).exists():
        return env
    for name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "msedge",
                 "microsoft-edge", "chrome"):
        p = shutil.which(name)
        if p:
            return p
    candidates = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "/Applications/Chromium.app/Contents/MacOS/Chromium",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    pw = Path.home() / ".cache" / "ms-playwright"
    for pattern in ("chromium_headless_shell-*/chrome-headless-shell-*/chrome-headless-shell",
                    "chromium-*/chrome-linux*/chrome"):
        for c in sorted(pw.glob(pattern), reverse=True):
            return str(c)
    return None


def to_pdf(html_path, pdf_path, limit=90):
    """Печать через браузер без окна. Браузер иногда не завершается сам после печати,
    поэтому ждём готовый файл и закрываем процесс."""
    browser = find_browser()
    if not browser:
        return False, "браузер Chrome, Chromium или Edge не найден"
    modes = ("",) if "headless-shell" in browser else ("--headless=new", "--headless")
    last = "браузер не создал PDF"
    for headless in modes:
        with tempfile.TemporaryDirectory() as profile:
            tmp_pdf = Path(profile) / "out.pdf"
            cmd = [browser] + ([headless] if headless else []) + [
                "--disable-gpu", "--no-sandbox", "--no-first-run", "--disable-extensions",
                "--disable-background-networking", f"--user-data-dir={profile}/u",
                "--no-pdf-header-footer", f"--print-to-pdf={tmp_pdf}", html_path.resolve().as_uri()]
            try:
                proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except OSError as exc:
                last = str(exc)
                continue
            size, stable, t0 = -1, 0, time.time()
            while time.time() - t0 < limit:
                if tmp_pdf.exists():
                    s = tmp_pdf.stat().st_size
                    stable = stable + 1 if s == size and s > 1000 else 0
                    size = s
                    if stable >= 2:
                        break
                if proc.poll() is not None and not tmp_pdf.exists():
                    break
                time.sleep(0.5)
            if proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
            if tmp_pdf.exists() and tmp_pdf.stat().st_size > 1000:
                shutil.copy(tmp_pdf, pdf_path)
                return True, browser
    return False, last


# ---------- команды ----------

def cmd_novoe(args):
    case = Path(args.case).expanduser()
    try:
        case.resolve().relative_to(SKILL_DIR)
        inside = True
    except ValueError:
        inside = False
    if inside:
        sys.exit("Дело нельзя создавать внутри папки скилла: держите личные данные отдельно.")
    case.mkdir(parents=True, exist_ok=True)
    made = []
    for name in TABLES + ["delo.md"]:
        target = case / name
        if target.exists():
            continue
        shutil.copy(TEMPLATES / name, target)
        made.append(name)
    (case / "kopii").mkdir(exist_ok=True)
    print(f"Дело: {case}")
    print("Создано: " + (", ".join(made) if made else "ничего нового, файлы уже были"))
    print("Сканы и фото документов кладите в папку kopii/.")


def cmd_proverka(args):
    data = load(args.case)
    errors, warnings = check(data)
    print(f"Людей {len(data['lyudi'])}, документов {len(data['dokumenty'])}, связей {len(data['svyazi'])}, "
          f"записей поиска {len(data['poisk'])}.")
    for e in errors:
        print("ОШИБКА: " + e)
    for w in warnings:
        print("внимание: " + w)
    if not errors and not warnings:
        print("Замечаний нет.")
    return 1 if errors else 0


def cmd_otchet(args):
    data = load(args.case)
    errors, warnings = check(data)
    for e in errors:
        print("ОШИБКА: " + e)
    if errors and not args.vse_ravno:
        print("Отчёт не собран: сначала исправьте ошибки (или ключ --vse-ravno).")
        return 1
    out_dir = data["path"] / "otchet"
    out_dir.mkdir(exist_ok=True)
    html_path = out_dir / "otchet.html"
    html_path.write_text(build_html(data, args.polnyi), encoding="utf-8")
    print(f"HTML: {html_path}")
    if args.bez_pdf:
        return 0
    pdf_path = out_dir / "otchet.pdf"
    ok, info = to_pdf(html_path, pdf_path)
    if ok:
        print(f"PDF: {pdf_path}")
    else:
        print(f"PDF не собран ({info}). Откройте HTML в браузере → Печать → Сохранить как PDF.")
    return 0


def main():
    ap = argparse.ArgumentParser(description="Дело по поиску предков: таблицы, проверка, отчёт.")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("novoe", help="создать папку дела с пустыми таблицами")
    p.add_argument("case")
    p.set_defaults(fn=cmd_novoe)
    p = sub.add_parser("proverka", help="проверить таблицы и статусы связей")
    p.add_argument("case")
    p.set_defaults(fn=cmd_proverka)
    p = sub.add_parser("otchet", help="собрать отчёт HTML и PDF")
    p.add_argument("case")
    p.add_argument("--polnyi", action="store_true", help="показывать подробности о живых людях")
    p.add_argument("--bez-pdf", action="store_true", help="только HTML")
    p.add_argument("--vse-ravno", action="store_true", help="собрать даже при ошибках проверки")
    p.set_defaults(fn=cmd_otchet)
    args = ap.parse_args()
    sys.exit(args.fn(args) or 0)


if __name__ == "__main__":
    main()
