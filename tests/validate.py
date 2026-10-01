"""End-to-end harness check on complex documents. Exit 1 on the first failed expectation.

    python tests/validate.py            # outputs in out/validation/ (look at */sheet.png)
    python tests/validate.py --offline  # skip the downloaded real-world PDF
"""
import os
import pathlib
import re
import subprocess
import sys
import urllib.request
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
from docx import Document  # noqa: E402
from docx_kit import md_to_docx  # noqa: E402
import fitz  # noqa: E402

OUT = ROOT / "out" / "validation"
OUT.mkdir(parents=True, exist_ok=True)
PY = sys.executable
fails = []


def tool(*args, expect=0):
    r = subprocess.run([PY, *map(str, args)], capture_output=True, text=True, encoding="utf-8", cwd=ROOT,
                       env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"})
    out = r.stdout + r.stderr
    ok = expect is None or r.returncode == expect
    name = " ".join(str(a) for a in args[:2])
    print(f"{'ok  ' if ok else 'FAIL'} [{r.returncode}] {name}")
    if not ok:
        print("     " + out.strip().replace("\n", "\n     ")[-1500:])
        fails.append(name)
    return out


def check(cond, msg):
    print(f"{'ok  ' if cond else 'FAIL'} {msg}")
    if not cond:
        fails.append(msg)


# Fast behavioral checks cover adversarial cases before the long rendering run.
tool("tests/regression.py")
tool("tests/torture.py")

P = [
    "Компания продолжила реализацию стратегии, утверждённой советом директоров в начале года. "
    "Ключевые инициативы — расширение продуктовой линейки, выход в новые регионы и повышение "
    "операционной эффективности — выполнены в срок и в рамках бюджета.",
    "Отдельное внимание уделялось качеству обслуживания: среднее время ответа поддержки сократилось "
    "с 4,2 до 2,7 часа, а доля обращений, решённых с первого контакта, выросла до 78 %. Это напрямую "
    "отразилось на удержании клиентов и рекомендациях (NPS +9 пунктов).",
    "Риски, выявленные в предыдущем периоде, в целом удалось снизить. Тем не менее сохраняется "
    "зависимость от двух крупнейших поставщиков, на которых приходится 41 % закупок; в следующем "
    "году планируется диверсификация.",
]
REGIONS = ["Москва", "Санкт-Петербург", "Новосибирск", "Екатеринбург", "Казань", "Нижний Новгород",
           "Челябинск", "Самара", "Омск", "Ростов-на-Дону", "Уфа", "Красноярск", "Воронеж", "Пермь"]


def long_rows(n=70):
    return [(f"{i + 1}", REGIONS[i % len(REGIONS)], f"{100 + i * 7.3:,.1f}".replace(",", " ").replace(".", ","),
             f"{(i * 37) % 23 + 2},{i % 10} %") for i in range(n)]


def complex_md():
    md = ["---", "title: Годовой отчёт о деятельности", "subtitle: Проверочный документ харнеса",
          "author: Отдел аналитики", "date: 1 октября 2026", "lang: ru-RU", "---", ""]
    for sec in range(1, 7):
        md += [f"# Раздел {sec}. Результаты направления", ""]
        md += [P[sec % 3], "", P[(sec + 1) % 3] + f"[^{sec}]", "", f"[^{sec}]: Примечание к разделу {sec}: данные управленческого учёта.", ""]
        md += ["## Ключевые выводы", "", "1. Выручка выросла быстрее рынка.", "2. Маржинальность улучшилась:",
               "    - за счёт подписок;", "    - за счёт автоматизации.", "3. Долговая нагрузка снижена.", ""]
        if sec == 2:
            md += ["Таблица: Динамика по регионам (длинная таблица на несколько страниц)", "",
                   "| № | Регион | Выручка, тыс. ₽ | Доля |", "|---|---|---:|---:|"]
            md += [f"| {a} | {b} | {c} | {d} |" for a, b, c, d in long_rows()]
            md += [""]
        if sec == 4:
            md += ["Таблица: Широкая таблица с семью колонками", "",
                   "| Показатель | I кв. | II кв. | III кв. | IV кв. | Итого | Δ г/г |",
                   "|---|---:|---:|---:|---:|---:|---:|"]
            for name in ("Выручка, млн ₽", "Себестоимость, млн ₽", "Валовая прибыль, млн ₽", "EBITDA, млн ₽", "Чистая прибыль, млн ₽"):
                md += [f"| {name} | 101,2 | 112,9 | 118,4 | 131,0 | 463,5 | +14,2 % |"]
            md += [""]
        md += ["> «Наша цель — не рост любой ценой, а устойчивое развитие». — из обращения CEO", ""]
        md += ["## Детали", "", P[(sec + 2) % 3], ""]
    return "\n".join(md)


def complex_typ(fixed=True):
    rows = ",\n    ".join(f"[{a}], [{b}], [{c}], [{d}]" for a, b, c, d in long_rows())
    secs = []
    for sec in range(1, 7):
        s = f"= Результаты направления {sec}\n\n{P[sec % 3]}\n\n{P[(sec + 1) % 3]}#footnote[Данные управленческого учёта, раздел {sec}.]\n\n"
        s += "== Ключевые выводы\n\n+ Выручка выросла быстрее рынка.\n+ Маржинальность улучшилась:\n  - за счёт подписок;\n  - за счёт автоматизации.\n+ Долговая нагрузка снижена.\n\n"
        if sec == 2:
            s += "#pagebreak(weak: true)\n" if fixed else ""  # the fix QA asks for
            s += ("#figure(table(columns: (auto, 1fr, auto, auto), align: (right, left, right, right),\n"
                  "    table.header[№][Регион][Выручка, тыс. ₽][Доля],\n    " + rows +
                  "),\n  caption: [Динамика по регионам]) <long>\n\n")
        if sec == 4:
            s += ("#figure(table(columns: (1fr,) + (auto,) * 6, align: (left,) + (right,) * 6,\n"
                  "    table.header[Показатель][I кв.][II кв.][III кв.][IV кв.][Итого][Δ г/г],\n"
                  "    [Выручка, млн ₽], [101,2], [112,9], [118,4], [131,0], [486,2], [+14,2 %],\n"
                  "    [EBITDA, млн ₽], [31,0], [33,4], [36,9], [40,2], [141,5], [+18,9 %],\n"
                  "  ), caption: [Широкая таблица])\n\n")
        s += "#callout(title: [Вывод])[Показатели направления выполнены; отклонения в пределах 3 %.]\n\n"
        s += f"== Детали\n\n{P[(sec + 2) % 3]}\n\n"
        secs.append(s)
    return ('#import "/templates/report.typ": *\n'
            '#show: report.with(title: "Годовой отчёт о деятельности", subtitle: "Проверочный документ харнеса",\n'
            '  author: "Отдел аналитики", date: "1 октября 2026", toc: true)\n\n' + "".join(secs))


def pages_changed(qa_out):
    m = re.search(r"pages with text changes: (.*)", qa_out)
    return m.group(1) if m else None


# ---------------------------------------------------------------- 1. new DOCX from Markdown
md = OUT / "complex.md"
md.write_text(complex_md(), encoding="utf-8")
cdocx = md_to_docx(md, OUT / "complex.docx", toc=True)
out = tool("tools/qa.py", cdocx, expect=None)
if "table starts at the page bottom" in out:
    # LibreOffice ignores keep-with-next inside tables (Word honours it): fix = page break before caption
    pg = re.search(r"p(\d+): table starts", out).group(1)
    with fitz.open(ROOT / "out/qa/new/complex.docx/complex.pdf") as d:
        cap = next(l for l in d[int(pg) - 1].get_text().splitlines() if l.startswith("Таблица:"))
    fixed = OUT / "complex_fixed.docx"
    tool("tools/docx_edit.py", cdocx, fixed, "--page-break-before", cap)
    cdocx = fixed
    out = tool("tools/qa.py", cdocx)
n_pages = int(re.search(r"(\d+) pages", out).group(1))
check(n_pages >= 6, f"complex.docx has many pages ({n_pages})")
with fitz.open(ROOT / "out/qa/new" / cdocx.name / (cdocx.stem + ".pdf")) as d:
    heads = sum("Регион" in d[i].get_text() and "Доля" in d[i].get_text() for i in range(len(d)))
check(heads >= 2, f"long docx table repeats its header on every page it spans ({heads} pages)")

# ---------------------------------------------------------------- 2. new PDF from Typst
typ = ROOT / "out" / "validation" / "complex.typ"
typ.write_text(complex_typ(fixed=False), encoding="utf-8")
out = tool("tools/qa.py", typ, expect=1)
check("table starts at the page bottom" in out or "overprinted" in out, "QA catches a broken long table (orphaned start or overprint)")
typ.write_text(complex_typ(fixed=True), encoding="utf-8")
out = tool("tools/qa.py", typ)  # the fixed layout passes
tpdf = ROOT / "out/qa/new/complex.typ/complex.pdf"
with fitz.open(tpdf) as d:
    heads = sum("Регион" in p.get_text() and "Доля" in p.get_text() for p in d)
    toc = "Содержание" in d[0].get_text()
check(heads >= 2, f"long Typst table repeats its header across pages ({heads} pages)")
check(toc, "Typst TOC present")
base_pdf = OUT / "complex_typst.pdf"
base_pdf.write_bytes(tpdf.read_bytes())

# ---------------------------------------------------------------- 3. careful DOCX edits
doc = Document(str(cdocx))
p = doc.paragraphs[[i for i, x in enumerate(doc.paragraphs) if x.style.name == "Heading 1"][0] + 1]
anchor = p.insert_paragraph_before("")
anchor.style = p.style
anchor.add_run("Срок ")
anchor.add_run("оплаты").bold = True
anchor.add_run(" составляет 30 дней с даты подписания акта.")
base = OUT / "base.docx"
doc.save(str(base))

tool("tools/docx_edit.py", base, "--tables")
plain = OUT / "edit_plain.docx"
tool("tools/docx_edit.py", base, plain, "--replace", "оплаты составляет 30 дней", "оплаты составляет 45 дней",
     "--cell", "0", "3", "2", "999,9")
d2 = Document(str(plain))
txt = "\n".join(x.text for x in d2.paragraphs)
check("оплаты составляет 45 дней" in txt and "30 дней" not in txt, "cross-run replace applied")
run = next(r for x in d2.paragraphs for r in x.runs if r.text.startswith("оплаты"))
check(run.bold, "formatting of the run where the match starts is kept (bold)")
check(d2.tables[0].rows[3].cells[2].text == "999,9", "table cell edited")
with zipfile.ZipFile(base) as a, zipfile.ZipFile(plain) as b:
    same = [n for n in a.namelist() if n != "word/document.xml" and a.read(n) == b.read(n)]
    check(len(same) == len(a.namelist()) - 1, "all other ZIP parts byte-identical")
out = tool("tools/qa.py", plain, "--original", base)
check(pages_changed(out) not in (None, "none"), f"edit visible only on expected pages: {pages_changed(out)}")

tool("tools/docx_edit.py", base, OUT / "x.docx", "--replace", "Ключевые выводы", "Итоги", expect=1)  # 6 matches -> refuse
tool("tools/docx_edit.py", base, base, "--replace", "Срок", "Период", expect=1)                       # overwrite -> refuse

tracked = OUT / "edit_tracked.docx"
tool("tools/docx_edit.py", base, tracked, "--track", "Claude",
     "--replace", "оплаты составляет 30 дней", "оплаты составляет 45 дней",
     "--cell", "0", "1", "1", "Москва (уточнено)",
     "--add-row", "0", "-1", "71|Тюмень|612,0|3,1 %",
     "--del-row", "0", "2")
xml = zipfile.ZipFile(tracked).read("word/document.xml").decode("utf8")
check(xml.count("<w:ins ") >= 3 and xml.count("<w:del ") >= 3 and "w:delText" in xml,
      "tracked changes written (w:ins/w:del/w:delText)")
skill_val = next(pathlib.Path.home().glob(".claude/skills/synced/*/docx/scripts/office/validate.py"), None)
if skill_val:
    tool(skill_val, tracked, "--original", base, "--author", "Claude")
tool("tools/qa.py", tracked, "--original", base, "--allow-reflow")
added = OUT / "edit_rows.docx"
tool("tools/docx_edit.py", base, added, "--add-row", "0", "-1", "71|Тюмень|612,0|3,1 %", "--del-row", "0", "1")
t = Document(str(added)).tables[0]
check(t.rows[-1].cells[1].text == "Тюмень" and t.rows[1].cells[1].text != "Москва", "row added (cloned format) and row deleted")

# ---------------------------------------------------------------- 3b. block operations on the report
import json  # noqa: E402
ops = [
    {"op": "insert", "before": "Раздел 3. Результаты направления", "like": "Раздел 1. Результаты направления",
     "text": "Раздел 2а. Новое направление"},
    {"op": "insert", "after": "Раздел 2а. Новое направление", "like": "Срок оплаты",
     "texts": ["Первый абзац нового раздела.", "Второй абзац нового раздела."]},
    {"op": "insert_table", "after": "Второй абзац нового раздела.", "like": "Себестоимость",
     "rows": [["Показатель", "I кв.", "II кв.", "III кв.", "IV кв.", "Итого", "Δ г/г"],
              ["Новая метрика", "1,0", "2,0", "3,0", "4,0", "10,0", "+5,0 %"]]},
    {"op": "insert_table", "after": "Первый абзац нового раздела.", "rows": [["Ключ", "Значение"], ["A", "1,5"]]},
    {"op": "delete_section", "heading": "Раздел 5. Результаты направления"},
    {"op": "set_text", "block": "Срок оплаты", "text": "Срок оплаты — 45 банковских дней."},
    {"op": "move", "block": "Срок оплаты", "after": "Первый абзац нового раздела."},
    {"op": "replace", "find": "из обращения CEO", "replace": "из письма CEO", "count": 5},
]
opsf = OUT / "ops.json"
opsf.write_text(json.dumps(ops, ensure_ascii=False, indent=1), encoding="utf-8")
tool("tools/docx_edit.py", base, "--blocks")
blk = OUT / "edit_blocks.docx"
tool("tools/docx_edit.py", base, blk, "--ops", opsf)
d3 = Document(str(blk))
heads = [p.text for p in d3.paragraphs if p.style.name.lower().startswith("heading 1")]
check("Раздел 2а. Новое направление" in heads and not any(h.startswith("Раздел 5.") for h in heads),
      f"section inserted as Heading 1 and section 5 deleted: {heads}")
paras = {p.text: p for p in d3.paragraphs}
like_style = Document(str(base)).paragraphs[[p.text for p in Document(str(base)).paragraphs].index(
    "Срок оплаты составляет 30 дней с даты подписания акта.")].style.name
check(paras["Первый абзац нового раздела."].style.name == like_style, "inserted paragraph cloned the 'like' style")
order = [p.text for p in d3.paragraphs]
check(order.index("Срок оплаты — 45 банковских дней.") == order.index("Первый абзац нового раздела.") + 1,
      "set_text + move: paragraph rewritten and relocated")
check(len(d3.tables) == len(Document(str(base)).tables) + 2, "two tables inserted (cloned + house style)")
newt = next(t for t in d3.tables if t.rows[-1].cells[0].text == "Новая метрика")
check(any(r.bold for r in newt.rows[0].cells[0].paragraphs[0].runs), "cloned table keeps header format")
tool("tools/qa.py", blk, "--original", base, "--allow-reflow")
blk_t = OUT / "edit_blocks_tracked.docx"
tool("tools/docx_edit.py", base, blk_t, "--ops", opsf, "--track", "Claude")
if skill_val:
    tool(skill_val, blk_t, "--original", base, "--author", "Claude")
tool("tools/qa.py", blk_t, "--original", base, "--allow-reflow")
bad = OUT / "ops_bad.json"
bad.write_text(json.dumps([{"op": "delete", "block": "Ключевые выводы"}], ensure_ascii=False), encoding="utf-8")
out = tool("tools/docx_edit.py", base, OUT / "x.docx", "--ops", bad, expect=1)
check("need exactly 1" in out, "ambiguous anchor refused (transaction rollback tested in regression.py)")

# mixed formatting inside the replaced phrase survives (only the changed chars are rewritten)
mix = OUT / "edit_mixed.docx"
tool("tools/docx_edit.py", base, mix, "--replace", "Срок оплаты составляет 30 дней", "Срок оплаты составляет 60 дней")
pm = next(p for p in Document(str(mix)).paragraphs if p.text.startswith("Срок"))
check(pm.text.startswith("Срок оплаты составляет 60 дней") and any(r.bold and r.text == "оплаты" for r in pm.runs),
      "minimal-diff replace keeps a bold word inside the phrase bold")

# ---------------------------------------------------------------- 4. careful PDF edits (Typst output)
pe = OUT / "pdf_edit.pdf"
tool("tools/pdf_edit.py", base_pdf, pe, "--replace", "486,2", "491,7")
out = tool("tools/qa.py", pe, "--original", base_pdf)
ch = pages_changed(out)
check(ch is not None and ch.count(",") == 0 and ch != "none", f"PDF number edit changed exactly one page: {ch}")
with fitz.open(pe) as d:
    check(sum(len(p.search_for("491,7")) for p in d) == 1 and not any(p.search_for("486,2") for p in d),
          "old value gone, new value present and extractable")
tool("tools/pdf_edit.py", base_pdf, OUT / "x.pdf", "--replace", "486,2", "486,2 млн рублей (оценка)", expect=1)

# ---------------------------------------------------------------- 5. real-world complex PDF
if "--offline" not in sys.argv:
    paper = OUT / "attention.pdf"
    if not paper.exists():
        try:
            urllib.request.urlretrieve("https://arxiv.org/pdf/1706.03762v7", paper)
        except OSError as e:
            print(f"skip real-world PDF: {e}")
    if paper.exists():
        tool("tools/qa.py", paper, expect=1)  # third-party PDF: QA must flag its non-embedded Times-Roman
        check("overlapping" not in tool("tools/qa.py", paper, expect=1), "no overprint false positives on LaTeX math")
        with fitz.open(paper) as d:
            tabs = [t for p in d for t in p.find_tables().tables]
            check(len(tabs) >= 2, f"PyMuPDF find_tables extracts tables from the paper ({len(tabs)})")
            if tabs:
                print("     first table:", tabs[0].extract()[:3])
        pe2 = OUT / "attention_edit.pdf"
        tool("tools/pdf_edit.py", paper, pe2, "--replace", "Attention Is All You Need", "Attention Is All We Need")
        out = tool("tools/qa.py", pe2, "--original", paper)
        check(pages_changed(out) == "[1]", f"real PDF: title edit touched only page 1 ({pages_changed(out)})")

print(f"\n{'FAILED: ' + str(len(fails)) if fails else 'ALL GREEN'}")
sys.exit(1 if fails else 0)
