"""Layout gate. Exit 1 on any ERROR. Run on every deliverable before handing it over.

    python tools/qa.py FILE                       # new document
    python tools/qa.py FILE --original ORIG       # careful edit: prove nothing else moved
    python tools/qa.py FILE --pages 3             # expected page count (exact)

FILE/ORIG: anything tools/render.py renders (.docx .typ .html .pdf ...).

Checks (rendered PDF):
  ERROR font not embedded / declared docx font substituted by renderer / tofu (missing glyph)
  ERROR text or image outside the page / in the margin beyond tolerance
  ERROR blank page, leftover placeholders ({{ }}, TODO, Lorem, broken refs)
  ERROR --original: page count changed, or pages whose text did not change moved visually
  WARN  heading stranded at the bottom of a page, nearly empty last page
"""
import argparse
import difflib
import pathlib
import re
import sys
import zipfile

import fitz

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from render import ROOT, render  # noqa: E402

PLACEHOLDERS = re.compile(
    r"\{\{|\}\}|\{%|TODO|TBD|XXX|Lorem ipsum|Error! Reference|Ошибка! Источник|"
    r"Ошибка! Закладка|Error! Bookmark|\?\?\?", re.I)
MARGIN_TOL = 2.0  # pt a glyph may poke into the margin before we call it overflow


class Report:
    def __init__(self):
        self.errors, self.warns = [], []

    def err(self, m): self.errors.append(m)
    def warn(self, m): self.warns.append(m)


def docx_fonts(path: pathlib.Path) -> set[str]:
    """Latin fonts the docx asks for (styles, theme, runs). Used to detect silent substitution."""
    if path.suffix.lower() != ".docx":
        return set()
    fonts = set()
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name in ("word/document.xml", "word/styles.xml") or name.startswith(("word/header", "word/footer")):
                fonts |= set(re.findall(r'w:(?:ascii|hAnsi|cs)="([^"]+)"', z.read(name).decode("utf8", "ignore")))
            elif name.startswith("word/theme/"):
                xml = z.read(name).decode("utf8", "ignore")
                fonts |= set(re.findall(r'<a:(?:major|minor)Font>\s*<a:latin typeface="([^"]+)"', xml))
    return {f for f in fonts if f and not f.startswith("+")}


FALLBACKS = ("dejavu", "liberation", "opensymbol", "notosans", "notoserif", "carlito", "caladea")


def installed_fonts() -> set[str]:
    """Normalised family names registered in Windows (machine + per-user). Empty off Windows."""
    try:
        import winreg
    except ImportError:
        return set()
    names = set()
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            k = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
        except OSError:
            continue
        for i in range(winreg.QueryInfoKey(k)[1]):
            n = winreg.EnumValue(k, i)[0]
            for part in re.sub(r"\s*\((TrueType|OpenType)\)", "", n).split(" & "):
                fam = re.sub(r"\s+(Bold|Italic|Light|Semibold|SemiBold|Black|Regular|Medium|Thin)\b.*", "", part)
                names.add(re.sub(r"[^a-z0-9]", "", fam.lower()))
    return names


def norm(name: str) -> str:
    name = name.split("+", 1)[-1]  # subset prefix ABCDEF+
    return re.sub(r"[^a-z0-9]", "", re.split(r"[-,]", name)[0].lower())


def text_margins(doc) -> tuple[float, float, float, float]:
    """Typical body margins = 5th percentile of text-line x0/x1 over the document."""
    xs0, xs1 = [], []
    for p in doc:
        for b in p.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                xs0.append(ln["bbox"][0]); xs1.append(ln["bbox"][2])
    if not xs0:
        return 0, 0, 0, 0
    xs0.sort(); xs1.sort()
    return xs0[len(xs0) // 20], xs1[-1 - len(xs1) // 20], 0, 0


def check_pdf(pdf: pathlib.Path, rep: Report, declared: set[str], expect_pages: int | None):
    doc = fitz.open(pdf)
    if expect_pages and len(doc) != expect_pages:
        rep.err(f"page count {len(doc)} != expected {expect_pages}")

    used = set()
    for i, p in enumerate(doc, 1):
        for f in p.get_fonts():
            used.add(norm(f[3]))
            if f[1] == "n/a" or f[1] == "":  # ext n/a => not embedded
                rep.err(f"p{i}: font '{f[3]}' is not embedded")
    have = installed_fonts()
    for f in sorted(declared):
        if have and norm(f) not in have:
            rep.err(f"docx font '{f}' is not installed here — renderer substitutes it, the page layout "
                    f"you see is NOT what Word shows. Install it or switch to a TOKENS font")
    fallback = {u for u in used if u.startswith(FALLBACKS)} - {norm(f) for f in declared}
    if declared and fallback:
        rep.err(f"renderer fell back to {sorted(fallback)}: some text uses a font that is not available")

    body_size = []
    for p in doc:
        for b in p.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                for s in ln["spans"]:
                    if s["text"].strip():
                        body_size.append(round(s["size"], 1))
    body = max(set(body_size), key=body_size.count) if body_size else 10

    for i, p in enumerate(doc, 1):
        W, H = p.rect.width, p.rect.height
        d = p.get_text("dict")
        txt = p.get_text()
        check_overprint(p, i, rep)
        if not txt.strip() and not p.get_images() and not p.get_drawings():
            rep.err(f"p{i}: blank page")
        if "�" in txt or "□" in txt:
            rep.err(f"p{i}: missing glyph (tofu/replacement char) — font lacks Cyrillic/symbol?")
        for m in PLACEHOLDERS.finditer(txt):
            rep.err(f"p{i}: placeholder/broken ref: …{txt[max(0, m.start()-30):m.end()+30]!r}…")
        for b in d["blocks"]:
            x0, y0, x1, y1 = b["bbox"]
            if x0 < -MARGIN_TOL or y0 < -MARGIN_TOL or x1 > W + MARGIN_TOL or y1 > H + MARGIN_TOL:
                kind = "image" if b["type"] == 1 else "text"
                rep.err(f"p{i}: {kind} block outside page bbox={tuple(round(v) for v in b['bbox'])}")
        for (x0, y0, x1, y1) in [r["rect"] for r in p.get_drawings() if r.get("rect")]:
            if x1 > W + MARGIN_TOL or x0 < -MARGIN_TOL:
                rep.err(f"p{i}: drawing (table border?) wider than page x={round(x0)}..{round(x1)}")
                break
        # Stranded heading: last text line on the page is larger/bolder than body, with more pages after.
        lines = [ln for b in d["blocks"] for ln in b.get("lines", []) if any(s["text"].strip() for s in ln["spans"])]
        if lines and i < len(doc):
            last = max(lines, key=lambda ln: ln["bbox"][3])
            sp = [s for s in last["spans"] if s["text"].strip()]
            if sp and sp[0]["size"] > body * 1.15 and last["bbox"][3] > H * 0.6:
                rep.warn(f"p{i}: heading '{''.join(s['text'] for s in sp)[:40]}' stranded at page bottom "
                         f"(set keep-with-next / break-after: avoid)")
    for i in range(len(doc) - 1):  # big hole mid-document: stale forced page break / kept block jumped
        p = doc[i]
        ys = [b["bbox"][3] for b in p.get_text("dict")["blocks"] if b["bbox"][3] < p.rect.height * 0.9]
        if ys and max(ys) < p.rect.height * 0.55 and i > 0:
            rep.warn(f"p{i+1}: content ends at {max(ys)/p.rect.height:.0%} of the page — check for a stale "
                     f"page break or a table/figure that jumped to the next page")
    if len(doc) > 1:
        last = doc[-1]
        ys = [b["bbox"][3] for b in last.get_text("dict")["blocks"]]
        if ys and max(ys) < last.rect.height * 0.15:
            rep.warn(f"last page is almost empty (content ends at {max(ys)/last.rect.height:.0%}) — tighten or pad")
    check_table_splits(doc, rep)
    # Per-document right edge overflow beyond the body column (e.g. a too-wide table)
    l, r, _, _ = text_margins(doc)
    for i, p in enumerate(doc, 1):
        for b in p.get_text("dict")["blocks"]:
            if b["type"] == 0 and b["bbox"][2] > r + 36 and b["bbox"][2] > p.rect.width - 20:
                rep.err(f"p{i}: text runs into right page edge (x1={round(b['bbox'][2])}, column ends ~{round(r)})")
    doc.close()


def table_rows(page):
    """Visual rows in the content zone: [(cells_count, text)], cells = lines sharing a baseline."""
    H = page.rect.height
    lines = sorted(((ln["bbox"][3], ln["bbox"][0], "".join(s["text"] for s in ln["spans"]).strip())
                    for b in page.get_text("dict")["blocks"] for ln in b.get("lines", [])
                    if H * 0.07 < ln["bbox"][3] < H * 0.93), key=lambda x: x[0])
    rows, cur, y0 = [], [], None
    for y, x, t in lines:
        if not t:
            continue
        if y0 is not None and abs(y - y0) > 2:
            rows.append(cur); cur = []
        if not cur:
            y0 = y
        cur.append((x, t))
    if cur:
        rows.append(cur)
    return [(len(r), " ".join(t for _, t in sorted(r))) for r in rows]


def check_overprint(page, i, rep: Report):
    """Text lines drawn on top of each other = content that did not fit and was stacked/clipped."""
    # ponytail: O(n^2) over lines per page; fine below a few thousand lines.
    # Only same-height lines of real text: math (fractions, radicals, sub/superscripts) legitimately
    # overlaps with differently sized boxes.
    boxes = [fitz.Rect(ln["bbox"]) for b in page.get_text("dict")["blocks"] for ln in b.get("lines", [])
             if len("".join(s["text"] for s in ln["spans"]).strip()) > 4]
    hits = 0
    for a in range(len(boxes)):
        for b in range(a + 1, len(boxes)):
            ha, hb = boxes[a].height, boxes[b].height
            if min(ha, hb) < 0.7 * max(ha, hb):
                continue
            inter = boxes[a] & boxes[b]
            if inter.is_valid and not inter.is_empty:
                small = min(boxes[a].get_area(), boxes[b].get_area()) or 1
                if inter.get_area() / small > 0.3:
                    hits += 1
    if hits:
        rep.err(f"p{i}: {hits} overlapping text line pair(s) — content is overprinted (overflowing "
                f"non-breakable block, negative spacing or absolute positioning)")


def check_table_splits(doc, rep: Report, min_rows=3):
    """A table that breaks across pages must leave >= min_rows body rows before the break.
    Detected via the repeated header row at the top of the next page."""
    # ponytail: needs a repeating header to see the split; tables without one are not checked.
    rows = [table_rows(p) for p in doc]
    for i in range(len(rows) - 1):
        nxt = [r for r in rows[i + 1] if r[0] >= 2]
        if not nxt or rows[i + 1].index(nxt[0]) > 1:
            continue
        head = nxt[0][1]
        cur = rows[i]
        idx = [k for k, r in enumerate(cur) if r[1] == head]
        if not idx:
            continue
        after = 0
        for r in cur[idx[-1] + 1:]:
            if r[0] < 2:
                break
            after += 1
        if after < min_rows:
            rep.err(f"p{i+1}: table starts at the page bottom with only {after} row(s) before the break "
                    f"— move it to the next page or keep its first rows together")


def page_texts(pdf):
    with fitz.open(pdf) as d:
        return [re.sub(r"\s+", " ", p.get_text()).strip() for p in d]


def pixel_diff(a, b, dpi=40) -> float:
    pa, pb = a.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY), b.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
    if (pa.width, pa.height) != (pb.width, pb.height):
        return 1.0
    sa, sb = pa.samples, pb.samples
    return sum(1 for x, y in zip(sa, sb) if abs(x - y) > 40) / len(sa)


def check_edit(new_pdf, old_pdf, rep: Report, allow_reflow: bool):
    tn, to = page_texts(new_pdf), page_texts(old_pdf)
    if len(tn) != len(to):
        (rep.warn if allow_reflow else rep.err)(f"page count changed {len(to)} -> {len(tn)}")
    with fitz.open(new_pdf) as dn, fitz.open(old_pdf) as do:
        changed = []
        for i in range(min(len(dn), len(do))):
            px = pixel_diff(dn[i], do[i])
            if tn[i] != to[i]:
                changed.append(i + 1)
            elif px > 0.0005:
                (rep.warn if allow_reflow else rep.err)(
                    f"p{i+1}: text identical but layout moved ({px:.2%} pixels) — unintended reflow/format change")
    print(f"pages with text changes: {changed or 'none'}")
    old, new = " ".join(to), " ".join(tn)
    sm = difflib.SequenceMatcher(None, old.split(" "), new.split(" "), autojunk=False)
    for op, a0, a1, b0, b1 in sm.get_opcodes():
        if op != "equal":
            print(f"  {op}: {' '.join(old.split(' ')[a0:a1])[:120]!r} -> {' '.join(new.split(' ')[b0:b1])[:120]!r}")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--original")
    ap.add_argument("--pages", type=int)
    ap.add_argument("--allow-reflow", action="store_true", help="edit is expected to move later pages")
    a = ap.parse_args()
    src = pathlib.Path(a.file)
    rep = Report()
    pdf, pages = render(src, ROOT / "out" / "qa" / "new" / src.name)
    check_pdf(pdf, rep, docx_fonts(src), a.pages)
    if src.suffix.lower() == ".docx" and re.search(r"TOC\s+\\o", zipfile.ZipFile(src).read("word/document.xml").decode("utf8", "ignore")):
        with fitz.open(pdf) as d:
            head = "\n".join(d[i].get_text() for i in range(min(3, len(d))))
        if len(re.findall(r"\.{5,}\s*\d+\s*$", head, re.M)) < 2:
            rep.err("document has a TOC field but the rendered TOC is empty (fields not updated)")
    if a.original:
        o = pathlib.Path(a.original)
        opdf, _ = render(o, ROOT / "out" / "qa" / "orig" / o.name)
        # Defects already in the original are not the edit's fault: report, don't block.
        orep = Report()
        check_pdf(opdf, orep, docx_fonts(o), None)
        old = set(orep.errors)
        for e in [e for e in rep.errors if e in old]:
            rep.errors.remove(e)
            rep.warn(f"(pre-existing in original) {e}")
        check_edit(pdf, opdf, rep, a.allow_reflow)
    for w in rep.warns:
        print("WARN ", w)
    for e in rep.errors:
        print("ERROR", e)
    print(f"{'FAIL' if rep.errors else 'PASS'}: {len(pages)} pages, {len(rep.errors)} errors, "
          f"{len(rep.warns)} warnings. Look at: {pdf.parent / 'sheet.png'}")
    sys.exit(1 if rep.errors else 0)


if __name__ == "__main__":
    main()
