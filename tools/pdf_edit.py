"""In-place text replacement in a PDF that has no editable source. Last resort: if a .docx /
.typ / .html source exists, edit that and re-render instead.

    python tools/pdf_edit.py IN.pdf OUT.pdf --replace "old" "new" [--replace ...] [--count N] [--dry-run]

Per edit: must match exactly --count times (default 1). Old text is really removed (redaction,
not a white box), images and vector art under it are kept, the new text is set in the same
embedded font / size / colour on the same baseline. Refuses when the new text is wider than
the old one (+3%) or the embedded subset lacks a glyph: PDF text does not reflow, so that
would break the layout. Use a shorter wording or regenerate the PDF.
"""
import argparse
import pathlib
import sys

from safe_output import staged_output

import fitz


def span_at(page, rect):
    best = None
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            for s in ln["spans"]:
                r = fitz.Rect(s["bbox"])
                if r.intersects(rect) and (best is None or (r & rect).get_area() > (fitz.Rect(best["bbox"]) & rect).get_area()):
                    best = s
    return best


def font_for(doc, page, span):
    for xref, ext, _, name, *_ in page.get_fonts(full=True):
        if name.split("+")[-1] == span["font"].split("+")[-1] or name == span["font"]:
            _, ext, _, buf = doc.extract_font(xref)
            if buf and ext not in ("n/a", ""):
                return fitz.Font(fontbuffer=buf), buf
    return None, None


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--replace", nargs=2, action="append", required=True, metavar=("OLD", "NEW"))
    ap.add_argument("--count", type=int, default=1)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    if pathlib.Path(a.src).resolve() == pathlib.Path(a.dst).resolve():
        sys.exit("refusing to overwrite the original")
    if a.count < 1 or any(not old.strip() for old, _ in a.replace):
        sys.exit("ERROR count must be positive and search text nonempty")
    doc = fitz.open(a.src)
    plan, errors = [], []
    for old, new in a.replace:
        hits = [(p, r) for p in doc for r in p.search_for(old)]
        print(f"{old!r} -> {new!r}: {len(hits)} match(es) on pages {sorted({p.number + 1 for p, _ in hits})}")
        if len(hits) != a.count:
            errors.append(f"{old!r}: {len(hits)} matches, expected {a.count}")
            continue
        for page, rect in hits:
            if page.rotation:
                errors.append("rotated pages: edit the source instead")
                continue
            s = span_at(page, rect)
            lines = [ln for b in page.get_text("dict")["blocks"] for ln in b.get("lines", []) if any(fitz.Rect(sp["bbox"]).intersects(rect) for sp in ln["spans"])]
            spans = [sp for ln in lines for sp in ln["spans"] if fitz.Rect(sp["bbox"]).intersects(rect)]
            if len(lines) != 1 or lines[0].get("dir") != (1.0, 0.0) or len({(sp["font"], round(sp["size"], 3), sp["color"]) for sp in spans}) != 1:
                errors.append(f"{old!r}: multiline, rotated or mixed-format match is unsupported")
                continue
            font, buf = font_for(doc, page, s) if s else (None, None)
            if not font:
                errors.append(f"{old!r} p{page.number+1}: embedded font not extractable")
                continue
            miss = [c for c in new if c.strip() and not font.has_glyph(ord(c))]
            if miss:
                errors.append(f"{old!r} p{page.number+1}: subset font {s['font']} lacks glyphs {miss}")
                continue
            size, n_sp = s["size"], old.count(" ")
            ink = sum(font.text_length(w, fontsize=size) for w in old.split(" "))
            sw = (rect.width - ink) / n_sp if n_sp else size * 0.25  # the original's word spacing
            w_new = sum(font.text_length(w, fontsize=size) for w in new.split(" ")) + new.count(" ") * sw
            if w_new > rect.width * 1.03:
                errors.append(f"{old!r} p{page.number+1}: new text {w_new:.1f}pt wider than old {rect.width:.1f}pt — would overflow")
                continue
            plan.append((page, rect, s, buf, new, sw))
    for i, (page, rect, *_) in enumerate(plan):
        if any(page.number == other.number and rect.intersects(r) for other, r, *_ in plan[:i]):
            errors.append("overlapping replacement targets")
    if errors:
        print("\n".join("ERROR " + e for e in errors))
        sys.exit(1)
    if a.dry_run:
        return
    for page, rect, *_ in plan:
        page.add_redact_annot(rect + (1, 1, -1, -1), fill=False)  # shrink: don't eat neighbour glyphs
    for page in {p for p, *_ in plan}:
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    for i, (page, rect, s, buf, new, sw) in enumerate(plan):
        name = f"edit{i}"
        page.insert_font(fontname=name, fontbuffer=buf)
        c = s["color"]
        font, x = fitz.Font(fontbuffer=buf), rect.x0
        # Subset fonts often lack the space glyph (PDF producers position words instead), so do the same.
        for word in new.split(" "):
            if word:
                page.insert_text((x, s["origin"][1]), word, fontname=name, fontsize=s["size"],
                                 color=((c >> 16 & 255) / 255, (c >> 8 & 255) / 255, (c & 255) / 255))
            x += font.text_length(word, fontsize=s["size"]) + sw
    with staged_output(a.src, a.dst) as temp:
        doc.save(temp, garbage=3, deflate=True)
        with fitz.open(temp) as chk:
            if len(chk) != len(doc):
                sys.exit("ERROR page count changed")
            for old, new in a.replace:
                if old not in new and any(p.search_for(old) for p in chk):
                    sys.exit(f"ERROR old text still extractable: {old!r}")
                if new.strip() and sum(len(p.search_for(new)) for p in chk) < a.count:
                    sys.exit(f"ERROR replacement not extractable: {new!r}")
    doc.close()
    print(f"wrote {a.dst}")


if __name__ == "__main__":
    main()
