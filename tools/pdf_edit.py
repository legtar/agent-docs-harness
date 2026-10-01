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
import re
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


def font_name(name):
    return re.sub(r"[^a-z0-9]", "", name.split("+")[-1].lower())


def font_for(doc, page, span):
    for xref, ext, _, name, *_ in page.get_fonts(full=True):
        if font_name(name) == font_name(span["font"]):
            _, ext, _, buf = doc.extract_font(xref)
            if buf and ext not in ("n/a", ""):
                return fitz.Font(fontbuffer=buf), buf
    return None, None


def link_key(link):
    return (link.get("kind"), tuple(round(v, 3) for v in link["from"]), link.get("uri"), link.get("page"), str(link.get("to")), link.get("file"), link.get("nameddest"))


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
    if any(any(c in text for c in "\n\r\t") for pair in a.replace for text in pair):
        sys.exit("ERROR text replacement is single-line; tabs and line breaks are unsupported")
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
            if any(rect.intersects(widget.rect) for widget in page.widgets() or []):
                errors.append(f"{old!r}: text overlaps an interactive form field; edit the field value instead")
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
            plan.append((page.number, rect, s, buf, new, sw))
    for i, (number, rect, *_) in enumerate(plan):
        if any(number == other and rect.intersects(r) for other, r, *_ in plan[:i]):
            errors.append("overlapping replacement targets")
    if errors:
        print("\n".join("ERROR " + e for e in errors))
        sys.exit(1)
    if a.dry_run:
        return
    links = {number: doc[number].get_links() for number, *_ in plan}
    for number, rect, *_ in plan:
        doc[number].add_redact_annot(rect + (1, 1, -1, -1), fill=False)  # shrink: don't eat neighbour glyphs
    for number in {n for n, *_ in plan}:
        page = doc[number]
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    # Redaction invalidates PyMuPDF link caches; reopen the in-memory PDF before restoring links.
    if any(links.values()):
        refreshed = fitz.open(stream=doc.tobytes(), filetype="pdf")
        doc.close()
        doc = refreshed
    for number in links:
        page = doc[number]
        current = {link_key(link) for link in page.get_links()}
        for link in links[page.number]:
            if link_key(link) not in current:
                page.insert_link({key: value for key, value in link.items() if key not in ("xref", "id")})
    for i, (number, rect, s, buf, new, sw) in enumerate(plan):
        page = doc[number]
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
            for number, original_links in links.items():
                actual = {link_key(link) for link in chk[number].get_links()}
                if any(link_key(link) not in actual for link in original_links):
                    sys.exit("ERROR hyperlink preservation failed")
            for old, new in a.replace:
                if old not in new and any(p.search_for(old) for p in chk):
                    sys.exit(f"ERROR old text still extractable: {old!r}")
                if new.strip() and sum(len(p.search_for(new)) for p in chk) < a.count:
                    sys.exit(f"ERROR replacement not extractable: {new!r}")
    doc.close()
    print(f"wrote {a.dst}")


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, RuntimeError) as exc:
        print(f"ERROR {exc}; nothing published")
        sys.exit(1)
