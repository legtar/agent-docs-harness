"""In-place text replacement in a PDF that has no editable source. Last resort: if a .docx /
.typ / .html source exists, edit that and re-render instead.

    python tools/pdf_edit.py IN.pdf OUT.pdf --replace "old" "new" [--replace ...] [--count N] [--dry-run]

Per edit: must match exactly --count times (default 1, or a third argument per pair). Old text
is really removed (redaction, not a white box), images and vector art under it are kept, the new
text is set in the same font family / size / colour on the same baseline.

Layout rules:
* new text must fit the old box (+3%), otherwise it may grow to the LEFT only when the span ends
  at its line's right edge and the space to the left is verifiably free (a tab stop, a table
  column edge, a right-aligned number). PDF text does not reflow.
* the embedded subset is used when it has every needed glyph. If not, the same family is loaded
  from the system fonts and embedded instead, so Cyrillic or digits missing from the subset no
  longer block the edit. A different family is refused.
"""
import argparse
import os
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


# Subset -> system font of the same family, so a missing Cyrillic/digit glyph does not block an edit.
SYSTEM_FAMILIES = {
    "arial": ["arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"],
    "helvetica": ["arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"],
    "timesnewroman": ["times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"],
    "times": ["times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"],
    "couriernew": ["cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"],
    "calibri": ["calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"],
    "cambria": ["cambria.ttc", "cambriab.ttf", "cambriai.ttf", "cambriaz.ttf"],
    "georgia": ["georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf"],
    "verdana": ["verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf"],
    "tahoma": ["tahoma.ttf", "tahomabd.ttf", "tahomai.ttf", "tahomaz.ttf"],
    "segoeui": ["segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"],
}
SYSTEM_FONT_DIRS = [pathlib.Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts", pathlib.Path.home() / "AppData/Local/Microsoft/Windows/Fonts"]


def system_font(span, needed):
    """Regular/italic/bold/bold-italic file of the span's family from the system fonts."""
    files = SYSTEM_FAMILIES.get(font_name(span["font"]))
    if not files:
        return None, None
    flags = bool(span["flags"]) & (1 << 4)  # bold
    italic = "Italic" in span["font"] or "Oblique" in span["font"] or (bool(span["flags"]) & (1 << 1))
    order = ["arialbd.ttf" if flags else "arial.ttf", "arialbi.ttf" if flags and italic else "ariali.ttf" if italic else "arial.ttf"]
    order += files
    seen = set()
    for name in order:
        if name in seen:
            continue
        seen.add(name)
        for folder in SYSTEM_FONT_DIRS:
            path = folder / name
            if not path.exists():
                continue
            font = fitz.Font(fontfile=str(path))
            if all(font.has_glyph(ord(c)) for c in needed if c.strip()):
                return font, font.buffer
    return None, None


def embedded_font(doc, page, span, needed):
    """Font buffer to typeset `needed` with: the document's own subset when complete, else the
    system font of the same family. Returns (font, buffer, how)."""
    for xref, ext, _, name, *_ in page.get_fonts(full=True):
        if font_name(name) == font_name(span["font"]):
            _, ext, _, buf = doc.extract_font(xref)
            if buf and ext not in ("n/a", ""):
                font = fitz.Font(fontbuffer=buf)
                if all(font.has_glyph(ord(c)) for c in needed if c.strip()):
                    return font, buf, "embedded"
    font, buf = system_font(span, needed)
    if buf:
        return font, buf, f"system:{span['font']}"
    return None, None, ""


def same_line(r, rect):
    """True when two span boxes sit on one text line: they overlap vertically by most of their height."""
    top, bottom = max(r.y0, rect.y0), min(r.y1, rect.y1)
    if bottom <= top:
        return False
    return (bottom - top) >= 0.5 * min(r.height, rect.height)


def neighbours(page, rect):
    """Every span box on `rect`'s own line, `rect` included."""
    out = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            boxes = [fitz.Rect(sp["bbox"]) for sp in ln["spans"]]
            if any(same_line(r, rect) for r in boxes):
                out.extend(boxes)
    return out


def right_column_edge(page, rect):
    """The x1 of the column `rect` is right-aligned to, or None when its right edge is not a column edge.

    Evidence: at least one span on a nearby line ends at the same x (within 1pt), i.e. the document
    itself aligns that edge. Otherwise the text is left-aligned prose and must not be re-anchored.
    """
    edge, supporters = None, 0
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            for sp in ln["spans"]:
                r = fitz.Rect(sp["bbox"])
                if r.y0 > rect.y1 + 24 or r.y1 < rect.y0 - 24:
                    continue
                if abs(r.x1 - rect.x1) <= 1.0:
                    supporters += 1
                    edge = r.x1
    return edge if supporters >= 2 else None


def free_space_left(page, rect, extra=0.0):
    """Width of the free run immediately left of `rect` on its own line.

    The run starts at the leftmost glyph of the line (so a label on the same line bounds it) and
    ends at the nearest glyph that finishes before `rect`. Returns 0 when the run is unusable.
    """
    line = neighbours(page, rect)
    if not line:
        return 0.0
    limit = max((r.x1 for r in line if r.x1 <= rect.x0 + 0.5), default=min(r.x0 for r in line))
    return max(0.0, rect.x0 - limit - extra)


def link_key(link):
    return (link.get("kind"), tuple(round(v, 3) for v in link["from"]), link.get("uri"), link.get("page"), str(link.get("to")), link.get("file"), link.get("nameddest"))


def repair_form_resources(doc, page):
    """Restore font resources that apply_redactions dropped from the Form XObjects it creates.

    MuPDF moves the surviving content stream of a redacted page into a Form XObject but copies only
    part of the page's /Font dict into it, so untouched text set in the missing fonts silently
    disappears (e.g. a page number in a bold face). Every font the form's stream actually selects is
    put back; names the page does not define are left alone.
    """
    font_map = font_map_of(doc, page.xref)
    if not font_map:
        return []
    fixed = []
    for xref in form_xobjects(doc, page.xref):
        body = doc.xref_stream(xref) or b""
        used = {m.group(1).decode() for m in re.finditer(rb"/([A-Za-z0-9_.+-]+)\s+[\d.]+\s+Tf", body)}
        missing = {k: v for k, v in font_map.items() if k in used and k not in font_map_of(doc, xref)}
        if not missing:
            continue
        add_fonts(doc, xref, missing)
        fixed.append((xref, sorted(missing)))
    return fixed


def inner_dict(text, key):
    """(body, start, end) of the << >> dictionary stored under `key` inside an inline dictionary."""
    start = re.search(r"/" + key + r"\s*<<", text)
    if not start:
        return None
    i, depth = start.end(), 1
    while i < len(text):
        pair = text[i:i + 2]
        if pair == "<<":
            depth, i = depth + 1, i + 2
        elif pair == ">>":
            depth -= 1
            if not depth:
                return text[start.end():i], start.start(), i + 2
            i += 2
        else:
            i += 1
    return None


def subdicts(doc, holder):
    """(name, xref) pairs of the direct-referenced sub-dictionaries of `holder` (object or inline << >>)."""
    if isinstance(holder, int):
        return [(n, int(v[1].split()[0])) for n in doc.xref_get_keys(holder)
                if (v := doc.xref_get_key(holder, n))[0] == "xref"]
    return [(m.group(1), int(m.group(2)))
            for m in re.finditer(r"/([A-Za-z0-9_.+-]+)\s+(\d+)\s+0\s+R", holder)]


def resource_part(doc, xref, key):
    """The sub-dictionary `key` of an object's /Resources, as (kind, xref_or_text)."""
    res = doc.xref_get_key(xref, "Resources")
    if res[0] == "xref":
        found = doc.xref_get_key(int(res[1].split()[0]), key)
    elif res[0] == "dict":
        found = inner_dict(res[1], key)
        if found is None:
            return ("null", None)
        return ("inline", found[0])
    else:
        return ("null", None)
    if found[0] == "xref":
        return ("xref", int(found[1].split()[0]))
    if found[0] == "dict":
        return ("inline", found[1])
    return ("null", None)


def font_map_of(doc, xref):
    """/Font name -> font object xref for a page or Form XObject, inline resources included."""
    kind, part = resource_part(doc, xref, "Font")
    if kind == "xref":
        return dict(subdicts(doc, part))
    if kind == "inline":
        return dict(subdicts(doc, f"<<{part}>>"))
    return {}


def add_fonts(doc, xref, missing):
    """Add /Font entries to an object's /Resources, creating the nested dictionaries when absent."""
    entries = "".join(f"/{n} {v} 0 R" for n, v in sorted(missing.items()))
    kind, part = resource_part(doc, xref, "Font")
    if kind == "xref":
        for n, v in missing.items():
            doc.xref_set_key(part, n, f"{v} 0 R")
        return
    body = f"<<{part}{entries}>>" if kind == "inline" else f"<<{entries}>>"
    res = doc.xref_get_key(xref, "Resources")
    if res[0] == "xref":
        doc.xref_set_key(int(res[1].split()[0]), "Font", body)
        return
    text = res[1] if res[0] == "dict" else "<<>>"
    found = inner_dict(text, "Font")
    text = text[:found[1]] + f"/Font{body}" + text[found[2]:] if found else \
        re.sub(r"^<<", f"<< /Font{body}", text, count=1)
    doc.xref_set_key(xref, "Resources", text.strip())


def form_xobjects(doc, xref, seen=None):
    """Every Form XObject reachable from a page or another form, recursively."""
    seen = set() if seen is None else seen
    out, queue = [], [xref]
    while queue:
        current = queue.pop()
        if current in seen:
            continue
        seen.add(current)
        kind, part = resource_part(doc, current, "XObject")
        if kind == "null":
            continue
        for _, child in subdicts(doc, part if kind == "xref" else f"<<{part}>>"):
            if doc.xref_get_key(child, "Subtype")[1] == "/Form":
                out.append(child)
                queue.append(child)
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--replace", nargs="+", action="append", required=True, metavar="VALUE")
    ap.add_argument("--count", type=int, default=None, help="default match count for pairs without a third argument")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    if pathlib.Path(a.src).resolve() == pathlib.Path(a.dst).resolve():
        sys.exit("refusing to overwrite the original")
    if a.count is not None and a.count < 1:
        sys.exit("ERROR --count must be positive")
    pairs = []
    for item in a.replace:
        if len(item) == 2:
            old, new, count = item[0], item[1], a.count if a.count is not None else 1
        elif len(item) == 3:
            old, new, count = item
            if not count.isdigit() or int(count) < 1:
                sys.exit(f"ERROR count must be a positive integer: {count!r}")
            count = int(count)
        else:
            sys.exit("ERROR --replace takes OLD NEW [COUNT]")
        if not old.strip():
            sys.exit("ERROR search text must be nonempty")
        pairs.append((old, new, count))
    if any(any(c in text for c in "\n\r\t") for pair in pairs for text in pair[:2]):
        sys.exit("ERROR text replacement is single-line; tabs and line breaks are unsupported")
    doc = fitz.open(a.src)
    before_words = word_counts(doc)  # snapshot: redaction and font insertion change the document
    plan, errors = [], []
    for old, new, want in pairs:
        hits = [(p, r) for p in doc for r in p.search_for(old)]
        print(f"{old!r} -> {new!r}: {len(hits)} match(es) on pages {sorted({p.number + 1 for p, _ in hits})}, expected {want}")
        if len(hits) != want:
            errors.append(f"{old!r}: {len(hits)} matches, expected {want}")
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
            font, buf, how = embedded_font(doc, page, s, new) if s else (None, None, "")
            if not font or not buf:
                errors.append(f"{old!r} p{page.number+1}: no font of family {s['font']!r} covers the new text "
                              f"(subset incomplete, system font unavailable)")
                continue
            size, n_sp = s["size"], old.count(" ")
            ink = sum(font.text_length(w, fontsize=size) for w in old.split(" "))
            sw = (rect.width - ink) / n_sp if n_sp else size * 0.25  # the original's word spacing
            w_new = sum(font.text_length(w, fontsize=size) for w in new.split(" ")) + new.count(" ") * sw
            box, note = rect, ""
            if w_new > rect.width * 1.03:
                # May only grow into free space on the left, keeping its right edge where it was.
                edge = right_column_edge(page, rect)
                room = free_space_left(page, rect) if edge is not None else 0.0
                if edge is None or w_new > rect.width + room:
                    why = "right edge is not a column edge" if edge is None else f"free room to the left: {room:.1f}pt"
                    errors.append(f"{old!r} p{page.number+1}: new text {w_new:.1f}pt wider than old {rect.width:.1f}pt "
                                  f"— would overflow ({why})")
                    continue
                box = fitz.Rect(rect.x1 - w_new, rect.y0, rect.x1, rect.y1)
                note = f" [grew {rect.width:.1f}->{w_new:.1f}pt leftwards, right edge kept at x={rect.x1:.1f}]"
            plan.append((page.number, box, s, buf, new, sw, how, note))
            print(f"  p{page.number+1} {old!r} @ {tuple(round(v,1) for v in rect)} font={how} "
                  f"width {rect.width:.1f}->{w_new:.1f}pt{note}")
    for i, (number, box, *_) in enumerate(plan):
        if any(number == other and box.intersects(b) for other, b, *_ in plan[:i]):
            errors.append("overlapping replacement targets")
    if errors:
        print("\n".join("ERROR " + e for e in errors))
        sys.exit(1)
    if a.dry_run:
        return
    links = {number: doc[number].get_links() for number, *_ in plan}
    for number, box, *_ in plan:
        doc[number].add_redact_annot(box + (1, 1, -1, -1), fill=False)  # shrink: don't eat neighbour glyphs
    for number in {n for n, *_ in plan}:
        page = doc[number]
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
        repair_form_resources(doc, page)
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
    for i, (number, box, s, buf, new, sw, how, note) in enumerate(plan):
        page = doc[number]
        name = f"edit{i}"
        page.insert_font(fontname=name, fontbuffer=buf)
        c = s["color"]
        font, x = fitz.Font(fontbuffer=buf), box.x0
        # Subset fonts often lack the space glyph (PDF producers position words instead), so do the same.
        for word in new.split(" "):
            if word:
                page.insert_text((x, s["origin"][1]), word, fontname=name, fontsize=s["size"],
                                 color=((c >> 16 & 255) / 255, (c >> 8 & 255) / 255, (c & 255) / 255))
            x += font.text_length(word, fontsize=s["size"]) + sw
        print(f"  wrote p{number+1} {new!r} in {how}{note}")
    with staged_output(a.src, a.dst) as temp:
        doc.save(temp, garbage=3, deflate=True)
        with fitz.open(temp) as chk:
            if len(chk) != len(doc):
                sys.exit("ERROR page count changed")
            for number, original_links in links.items():
                actual = {link_key(link) for link in chk[number].get_links()}
                if any(link_key(link) not in actual for link in original_links):
                    sys.exit("ERROR hyperlink preservation failed")
            for old, new, want in pairs:
                if old not in new and any(p.search_for(old) for p in chk):
                    sys.exit(f"ERROR old text still extractable: {old!r}")
                if new.strip() and sum(len(p.search_for(new)) for p in chk) < want:
                    sys.exit(f"ERROR replacement not extractable: {new!r}")
            lost = lost_text(before_words, word_counts(chk), pairs)
            if lost:
                sys.exit("ERROR text lost outside the edited boxes: " + "; ".join(lost))
    doc.close()
    print(f"wrote {a.dst}")


def word_counts(doc):
    """(page, word) -> how many times it is extractable."""
    out = {}
    for page in doc:
        for w in page.get_text("words"):
            key = (page.number + 1, w[4])
            out[key] = out.get(key, 0) + 1
    return out


def lost_text(before, after, pairs):
    """Words extractable from the source but gone from the result, ignoring the edited strings.

    Guards against collateral damage: a redaction that eats a neighbouring run, or a Form XObject
    that lost a font resource, silently removes text nobody asked to change.
    """
    edited = set()
    for old, new, _ in pairs:
        edited.update(old.split())
        edited.update(new.split())
    lost = []
    for (number, text), n in sorted(before.items()):
        if text not in edited and after.get((number, text), 0) < n:
            lost.append(f"{text!r} p{number}")
    return lost[:10]


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, RuntimeError) as exc:
        print(f"ERROR {exc}; nothing published")
        sys.exit(1)
