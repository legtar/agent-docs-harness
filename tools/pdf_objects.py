"""Transactional PDF page and object edits. Coordinates are points on unrotated pages.
Tables/images/text can only be inserted into a verified empty rectangle. Existing PDF
text does not reflow; edit the source for growing an existing table or moving complex art.
Usage: python tools/pdf_objects.py IN OUT --ops ops.json [--dry-run]
"""
import argparse
import json
import math
import pathlib
import sys
import pymupdf as fitz
from safe_output import staged_output


def rectangle(page, value):
    if not isinstance(value, list) or len(value) != 4 or any(type(x) not in (int, float) or not math.isfinite(x) for x in value):
        raise ValueError("rect must contain four finite coordinates")
    rect = fitz.Rect(value)
    if rect.is_empty or not page.rect.contains(rect):
        raise ValueError("rectangle is empty or outside the page")
    if page.rotation:
        raise ValueError("object insertion on rotated pages is unsupported")
    return rect


def require_empty(page, rect):
    objects = [fitz.Rect(b[:4]) for b in page.get_text("blocks")]
    objects += [fitz.Rect(x["bbox"]) for x in page.get_image_info()]
    objects += [x["rect"] for x in page.get_drawings()]
    objects += [w.rect for w in page.widgets() or []]
    objects += [a.rect for a in page.annots() or []]
    objects += [link["from"] for link in page.get_links()]
    if any(rect.intersects(r) for r in objects):
        raise ValueError("target overlaps existing content; edit the source or choose empty space")


def apply_ops(doc, ops, base):
    if not isinstance(ops, list) or not ops or not all(isinstance(o, dict) for o in ops):
        raise ValueError("ops must be a nonempty list of objects")
    for n, op in enumerate(ops, 1):
        try:
            kind = op["op"]
            index = op.get("page", 1)
            if type(index) is not int or not 1 <= index <= len(doc):
                raise ValueError("page is a 1-based index into the current document")
            page = doc[index - 1]
            if kind == "rotate_page":
                angle = op["angle"]
                if type(angle) is not int or angle not in (0, 90, 180, 270):
                    raise ValueError("angle must be 0, 90, 180 or 270")
                page.set_rotation(angle)
            elif kind == "delete_page":
                if len(doc) == 1:
                    raise ValueError("cannot delete the last page")
                doc.delete_page(index - 1)
            elif kind == "duplicate_page":
                doc.fullcopy_page(index - 1, to=-1 if index == len(doc) else index)
            elif kind in ("insert_image", "insert_text", "insert_table"):
                rect = rectangle(page, op["rect"])
                require_empty(page, rect)
                if kind == "insert_image":
                    path = (base / op["file"]).resolve()
                    page.insert_image(rect, filename=str(path), keep_proportion=True)
                elif kind == "insert_text":
                    text = op["text"]
                    if not isinstance(text, str) or not text.strip():
                        raise ValueError("text must be nonempty")
                    # HTML renderer provides Unicode shaping, embedded fonts and explicit no-shrink.
                    import html
                    size = op.get("fontsize", 11)
                    if type(size) not in (int, float) or not math.isfinite(size) or not 6 <= size <= 72:
                        raise ValueError("fontsize must be in 6..72 pt")
                    spare, scale = page.insert_htmlbox(rect, '<div>' + html.escape(text).replace('\n', '<br>') + '</div>',
                        css=f"div {{font-family: sans-serif; font-size: {size}pt;}}", scale_low=1)
                    if spare < 0 or scale != 1:
                        raise ValueError("text does not fit; no font shrinking allowed")
                else:
                    import html
                    rows = op["rows"]
                    if not isinstance(rows, list) or not rows or not isinstance(rows[0], list) or not rows[0] or any(not isinstance(r, list) or len(r) != len(rows[0]) for r in rows):
                        raise ValueError("rows must be a nonempty rectangular list")
                    markup = '<table width="100%">' + ''.join('<tr>' + ''.join(('<th>' if i == 0 else '<td>') + html.escape(str(v)) + ('</th>' if i == 0 else '</td>') for v in row) + '</tr>' for i, row in enumerate(rows)) + '</table>'
                    css = "table {width:100%;border-collapse:collapse;font-family:sans-serif;font-size:10pt;} th,td {padding:5pt;border-bottom:0.5pt solid #adb5bd;} th {background:#e9eef4;text-align:left;}"
                    spare, scale = page.insert_htmlbox(rect, markup, css=css, scale_low=1)
                    if spare < 0 or scale != 1:
                        raise ValueError("table does not fit; use more space or create another page")
            else:
                raise ValueError(f"unknown op {kind!r}")
        except (KeyError, ValueError, TypeError, RuntimeError) as exc:
            raise ValueError(f"op {n}: {exc}") from exc


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("src"); ap.add_argument("dst")
    ap.add_argument("--ops", required=True); ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    try:
        path = pathlib.Path(args.ops).resolve()
        ops = json.loads(path.read_text(encoding="utf-8-sig"))
        with fitz.open(args.src) as doc:
            if doc.is_encrypted:
                raise ValueError("encrypted PDFs are unsupported")
            apply_ops(doc, ops, path.parent)
            if args.dry_run:
                print("dry run: all operations resolved, nothing written"); return
            with staged_output(args.src, args.dst) as temp:
                doc.save(temp, garbage=3, deflate=True)
                with fitz.open(temp) as check:
                    if len(check) != len(doc):
                        raise ValueError("page count failed validation")
                    for p in check:
                        p.get_pixmap(matrix=fitz.Matrix(.2, .2))
        print(f"wrote {args.dst}")
    except (ValueError, OSError, RuntimeError) as exc:
        print(f"ERROR {exc}; nothing published")
        sys.exit(1)


if __name__ == "__main__":
    main()
