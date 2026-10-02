"""Structural edits inside a PDF that has no editable source: add / delete table rows, rewrite a
cell, delete a text block, insert a paragraph, open or close vertical space. If the source
(.docx / .typ / .html) exists, edit that and re-render instead — this is the last resort.

Look first:
    python tools/pdf_flow.py IN.pdf --rows [PAGE]     # visual rows: page, row index, cells
Edit:
    python tools/pdf_flow.py IN.pdf OUT.pdf --ops ops.json [--dry-run]

ops.json — list, applied in order. Pages are 1-based. ROW = text that occurs in exactly one
visual row of the page, or the row index from --rows.
  {"op":"del_row", "page":1, "row":ROW}
  {"op":"add_row", "page":1, "after":ROW, "values":["…", …]}   # one value per cell of that row
  {"op":"cell",    "page":1, "row":ROW, "col":2, "text":"…"}
  {"op":"format",  "page":1, "row":ROW, "col":2?, "color":"C00000"?, "fill":"FFF2CC"?}   # text colour / background
  {"op":"delete",  "page":1, "block":"text inside one text block"}
  {"op":"insert",  "page":1, "after":"text inside one text block", "text":"new paragraph"}
  {"op":"insert_image", "page":1, "after":"text inside one text block", "file":"chart.png", "width_cm":10?}
  {"op":"open_gap",   "page":1, "y":350, "height":24}           # points from the page top
  {"op":"close_band", "page":1, "y0":350, "y1":374}             # deletes what is inside

How it works: the page content stream is rewritten so that everything between the edit and
the bottom of the body moves as one rigid unit; rules and borders that cross the edit are
stretched or shortened; a new row is typeset in the font, size, colour and alignment of the
row it is cloned from. Nothing is rasterised and nothing else is re-typeset.

Guarantees (otherwise nothing is written):
  * the page is re-rendered and compared: everything outside the edit must be pixel-identical,
    everything below it identical after the shift; no text may be lost or duplicated;
  * a PDF page does not reflow onto the next one. If the page has no room for the new row /
    paragraph the op is refused — rebuild from the source, or convert (render.py --to-docx).
Not supported: rotated pages, scans, multi-column bands (content beside the edited band).
"""
import argparse
import collections
import io
import json
import pathlib
import re
import sys
from decimal import Decimal

import pymupdf as fitz
import pikepdf

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from pdf_edit import embedded_font  # noqa: E402
from safe_output import staged_output  # noqa: E402

EPS = 0.3  # pt: rules sitting exactly on a band edge


class OpError(Exception):
    pass


# ---------------------------------------------------------------- matrices (a b c d e f)
IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def mul(m, n):
    """m applied first, then n."""
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F)


def apply(m, x, y):
    return x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]


def user_delta(ctm, dy):
    """Translation in user space that moves a point by (0, dy) in page space."""
    a, b, c, d = ctm[:4]
    det = a * d - b * c
    if abs(det) < 1e-12:
        return 0.0, 0.0
    return -c * dy / det, a * dy / det


def D(x):
    """Number as a PDF operand (no exponent notation, no float noise)."""
    s = f"{x:.5f}".rstrip("0").rstrip(".")
    return Decimal(s if s not in ("", "-", "-0") else "0")


def ins(op, *vals):
    return pikepdf.ContentStreamInstruction([D(v) if isinstance(v, float) else v for v in vals], pikepdf.Operator(op))


# ---------------------------------------------------------------- the vertical remap
class Mapper:
    """Remap of page space (PDF y, growing upwards). `low` is the bottom of the moving body:
    whatever lies below it (footer, page frame) stays where it is.
      open : y in [low, cut)       -> y - h          (a gap of height h opens below `cut`)
      close: y in [bottom, top)    -> top            (band collapses, its text is dropped)
             y in [low, bottom)    -> y + h
      kill : text whose origin lies in `kill_rect` is dropped (cell rewrite), nothing moves
    """
    def __init__(self, mode, h=0.0, cut=None, top=None, bottom=None, low=-1e9, kill_rect=None, clone=None,
                 paint_rect=None, rgb=None, shade_rect=None, shade_rgb=None):
        self.mode, self.h, self.cut, self.top, self.bottom, self.low = mode, h, cut, top, bottom, low
        # clone = (low, high, high_ext): paths lying in [low, high_ext] and reaching below `high` are
        # copied h further down — the frames, fills and bottom rule of a row, not the rule above it
        self.kill_rect, self.clone = kill_rect, clone
        self.paint_rect, self.rgb = paint_rect, rgb   # recolour text whose origin lies in paint_rect
        self.shade_rect, self.shade_rgb = shade_rect, shade_rgb   # recolour fills lying in shade_rect

    def paint(self, X, Y):
        if not self.paint_rect:
            return False
        x0, y0, x1, y1 = self.paint_rect
        return x0 <= X <= x1 and y0 <= Y <= y1

    def y(self, Y):
        if self.mode == "open":
            return Y - self.h if self.low <= Y < self.cut - EPS else Y
        if self.mode == "close":
            if Y >= self.top - EPS or Y < self.low:
                return Y
            return self.top if Y >= self.bottom - EPS else Y + self.h
        return Y

    def kill(self, X, Y):
        if self.mode == "close":
            return self.bottom + EPS <= Y < self.top - EPS
        if self.kill_rect:
            x0, y0, x1, y1 = self.kill_rect
            return x0 <= X <= x1 and y0 <= Y <= y1
        return False


PAINT = {"S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n"}
FILL = {"g", "rg", "k", "cs", "sc", "scn"}
PATH = {"m": 1, "l": 1, "c": 3, "v": 2, "y": 2, "h": 0, "re": 0}


class Rewriter:
    """Walks a content stream with the graphics state it needs (CTM, text line matrix) and
    re-emits it under a Mapper. Text lines and images move rigidly by their origin; path points
    move one by one, so rules crossing the edit stretch instead of breaking."""

    def __init__(self, pdf, mapper):
        self.pdf, self.mp = pdf, mapper
        self.moved = self.killed = self.cloned = self.painted = self.shaded = 0
        self.n = 0
        self.fill = []   # instructions that set the current non-stroking colour

    def run(self, instructions, ctm, resources):
        out = []
        stack, leading = [], 0.0
        path, clip = [], None
        i, count = 0, len(instructions)
        while i < count:
            it = instructions[i]
            i += 1
            if isinstance(it, pikepdf.ContentStreamInlineImage):
                out += self.placed(it, ctm, [(0, 0), (1, 0), (0, 1), (1, 1)], None, resources)
                continue
            op, vals = str(it.operator), it.operands
            if op == "q":
                stack.append((ctm, leading, list(self.fill)))
                out.append(it)
            elif op == "Q":
                if stack:
                    ctm, leading, self.fill = stack.pop()
                out.append(it)
            elif op in FILL:
                self.note_fill(it)
                out.append(it)
            elif op == "cm":
                ctm = mul(tuple(float(v) for v in vals), ctm)
                out.append(it)
            elif op == "TL":
                leading = float(vals[0])
                out.append(it)
            elif op == "BT":
                j = i
                while j < count and (isinstance(instructions[j], pikepdf.ContentStreamInlineImage)
                                     or str(instructions[j].operator) != "ET"):
                    j += 1
                block, leading = self.text(instructions[i:j], ctm, leading)
                out += [it] + block + ([instructions[j]] if j < count else [])
                i = j + 1
            elif op in PATH:
                path.append((op, [float(v) for v in vals]))
            elif op in ("W", "W*"):
                clip = it
            elif op in PAINT:
                out += self.path(path, clip, it, ctm)
                path, clip = [], None
            elif op == "Do":
                out += self.xobject(it, ctm, resources)
            else:
                out.append(it)
        return out

    def note_fill(self, it):
        op = str(it.operator)
        if op == "cs":
            self.fill = [it]
        elif op in ("sc", "scn"):
            self.fill = [i for i in self.fill if str(i.operator) == "cs"] + [it]
        else:
            self.fill = [it]

    # ---- paths
    def path(self, path, clip, paint, ctm):
        mp, out, pts, xs = self.mp, [], [], []
        axis = abs(ctm[1]) < 1e-9 and abs(ctm[2]) < 1e-9 and abs(ctm[3]) > 1e-12

        def moved(x, y):
            X, Y = apply(ctm, x, y)
            pts.append(Y)
            xs.append(X)
            Y2 = mp.y(Y)
            if Y2 == Y:
                return x, y
            du, dv = user_delta(ctm, Y2 - Y)
            return x + du, y + dv

        for op, v in path:
            if op == "re":
                x, y, w, h = v
                if axis:
                    (_, ya), (_, yb) = moved(x, y), moved(x, y + h)
                    out.append(ins("re", x, ya, w, yb - ya))
                else:
                    corners = [moved(x, y), moved(x + w, y), moved(x + w, y + h), moved(x, y + h)]
                    out.append(ins("m", *corners[0]))
                    out += [ins("l", *c) for c in corners[1:]]
                    out.append(ins("h"))
            elif op == "h":
                out.append(ins("h"))
            else:
                flat = []
                for k in range(0, len(v), 2):
                    flat += moved(v[k], v[k + 1])
                out.append(ins(op, *flat))
        if clip is not None:
            out.append(clip)
        out.append(paint)
        if mp.shade_rect and pts and clip is None and str(paint.operator) in ("f", "F", "f*"):
            x0, y0, x1, y1 = mp.shade_rect  # an existing background of the row / cell takes the new colour
            if (min(pts) >= y0 - 1.5 and max(pts) <= y1 + 1.5 and min(xs) >= x0 - 2 and max(xs) <= x1 + 2
                    and max(pts) - min(pts) > 3):
                out = [ins("rg", *mp.shade_rgb)] + out + (list(self.fill) or [ins("g", 0.0)])
                self.shaded += 1
        if any(abs(mp.y(Y) - Y) > 1e-9 for Y in pts):
            self.moved += 1
        # add_row: rules and fills of the source row are copied into the gap below it
        band = mp.clone
        if band and pts and clip is None and str(paint.operator) != "n":
            lo, hi, hi_ext = band
            if min(pts) >= lo - EPS and max(pts) <= hi_ext + EPS and min(pts) < hi - EPS:
                du, dv = user_delta(ctm, -mp.h)
                for op, v in path:
                    if op == "re":
                        out.append(ins("re", v[0] + du, v[1] + dv, v[2], v[3]))
                    elif op == "h":
                        out.append(ins("h"))
                    else:
                        out.append(ins(op, *[c + (du if k % 2 == 0 else dv) for k, c in enumerate(v)]))
                out.append(paint)
                self.cloned += 1
        return out

    # ---- text objects
    def text(self, block, ctm, leading):
        """Every positioning operator becomes an absolute Tm (Td/T* are relative to the previous
        line, so one moved line would otherwise drag all following ones)."""
        mp, out = self.mp, []
        tlm, changed = IDENT, False
        X, Y = apply(ctm, 0.0, 0.0)
        dead, tint = mp.kill(X, Y), mp.paint(X, Y)
        lead_in = None if mp.y(Y) == Y else self.tm(tlm, ctm, mp.y(Y) - Y)  # text shown before any positioning
        for it in block:
            if isinstance(it, pikepdf.ContentStreamInlineImage):
                out.append(it)
                continue
            op, vals = str(it.operator), it.operands
            show = None
            if op == "Td":
                tlm = mul((1, 0, 0, 1, float(vals[0]), float(vals[1])), tlm)
            elif op == "TD":
                leading = -float(vals[1])
                out.append(ins("TL", leading))
                tlm = mul((1, 0, 0, 1, float(vals[0]), float(vals[1])), tlm)
            elif op == "Tm":
                tlm = tuple(float(v) for v in vals)
            elif op == "T*":
                tlm = mul((1, 0, 0, 1, 0.0, -leading), tlm)
            elif op == "'":
                tlm = mul((1, 0, 0, 1, 0.0, -leading), tlm)
                show = vals[0]
            elif op == '"':
                out += [ins("Tw", vals[0]), ins("Tc", vals[1])]
                tlm = mul((1, 0, 0, 1, 0.0, -leading), tlm)
                show = vals[2]
            elif op == "TL":
                leading = float(vals[0])
                out.append(it)
                continue
            elif op in ("Tj", "TJ"):
                if dead:
                    self.killed += 1
                    changed = True
                    continue
                if lead_in is not None:
                    out.append(lead_in)
                    lead_in, changed = None, True
                out += self.shown(it, tint)
                changed = changed or tint
                continue
            elif op in FILL:
                self.note_fill(it)
                out.append(it)
                continue
            else:
                out.append(it)
                continue
            # a positioning operator: place the line where the mapper wants it
            lead_in = None
            X, Y = apply(ctm, tlm[4], tlm[5])
            dead, tint = mp.kill(X, Y), mp.paint(X, Y)
            dy = mp.y(Y) - Y
            if dy:
                changed = True
                self.moved += 1
            out.append(self.tm(tlm, ctm, dy))
            if show is not None:
                if dead:
                    self.killed += 1
                    changed = True
                else:
                    out += self.shown(pikepdf.ContentStreamInstruction([show], pikepdf.Operator("Tj")), tint)
                    changed = changed or tint
        return (out if changed else list(block)), leading

    def shown(self, it, tint):
        """A show operator, in the new colour when its line is being recoloured; the colour that
        was in force is put back right after, so nothing else on the page changes."""
        if not tint:
            return [it]
        self.painted += 1
        return [ins("rg", *self.mp.rgb), it] + (list(self.fill) or [ins("g", 0.0)])

    @staticmethod
    def tm(tlm, ctm, dy):
        du, dv = user_delta(ctm, dy) if dy else (0.0, 0.0)
        return ins("Tm", tlm[0], tlm[1], tlm[2], tlm[3], tlm[4] + du, tlm[5] + dv)

    # ---- images and form XObjects
    def placed(self, it, ctm, corners, form, resources):
        """An object drawn through the CTM: moves rigidly; a form crossing the edit is rewritten."""
        mp = self.mp
        pts = [apply(ctm, x, y) for x, y in corners]
        ys = [p[1] for p in pts]
        cx, cy = sum(p[0] for p in pts) / 4, sum(ys) / 4
        if mp.kill(cx, cy) and (mp.mode != "close" or (min(ys) >= mp.bottom - EPS and max(ys) <= mp.top + EPS)):
            self.killed += 1
            return []
        deltas = {round(mp.y(Y) - Y, 4) for Y in (min(ys), max(ys))}
        if len(deltas) == 1:
            dy = deltas.pop()
            if not dy:
                return [it]
            du, dv = user_delta(ctm, dy)
            self.moved += 1
            return [ins("cm", 1.0, 0.0, 0.0, 1.0, du, dv), it, ins("cm", 1.0, 0.0, 0.0, 1.0, -du, -dv)]
        if form is None:
            return [it]  # a background image spanning the edit stays put; verification decides
        return [self.rewrite_form(form, ctm, resources)]

    def xobject(self, it, ctm, resources):
        name = it.operands[0]
        xobjects = resources.get("/XObject") if resources is not None else None
        obj = xobjects.get(str(name)) if xobjects is not None else None
        if obj is None:
            return [it]
        if obj.get("/Subtype") == "/Form":
            m = tuple(float(v) for v in obj.get("/Matrix", IDENT))
            bb = [float(v) for v in obj.BBox]
            corners = [apply(m, x, y) for x, y in ((bb[0], bb[1]), (bb[2], bb[1]), (bb[0], bb[3]), (bb[2], bb[3]))]
            return self.placed(it, ctm, corners, obj, resources)
        return self.placed(it, ctm, [(0, 0), (1, 0), (0, 1), (1, 1)], None, resources)

    def rewrite_form(self, form, ctm, resources):
        m = tuple(float(v) for v in form.get("/Matrix", IDENT))
        inner = mul(m, ctm)
        res = own_resources(form.get("/Resources") if form.get("/Resources") is not None else resources)
        body = self.run(pikepdf.parse_content_stream(form), inner, res)
        new = self.pdf.make_stream(pikepdf.unparse_content_stream(body))
        for key, value in form.items():
            if key not in ("/Length", "/Filter", "/DecodeParms"):
                new[key] = value
        new["/Resources"] = res
        bb = [float(v) for v in form.BBox]
        xs, ys = [bb[0], bb[2]], [bb[1], bb[3]]
        for x in (bb[0], bb[2]):
            for y in (bb[1], bb[3]):  # the form's clip box must cover content that moved
                X, Y = apply(inner, x, y)
                du, dv = user_delta(inner, self.mp.y(Y) - Y)
                xs.append(x + du)
                ys.append(y + dv)
        new["/BBox"] = pikepdf.Array([D(min(xs)), D(min(ys)), D(max(xs)), D(max(ys))])
        if resources.get("/XObject") is None:
            resources["/XObject"] = pikepdf.Dictionary()
        self.n += 1
        name = f"/HFlow{self.n}"
        while name in resources["/XObject"]:
            self.n += 1
            name = f"/HFlow{self.n}"
        resources["/XObject"][name] = new
        return pikepdf.ContentStreamInstruction([pikepdf.Name(name)], pikepdf.Operator("Do"))


def own_resources(res):
    """Shallow private copy, so adding an XObject never touches a dictionary shared by other pages."""
    res = pikepdf.Dictionary(res) if res is not None else pikepdf.Dictionary()
    if res.get("/XObject") is not None:
        res["/XObject"] = pikepdf.Dictionary(res["/XObject"])
    return res


def page_resources(obj):
    node = obj
    while node is not None:
        if node.get("/Resources") is not None:
            return node["/Resources"]
        node = node.get("/Parent")
    return None


def rewrite_page(data, pno, mapper):
    """Apply `mapper` to page `pno` of the PDF bytes; returns (new bytes, Rewriter stats)."""
    pdf = pikepdf.open(io.BytesIO(data))
    page = pdf.pages[pno]
    res = own_resources(page_resources(page.obj))
    page.obj["/Resources"] = res
    rw = Rewriter(pdf, mapper)
    body = rw.run(pikepdf.parse_content_stream(page), IDENT, res)
    page.obj["/Contents"] = pdf.make_stream(pikepdf.unparse_content_stream(body))
    annots = page.obj.get("/Annots")
    if annots is not None:  # links and comments travel with the text they sit on
        keep = []
        for a in annots:
            r = [float(v) for v in a.get("/Rect", [0, 0, 0, 0])]
            if mapper.kill((r[0] + r[2]) / 2, (r[1] + r[3]) / 2) and mapper.mode == "close":
                continue
            a["/Rect"] = pikepdf.Array([D(r[0]), D(mapper.y(r[1])), D(r[2]), D(mapper.y(r[3]))])
            if a.get("/QuadPoints") is not None:
                qp = [float(v) for v in a["/QuadPoints"]]
                a["/QuadPoints"] = pikepdf.Array([D(mapper.y(v)) if k % 2 else D(v) for k, v in enumerate(qp)])
            keep.append(a)
        page.obj["/Annots"] = pikepdf.Array(keep)
    buf = io.BytesIO()
    pdf.save(buf)
    return buf.getvalue(), rw


def norm_ws(text):
    """Selectors are typed with plain spaces; PDFs carry no-break spaces and soft hyphens."""
    return re.sub(r"\s+", " ", text.replace("­", "").replace(" ", " ").replace(" ", " ").replace(" ", " ")).strip()


# ---------------------------------------------------------------- reading the page (PyMuPDF space: y down)
Row = collections.namedtuple("Row", "y0 y1 baseline cells spans")   # cells: list of (x0, x1, text, spans)


def page_spans(page):
    out = []
    for b in page.get_text("dict", flags=fitz.TEXTFLAGS_DICT & ~fitz.TEXT_PRESERVE_IMAGES)["blocks"]:
        for ln in b.get("lines", []):
            if abs(ln["dir"][0] - 1) > 1e-3:
                continue  # rotated text is not a table row
            out += [s for s in ln["spans"] if s["text"].strip()]
    return out


def visual_rows(page):
    """Text lines as the eye sees them: spans sharing a baseline, split into cells at wide gaps."""
    spans = sorted(page_spans(page), key=lambda s: (round(s["origin"][1], 1), s["bbox"][0]))
    lines = []
    for s in spans:
        if lines and abs(s["origin"][1] - lines[-1][0]["origin"][1]) <= 2.0:
            lines[-1].append(s)
        else:
            lines.append([s])
    rows = []
    for line in lines:
        line.sort(key=lambda s: s["bbox"][0])
        cells = []
        for s in line:
            gap = s["bbox"][0] - cells[-1][-1]["bbox"][2] if cells else 0
            if cells and gap <= max(0.75 * s["size"], 5.0):
                cells[-1].append(s)
            else:
                cells.append([s])
        packed = []
        for group in cells:
            text = ""
            for k, s in enumerate(group):
                glue = " " if k and s["bbox"][0] - group[k - 1]["bbox"][2] > 0.15 * s["size"] and not text.endswith(" ") else ""
                text += glue + s["text"]
            packed.append((group[0]["bbox"][0], group[-1]["bbox"][2], text.strip(), group))
        rows.append(Row(min(s["bbox"][1] for s in line), max(s["bbox"][3] for s in line),
                        line[0]["origin"][1], packed, line))
    return rows


def boundaries(page):
    """Horizontal separators of the page as clusters (y_min, y_max, thin): a thin cluster is a
    stand-alone rule (a line or hairline box); the others are edges of boxes — cell frames, row
    fills. Edges closer than 1.6pt belong to one separator (a frame drawn as outer + inner box)."""
    spans = []
    for d in page.get_drawings():
        half = (d.get("width") or 0) / 2 if "s" in (d.get("type") or "") else 0
        for item in d["items"]:
            if item[0] == "l":
                a, b = item[1], item[2]
                if abs(a.y - b.y) < 0.6 and abs(a.x - b.x) > 30:
                    y = (a.y + b.y) / 2
                    spans.append((y - half, y + half, True))
                continue
            r = item[1] if item[0] == "re" else item[1].rect if item[0] == "qu" else None
            if r is None or r.width <= 30:
                continue
            if r.height < 2.5:
                spans.append((r.y0 - half, r.y1 + half, True))
            else:
                spans += [(r.y0 - half, r.y0 + half, False), (r.y1 - half, r.y1 + half, False)]
    groups = []
    for y0, y1, thin in sorted(spans):
        if groups and y0 - groups[-1][1] <= 1.6:
            groups[-1] = (groups[-1][0], max(groups[-1][1], y1), groups[-1][2] and thin)
        else:
            groups.append((y0, y1, thin))
    return groups


def find_row(page, sel):
    rows = visual_rows(page)
    if type(sel) is int:
        if not 0 <= sel < len(rows):
            raise OpError(f"no row {sel} on page {page.number + 1} (0..{len(rows) - 1}, see --rows)")
        return rows, sel
    if not isinstance(sel, str) or not sel.strip():
        raise OpError("row: text of the row or its index from --rows")
    sel = norm_ws(sel)
    hits = [i for i, r in enumerate(rows) if sel in norm_ws(" ".join(c[2] for c in r.cells))
            or sel in norm_ws(" | ".join(c[2] for c in r.cells))]  # as printed by --rows
    if len(hits) != 1:
        raise OpError(f"row {sel!r}: {len(hits)} rows on page {page.number + 1} contain it (need exactly 1)")
    return rows, hits[0]


Band = collections.namedtuple("Band", "top bottom top_ext below_rule ruled thin_bottom")


def row_band(page, rows, i):
    """Vertical extent of row i. Both edges sit just below a separator, so a row owns the rule
    under it and the rule above belongs to the previous row:
      top, bottom  the band;  top_ext  top including the separator above (for cloning frames)
      below_rule   y between the row's text and the separator under it (None if there is none):
                   cutting here carries the separator, and row borders that end on it, downwards
      ruled        the band comes from separators around the row, not from the row pitch
      thin_bottom  the separator under the row is a stand-alone rule (it may close the table)
    """
    r = rows[i]
    groups = boundaries(page)
    above = [g for g in groups if g[1] <= r.y0 + 1.5]
    below = [g for g in groups if g[0] >= r.y1 - 1.5]
    height = r.y1 - r.y0
    if above and below and below[0][1] - above[-1][1] <= 4 * height + 20:
        a, b = above[-1], below[0]
        siblings = [o for j, o in enumerate(rows) if j != i and a[1] < o.baseline < b[0]
                    and len(o.cells) == len(r.cells) >= 2]
        if not siblings:  # other lines in the band are wrapped cell text, not rows of their own
            return Band(a[1] + 0.15, b[1] + 0.15, a[0] - 0.15, (r.y1 + b[0]) / 2, True, b[2])
    nxt = rows[i + 1] if i + 1 < len(rows) else None
    prv = rows[i - 1] if i else None
    if nxt is not None and nxt.baseline - r.baseline <= 3 * height:
        pitch = nxt.baseline - r.baseline
        bottom = (r.y1 + nxt.y0) / 2
    elif prv is not None and r.baseline - prv.baseline <= 3 * height:
        pitch = r.baseline - prv.baseline
        bottom = r.y1 + (r.y0 - prv.y1) / 2
    else:
        raise OpError("cannot tell where this row ends: no rules around it and no neighbouring row")
    under = [g for g in below if g[0] <= bottom + 2]  # a rule right under an otherwise unruled row
    return Band(bottom - pitch, bottom, bottom - pitch,
                (r.y1 + under[0][0]) / 2 if under else None, False, bool(under) and under[0][2])


def content_boxes(page):
    boxes = [fitz.Rect(s["bbox"]) for s in page_spans(page)]
    boxes += [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
    boxes += [fitz.Rect(d["rect"]) for d in page.get_drawings() if d.get("rect") is not None]
    return [b for b in boxes if b.is_valid and not b.is_empty or b.width > 0 or b.height > 0]


def furniture_top(doc, page):
    """Top of the running footer: text in the bottom 15% that sits at the same height on other pages."""
    H = page.rect.height
    mine = [s for s in page_spans(page) if s["bbox"][1] > H * 0.85]
    if not mine or len(doc) < 2:
        return H
    # furniture = the same text (page numbers aside) at the same height on another page;
    # body lines also share baselines across pages, but never their text
    def key(s):
        return round(s["origin"][1]), re.sub(r"\d+", "#", s["text"].strip())
    others = set()
    for p in doc:
        if p.number != page.number and abs(p.rect.height - H) < 1:
            others.update(key(s) for s in page_spans(p) if s["bbox"][1] > H * 0.85)
    hits = [s["bbox"][1] for s in mine
            if any((y, key(s)[1]) in others for y in (key(s)[0] - 1, key(s)[0], key(s)[0] + 1))]
    return min(hits) if hits else H


def body_limits(doc, page, cut):
    """(moving_bottom, floor): the lowest body content below `cut`, and how far down the body
    may reach — above the footer if the page has one, else the deepest body of the document."""
    H = page.rect.height
    boxes = content_boxes(page)
    below = sorted((b for b in boxes if b.y0 >= cut - 1.0 and b.y1 <= H), key=lambda b: b.y0)
    footer_top = furniture_top(doc, page)
    tail = [b for b in below if b.y0 > H * 0.88 and b.y0 < footer_top]
    if tail:
        first = min(b.y0 for b in tail)
        body = [b.y1 for b in boxes if b.y1 <= first - 0.5 and not (b.y0 < cut and b.height > H * 0.5)]
        if not body or first - max(body) > 14:
            footer_top = first
    moving = [b for b in below if b.y1 < footer_top - 0.5]
    moving_bottom = max((b.y1 for b in moving), default=cut)
    if footer_top < H:
        floor = footer_top - 8
    else:  # no footer here: the body may go as deep as it does on the fullest page, or the top margin mirrored
        deepest = 0.0
        for p in doc:
            ys = [fitz.Rect(s["bbox"]).y1 for s in page_spans(p)]
            deepest = max(deepest, max(ys, default=0.0))
        tops = [b.y0 for b in boxes]
        floor = max(deepest, H - (min(tops) if tops else 36.0))
    return moving_bottom, floor


def guard_band(page, top, bottom, own):
    """Refuse when text that is not ours sits beside the band (multi-column pages). Ours = whatever
    lies within the horizontal extent of `own` (wrapped lines of the same row or block)."""
    x0, x1 = min(o.x0 for o in own) - 12, max(o.x1 for o in own) + 12
    for s in page_spans(page):
        r = fitz.Rect(s["bbox"])
        base = s["origin"][1]
        inside = top + EPS < base < bottom - EPS
        if inside and not (x0 <= r.x0 and r.x1 <= x1):
            raise OpError(f"other content shares this band ({s['text'][:30]!r} at x={r.x0:.0f}); "
                          f"multi-column bands are not supported")


def guard_cut(page, y):
    for s in page_spans(page):
        r = fitz.Rect(s["bbox"])
        if r.y0 + 0.25 * r.height < y < s["origin"][1] - 0.5:
            raise OpError(f"the cut at y={y:.1f} would slice through text {s['text'][:30]!r}")


def pdf_y(page, y):
    return (fitz.Point(0, y) * ~page.transformation_matrix).y


# ---------------------------------------------------------------- verification
def strips_equal(a, b, ya, yb, height, zoom, what):
    """Rows [ya, ya+height) of pixmap a must equal rows [yb, …) of pixmap b (page points)."""
    r0a, r0b, n = round(ya * zoom) + 2, round(yb * zoom) + 2, round(height * zoom) - 4
    if n <= 0:
        return
    stride = a.width * a.n
    if r0a < 0 or r0b < 0 or (r0a + n) * stride > len(a.samples) or (r0b + n) * stride > len(b.samples):
        n = min(n, a.height - r0a, b.height - r0b)
        if n <= 0:
            return
    sa, sb = a.samples[r0a * stride:(r0a + n) * stride], b.samples[r0b * stride:(r0b + n) * stride]
    if sa == sb:
        return
    diff = sum(1 for x, y in zip(sa, sb) if abs(x - y) > 48)
    if diff > len(sa) * 0.0004:
        raise OpError(f"verification failed: {what} changed ({diff / len(sa):.2%} of pixels). This PDF's "
                      f"structure is not supported for in-place edits; rebuild from the source")


def words(page, clip=None):
    return collections.Counter(w[4] for w in page.get_text("words", clip=clip))


def words_in(page, rect):
    """Words (as the extractor sees them) whose centre lies in rect — what an edit may remove."""
    return collections.Counter(w[4] for w in page.get_text("words")
                               if rect.x0 <= (w[0] + w[2]) / 2 <= rect.x1 and rect.y0 <= (w[1] + w[3]) / 2 <= rect.y1)


def verify(before_page, after_page, regions, removed, h):
    """regions: (y_before, y_after, height, label) strips that must match pixel for pixel."""
    zoom = max(1, round(1.5 * h)) / h if h else 1.5   # the shift is a whole number of pixels at this zoom
    m = fitz.Matrix(zoom, zoom)
    a = before_page.get_pixmap(matrix=m, colorspace=fitz.csGRAY, alpha=False)
    b = after_page.get_pixmap(matrix=m, colorspace=fitz.csGRAY, alpha=False)
    if (a.width, a.height) != (b.width, b.height):
        raise OpError("verification failed: page size changed")
    for ya, yb, height, label in regions:
        strips_equal(a, b, ya, yb, height, zoom, label)
    want = words(before_page) - removed
    got = words(after_page)
    if want != got:
        lost, extra = list((want - got).elements())[:6], list((got - want).elements())[:6]
        raise OpError(f"verification failed: text changed outside the edit (lost {lost}, extra {extra})")


# ---------------------------------------------------------------- typesetting new text like existing text
def colour(span):
    c = span["color"]
    return ((c >> 16 & 255) / 255, (c >> 8 & 255) / 255, (c & 255) / 255)


class Typesetter:
    def __init__(self, doc):
        self.doc, self.n = doc, 0

    def font(self, page, span, text):
        font, buf, how = embedded_font(self.doc, page, span, text)
        if not font or not buf:
            raise OpError(f"no font of family {span['font']!r} covers {text!r} (subset incomplete, "
                          f"system font unavailable)")
        return font, buf

    def width(self, font, text, size):
        return sum(font.text_length(w, fontsize=size) for w in text.split(" ")) + text.count(" ") * self.space(font, size)

    @staticmethod
    def space(font, size):
        return font.text_length(" ", fontsize=size) if font.has_glyph(32) else size * 0.25

    def put(self, page, x, baseline, text, span, font, buf, stretch=0.0):
        self.n += 1
        name = f"hflow{self.n}"
        page.insert_font(fontname=name, fontbuffer=buf)
        size = span["size"]
        for word in text.split(" "):  # subset fonts often lack the space glyph: position words instead
            if word:
                page.insert_text((x, baseline), word, fontname=name, fontsize=size, color=colour(span))
            x += font.text_length(word, fontsize=size) + self.space(font, size) + stretch


def column_extent(rows, i, cell):
    """How the column of `cell` is aligned, judged from cells of other widths in the same column.
    With no such evidence the value is centred on the old one (the smallest possible error)."""
    x0, x1 = cell[0], cell[1]
    votes = collections.Counter()
    for j in range(max(0, i - 8), min(len(rows), i + 9)):
        if j == i or len(rows[j].cells) < 2:
            continue
        for c in rows[j].cells:
            if c[1] <= x0 or c[0] >= x1 or abs((c[1] - c[0]) - (x1 - x0)) <= 1.2:
                continue  # another column, or the same width: says nothing about alignment
            if abs(c[1] - x1) <= 1.2:
                votes["right"] += 1
            elif abs(c[0] - x0) <= 1.2:
                votes["left"] += 1
            elif abs((c[0] + c[1]) - (x0 + x1)) <= 2.4:
                votes["center"] += 1
    return votes.most_common(1)[0][0] if votes else "center"


def place_x(align, cell, width):
    if align == "right":
        return cell[1] - width
    if align == "center":
        return (cell[0] + cell[1]) / 2 - width / 2
    return cell[0]


def column_room(rows, i, k):
    """(left, right) limits for a value in cell k of row i: the nearest edges of cells in the
    neighbouring columns, in this row and the rows around it. Cells that overlap ours (the same
    column, or a spanning cell) do not bound it."""
    x0, x1 = rows[i].cells[k][0], rows[i].cells[k][1]
    left, right = x0 - 200.0, x1 + 200.0
    for j in range(max(0, i - 6), min(len(rows), i + 7)):
        if len(rows[j].cells) < 2:
            continue  # a paragraph line or caption, not a table row
        for c in rows[j].cells:
            if c[1] <= x0 + 0.5:
                left = max(left, c[1] + 5)   # keep a visible gap to the neighbouring column
            elif c[0] >= x1 - 0.5:
                right = min(right, c[0] - 5)
    return left, right


# ---------------------------------------------------------------- operations
class Flow:
    def __init__(self, data):
        self.data = data
        self.doc = fitz.open(stream=data, filetype="pdf")
        if self.doc.needs_pass:
            raise OpError("PDF is encrypted")
        self.ts = Typesetter(self.doc)
        self.base = pathlib.Path.cwd()   # relative image paths resolve against the ops file

    def page(self, op):
        n = op.get("page", 1)
        if type(n) is not int or not 1 <= n <= len(self.doc):
            raise OpError(f"page must be 1..{len(self.doc)}")
        page = self.doc[n - 1]
        if page.rotation:
            raise OpError("rotated pages are not supported")
        if not page_spans(page):
            raise OpError("page has no text layer (a scan?); nothing to edit in place")
        return page

    def reload(self, data):
        self.doc.close()
        self.data = data
        self.doc = fitz.open(stream=data, filetype="pdf")
        self.ts.doc = self.doc

    def commit(self):
        """Fold PyMuPDF-side insertions back into the byte image the next op starts from."""
        self.reload(self.doc.tobytes(garbage=3, deflate=True))

    def shift(self, page, mapper, regions, removed, h):
        """Rewrite the page under `mapper`, prove nothing else changed, make it the current state.
        `page` is invalid afterwards — fetch self.doc[pno] again."""
        pno = page.number
        data, rw = rewrite_page(self.data, pno, mapper)
        after = fitz.open(stream=data, filetype="pdf")
        try:
            verify(page, after[pno], regions, removed, h)
        finally:
            after.close()
        self.reload(data)
        return rw

    # -- vertical space
    def open_gap(self, page, y, h, clone=None):
        """Everything of the body below y moves down by h. Returns nothing; raises when there is no room."""
        if h <= 0:
            raise OpError("height must be positive")
        guard_cut(page, y)
        moving_bottom, floor = body_limits(self.doc, page, y)
        if moving_bottom + h > floor + 0.5:
            raise OpError(f"no room on page {page.number + 1}: {h:.1f}pt needed, {max(0.0, floor - moving_bottom):.1f}pt free "
                          f"above the bottom margin. A PDF page does not reflow — rebuild from the source")
        H = page.rect.height
        low = moving_bottom + 1.0
        mp = Mapper("open", h=h, cut=pdf_y(page, y), low=pdf_y(page, low),
                    clone=tuple(pdf_y(page, v) for v in clone) if clone else None)
        regions = [(0, 0, y, "content above the edit"), (y, y + h, moving_bottom - y, "content below the edit"),
                   (low + h + 1, low + h + 1, H - low - h - 1, "footer")]
        return self.shift(page, mp, regions, collections.Counter(), h)

    def close_band(self, page, y0, y1):
        if y1 - y0 <= 0:
            raise OpError("band is empty")
        h = y1 - y0
        guard_cut(page, y0)
        guard_cut(page, y1)
        moving_bottom, _ = body_limits(self.doc, page, y1)
        H = page.rect.height
        low = moving_bottom + 1.0
        removed = words_in(page, fitz.Rect(0, y0, page.rect.width, y1))
        mp = Mapper("close", h=h, top=pdf_y(page, y0), bottom=pdf_y(page, y1), low=pdf_y(page, low))
        regions = [(0, 0, y0, "content above the edit"), (y1, y0, moving_bottom - y1, "content below the edit"),
                   (low + 1, low + 1, H - low - 1, "footer")]
        return self.shift(page, mp, regions, removed, h)

    # -- table rows
    def del_row(self, op):
        page = self.page(op)
        rows, i = find_row(page, op["row"])
        band = row_band(page, rows, i)
        guard_band(page, band.top, band.bottom, [fitz.Rect(s["bbox"]) for s in rows[i].spans])
        self.close_band(page, band.top, band.bottom)
        return f"row {i} removed ({band.bottom - band.top:.1f}pt closed)"

    def add_row(self, op):
        page = self.page(op)
        rows, i = find_row(page, op["after"])
        src = rows[i]
        values = op["values"]
        if not isinstance(values, list) or not all(isinstance(v, str) for v in values):
            raise OpError("values must be a list of strings")
        if len(values) != len(src.cells):
            raise OpError(f"row has {len(src.cells)} cells {[c[2] for c in src.cells]}, got {len(values)} values")
        band = row_band(page, rows, i)
        h = band.bottom - band.top
        nxt = rows[i + 1] if i + 1 < len(rows) else None
        has_next = nxt is not None and nxt.baseline - src.baseline <= 1.8 * h and len(nxt.cells) >= 2
        # Where the gap opens and what is copied into it (mupdf y: low edge, top, top incl. separator):
        #  - normally below the row's own separator, which is copied under the new row together
        #    with the row's frames and fills;
        #  - last row closed by a stand-alone rule (often heavier): the rule moves down and the
        #    separator above the row is copied between the two rows;
        #  - an unruled row with a rule right under it (end of a group): above that rule, no copy.
        cut_y, clone = band.bottom, (band.bottom, band.top, band.top_ext)
        if band.below_rule is not None and band.thin_bottom and not has_next:
            cut_y = band.below_rule
            clone = (cut_y, band.top_ext - 5.0, band.top_ext) if band.ruled else None
        elif band.below_rule is not None and not band.ruled:
            cut_y, clone = band.below_rule, None
        plan = []
        for k, (cell, value) in enumerate(zip(src.cells, values)):
            if not value:
                continue
            like = cell[3][0]
            font, buf = self.ts.font(page, like, value)
            w = self.ts.width(font, value, like["size"])
            align = column_extent(rows, i, cell)
            x = place_x(align, cell, w)
            left, right = column_room(rows, i, k)
            if x < left - 0.5 or x + w > right + 0.5:
                raise OpError(f"value {value!r} ({w:.0f}pt) does not fit column {k} ({right - left:.0f}pt free)")
            plan.append((x, like["origin"][1] + h, value, like, font, buf))
        pno = page.number
        self.open_gap(page, cut_y, h, clone=clone)
        page = self.doc[pno]
        for x, baseline, value, like, font, buf in plan:
            self.ts.put(page, x, baseline, value, like, font, buf)
        self.commit()
        return f"row added below row {i} ({h:.1f}pt opened)"

    def cell(self, op):
        page = self.page(op)
        rows, i = find_row(page, op["row"])
        cells = rows[i].cells
        k = op["col"]
        if type(k) is not int or not -len(cells) <= k < len(cells):
            raise OpError(f"col must be an index into the row's {len(cells)} cells {[c[2] for c in cells]}")
        k %= len(cells)
        cell, text = cells[k], op["text"]
        if not isinstance(text, str) or "\n" in text:
            raise OpError("text must be a single line")
        like = cell[3][0]
        plan = None
        if text:
            font, buf = self.ts.font(page, like, text)
            w = self.ts.width(font, text, like["size"])
            x = place_x(column_extent(rows, i, cell), cell, w)
            left, right = column_room(rows, i, k)
            if x < left - 0.5 or x + w > right + 0.5:
                raise OpError(f"text {text!r} ({w:.0f}pt) does not fit the column ({right - left:.0f}pt free)")
            plan = (x, like["origin"][1], text, like, font, buf)
        box = fitz.Rect(cell[0] - 1, rows[i].y0 - 1, cell[1] + 1, rows[i].y1 + 1)
        removed = words_in(page, box)
        kill = (box.x0, pdf_y(page, box.y1), box.x1, pdf_y(page, box.y0))
        H = page.rect.height
        regions = [(0, 0, rows[i].y0 - 2, "content above the cell"), (rows[i].y1 + 2, rows[i].y1 + 2, H - rows[i].y1 - 2, "content below the cell")]
        pno = page.number
        rw = self.shift(page, Mapper("kill", kill_rect=kill), regions, removed, 0)
        if not rw.killed:
            raise OpError("could not isolate the cell text in the page content")
        if plan:
            self.ts.put(self.doc[pno], *plan)
        self.commit()
        return f"cell {k} of row {i}: {cell[2]!r} -> {text!r}"

    # -- text blocks
    def block(self, page, sel, what="block"):
        if not isinstance(sel, str) or not sel.strip():
            raise OpError(f"{what}: text that occurs in exactly one text block")
        blocks = [b for b in page.get_text("dict")["blocks"] if b["type"] == 0 and b.get("lines")]
        flat = lambda b: " ".join("".join(s["text"] for s in ln["spans"]) for ln in b["lines"])  # noqa: E731
        hits = [b for b in blocks if norm_ws(sel) in norm_ws(flat(b))]
        if len(hits) != 1:
            raise OpError(f"{what} {sel!r}: {len(hits)} text blocks on page {page.number + 1} contain it (need exactly 1)")
        return blocks, hits[0]

    def delete(self, op):
        page = self.page(op)
        blocks, b = self.block(page, op["block"])
        r = fitz.Rect(b["bbox"])
        tops = sorted(c.y0 for c in content_boxes(page) if c.y0 >= r.y1 - 0.5)
        bottoms = sorted(c.y1 for c in content_boxes(page) if c.y1 <= r.y0 + 0.5)
        if tops and tops[0] - r.y1 < 3 * (r.height / max(1, len(b["lines"]))):
            y0, y1 = r.y0, tops[0]              # the block and the space after it
        elif bottoms:
            y0, y1 = bottoms[-1], r.y1          # last block: take the space before it
        else:
            y0, y1 = r.y0, r.y1
        guard_band(page, y0, y1, [r])
        self.close_band(page, y0 - 0.2, y1 - 0.2)
        return f"block removed ({y1 - y0:.1f}pt closed)"

    def insert(self, op):
        page = self.page(op)
        blocks, b = self.block(page, op["after"], "after")
        text = op["text"]
        if not isinstance(text, str) or not text.strip():
            raise OpError("text must be non-empty")
        r = fitz.Rect(b["bbox"])
        like = next(s for ln in b["lines"] for s in ln["spans"] if s["text"].strip())
        bases = [ln["spans"][0]["origin"][1] for ln in b["lines"] if ln["spans"]]
        pitch = (bases[-1] - bases[0]) / (len(bases) - 1) if len(bases) > 1 and bases[-1] > bases[0] else like["size"] * 1.25
        # paragraph spacing as the page itself uses it: the usual extra distance between body blocks
        body = sorted((c for c in blocks if abs(c["bbox"][0] - r.x0) < 2.0 and c["lines"][0]["spans"]
                       and abs(c["lines"][0]["spans"][0]["size"] - like["size"]) < 0.3), key=lambda c: c["bbox"][1])
        gaps = sorted(g for a_, b_ in zip(body, body[1:])
                      if 0 <= (g := b_["lines"][0]["spans"][0]["origin"][1] - a_["lines"][-1]["spans"][0]["origin"][1] - pitch) <= pitch)
        para_gap = gaps[len(gaps) // 2] if gaps else pitch * 0.4
        justified = len(b["lines"]) > 1 and all(abs(ln["bbox"][2] - r.x1) < 1.5 for ln in b["lines"][:-1])
        font, buf = self.ts.font(page, like, text)
        size = like["size"]
        width = max(c["bbox"][2] for c in blocks if abs(c["bbox"][0] - r.x0) < 2.0) - r.x0  # the column, not a short anchor
        lines, cur = [], ""
        for word in text.split():
            trial = (cur + " " + word).strip()
            if cur and self.ts.width(font, trial, size) > width:
                lines.append(cur)
                cur = word
            else:
                cur = trial
            if self.ts.width(font, word, size) > width:
                raise OpError(f"word {word[:30]!r} is wider than the text block")
        lines.append(cur)
        h = len(lines) * pitch + para_gap
        first = bases[-1] + pitch + para_gap
        cut = r.y1 + min(para_gap, 1.0) * 0.5
        pno = page.number
        self.open_gap(page, cut, h)
        page = self.doc[pno]
        for k, line in enumerate(lines):
            stretch = 0.0
            if justified and k < len(lines) - 1 and " " in line:  # justified like its neighbour, last line ragged
                stretch = (width - self.ts.width(font, line, size)) / line.count(" ")
            self.ts.put(page, r.x0, first + k * pitch, line, like, font, buf, stretch)
        self.commit()
        return f"paragraph of {len(lines)} line(s) inserted ({h:.1f}pt opened)"

    def format(self, op):
        """Text colour and background fill of a row or one of its cells. Font, size and position
        are not touched (a PDF cannot change weight without re-typesetting: do that in the source)."""
        page = self.page(op)
        rows, i = find_row(page, op["row"])
        row = rows[i]
        cells = row.cells
        if "col" in op:
            k = op["col"]
            if type(k) is not int or not -len(cells) <= k < len(cells):
                raise OpError(f"col must be an index into the row's {len(cells)} cells")
            cells = [cells[k % len(cells)]]
        unknown = set(op) - {"op", "page", "row", "col", "color", "fill"}
        if unknown or not ({"color", "fill"} & set(op)):
            raise OpError("format takes color and/or fill (hex like 'C00000'); other properties need the source document")
        rgb = {}
        for key in ("color", "fill"):
            if key in op:
                if not isinstance(op[key], str) or not re.fullmatch(r"#?[0-9A-Fa-f]{6}", op[key]):
                    raise OpError(f"{key} must be a 6-digit hex colour like 'C00000'")
                v = op[key].lstrip("#")
                rgb[key] = tuple(int(v[j:j + 2], 16) / 255 for j in (0, 2, 4))
        pno, H = page.number, page.rect.height
        top, bottom = row.y0 - 2, row.y1 + 2
        shade = area = None
        if "fill" in rgb:
            band = row_band(page, rows, i)
            if "col" in op:
                left, right = column_room(rows, i, row.cells.index(cells[0]))
                x0, x1 = max(left - 4, cells[0][0] - 40), min(right + 4, cells[0][1] + 40)
            else:  # the row as wide as the rules around it, else as its text
                wide = [(d["rect"].x0, d["rect"].x1) for d in page.get_drawings()
                        if d["rect"].width > 30 and d["rect"].y1 >= band.top_ext - 1 and d["rect"].y0 <= band.bottom + 1]
                x0 = min([a for a, _ in wide] + [row.cells[0][0] - 4])
                x1 = max([b for _, b in wide] + [row.cells[-1][1] + 4])
            area = fitz.Rect(x0, band.top, x1, band.bottom - 0.3)
            shade = (x0, pdf_y(page, band.bottom), x1, pdf_y(page, band.top_ext))
            top, bottom = band.top_ext - 1, band.bottom + 1
        box = fitz.Rect(min(c[0] for c in cells) - 1, row.y0 - 1, max(c[1] for c in cells) + 1, row.y1 + 1)
        mp = Mapper("paint", rgb=rgb.get("color"), shade_rect=shade, shade_rgb=rgb.get("fill"),
                    paint_rect=(box.x0, pdf_y(page, box.y1), box.x1, pdf_y(page, box.y0)) if "color" in rgb else None)
        regions = [(0, 0, top, "content above the row"), (bottom, bottom, H - bottom, "content below the row")]
        rw = self.shift(page, mp, regions, collections.Counter(), 0)
        note = []
        if "color" in rgb:
            if not rw.painted:
                raise OpError("could not isolate the row text in the page content")
            want = int(rgb["color"][0] * 255) << 16 | int(rgb["color"][1] * 255) << 8 | int(rgb["color"][2] * 255)
            after = {s["color"] for r in visual_rows(self.doc[pno]) if abs(r.baseline - row.baseline) < 0.5
                     for c in r.cells if any(abs(c[0] - t[0]) < 1 for t in cells) for s in c[3]}
            if after != {want}:
                raise OpError("verification failed: the text did not take the new colour (unusual colour space or text mode)")
            note.append("colour")
        if "fill" in rgb:
            if not rw.shaded:  # the row had no background of its own: lay one under the text
                self.doc[pno].draw_rect(area, color=None, fill=rgb["fill"], overlay=False)
                self.commit()
            pix = self.doc[pno].get_pixmap(clip=area, colorspace=fitz.csRGB, alpha=False)
            want = tuple(round(v * 255) for v in rgb["fill"])
            px = pix.samples
            hit = sum(1 for k in range(0, len(px), 3) if all(abs(px[k + j] - want[j]) <= 12 for j in range(3)))
            if hit < len(px) / 3 * 0.35:
                raise OpError("verification failed: the fill is not visible (the row is covered by another opaque shape)")
            note.append("fill" if not rw.shaded else f"fill ({rw.shaded} existing background(s) recoloured)")
        return f"row {i}{' cell ' + str(op['col']) if 'col' in op else ''}: {' + '.join(note)} applied"

    def insert_image(self, op):
        """Open room under a text block and place a picture there, centred in the column."""
        page = self.page(op)
        blocks, b = self.block(page, op["after"], "after")
        path = pathlib.Path(op["file"])
        path = path if path.is_absolute() else (self.base / path).resolve()
        if not path.is_file():
            raise OpError(f"image file not found: {path}")
        try:
            pix = fitz.Pixmap(str(path))
        except Exception as e:
            raise OpError(f"cannot read image {path.name}: {e}") from None
        r = fitz.Rect(b["bbox"])
        x1 = max(c["bbox"][2] for c in blocks if abs(c["bbox"][0] - r.x0) < 2.0)   # the column
        column = x1 - r.x0
        dpi = pix.xres if pix.xres and pix.xres > 1 else 96
        width = op.get("width_cm")
        if width is not None and (type(width) not in (int, float) or width <= 0):
            raise OpError("width_cm must be a positive number")
        w = width * 72 / 2.54 if width else min(pix.width / dpi * 72, column)
        if w > column + 0.5:
            raise OpError(f"image {w / 72 * 2.54:.1f} cm is wider than the text column ({column / 72 * 2.54:.1f} cm)")
        h = w * pix.height / pix.width
        pad = 8.0
        pno = page.number
        cut = r.y1 + 0.5
        self.open_gap(page, cut, h + 2 * pad)
        x = r.x0 + (column - w) / 2
        self.doc[pno].insert_image(fitz.Rect(x, cut + pad, x + w, cut + pad + h), filename=str(path), keep_proportion=True)
        self.commit()
        return f"image {w / 72 * 2.54:.1f} x {h / 72 * 2.54:.1f} cm inserted ({h + 2 * pad:.1f}pt opened)"

    def run(self, op):
        kind = op.get("op")
        need = {"del_row": ("row",), "add_row": ("after", "values"), "cell": ("row", "col", "text"),
                "delete": ("block",), "insert": ("after", "text"), "insert_image": ("after", "file"),
                "format": ("row",),
                "open_gap": ("y", "height"),
                "close_band": ("y0", "y1")}
        if kind not in need:
            raise OpError(f"unknown op {kind!r}; have {sorted(need)}")
        for key in need[kind]:
            if key not in op:
                raise OpError(f"missing {key!r}")
        if kind in ("open_gap", "close_band"):
            page = self.page(op)
            for key in need[kind]:
                if type(op[key]) not in (int, float) or not 0 <= op[key] <= page.rect.height:
                    raise OpError(f"{key} must be a y position on the page, in points from the top")
            if kind == "open_gap":
                self.open_gap(page, float(op["y"]), float(op["height"]))
                return f"gap of {op['height']}pt opened at y={op['y']}"
            self.close_band(page, float(op["y0"]), float(op["y1"]))
            return f"band y={op['y0']}..{op['y1']} closed"
        return getattr(self, kind)(op)


def list_rows(doc, only=None):
    for page in doc:
        if only and page.number + 1 != only:
            continue
        print(f"page {page.number + 1}:")
        for i, r in enumerate(visual_rows(page)):
            print(f"  r{i:<3} y={r.y0:6.1f}  " + " | ".join(c[2] for c in r.cells)[:150])


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src")
    ap.add_argument("dst", nargs="?")
    ap.add_argument("--rows", nargs="?", const=0, type=int, metavar="PAGE", help="list visual rows and exit")
    ap.add_argument("--ops")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    data = pathlib.Path(a.src).read_bytes()
    if a.rows is not None:
        with fitz.open(stream=data, filetype="pdf") as doc:
            list_rows(doc, a.rows or None)
        return
    if not a.dst or not a.ops:
        sys.exit("ERROR give OUT.pdf and --ops ops.json (or --rows to look)")
    ops = json.loads(pathlib.Path(a.ops).read_text(encoding="utf-8-sig"))
    if not isinstance(ops, list) or not all(isinstance(o, dict) for o in ops) or not ops:
        sys.exit("ERROR ops must be a non-empty JSON list of objects; nothing written")
    flow = Flow(data)
    flow.base = pathlib.Path(a.ops).resolve().parent
    pages = len(flow.doc)
    for n, op in enumerate(ops, 1):
        try:
            print(f"ok   op {n} ({op.get('op')}): {flow.run(op)}")
        except (OpError, KeyError, TypeError, pikepdf.PdfError) as e:
            print(f"ERROR op {n} ({op.get('op')}): {e}\nnothing written")
            sys.exit(1)
    if a.dry_run:
        print("dry run: all ops verified, nothing written")
        return
    with staged_output(a.src, a.dst) as temp:
        flow.doc.save(temp, garbage=3, deflate=True)
        with fitz.open(temp) as chk:
            if len(chk) != pages:
                raise OpError("page count changed")
    print(f"wrote {a.dst}")


if __name__ == "__main__":
    try:
        main()
    except (OpError, OSError, ValueError, RuntimeError) as exc:
        print(f"ERROR {exc}; nothing published")
        sys.exit(1)
