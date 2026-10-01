"""Careful edits of an existing .docx: text, blocks (paragraphs/headings/lists/tables/sections),
table cells/rows. New content is CLONED from existing blocks, so it inherits the document's own
styles, numbering and fonts. Untouched ZIP parts are copied byte-for-byte; the original is never
overwritten; any failed op aborts the whole run (nothing written).

Look first:
    python tools/docx_edit.py IN.docx --blocks          # body outline: #index kind style text
    python tools/docx_edit.py IN.docx --tables          # tables with row/cell indexes

Edit (all flags combine; executed in order: ops file, then flags):
    python tools/docx_edit.py IN OUT --ops ops.json [--track "Author"] [--dry-run]
    python tools/docx_edit.py IN OUT --replace "old" "new"            # text inside paragraphs, any part
    python tools/docx_edit.py IN OUT --set-text BLOCK "whole new paragraph text"
    python tools/docx_edit.py IN OUT --insert-after BLOCK "text" [--like BLOCK]
    python tools/docx_edit.py IN OUT --delete BLOCK | --delete-section "Heading text"
    python tools/docx_edit.py IN OUT --cell T R C "text" | --add-row T AFTER "a|b|c" | --del-row T R

BLOCK = "#12" (index from --blocks) or a text fragment that occurs in exactly ONE block.
T (table) = table index from --tables, or a text fragment inside that table.

ops.json — list of objects, applied in order (anchors resolved after previous ops):
  {"op":"replace", "find":"…", "replace":"…", "count":1}
  {"op":"set_text", "block":B, "text":"…"}
  {"op":"insert", "after"|"before":B, "text":"…" | "texts":["…",…], "like":B?, "style":"Heading 2"?}
  {"op":"insert_table", "after"|"before":B, "rows":[["h1","h2"],["a","b"]], "like":T?}
  {"op":"delete", "block":B, "to":B?}             # inclusive range
  {"op":"delete_section", "heading":B}             # heading + everything until next heading of same/higher level
  {"op":"move", "block":B, "to":B?, "after"|"before":B}
  {"op":"page_break_before", "block":B}            # e.g. a table caption QA flagged as orphaned
  {"op":"cell", "table":T, "row":R, "col":C, "text":"…"}
  {"op":"add_row", "table":T, "after":R, "values":["…",…]}
  {"op":"del_row", "table":T, "row":R}
`like` = the block to clone formatting from (default: the anchor block for paragraphs). Inserting a
list item next to a list item continues its numbering; inserting a heading: like an existing heading.
"""
import argparse
import copy
import datetime
import difflib
import json
import os
import pathlib
import re
import sys
import zipfile
import math

from safe_output import staged_output

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
q = lambda t: f"{{{W}}}{t}"  # noqa: E731
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
PARTS = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")
OPAQUE = "\x00"  # stands for non-text run content; never matches


class OpError(Exception):
    pass


# ---------- run-level text model ----------
def simple(r) -> bool:
    return all(c.tag in (q("rPr"), q("t")) for c in r)


def runs_of(p):
    """Visible runs whose own paragraph is p (skip text boxes' nested paragraphs and deletions)."""
    out = []
    for r in p.iter(q("r")):
        anc = r.getparent()
        ok = True
        while anc is not p:
            if anc.tag in (q("p"), q("del"), q("moveFrom")):
                ok = False
                break
            anc = anc.getparent()
        if ok:
            out.append(r)
    return out


def model(p):
    """Paragraph text + list of (start, end, run). Non-simple runs become OPAQUE chars."""
    text, spans = "", []
    for r in runs_of(p):
        t = "".join(x.text or "" for x in r.findall(q("t"))) if simple(r) else OPAQUE
        spans.append((len(text), len(text) + len(t), r, simple(r)))
        text += t
    return text, spans


def set_text(r, s):
    for t in r.findall(q("t")):
        r.remove(t)
    t = etree.SubElement(r, q("t"))
    t.text = s
    t.set(XML_SPACE, "preserve")


def split(r, k):
    """Split simple run r at char k; returns the right half (inserted after r)."""
    s = "".join(x.text or "" for x in r.findall(q("t")))
    right = copy.deepcopy(r)
    set_text(r, s[:k])
    set_text(right, s[k:])
    r.addnext(right)
    return right


class Editor:
    def __init__(self, author=None):
        self.author = author
        self.date = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.next_id = 900000

    def rev(self):
        self.next_id += 1
        return {q("id"): str(self.next_id), q("author"): self.author, q("date"): self.date}

    def kill_run(self, r):
        """Remove a run, or mark it deleted when tracking."""
        if not self.author:
            r.getparent().remove(r)
            return
        d = etree.Element(q("del"), self.rev())
        r.addprevious(d)
        d.append(r)
        for t in r.findall(q("t")):
            t.tag = q("delText")
        for t in r.findall(q("instrText")):
            t.tag = q("delInstrText")

    def wrap_ins(self, r):
        if self.author:
            w = etree.Element(q("ins"), self.rev())
            r.addprevious(w)
            w.append(r)

    def mark_para(self, p, kind):
        """Tracked insertion/deletion of the paragraph mark itself (pPr/rPr/w:ins|w:del, first child)."""
        pPr = p.find(q("pPr"))
        if pPr is None:
            pPr = etree.Element(q("pPr")); p.insert(0, pPr)
        rPr = pPr.find(q("rPr"))
        if rPr is None:
            rPr = etree.Element(q("rPr"))
            tail = [c for c in pPr if c.tag in (q("sectPr"), q("pPrChange"))]
            (tail[0].addprevious if tail else pPr.append)(rPr)
        rPr.insert(0, etree.Element(q(kind), self.rev()))

    def apply_ins(self, p, start, end, new):
        """apply(), plus pure insertion (start == end): new text takes the preceding char's format."""
        if start < end:
            return self.apply(p, start, end, new)
        _, spans = model(p)
        cands = [(s0, e0, r) for s0, e0, r, ok in spans if ok and s0 < start <= e0] or                 [(s0, e0, r) for s0, e0, r, ok in spans if ok and s0 == start]
        if not cands:
            raise OpError("insertion point is next to a tab/field/image; widen the find text")
        s0, e0, r = cands[0]
        if s0 == start:  # at paragraph start: format of the following run, placed before it
            nr = copy.deepcopy(r); set_text(nr, new); r.addprevious(nr)
        else:
            if start < e0:
                split(r, start - s0)
            nr = copy.deepcopy(r); set_text(nr, new); r.addnext(nr)
        self.wrap_ins(nr)

    def apply(self, p, start, end, new):
        _, spans = model(p)
        hit = []
        for s0, e0, r, ok in spans:
            if e0 <= start or s0 >= end or s0 == e0:
                continue
            if s0 < start:
                r = split(r, start - s0)
                s0 = start
            if e0 > end:
                split(r, end - s0)
            hit.append(r)
        first = hit[0]
        if not self.author:
            if new:
                set_text(first, new)
            else:
                first.getparent().remove(first)
            for r in hit[1:]:
                r.getparent().remove(r)
            return
        ins_src = copy.deepcopy(first)
        for r in hit:
            self.kill_run(r)
        if new:
            set_text(ins_src, new)
            hit[-1].getparent().addnext(ins_src)
            self.wrap_ins(ins_src)


# ---------- blocks ----------
def styles_map(zin):
    """styleId -> (name, outline level or None), following basedOn."""
    try:
        root = etree.fromstring(zin.read("word/styles.xml"))
    except KeyError:
        return {}
    raw = {}
    for s in root.findall(q("style")):
        sid = s.get(q("styleId"))
        name = s.find(q("name"))
        ol = s.find(f"{q('pPr')}/{q('outlineLvl')}")
        based = s.find(q("basedOn"))
        raw[sid] = (name.get(q("val")) if name is not None else sid,
                    int(ol.get(q("val"))) if ol is not None else None,
                    based.get(q("val")) if based is not None else None)
    out = {}
    for sid, (name, ol, based) in raw.items():
        seen = 0
        while ol is None and based in raw and seen < 10:
            ol, based, seen = raw[based][1], raw[based][2], seen + 1
        m = re.match(r"heading\s*(\d)", name, re.I)
        if ol is None and m:
            ol = int(m.group(1)) - 1
        out[sid] = (name, ol)
    return out


CONTENT = {q("p"), q("tbl"), q("sdt"), q("customXml")}


def blocks(body):
    """Content blocks of the body (bookmark/permission markers between paragraphs are skipped)."""
    return [b for b in body if b.tag in CONTENT]


def block_text(b):
    if b.tag == q("p"):
        return model(b)[0].replace(OPAQUE, "")
    if b.tag == q("tbl"):
        return " | ".join(cell_text(tc).replace("\n", " ") for tc in b.iter(q("tc")))
    return " ".join(model(p)[0].replace(OPAQUE, "") for p in b.iter(q("p")))


def p_style(p):
    s = p.find(f"{q('pPr')}/{q('pStyle')}")
    return s.get(q("val")) if s is not None else "Normal"


def level(b, smap):
    """Outline level (0 = H1) or None for body text."""
    if b.tag != q("p"):
        return None
    ol = b.find(f"{q('pPr')}/{q('outlineLvl')}")
    if ol is not None:
        return int(ol.get(q("val")))
    return smap.get(p_style(b), (None, None))[1]


def list_blocks(body, smap):
    for i, b in enumerate(blocks(body)):
        kind = {q("p"): "p", q("tbl"): "table", q("sdt"): "sdt"}.get(b.tag, b.tag.split("}")[1])
        style = smap.get(p_style(b), (p_style(b),))[0] if b.tag == q("p") else ""
        lv = level(b, smap)
        num = " list" if b.find(f"{q('pPr')}/{q('numPr')}") is not None else ""
        tag = f"H{lv + 1}" if lv is not None and lv < 9 else kind
        print(f"#{i:<4} {tag:<6} {style[:18]:<18}{num:<5} {block_text(b)[:90]!r}")


def find_block(body, sel, what="block"):
    bl = blocks(body)
    if isinstance(sel, int) or (isinstance(sel, str) and re.fullmatch(r"#\d+", sel)):
        i = int(str(sel).lstrip("#"))
        if not 0 <= i < len(bl):
            raise OpError(f"{what} {sel}: index out of range (0..{len(bl) - 1})")
        return bl[i]
    if not isinstance(sel, str) or not sel.strip():
        raise OpError(f"{what}: nonempty text selector required")
    hits = [b for b in bl if sel in block_text(b)]
    if len(hits) != 1:
        close = difflib.get_close_matches(sel, [block_text(b)[:120] for b in bl if block_text(b).strip()], n=3, cutoff=0.3)
        raise OpError(f"{what} {sel!r}: {len(hits)} blocks contain it (need exactly 1). "
                      + (f"candidates: {[block_text(b)[:60] for b in hits[:4]]}" if hits else f"closest: {close}"))
    return hits[0]


def fresh(el):
    """Strip identity from a clone: paraIds, bookmarks, comment anchors, revision marks."""
    for e in el.iter():
        for a in (f"{{{W14}}}paraId", f"{{{W14}}}textId"):
            e.attrib.pop(a, None)
    for tag in ("bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "commentReference",
                "ins", "del", "pPrChange", "rPrChange", "trPrChange", "tcPrChange"):
        for e in list(el.iter(q(tag))):
            if tag == "ins" and len(e):  # unwrap inserted content, keep it
                for c in list(e):
                    e.addprevious(c)
            e.getparent().remove(e)
    return el


def first_rpr(p):
    for r in runs_of(p):
        if r.find(q("t")) is not None:
            rp = r.find(q("rPr"))
            return copy.deepcopy(rp) if rp is not None else None
    return None


def make_para(like, text, style_id=None):
    """New paragraph: like's paragraph props (style, numbering, spacing) + like's first run format."""
    p = etree.Element(q("p"))
    pPr = like.find(q("pPr"))
    if pPr is not None:
        pPr = fresh(copy.deepcopy(pPr))
        for e in pPr.findall(q("sectPr")):  # never duplicate a section break
            pPr.remove(e)
        p.append(pPr)
    if style_id:
        if pPr is None:
            pPr = etree.SubElement(p, q("pPr"))
        st = pPr.find(q("pStyle"))
        if st is None:
            st = etree.Element(q("pStyle")); pPr.insert(0, st)
        st.set(q("val"), style_id)
    r = etree.SubElement(p, q("r"))
    rpr = first_rpr(like)
    if rpr is not None:
        r.append(fresh(rpr))
    for i, line in enumerate(text.split("\n")):
        if i:
            etree.SubElement(r, q("br"))
        t = etree.SubElement(r, q("t")); t.text = line; t.set(XML_SPACE, "preserve")
    return p


def track_new_block(b, ed):
    if not ed.author:
        return
    paras = [b] if b.tag == q("p") else list(b.iter(q("p")))
    for p in paras:
        for r in runs_of(p):
            ed.wrap_ins(r)
        if b.tag == q("p"):
            ed.mark_para(p, "ins")
    if b.tag == q("tbl"):
        for tr in b.findall(q("tr")):
            trPr(tr).append(etree.Element(q("ins"), ed.rev()))


def remove_block(b, ed):
    if not ed.author:
        b.getparent().remove(b)
        return
    paras = [b] if b.tag == q("p") else list(b.iter(q("p")))
    for p in paras:
        for r in runs_of(p):
            ed.kill_run(r)
        if b.tag == q("p"):
            ed.mark_para(p, "del")
    if b.tag == q("tbl"):
        for tr in b.findall(q("tr")):
            trPr(tr).append(etree.Element(q("del"), ed.rev()))


def place(body, new, op):
    if "after" in op:
        find_block(body, op["after"], "after").addnext(new)
    elif "before" in op:
        find_block(body, op["before"], "before").addprevious(new)
    else:
        raise OpError(f"{op['op']}: needs 'after' or 'before'")


def block_range(body, op):
    bl = blocks(body)
    a = bl.index(find_block(body, op["block"]))
    b = bl.index(find_block(body, op["to"], "to")) if op.get("to") is not None else a
    if b < a:
        raise OpError(f"range {op['block']!r}..{op['to']!r} is reversed")
    return bl[a:b + 1]


# ---------- tables ----------
def cell_text(tc):
    return "\n".join(model(p)[0].replace(OPAQUE, "") for p in tc.findall(q("p")))


def list_tables(root):
    for ti, tbl in enumerate(root.iter(q("tbl"))):
        rows = tbl.findall(q("tr"))
        print(f"table {ti}: {len(rows)} rows")
        for ri, tr in enumerate(rows[:60]):
            cells = [cell_text(tc).replace("\n", "⏎")[:28] for tc in tr.findall(q("tc"))]
            print(f"  r{ri}: " + " | ".join(cells))


def get_table(root, sel):
    tbls = list(root.iter(q("tbl")))
    if isinstance(sel, int) or (isinstance(sel, str) and sel.lstrip("-").isdigit()):
        ti = int(sel)
        if not 0 <= ti < len(tbls):
            raise OpError(f"no table {ti} (have {len(tbls)})")
        return tbls[ti]
    if not isinstance(sel, str) or not sel.strip():
        raise OpError("table: nonempty text selector required")
    hits = [t for t in tbls if sel in block_text(t)]
    if len(hits) != 1:
        raise OpError(f"table {sel!r}: {len(hits)} tables contain it (need exactly 1)")
    return hits[0]


def get_row(tbl, ri):
    rows = tbl.findall(q("tr"))
    ri = int(ri)
    if not -len(rows) <= ri < len(rows):
        raise OpError(f"no row {ri} (have {len(rows)})")
    return rows[ri]


def trPr(tr):
    t = tr.find(q("trPr"))
    if t is None:
        t = etree.Element(q("trPr"))
        ex = tr.find(q("tblPrEx"))
        if ex is not None:
            ex.addnext(t)
        else:
            tr.insert(0, t)
    return t


def set_para_text(p, text, ed: Editor, rpr=None, extra=()):
    """Replace the text of paragraph p (and wipe text of `extra` paragraphs), keeping paragraph
    props and the first text run's formatting. Images/fields-only runs are left alone."""
    rpr = rpr if rpr is not None else first_rpr(p)
    for para in (p, *extra):
        for r in runs_of(para):
            if simple(r):
                ed.kill_run(r)
            elif r.find(q("t")) is not None:
                if ed.author:
                    raise OpError("mixed text/object run cannot be rewritten with track changes")
                for t in r.findall(q("t")):
                    r.remove(t)
    new = etree.Element(q("r"))
    if rpr is not None:
        new.append(rpr)
    for i, line in enumerate(text.split("\n")):
        if i:
            etree.SubElement(new, q("br"))
        t = etree.SubElement(new, q("t")); t.text = line; t.set(XML_SPACE, "preserve")
    p.append(new)
    ed.wrap_ins(new)


def set_cell(tc, text, ed: Editor):
    """Replace cell content, keeping paragraph props and the first run's formatting."""
    paras = tc.findall(q("p"))
    rpr = next((first_rpr(p) for p in paras if first_rpr(p) is not None), None)
    set_para_text(paras[0], text, ed, rpr, paras[1:])
    if not ed.author:
        for p in paras[1:]:
            tc.remove(p)


def add_row(tbl, after, values, ed: Editor):
    """Clone row `after` (formatting, borders, merges) below itself and fill it."""
    src = get_row(tbl, after)
    tr = fresh(copy.deepcopy(src))
    for e in list(tr.iter(q("tblHeader"))):  # a cloned header row must not become a repeating header
        e.getparent().remove(e)
    for e in list(tr.iter(q("vMerge"))):
        e.getparent().remove(e)
    if any(tc.find(q("tbl")) is not None or tc.find(".//" + q("drawing")) is not None for tc in tr.findall(q("tc"))):
        raise OpError("add_row: template row contains nested tables or drawings; use a plain row")
    cells = tr.findall(q("tc"))
    if len(values) != len(cells):
        raise OpError(f"row has {len(cells)} cells, got {len(values)} values")
    for tc, v in zip(cells, values):
        set_cell(tc, v, Editor(None))
    src.addnext(tr)
    if ed.author:
        trPr(tr).append(etree.Element(q("ins"), ed.rev()))
        for tc in cells:
            for p in tc.findall(q("p")):
                for r in runs_of(p):
                    ed.wrap_ins(r)
    return tr


def del_row(tbl, ri, ed: Editor):
    tr = get_row(tbl, ri)
    if not ed.author:
        tbl.remove(tr)
        return
    trPr(tr).append(etree.Element(q("del"), ed.rev()))
    for tc in tr.findall(q("tc")):
        for p in tc.findall(q("p")):
            for r in runs_of(p):
                ed.kill_run(r)


def table_like(like, rows):
    """New table cloned from an existing one: same tblPr/grid/borders; header row from its
    first row, body rows from its second row."""
    if like.find(".//" + q("drawing")) is not None or like.find(".//" + q("pict")) is not None or any(t is not like for t in like.iter(q("tbl"))):
        raise OpError("insert_table: template contains drawings or nested tables; choose a plain table")
    rws = like.findall(q("tr"))
    ncol = len(rws[0].findall(q("tc")))
    if any(len(r) != ncol for r in rows):
        raise OpError(f"insert_table: template table has {ncol} columns, rows must have {ncol} values")
    t = fresh(copy.deepcopy(like))
    for tr in t.findall(q("tr")):
        t.remove(tr)
    head_src, body_src = rws[0], rws[1] if len(rws) > 1 else rws[0]
    for i, vals in enumerate(rows):
        tr = fresh(copy.deepcopy(head_src if i == 0 else body_src))
        if i:
            for e in list(tr.iter(q("tblHeader"))):
                e.getparent().remove(e)
        for tc in tr.findall(q("tc")):  # drop vertical merges from the template row
            for e in list(tc.iter(q("vMerge"))):
                e.getparent().remove(e)
        cells = tr.findall(q("tc"))
        if len(cells) != ncol:
            raise OpError("insert_table: template rows have merged cells; pick a simpler 'like' table")
        for tc, v in zip(cells, vals):
            set_cell(tc, str(v), Editor(None))
        t.append(tr)
    return t


def table_plain(body, rows):
    """No template table in the document: booktabs table at full text width (house style)."""
    sys.path.insert(0, str(pathlib.Path(__file__).parent))
    from docx_kit import TOKENS, is_num
    sect = body.find(q("sectPr"))
    pg, mar = sect.find(q("pgSz")), sect.find(q("pgMar"))
    width = int(pg.get(q("w"))) - int(mar.get(q("left"))) - int(mar.get(q("right")))
    ncol = len(rows[0])
    colw = [width // ncol] * ncol
    colw[-1] += width - sum(colw)
    T = TOKENS
    x = [f'<w:tbl xmlns:w="{W}"><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/><w:tblBorders>'
         f'<w:top w:val="single" w:sz="8" w:space="0" w:color="{T["ink"]}"/><w:left w:val="nil"/>'
         f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{T["ink"]}"/><w:right w:val="nil"/>'
         f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{T["rule"]}"/><w:insideV w:val="nil"/>'
         f'</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="50" w:type="dxa"/>'
         f'<w:left w:w="100" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
         f'</w:tblCellMar></w:tblPr><w:tblGrid>' + "".join(f'<w:gridCol w:w="{w}"/>' for w in colw) + "</w:tblGrid>"]
    numeric = [all(is_num(str(r[c])) for r in rows[1:]) and len(rows) > 1 for c in range(ncol)]
    n = len(rows)
    for i, vals in enumerate(rows):
        keep = "<w:keepNext/>" if i < n - 1 and (n <= 8 or i <= 3 or i >= n - 3) else ""
        x.append("<w:tr><w:trPr><w:cantSplit/>" + ("<w:tblHeader/>" if i == 0 else "") + "</w:trPr>")
        for c, v in enumerate(vals):
            bd = f'<w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="{T["ink"]}"/></w:tcBorders>' if i == 0 else ""
            jc = '<w:jc w:val="right"/>' if numeric[c] else ""
            b = "<w:b/>" if i == 0 else ""
            txt = str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            x.append(f'<w:tc><w:tcPr><w:tcW w:w="{colw[c]}" w:type="dxa"/>{bd}</w:tcPr><w:p><w:pPr>{keep}'
                     f'<w:spacing w:before="0" w:after="0"/>{jc}</w:pPr><w:r><w:rPr>{b}</w:rPr>'
                     f'<w:t xml:space="preserve">{txt}</w:t></w:r></w:p></w:tc>')
        x.append("</w:tr>")
    x.append("</w:tbl>")
    return etree.fromstring("".join(x))



WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def validate_op(op):
    required = {
        "replace": ("find", "replace"), "set_text": ("block", "text"),
        "insert": (), "insert_table": ("rows",), "delete": ("block",),
        "delete_section": ("heading",), "move": ("block",),
        "page_break_before": ("block",), "cell": ("table", "row", "col", "text"),
        "add_row": ("table", "values"), "del_row": ("table", "row"),
        "resize_image": ("image",),
    }
    kind = op.get("op", "replace")
    if kind not in required:
        raise OpError(f"unknown op {kind!r}")
    for key in required[kind]:
        if key not in op:
            raise OpError(f"missing {key!r}")
    if kind in ("insert", "insert_table", "move") and (("after" in op) == ("before" in op)):
        raise OpError("exactly one of after/before is required")
    if kind == "insert" and (("text" in op) == ("texts" in op)):
        raise OpError("exactly one of text/texts is required")
    if kind == "replace":
        if not all(isinstance(op[k], str) for k in ("find", "replace")):
            raise OpError("find and replace must be strings")
        count = op.get("count", 1)
        if type(count) is not int or count < 1:
            raise OpError("count must be a positive integer")
    for key in ("text",):
        if key in op and not isinstance(op[key], str):
            raise OpError(f"{key} must be a string")
    for key in ("texts", "values"):
        if key in op and (not isinstance(op[key], list) or not op[key] or not all(isinstance(v, str) for v in op[key])):
            raise OpError(f"{key} must be a nonempty list of strings")


def drawing_frames(root):
    return root.xpath("//wp:inline | //wp:anchor", namespaces={"wp": WP})


def list_images(root):
    for i, frame in enumerate(drawing_frames(root)):
        ext, props = frame.find(f"{{{WP}}}extent"), frame.find(f"{{{WP}}}docPr")
        if ext is not None:
            print(f"image {i}: {int(ext.get('cx')) / 360000:.2f} x {int(ext.get('cy')) / 360000:.2f} cm "
                  f"{props.get('name', '') if props is not None else ''}")


def resize_image(root, op, ed):
    if ed.author:
        raise OpError("resize_image: drawing revisions are unsupported; use an untracked copy")
    frames = drawing_frames(root)
    i = op["image"]
    if type(i) is not int or not 0 <= i < len(frames):
        raise OpError(f"image index must be in 0..{len(frames)-1}")
    frame = frames[i]
    ext = frame.find(f"{{{WP}}}extent")
    if ext is None:
        raise OpError("drawing has no extent")
    width, height = int(ext.get("cx")), int(ext.get("cy"))
    if not any(k in op for k in ("width_cm", "height_cm")):
        raise OpError("provide width_cm or height_cm")
    for key in ("width_cm", "height_cm"):
        if key in op and (type(op[key]) not in (int, float) or not math.isfinite(op[key]) or op[key] <= 0):
            raise OpError(f"{key} must be finite and positive")
    if width <= 0 or height <= 0:
        raise OpError("invalid source drawing size")
    nw = round(op.get("width_cm", width / 360000) * 360000)
    nh = round(op.get("height_cm", height / 360000) * 360000)
    if "height_cm" not in op:
        nh = round(height * nw / width)
    if "width_cm" not in op:
        nw = round(width * nh / height)
    # Use the section enclosing the drawing, including section breaks in paragraphs.
    body = root.find(q("body"))
    top = frame
    while top.getparent() is not body:
        top = top.getparent()
    following = list(body)[list(body).index(top):]
    sect = next((s for b in following for s in b.iter(q("sectPr"))), None)
    if sect is not None:
        pg, mar = sect.find(q("pgSz")), sect.find(q("pgMar"))
        if pg is not None and mar is not None:
            maxw = (int(pg.get(q("w"))) - int(mar.get(q("left"))) - int(mar.get(q("right")))) * 635
            maxh = (int(pg.get(q("h"))) - int(mar.get(q("top"))) - int(mar.get(q("bottom")))) * 635
            if nw > maxw or nh > maxh:
                raise OpError("drawing exceeds section text area")
    ext.set("cx", str(nw)); ext.set("cy", str(nh))
    for x in frame.xpath(".//a:xfrm/a:ext", namespaces={"a": A}):
        x.set("cx", str(nw)); x.set("cy", str(nh))


# ---------- op executor ----------
def run_ops(ops, trees, ed, smap, changed):
    body = trees["word/document.xml"].find(q("body"))
    if not isinstance(ops, list) or not all(isinstance(op, dict) for op in ops):
        raise OpError("ops must be a JSON list of objects")
    for n, op in enumerate(ops, 1):
        kind = op.get("op", "replace")
        tag = f"op {n} ({kind})"
        try:
            validate_op(op)
            if kind == "replace":
                do_replace(op, trees, ed, changed)
                continue
            changed.add("word/document.xml")
            if kind == "resize_image":
                resize_image(trees["word/document.xml"], op, ed)
            elif kind == "set_text":
                b = find_block(body, op["block"])
                if b.tag != q("p"):
                    raise OpError("set_text works on paragraphs; use cell/insert_table for tables")
                set_para_text(b, op["text"], ed)
            elif kind == "insert":
                like = find_block(body, op["like"], "like") if op.get("like") is not None else \
                    find_block(body, op.get("after", op.get("before")), "anchor")
                if like.tag != q("p"):
                    raise OpError("insert: 'like' must be a paragraph (use insert_table for tables)")
                style_id = None
                if op.get("style"):
                    style_id = next((sid for sid, (nm, _) in smap.items() if nm.lower() == op["style"].lower() or sid == op["style"]), None)
                    if not style_id:
                        raise OpError(f"style {op['style']!r} not in document; have: {sorted(nm for nm, _ in smap.values())[:40]}")
                texts = op.get("texts") or [op["text"]]
                new = [make_para(like, t, style_id) for t in texts]
                if "after" in op:
                    anchor = find_block(body, op["after"], "after")
                    for p in reversed(new):
                        anchor.addnext(p)
                else:
                    anchor = find_block(body, op["before"], "before")
                    for p in new:
                        anchor.addprevious(p)
                for p in new:
                    track_new_block(p, ed)
            elif kind == "insert_table":
                rows = op["rows"]
                if not isinstance(rows, list) or not rows or not isinstance(rows[0], list) or not rows[0] or any(not isinstance(r, list) or len(r) != len(rows[0]) for r in rows):
                    raise OpError("rows must be a nonempty rectangular list")
                like = get_table(trees["word/document.xml"], op["like"]) if op.get("like") is not None else None
                t = table_like(like, rows) if like is not None else table_plain(body, rows)
                place(body, t, op)
                track_new_block(t, ed)
            elif kind == "delete":
                for b in block_range(body, op):
                    remove_block(b, ed)
            elif kind == "delete_section":
                h = find_block(body, op["heading"], "heading")
                lv = level(h, smap)
                if lv is None:
                    raise OpError(f"{op['heading']!r} is not a heading (no outline level)")
                bl = blocks(body)
                i = bl.index(h)
                j = i + 1
                while j < len(bl) and not (level(bl[j], smap) is not None and level(bl[j], smap) <= lv):
                    j += 1
                for b in bl[i:j]:
                    remove_block(b, ed)
            elif kind == "move":
                seg = block_range(body, op)
                anchor = find_block(body, op.get("after", op.get("before")), "anchor")
                if anchor in seg:
                    raise OpError("move: anchor is inside the moved range")
                if ed.author and any(b.find(".//" + q("drawing")) is not None or b.find(".//" + q("pict")) is not None for b in seg):
                    raise OpError("tracked move of drawings is unsupported; use an untracked copy")
                if ed.author:  # tracked move = tracked delete + tracked insert of a copy
                    copies = [fresh(copy.deepcopy(b)) for b in seg]
                    for b in seg:
                        remove_block(b, ed)
                    seg = copies
                    for b in seg:
                        track_new_block(b, ed)
                for b in (reversed(seg) if "after" in op else seg):
                    (anchor.addnext if "after" in op else anchor.addprevious)(b)
            elif kind == "page_break_before":
                b = find_block(body, op["block"])
                if b.tag != q("p"):
                    raise OpError("page_break_before: target a paragraph (for a table: its caption)")
                pPr = b.find(q("pPr"))
                if pPr is None:
                    pPr = etree.Element(q("pPr")); b.insert(0, pPr)
                if pPr.find(q("pageBreakBefore")) is None:
                    pb = etree.Element(q("pageBreakBefore"))
                    lead = [c for c in pPr if c.tag in (q("pStyle"), q("keepNext"), q("keepLines"))]
                    (lead[-1].addnext if lead else lambda e: pPr.insert(0, e))(pb)
            elif kind == "cell":
                cells = get_row(get_table(trees["word/document.xml"], op["table"]), op["row"]).findall(q("tc"))
                c = int(op["col"])
                if not 0 <= c < len(cells):
                    raise OpError(f"no cell {c} (row has {len(cells)})")
                set_cell(cells[c], op["text"], ed)
            elif kind == "add_row":
                add_row(get_table(trees["word/document.xml"], op["table"]), op.get("after", -1), op["values"], ed)
            elif kind == "del_row":
                del_row(get_table(trees["word/document.xml"], op["table"]), op["row"], ed)
            else:
                raise OpError(f"unknown op {kind!r}")
            print(f"ok   {tag}")
        except (OpError, KeyError, ValueError, TypeError, IndexError) as e:
            raise OpError(f"{tag}: {e}") from None


def do_replace(e, trees, ed, changed):
    if not e.get("find") or "\n" in e["replace"] or "\n" in e["find"]:
        raise OpError(f"bad replace {e!r}: non-empty find, no newlines (one paragraph at a time)")
    found = []
    for name, root in trees.items():
        for p in root.iter(q("p")):
            text, _ = model(p)
            found += [(name, p, m.start(), text) for m in re.finditer(re.escape(e["find"]), text)]
    want = e.get("count", 1)
    print(f"     replace {e['find'][:50]!r} -> {e['replace'][:50]!r}: {len(found)} match(es), expected {want}")
    for name, _, s, text in found[:5]:
        print(f"       {name}: …{text[max(0, s - 40):s + len(e['find']) + 40]!r}…".replace(OPAQUE, "¤"))
    if len(found) != want:
        if not found:
            pool = [model(p)[0] for r in trees.values() for p in r.iter(q("p"))]
            close = difflib.get_close_matches(e["find"], [t for t in pool if t.strip()], n=2, cutoff=0.4)
            raise OpError(f"{e['find']!r}: 0 matches. closest paragraphs: {close}")
        raise OpError(f"{e['find']!r}: {len(found)} matches, expected {want} (set count or lengthen find)")
    # Minimal diff: only the differing middle is rewritten, so unchanged words keep their own
    # formatting (a bold word inside the phrase stays bold) and redlines stay small.
    old, new = e["find"], e["replace"]
    pre = len(os.path.commonprefix([old, new]))
    suf = len(os.path.commonprefix([old[pre:][::-1], new[pre:][::-1]]))
    for name, p, s, _ in sorted(found, key=lambda h: -h[2]):  # right-to-left keeps offsets valid
        if model(p)[0][s:s + len(old)] != old:
            raise OpError(f"overlapping matches of {old!r}")
        if pre + suf < len(old) or len(new) > pre + suf:
            ed.apply_ins(p, s + pre, s + len(old) - suf, new[pre:len(new) - suf])
        changed.add(name)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src")
    ap.add_argument("dst", nargs="?")
    ap.add_argument("--images", action="store_true", help="list drawings with indexes and dimensions")
    ap.add_argument("--blocks", action="store_true", help="list body blocks and exit")
    ap.add_argument("--tables", action="store_true", help="list tables with row/cell indexes and exit")
    ap.add_argument("--ops", "--edits", dest="ops", help="JSON list of ops (see top of file)")
    ap.add_argument("--replace", nargs=2, action="append", default=[], metavar=("OLD", "NEW"))
    ap.add_argument("--set-text", nargs=2, action="append", default=[], metavar=("BLOCK", "TEXT"))
    ap.add_argument("--insert-after", nargs=2, action="append", default=[], metavar=("BLOCK", "TEXT"))
    ap.add_argument("--like", help="clone formatting for --insert-after from this block")
    ap.add_argument("--delete", action="append", default=[], metavar="BLOCK")
    ap.add_argument("--delete-section", action="append", default=[], metavar="HEADING")
    ap.add_argument("--page-break-before", action="append", default=[], metavar="BLOCK")
    ap.add_argument("--cell", nargs=4, action="append", default=[], metavar=("T", "R", "C", "TEXT"),
                    help=r"set cell text (literal \n = line break), keeps cell formatting")
    ap.add_argument("--add-row", nargs=3, action="append", default=[], metavar=("T", "AFTER", "A|B|C"))
    ap.add_argument("--del-row", nargs=2, action="append", default=[], metavar=("T", "R"))
    ap.add_argument("--track", metavar="AUTHOR", help="record everything as tracked changes")
    ap.add_argument("--dry-run", action="store_true", help="run all ops in memory, write nothing")
    a = ap.parse_args()

    zin = zipfile.ZipFile(a.src)
    trees = {n: etree.fromstring(zin.read(n)) for n in zin.namelist() if PARTS.match(n)}
    smap = styles_map(zin)
    if a.images:
        list_images(trees["word/document.xml"])
        return
    if a.blocks or a.tables:
        (list_blocks(trees["word/document.xml"].find(q("body")), smap) if a.blocks
         else list_tables(trees["word/document.xml"]))
        return
    if not a.dst:
        sys.exit("dst required")
    src, dst = pathlib.Path(a.src).resolve(), pathlib.Path(a.dst).resolve()
    if src == dst:
        sys.exit("refusing to overwrite the original; write to a new file")

    ops = []
    if a.ops:
        loaded = json.loads(pathlib.Path(a.ops).read_text(encoding="utf-8-sig"))
        if not isinstance(loaded, list) or not all(isinstance(o, dict) for o in loaded):
            sys.exit("ERROR ops must be a JSON list of objects; nothing written")
        ops.extend(loaded)
    ops += [{"op": "replace", "find": o, "replace": n} for o, n in a.replace]
    ops += [{"op": "set_text", "block": b, "text": t} for b, t in a.set_text]
    ops += [{"op": "insert", "after": b, "text": t, **({"like": a.like} if a.like else {})} for b, t in a.insert_after]
    ops += [{"op": "delete", "block": b} for b in a.delete]
    ops += [{"op": "delete_section", "heading": h} for h in a.delete_section]
    ops += [{"op": "page_break_before", "block": b} for b in a.page_break_before]
    # Row deletes bottom-up first, so given indexes refer to the original table.
    ops += [{"op": "del_row", "table": t, "row": int(r)} for t, r in sorted(a.del_row, key=lambda x: -int(x[1]))]
    ops += [{"op": "cell", "table": t, "row": int(r), "col": int(c), "text": x.replace(r"\n", "\n")} for t, r, c, x in a.cell]
    ops += [{"op": "add_row", "table": t, "after": int(r), "values": v.split("|")} for t, r, v in a.add_row]
    if not ops:
        sys.exit("no edits given")

    ed = Editor(a.track)
    ids = [int(v) for t in trees.values() for v in t.xpath("//@w:id", namespaces={"w": W}) if v.lstrip("-").isdigit()]
    ed.next_id = max(ids + [0]) + 1000
    changed = set()
    try:
        run_ops(ops, trees, ed, smap, changed)
    except OpError as e:
        print(f"ERROR {e}\nnothing written")
        sys.exit(1)
    if a.dry_run:
        print("dry run: all ops resolved, nothing written")
        return
    with staged_output(src, dst) as temp:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename in changed:
                    data = etree.tostring(trees[info.filename], xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(info, data)
        with zipfile.ZipFile(temp) as chk:
            if chk.testzip():
                raise OpError("output ZIP failed CRC validation")
            for name in changed:
                etree.fromstring(chk.read(name))
        zin.close()
    print(f"wrote {dst}" + (f" (tracked as {a.track})" if a.track else ""))


if __name__ == "__main__":
    main()
