"""Careful edits of an existing .docx: text, blocks (paragraphs/headings/lists/tables/sections),
table cells/rows. New content is CLONED from existing blocks, so it inherits the document's own
styles, numbering and fonts. Untouched ZIP parts are copied byte-for-byte; the original is never
overwritten; any failed op aborts the whole run (nothing written).

Look first:
    python tools/docx_edit.py IN.docx --blocks          # body outline: #index kind style text
    python tools/docx_edit.py IN.docx --blocks --full   # the complete text, tables included
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
  {"op":"add_col", "table":T, "after"|"before":C, "values":["…" per row]?}   # clones column C's look; table width kept
  {"op":"del_col", "table":T, "col":C}
  {"op":"col_widths", "table":T, "widths":[3,1,1]}            # relative weights, table width kept
  {"op":"merge_cells", "table":T, "row":R, "col":C, "rows":1, "cols":2}
  {"op":"format", TARGET, bold|italic|underline|strike: true/false, color:"C00000", highlight:"yellow",
        size:11, font:"Georgia", align:"left|center|right|justify", style:"Heading 2",
        space_before|space_after: pt, keep_next: bool, fill:"FFF2CC", valign:"top|center|bottom"}
        TARGET = "find":"text"(+count) | "block":B(+"to") | "table":T(+"row":R)(+"col":C)
        Only the named properties change; everything else (fonts, colours, sizes) stays.
  {"op":"insert_image", "after"|"before":B, "file":"chart.png", "width_cm":12?, "align":"center"?}
  {"op":"replace_image", "image":N, "file":"logo.png"}        # N from --images; fits the old box
  {"op":"resize_image", "image":N, "width_cm":6}
B may also address a paragraph inside a table cell: {"table":T, "row":R, "col":C, "para":0}.
Columns C are grid columns (merged cells count by the columns they span); -1 = last.
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


# resolve_entities=False: a document must not be able to pull local files into itself (XXE).
# libxml2's depth/size limits stay on: files that exceed them are decompression bombs, not reports.
PARSER = etree.XMLParser(resolve_entities=False)
MAIN = "word/document.xml"   # key of the main part in the parts dict, whatever its real name is


def parse(data):
    return etree.fromstring(data, PARSER)


class Parts(dict):
    """{part name: tree}; main_part is the real package path of the document
    (some producers write word/document2.xml)."""
    main_part = MAIN


def open_package(path):
    """(zip, Parts) with the main document under MAIN and headers/footers/notes beside it."""
    zin = zipfile.ZipFile(path)
    names = set(zin.namelist())
    main = MAIN
    if "_rels/.rels" in names:
        for r in parse(zin.read("_rels/.rels")):
            if r.get("Type", "").endswith("/officeDocument"):
                main = r.get("Target", MAIN).lstrip("/")
    if main not in names:
        zin.close()
        raise OpError("not a Word document (main document part missing)")
    trees = Parts({n: parse(zin.read(n)) for n in names if PARTS.match(n) and n != main})
    trees.main_part = main
    trees[MAIN] = parse(zin.read(main))
    if trees[MAIN].find(q("body")) is None:
        zin.close()
        raise OpError("not a Word document (no w:body; a .docm/.dotx/glossary or Strict OOXML file?)")
    return zin, trees


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



_FIELD_CACHE = None

def protected_field_runs(p):
    """Field instructions/results are generated content, including fields spanning paragraphs."""
    root = p.getroottree().getroot()
    if _FIELD_CACHE is not None and root in _FIELD_CACHE:
        return _FIELD_CACHE[root]
    protected, depth = set(), 0
    for r in root.iter(q("r")):
        ancestors = {a.tag for a in r.iterancestors()}
        if ancestors & {q("del"), q("moveFrom")}:
            continue
        markers = r.findall(q("fldChar"))
        if depth or markers or q("fldSimple") in ancestors:
            protected.add(r)
        for marker in markers:
            kind = marker.get(q("fldCharType"))
            if kind == "begin": depth += 1
            elif kind == "end": depth = max(0, depth - 1)
    if _FIELD_CACHE is not None:
        _FIELD_CACHE[root] = protected
    return protected


def revision_guard(r, ed):
    if ed.author and any(a.tag in (q("ins"), q("moveTo")) and a.get(q("id")) not in ed.created_ids for a in r.iterancestors()):
        raise OpError("tracked editing inside an existing insertion/move revision is unsupported; resolve the previous revision first")


def section_guard(seg, op):
    if not op.get("allow_section_change", False) and any(list(b.iter(q("sectPr"))) for b in seg):
        raise OpError("range contains a section boundary; use allow_section_change=true only for an intentional section-layout change")


def model(p):
    """Paragraph text + list of (start, end, run). Non-simple runs become OPAQUE chars."""
    text, spans = "", []
    protected = protected_field_runs(p)
    for r in runs_of(p):
        editable = simple(r) and r not in protected
        t = "".join(x.text or "" for x in r.findall(q("t"))) if editable else OPAQUE
        spans.append((len(text), len(text) + len(t), r, editable))
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
        self.created_ids = set()
        self.zin = None                 # source package, for ops that add parts (pictures)
        self.parts = {}                 # part name -> bytes: new or rewritten non-text parts
        self.base = pathlib.Path.cwd()  # relative file paths in ops resolve against the ops file
        self.main_part = MAIN           # real package path of the document part

    def rev(self):
        self.next_id += 1
        self.created_ids.add(str(self.next_id))
        return {q("id"): str(self.next_id), q("author"): self.author, q("date"): self.date}

    def kill_run(self, r):
        """Remove a run, or mark it deleted when tracking."""
        if not self.author:
            r.getparent().remove(r)
            return
        revision_guard(r, self)
        owned = next((a for a in r.iterancestors() if a.tag == q("ins") and a.get(q("id")) in self.created_ids), None)
        if owned is not None:
            r.getparent().remove(r)
            if not len(owned): owned.getparent().remove(owned)
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
            revision_guard(r, self)
            if any(a.tag == q("ins") and a.get(q("id")) in self.created_ids for a in r.iterancestors()):
                return
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
        revision_guard(r, self)
        if s0 == start:  # at paragraph start: format of the following run, placed before it
            nr = copy.deepcopy(r); set_text(nr, new); r.addprevious(nr)
        else:
            if start < e0:
                split(r, start - s0)
            nr = copy.deepcopy(r); set_text(nr, new); r.addnext(nr)
        self.wrap_ins(nr)

    def apply(self, p, start, end, new):
        _, spans = model(p)
        targets = [r for s, e, r, ok in spans if s != e and e > start and s < end]
        owners = [next((a for a in r.iterancestors() if a.tag == q("ins") and a.get(q("id")) in self.created_ids), None) for r in targets]
        if self.author and any(a is not None for a in owners):
            if all(a is not None for a in owners):
                return Editor().apply(p, start, end, new)
            raise OpError("replacement crosses original text and an insertion created earlier in this batch; split the edit")
        hit = []
        for s0, e0, r, ok in spans:
            if e0 <= start or s0 >= end or s0 == e0:
                continue
            revision_guard(r, self)
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
        root = parse(zin.read("word/styles.xml"))
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
    """Content blocks of the body (bookmark/permission markers between paragraphs are skipped).
    Content controls that merely wrap blocks are looked through, so their paragraphs and tables
    are addressable; generated ones (TOC, cover-page galleries) stay one opaque block."""
    out = []
    for b in body:
        inner = b.find(q("sdtContent")) if b.tag == q("sdt") else None
        if inner is not None and b.find(f"{q('sdtPr')}/{q('docPartObj')}") is None \
                and any(c.tag in CONTENT for c in inner):
            out += blocks(inner)
        elif b.tag in CONTENT:
            out.append(b)
    return out


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


def list_blocks(body, smap, full=False):
    """Outline (one line per block) or, with full=True, the whole text: what an agent reads
    before deciding on edits."""
    tables = list(body.getparent().iter(q("tbl")))
    for i, b in enumerate(blocks(body)):
        kind = {q("p"): "p", q("tbl"): "table", q("sdt"): "sdt"}.get(b.tag, b.tag.split("}")[1])
        style = smap.get(p_style(b), (p_style(b),))[0] if b.tag == q("p") else ""
        lv = level(b, smap)
        num = " list" if b.find(f"{q('pPr')}/{q('numPr')}") is not None else ""
        tag = f"H{lv + 1}" if lv is not None and lv < 9 else kind
        if not full:
            print(f"#{i:<4} {tag:<6} {style[:18]:<18}{num:<5} {block_text(b)[:90]!r}")
        elif b.tag == q("tbl"):
            print(f"#{i} table {tables.index(b)}:")
            for ri, tr in enumerate(table_rows(b)):
                print(f"    r{ri}: " + " | ".join(cell_text(tc).replace("\n", "⏎") for tc in row_cells(tr)))
        else:
            print(f"#{i} {tag}{num}: {block_text(b)}")


def find_block(body, sel, what="block"):
    if isinstance(sel, dict):  # a paragraph inside a table cell
        unknown = set(sel) - {"table", "row", "col", "para"}
        if unknown or not {"table", "row", "col"} <= set(sel):
            raise OpError(f"{what}: cell address is {{table, row, col, para?}}; got {sorted(sel)}")
        tbl = get_table(body.getparent(), sel["table"])
        tr = get_row(tbl, sel["row"])
        cells = row_cells(tr)
        c, k = index_value(sel["col"], "col"), index_value(sel.get("para", 0), "para")
        if not 0 <= c < len(cells):
            raise OpError(f"{what}: no cell {c} (row has {len(cells)})")
        paras = vertical_source(tbl, tr, cells[c]).findall(q("p"))
        if not -len(paras) <= k < len(paras):
            raise OpError(f"{what}: no paragraph {k} in that cell (has {len(paras)})")
        return paras[k]
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
        for tr in table_rows(b):
            trPr(tr).append(etree.Element(q("ins"), ed.rev()))


def remove_block(b, ed):
    inserted_mark = b.find(f"{q('pPr')}/{q('rPr')}/{q('ins')}")
    if ed.author and inserted_mark is not None and inserted_mark.get(q("id")) in ed.created_ids:
        b.getparent().remove(b)
        return
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
        for tr in table_rows(b):
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
    first = find_block(body, op["block"])
    if first not in bl:  # paragraph inside a table cell
        if op.get("to") is not None:
            raise OpError("ranges are not supported inside table cells; address one paragraph per op")
        return [first]
    a = bl.index(first)
    b = bl.index(find_block(body, op["to"], "to")) if op.get("to") is not None else a
    if b < a:
        raise OpError(f"range {op['block']!r}..{op['to']!r} is reversed")
    return bl[a:b + 1]


# ---------- tables ----------
def _children(parent, tag):
    """Direct children `tag`, looking through content controls / customXml wrappers
    (forms and templates wrap whole rows and cells in w:sdt)."""
    out = []
    for c in parent:
        if c.tag == q(tag):
            out.append(c)
        elif c.tag == q("sdt"):
            inner = c.find(q("sdtContent"))
            if inner is not None:
                out += _children(inner, tag)
        elif c.tag == q("customXml"):
            out += _children(c, tag)
    return out


def table_rows(tbl):
    return _children(tbl, "tr")


def row_cells(tr):
    return _children(tr, "tc")


def outer(el, stop):
    """el, or the wrapper (w:sdt …) that holds it directly under `stop`."""
    while el.getparent() is not None and el.getparent().tag != stop:
        el = el.getparent()
    return el


def detach(el):
    """Remove a row/cell; a content control left empty by that goes too."""
    parent = el.getparent()
    parent.remove(el)
    while parent is not None and parent.tag in (q("sdtContent"), q("customXml")) and not len(parent):
        holder = parent.getparent() if parent.tag == q("sdtContent") else parent
        parent = holder.getparent()
        parent.remove(holder)

def cell_text(tc):
    return "\n".join(model(p)[0].replace(OPAQUE, "") for p in tc.findall(q("p")))


def list_tables(root):
    for ti, tbl in enumerate(root.iter(q("tbl"))):
        rows = table_rows(tbl)
        print(f"table {ti}: {len(rows)} rows")
        for ri, tr in enumerate(rows[:60]):
            cells = [cell_text(tc).replace("\n", "⏎")[:28] for tc in row_cells(tr)]
            print(f"  r{ri}: " + " | ".join(cells))


def get_table(root, sel):
    tbls = list(root.iter(q("tbl")))
    if isinstance(sel, bool):
        raise OpError("table index cannot be boolean")
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


def index_value(value, name):
    if type(value) is int:
        return value
    if isinstance(value, str) and re.fullmatch(r"-?\d+", value):
        return int(value)
    raise OpError(f"{name} must be an integer index")


def get_row(tbl, ri):
    rows = table_rows(tbl)
    ri = index_value(ri, "row")
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
    protected = protected_field_runs(p)
    for para in (p, *extra):
        for r in runs_of(para):
            if r in protected:
                continue
            revision_guard(r, ed)
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
            if all(child.tag == q("pPr") for child in p):
                tc.remove(p)


def add_row(tbl, after, values, ed: Editor):
    """Clone row `after` (formatting, borders, merges) below itself and fill it."""
    src = get_row(tbl, after)
    rows = table_rows(tbl)
    index = rows.index(src)
    if index + 1 < len(rows) and any(tc.find(f"{q('tcPr')}/{q('vMerge')}") is not None and tc.find(f"{q('tcPr')}/{q('vMerge')}").get(q("val")) != "restart" for tc in rows[index + 1].findall(q("tc"))):
        raise OpError("add_row: insertion inside a vertical merge would break its chain; insert after the merged region")
    src_cells = row_cells(src)
    if len(values) != len(src_cells):
        raise OpError(f"row has {len(src_cells)} cells, got {len(values)} values")
    if any(list(src.iter(q(tag))) for tag in ("drawing", "pict", "object", "tbl", "footnoteReference", "endnoteReference")):
        # Objects cannot be duplicated (IDs, relationships). Build the row from the look of
        # each cell instead: same borders/shading/spans/paragraph and run format, text only.
        tr = etree.Element(q("tr"))
        for child in src:
            if child.tag in (q("tblPrEx"), q("trPr")):
                tr.append(fresh(copy.deepcopy(child)))
        for tc, v in zip(src_cells, values):
            tr.append(empty_cell_like(tc, v, keep_span=True))
    else:
        tr = fresh(copy.deepcopy(src))
        ordered = row_cells(tr)  # a copy must not reuse content-control IDs: keep the cells, drop the wrappers
        for child in list(tr):
            if child.tag in (q("tc"), q("sdt"), q("customXml")):
                tr.remove(child)
        tr.extend(ordered)
        for e in list(tr.iter(q("vMerge"))):
            e.getparent().remove(e)
        for tc, v in zip(row_cells(tr), values):
            set_cell(tc, v, Editor(None))
    for e in list(tr.iter(q("tblHeader"))):  # a cloned header row must not become a repeating header
        e.getparent().remove(e)
    cells = row_cells(tr)
    outer(src, q("tbl")).addnext(tr)
    if ed.author:
        trPr(tr).append(etree.Element(q("ins"), ed.rev()))
        for tc in cells:
            for p in tc.findall(q("p")):
                for r in runs_of(p):
                    ed.wrap_ins(r)
    return tr



def row_grid_cells(tr):
    """Physical cells with their grid start and span (horizontal merges and gridBefore)."""
    before = tr.find(f"{q('trPr')}/{q('gridBefore')}")
    start = int(before.get(q("val"))) if before is not None else 0
    cells = []
    for tc in row_cells(tr):
        span = tc.find(f"{q('tcPr')}/{q('gridSpan')}")
        width = int(span.get(q("val"))) if span is not None else 1
        cells.append((start, width, tc)); start += width
    return cells


def vertical_source(tbl, tr, tc):
    merge = tc.find(f"{q('tcPr')}/{q('vMerge')}")
    if merge is None or merge.get(q("val")) == "restart":
        return tc
    start, span, _ = next(c for c in row_grid_cells(tr) if c[2] is tc)
    rows = table_rows(tbl)
    for prev in reversed(rows[:rows.index(tr)]):
        match = next((c for c in row_grid_cells(prev) if c[:2] == (start, span)), None)
        if match is None:
            break
        candidate = match[2]
        marker = candidate.find(f"{q('tcPr')}/{q('vMerge')}")
        if marker is None:
            break
        if marker.get(q("val")) == "restart":
            return candidate
    raise OpError("vertical merge has no matching restart cell")


def del_row(tbl, ri, ed: Editor):
    tr = get_row(tbl, ri)
    if ed.author and list(tr.iter(q("vMerge"))):
        raise OpError("tracked deletion of vertically merged rows is unsupported; use an untracked copy")
    if not ed.author:
        rows = table_rows(tbl); index = rows.index(tr)
        if index + 1 < len(rows):
            following = row_grid_cells(rows[index + 1])
            transfers = []
            for start, span, tc in row_grid_cells(tr):
                marker = tc.find(f"{q('tcPr')}/{q('vMerge')}")
                if marker is not None and marker.get(q("val")) == "restart":
                    match = next((c[2] for c in following if c[:2] == (start, span)), None)
                    continuation = match.find(f"{q('tcPr')}/{q('vMerge')}") if match is not None else None
                    if continuation is not None and continuation.get(q("val")) != "restart":
                        transfers.append((tc, match, continuation))
            for tc, match, continuation in transfers:
                continuation.set(q("val"), "restart")
                for child in list(match):
                    if child.tag != q("tcPr"): match.remove(child)
                for child in list(tc):
                    if child.tag != q("tcPr"): match.append(child)
        detach(tr)
        return
    trPr(tr).append(etree.Element(q("del"), ed.rev()))
    for tc in row_cells(tr):
        for p in tc.findall(q("p")):
            for r in runs_of(p):
                ed.kill_run(r)


def table_like(like, rows):
    """New table cloned from an existing one: same tblPr/grid/borders; header row from its
    first row, body rows from its second row."""
    if any(list(like.iter(q(tag))) for tag in ("drawing", "pict", "footnoteReference", "endnoteReference")) or any(t is not like for t in like.iter(q("tbl"))):
        raise OpError("insert_table: template contains drawings or nested tables; choose a plain table")
    rws = table_rows(like)
    ncol = len(row_cells(rws[0]))
    if any(len(r) != ncol for r in rows):
        raise OpError(f"insert_table: template table has {ncol} columns, rows must have {ncol} values")
    t = fresh(copy.deepcopy(like))
    for tr in table_rows(t):
        detach(tr)
    head_src, body_src = rws[0], rws[1] if len(rws) > 1 else rws[0]
    for i, vals in enumerate(rows):
        tr = fresh(copy.deepcopy(head_src if i == 0 else body_src))
        if i:
            for e in list(tr.iter(q("tblHeader"))):
                e.getparent().remove(e)
        for tc in row_cells(tr):  # drop vertical merges from the template row
            for e in list(tc.iter(q("vMerge"))):
                e.getparent().remove(e)
        cells = row_cells(tr)
        if len(cells) != ncol:
            raise OpError("insert_table: template rows have merged cells; pick a simpler 'like' table")
        for tc, v in zip(cells, vals):
            set_cell(tc, str(v), Editor(None))
        t.append(tr)
    return t


def section_for(body, block):
    top = block
    while top.getparent() is not body:
        top = top.getparent()
    for b in list(body)[list(body).index(top):]:
        if b.tag == q("sectPr"):
            return b
        section = b.find(f"{q('pPr')}/{q('sectPr')}")
        if section is not None:
            return section
    return body.find(q("sectPr"))


def table_plain(body, rows, section=None):
    """No template table in the document: booktabs table at full text width (house style)."""
    sys.path.insert(0, str(pathlib.Path(__file__).parent))
    from docx_kit import TOKENS, is_num
    sect = section if section is not None else body.find(q("sectPr"))
    pg, mar = sect.find(q("pgSz")), sect.find(q("pgMar"))
    width = (int(pg.get(q("w"), "12240")) if pg is not None else 12240) - (int(mar.get(q("left"), "1440")) if mar is not None else 1440) - (int(mar.get(q("right"), "1440")) if mar is not None else 1440)
    cols = sect.find(q("cols"))
    if cols is not None:
        individual = [int(c.get(q("w"))) for c in cols.findall(q("col")) if c.get(q("w"))]
        number = int(cols.get(q("num"), "1"))
        if individual:
            width = min(individual)
        elif number > 1:
            width = (width - (number - 1) * int(cols.get(q("space"), "720"))) // number
    if width <= 0:
        raise OpError("section has no usable table width")
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
        "resize_image": ("image",), "replace_image": ("image", "file"), "insert_image": ("file",),
        "format": (), "add_col": ("table",), "del_col": ("table", "col"),
        "col_widths": ("table", "widths"), "merge_cells": ("table", "row", "col"),
    }
    kind = op.get("op", "replace")
    if kind not in required:
        raise OpError(f"unknown op {kind!r}")
    for key in required[kind]:
        if key not in op:
            raise OpError(f"missing {key!r}")
    if kind in ("insert", "insert_table", "move", "insert_image") and (("after" in op) == ("before" in op)):
        raise OpError("exactly one of after/before is required")
    if kind == "insert" and (("text" in op) == ("texts" in op)):
        raise OpError("exactly one of text/texts is required")
    if kind == "replace":
        if not all(isinstance(op[k], str) for k in ("find", "replace")):
            raise OpError("find and replace must be strings")
        count = op.get("count", 1)
        if type(count) is not int or count < 1:
            raise OpError("count must be a positive integer")
    if "allow_section_change" in op and type(op["allow_section_change"]) is not bool:
        raise OpError("allow_section_change must be a boolean")
    for key in ("text",):
        if key in op and not isinstance(op[key], str):
            raise OpError(f"{key} must be a string")
    if kind == "add_col" and "after" in op and "before" in op:
        raise OpError("give after or before, not both")
    for key in ("file",):
        if key in op and (not isinstance(op[key], str) or not op[key].strip()):
            raise OpError(f"{key} must be a file path")
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
    if nw < 1 or nh < 1:
        raise OpError("drawing dimensions round to zero")
    for ancestor in frame.iterancestors():
        if ancestor.tag == q("tc"):
            cellw = ancestor.find(f"{q('tcPr')}/{q('tcW')}")
            if cellw is not None and cellw.get(q("type"), "dxa") == "dxa":
                limit = int(cellw.get(q("w"))) * 635
                if nw > limit:
                    raise OpError("drawing exceeds table cell width")
    if sect is not None:
        cols = sect.find(q("cols"))
        if cols is not None:
            explicit = [int(c.get(q("w"))) * 635 for c in cols.findall(q("col")) if c.get(q("w"))]
            num = int(cols.get(q("num"), "1"))
            if explicit:
                limit = min(explicit)
            elif num > 1 and pg is not None and mar is not None:
                limit = (maxw - (num - 1) * int(cols.get(q("space"), "720")) * 635) / num
            else:
                limit = nw
            if nw > limit:
                raise OpError("drawing exceeds section column width")
    ext.set("cx", str(nw)); ext.set("cy", str(nh))
    for x in frame.xpath(".//a:xfrm/a:ext", namespaces={"a": A}):
        x.set("cx", str(nw)); x.set("cy", str(nh))


# ---------- formatting ----------
# Schema child order (wml.xsd). Out-of-order children make Word report "unreadable content".
RPR_ORDER = ("rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof "
             "snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd "
             "fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange").split()
PPR_ORDER = ("pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd "
             "tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi "
             "adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc "
             "textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange").split()
TCPR_ORDER = ("cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign "
              "hideMark headers cellIns cellDel cellMerge tcPrChange").split()
RUN_KEYS = ("bold", "italic", "underline", "strike", "color", "highlight", "size", "font")
PARA_KEYS = ("align", "style", "space_before", "space_after", "keep_next")
CELL_KEYS = ("fill", "valign")
HIGHLIGHTS = ("yellow green cyan magenta blue red darkBlue darkCyan darkGreen darkMagenta darkRed darkYellow "
              "darkGray lightGray black white none").split()
ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}


def put(parent, tag, order, attrs=None, remove=False):
    """Set, replace or remove child `tag` of a properties element at its schema position."""
    for old in parent.findall(q(tag)):
        parent.remove(old)
    if remove:
        return None
    e = etree.Element(q(tag), {q(k): str(v) for k, v in (attrs or {}).items()})
    rank = order.index(tag)
    for sib in parent:
        name = etree.QName(sib).localname
        # elements outside the list are extensions (w14:…), which the schema puts last
        if sib.tag != q(name) or name not in order or order.index(name) > rank:
            sib.addprevious(e)
            return e
    parent.append(e)
    return e


def props_of(el, tag):
    """The rPr / pPr / tcPr of a run / paragraph / cell, created as first child if missing."""
    pr = el.find(q(tag))
    if pr is None:
        pr = etree.Element(q(tag))
        el.insert(0, pr)
    return pr


def track_props(pr, ed):
    """Tracked formatting: keep the pre-change properties in w:rPrChange / pPrChange / tcPrChange."""
    change = etree.QName(pr).localname + "Change"
    if not ed.author or pr.find(q(change)) is not None:
        return
    old = etree.Element(pr.tag)
    skip = {q(t) for t in ("rPrChange", "pPrChange", "tcPrChange", "sectPr", "ins", "del", "moveFrom", "moveTo",
                           "cellIns", "cellDel", "cellMerge")}
    if change == "pPrChange":
        skip.add(q("rPr"))
    for c in pr:
        if c.tag not in skip:
            old.append(copy.deepcopy(c))
    etree.SubElement(pr, q(change), ed.rev()).append(old)


def hex_color(v, key):
    if not isinstance(v, str) or not re.fullmatch(r"#?[0-9A-Fa-f]{6}", v):
        raise OpError(f"{key} must be a 6-digit hex colour like 'C00000'")
    return v.lstrip("#").upper()


def check_format(f, smap):
    for k in ("bold", "italic", "underline", "strike", "keep_next"):
        if k in f and type(f[k]) is not bool:
            raise OpError(f"{k} must be true or false")
    for k in ("color", "fill"):
        if k in f:
            f[k] = hex_color(f[k], k)
    if "highlight" in f and f["highlight"] not in HIGHLIGHTS:
        raise OpError(f"highlight must be one of {HIGHLIGHTS}")
    for k in ("size", "space_before", "space_after"):
        if k in f and (type(f[k]) not in (int, float) or not math.isfinite(f[k]) or not 0 <= f[k] <= 400
                       or (k == "size" and f[k] < 1)):
            raise OpError(f"{k} must be a number of points")
    if "font" in f and (not isinstance(f["font"], str) or not f["font"].strip()):
        raise OpError("font must be a font family name")
    if "align" in f and f["align"] not in ALIGN:
        raise OpError(f"align must be one of {sorted(ALIGN)}")
    if "valign" in f and f["valign"] not in ("top", "center", "bottom"):
        raise OpError("valign must be top, center or bottom")
    if "style" in f:
        sid = next((sid for sid, (nm, _) in smap.items() if nm.lower() == str(f["style"]).lower() or sid == f["style"]), None)
        if not sid:
            raise OpError(f"style {f['style']!r} not in document; have: {sorted(nm for nm, _ in smap.values())[:40]}")
        f["style"] = sid


def format_run(rpr, f):
    for key, tags in (("bold", ("b", "bCs")), ("italic", ("i", "iCs")), ("strike", ("strike",))):
        if key in f:
            for t in tags:
                put(rpr, t, RPR_ORDER, None if f[key] else {"val": "0"})
    if "underline" in f:
        put(rpr, "u", RPR_ORDER, {"val": "single" if f["underline"] else "none"})
    if "color" in f:
        put(rpr, "color", RPR_ORDER, {"val": f["color"]})
    if "highlight" in f:
        put(rpr, "highlight", RPR_ORDER, {"val": f["highlight"]}, remove=f["highlight"] == "none")
    if "size" in f:
        for t in ("sz", "szCs"):
            put(rpr, t, RPR_ORDER, {"val": round(f["size"] * 2)})
    if "font" in f:
        put(rpr, "rFonts", RPR_ORDER, {a: f["font"] for a in ("ascii", "hAnsi", "cs", "eastAsia")})


def format_para(ppr, f):
    if "style" in f:
        put(ppr, "pStyle", PPR_ORDER, {"val": f["style"]})
    if "align" in f:
        put(ppr, "jc", PPR_ORDER, {"val": ALIGN[f["align"]]})
    if "keep_next" in f:
        put(ppr, "keepNext", PPR_ORDER, remove=not f["keep_next"])
    if "space_before" in f or "space_after" in f:
        sp = ppr.find(q("spacing"))
        if sp is None:
            sp = put(ppr, "spacing", PPR_ORDER)
        for key, attr in (("space_before", "before"), ("space_after", "after")):
            if key in f:
                sp.set(q(attr), str(round(f[key] * 20)))
                sp.attrib.pop(q(attr + "Autospacing"), None)


def format_cell(tcpr, f):
    if "fill" in f:
        put(tcpr, "shd", TCPR_ORDER, {"val": "clear", "color": "auto", "fill": f["fill"]})
    if "valign" in f:
        put(tcpr, "vAlign", TCPR_ORDER, {"val": f["valign"]})


def isolate(p, start, end):
    """Split runs so that chars [start, end) of paragraph p are whole runs; return those runs."""
    hit = []
    for s0, e0, r, ok in model(p)[1]:
        if e0 <= start or s0 >= end or s0 == e0:
            continue
        if s0 < start:
            r = split(r, start - s0)
            s0 = start
        if e0 > end:
            split(r, end - s0)
        hit.append(r)
    return hit


def find_matches(trees, find, want):
    found = []
    for name, root in trees.items():
        for p in root.iter(q("p")):
            text, _ = model(p)
            found += [(name, p, m.start()) for m in re.finditer(re.escape(find), text)]
    if len(found) != want:
        raise OpError(f"{find!r}: {len(found)} matches, expected {want} (set count or lengthen find)")
    return found


def do_format(op, trees, ed, smap, changed):
    """Change formatting only; text, structure and every property not named stay as they are."""
    f = {k: op[k] for k in RUN_KEYS + PARA_KEYS + CELL_KEYS if k in op}
    if not f:
        raise OpError(f"format: give at least one of {RUN_KEYS + PARA_KEYS + CELL_KEYS}")
    check_format(f, smap)
    targets = [k for k in ("find", "block", "table") if k in op]
    if len(targets) != 1:
        raise OpError("format: exactly one target — find (text), block (paragraphs) or table (+row/col)")
    root = trees["word/document.xml"]
    body = root.find(q("body"))
    runs, paras, cells = [], [], []
    if "find" in op:
        extra = [k for k in f if k not in RUN_KEYS]
        if extra:
            raise OpError(f"format: {extra} apply to paragraphs/cells; target a block or table instead of find")
        if not isinstance(op["find"], str) or not op["find"] or "\n" in op["find"]:
            raise OpError("format: find must be non-empty text within one paragraph")
        found = find_matches(trees, op["find"], op.get("count", 1))
        for name, p, s in sorted(found, key=lambda h: -h[2]):  # right-to-left keeps offsets valid
            runs += isolate(p, s, s + len(op["find"]))
            changed.add(name)
    elif "table" in op:
        tbl = get_table(root, op["table"])
        rows = table_rows(tbl) if "row" not in op else [get_row(tbl, op["row"])]
        col = None
        if "col" in op:
            n = len(grid_widths(tbl)[1])
            col = index_value(op["col"], "col")
            if not -n <= col < n:
                raise OpError(f"no column {col} (table grid has {n})")
            col %= n
        for tr in rows:
            cells += [tc for start, span, tc in row_grid_cells(tr) if col is None or start <= col < start + span]
        if not cells:
            raise OpError("format: no cells selected")
        paras = [p for tc in cells for p in tc.findall(q("p"))]
        changed.add("word/document.xml")
    else:
        for b in block_range(body, op):
            if b.tag == q("p"):
                paras.append(b)
            elif b.tag == q("tbl"):
                found_cells = list(b.iter(q("tc")))
                cells += found_cells
                paras += [p for tc in found_cells for p in tc.findall(q("p"))]
            else:
                raise OpError("format: target paragraphs or tables")
        if not cells and any(k in f for k in CELL_KEYS):
            raise OpError("format: fill/valign apply to table cells; use a table target")
        changed.add("word/document.xml")
    runs += [r for p in paras for r in runs_of(p)]
    if any(k in f for k in RUN_KEYS):
        for r in runs:
            revision_guard(r, ed)
            rpr = props_of(r, "rPr")
            track_props(rpr, ed)
            format_run(rpr, f)
        for p in paras:  # the paragraph mark carries the format of text typed there later
            mark = p.find(f"{q('pPr')}/{q('rPr')}")
            if mark is None:
                mark = put(props_of(p, "pPr"), "rPr", PPR_ORDER)
            format_run(mark, f)
    if any(k in f for k in PARA_KEYS):
        for p in paras:
            ppr = props_of(p, "pPr")
            track_props(ppr, ed)
            format_para(ppr, f)
    if any(k in f for k in CELL_KEYS):
        for tc in cells:
            tcpr = props_of(tc, "tcPr")
            track_props(tcpr, ed)
            format_cell(tcpr, f)
    print(f"     format: {len(runs)} run(s), {len(paras)} paragraph(s), {len(cells)} cell(s)")


# ---------- table columns and merges ----------
def grid_widths(tbl):
    grid = tbl.find(q("tblGrid"))
    cols = grid.findall(q("gridCol")) if grid is not None else []
    if not cols:
        raise OpError("table has no column grid (w:tblGrid); column operations are not possible")
    try:
        return grid, [int(float(c.get(q("w"), "0"))) for c in cols]
    except ValueError:
        raise OpError("table grid has non-numeric widths") from None


def refit(tbl, total, weights=None):
    """Rescale the grid to `total` width and rewrite every cell width from it, so the table
    keeps its outer width (and stays inside the text column) after columns change."""
    grid, ws = grid_widths(tbl)
    ws = weights or ws
    if total <= 0 or sum(ws) <= 0:
        return  # auto-sized table: Word/LibreOffice lay the columns out themselves
    scaled = [max(1, round(w * total / sum(ws))) for w in ws]
    scaled[-1] += total - sum(scaled)
    for g, w in zip(grid.findall(q("gridCol")), scaled):
        g.set(q("w"), str(w))
    for tr in table_rows(tbl):
        for start, span, tc in row_grid_cells(tr):
            tcw = tc.find(f"{q('tcPr')}/{q('tcW')}")
            w = sum(scaled[start:start + span])
            if tcw is None or not w:
                continue
            kind = tcw.get(q("type"), "dxa")
            if kind == "dxa":
                tcw.set(q("w"), str(w))
            elif kind == "pct":
                tcw.set(q("w"), str(round(w * 5000 / total)))


def column_index(tbl, value, name="col"):
    n = len(grid_widths(tbl)[1])
    c = index_value(value, name)
    if not -n <= c < n:
        raise OpError(f"no column {c} (table grid has {n})")
    return c % n


def set_span(tc, span):
    put(props_of(tc, "tcPr"), "gridSpan", TCPR_ORDER, {"val": span}, remove=span == 1)


def empty_cell_like(tc, text, keep_span=False):
    """A new cell with tc's look (borders, shading, margins, paragraph and run format) and no
    objects; 1x1 unless keep_span."""
    new = etree.Element(q("tc"))
    tcpr = tc.find(q("tcPr"))
    if tcpr is not None:
        tcpr = fresh(copy.deepcopy(tcpr))
        for t in ("vMerge", "hMerge") + (() if keep_span else ("gridSpan",)):
            put(tcpr, t, TCPR_ORDER, remove=True)
        new.append(tcpr)
    paras = tc.findall(q("p"))
    first = next((x for x in paras if first_rpr(x) is not None), paras[0] if paras else None)
    new.append(make_para(first, text) if first is not None else etree.fromstring(f'<w:p xmlns:w="{W}"/>'))
    return new


def add_col(tbl, op, ed):
    """Clone a grid column (cell look and text format) next to itself; table width is kept."""
    if ed.author:
        raise OpError("add_col: column changes cannot be recorded as tracked changes; use an untracked copy")
    grid, ws = grid_widths(tbl)
    before = "before" in op
    c = column_index(tbl, op["before"] if before else op.get("after", -1), "before" if before else "after")
    rows = table_rows(tbl)
    values = op.get("values") or [""] * len(rows)
    if len(values) != len(rows):
        raise OpError(f"table has {len(rows)} rows, got {len(values)} values (one per row, '' to leave empty)")
    plan = []
    for ri, (tr, v) in enumerate(zip(rows, values)):
        hit = next((x for x in row_grid_cells(tr) if x[0] <= c < x[0] + x[1]), None)
        if hit is None:
            raise OpError(f"row {ri} has no cell at grid column {c} (ragged row); fix the row first")
        start, span, tc = hit
        inside = (start < c) if before else (c < start + span - 1)
        if inside and v:
            raise OpError(f"row {ri}: the new column falls inside a merged cell; its value must be ''")
        plan.append((tc, span, inside, v))
    for tc, span, inside, v in plan:
        if inside:
            set_span(tc, span + 1)  # the merged cell simply grows over the new column
        else:
            new = empty_cell_like(tc, v)
            edge = outer(tc, q("tr"))
            (edge.addprevious if before else edge.addnext)(new)
    cols = grid.findall(q("gridCol"))
    new_col = copy.deepcopy(cols[c])
    (cols[c].addprevious if before else cols[c].addnext)(new_col)
    refit(tbl, sum(ws))


def del_col(tbl, op, ed):
    if ed.author:
        raise OpError("del_col: column changes cannot be recorded as tracked changes; use an untracked copy")
    grid, ws = grid_widths(tbl)
    if len(ws) < 2:
        raise OpError("del_col: cannot delete the only column; delete the table block instead")
    c = column_index(tbl, op["col"])
    plan = []
    for ri, tr in enumerate(table_rows(tbl)):
        hit = next((x for x in row_grid_cells(tr) if x[0] <= c < x[0] + x[1]), None)
        if hit is None:
            continue  # ragged row that never reaches this column
        if hit[1] == 1 and len(row_cells(tr)) == 1:
            raise OpError(f"row {ri} would be left without cells")
        plan.append(hit)
    for start, span, tc in plan:
        if span > 1:
            set_span(tc, span - 1)
        else:
            detach(tc)
    grid.remove(grid.findall(q("gridCol"))[c])
    refit(tbl, sum(ws))


def col_widths(tbl, op, ed):
    if ed.author:
        raise OpError("col_widths: not recordable as a tracked change; use an untracked copy")
    _, ws = grid_widths(tbl)
    w = op["widths"]
    if not isinstance(w, list) or len(w) != len(ws) or any(type(x) not in (int, float) or not math.isfinite(x) or x <= 0 for x in w):
        raise OpError(f"widths must be {len(ws)} positive numbers (relative weights, one per grid column)")
    if sum(ws) <= 0:
        raise OpError("table is auto-sized (no grid widths); nothing to redistribute")
    refit(tbl, sum(ws), w)


def merge_cells(tbl, op, ed):
    """Merge a rectangle of plain cells into one; text of the absorbed cells moves into the first."""
    if ed.author:
        raise OpError("merge_cells: not recordable as a tracked change; use an untracked copy")
    rows = table_rows(tbl)
    r0 = rows.index(get_row(tbl, op["row"]))
    c0 = column_index(tbl, op["col"])
    nr, nc = op.get("rows", 1), op.get("cols", 1)
    if type(nr) is not int or type(nc) is not int or nr < 1 or nc < 1 or nr * nc < 2:
        raise OpError("merge_cells: rows and cols are positive integers covering at least two cells")
    if r0 + nr > len(rows):
        raise OpError("merge_cells: region runs past the last row")
    region = []
    for tr in rows[r0:r0 + nr]:
        line = []
        for c in range(c0, c0 + nc):
            hit = next((x for x in row_grid_cells(tr) if x[0] == c and x[1] == 1), None)
            if hit is None or hit[2].find(f"{q('tcPr')}/{q('vMerge')}") is not None:
                raise OpError("merge_cells: region must consist of unmerged cells aligned to the grid")
            line.append(hit[2])
        region.append(line)
    top = region[0][0]

    def absorb(tc):
        for child in list(tc):
            if child.tag == q("tcPr"):
                continue
            if child.tag == q("p") and not block_text(child).strip() and not list(child.iter(q("drawing"))):
                tc.remove(child)
            else:
                top.append(child)  # keeps runs, formatting and objects

    for ri, line in enumerate(region):
        for tc in line[1:]:
            absorb(tc)
            detach(tc)
        head = line[0]
        if nc > 1:
            set_span(head, nc)
        if nr > 1:
            put(props_of(head, "tcPr"), "vMerge", TCPR_ORDER, {"val": "restart"} if ri == 0 else None)
            if ri:
                absorb(head)
                head.append(etree.fromstring(f'<w:p xmlns:w="{W}"/>'))
    if top[-1].tag != q("p"):  # a cell must end with a paragraph
        top.append(etree.fromstring(f'<w:p xmlns:w="{W}"/>'))
    _, ws = grid_widths(tbl)
    refit(tbl, sum(ws))


# ---------- pictures ----------
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
IMAGE_TYPES = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
               "bmp": "image/bmp", "tif": "image/tiff", "tiff": "image/tiff"}
TYPES = "[Content_Types].xml"


def rels_name(ed):
    folder, _, base = ed.main_part.rpartition("/")
    return f"{folder}/_rels/{base}.rels" if folder else f"_rels/{base}.rels"


def part_xml(ed, name):
    return parse(ed.parts[name] if name in ed.parts else ed.zin.read(name))


def store_xml(ed, name, root):
    ed.parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def add_media(ed, file):
    """Add an image file to the package; returns (relationship id, width px, height px, dpi)."""
    path = (ed.base / file).resolve() if not pathlib.Path(file).is_absolute() else pathlib.Path(file)
    ext = path.suffix.lower().lstrip(".")
    if ext not in IMAGE_TYPES:
        raise OpError(f"image type .{ext} unsupported; use one of {sorted(IMAGE_TYPES)}")
    if not path.is_file():
        raise OpError(f"image file not found: {path}")
    import pymupdf as fitz
    try:
        pix = fitz.Pixmap(str(path))
        w, h, dpi = pix.width, pix.height, pix.xres if pix.xres and pix.xres > 1 else 96
    except Exception as e:
        raise OpError(f"cannot read image {path.name}: {e}") from None
    if w < 1 or h < 1:
        raise OpError("image has no pixels")
    names = set(ed.zin.namelist()) | set(ed.parts)
    folder = ed.main_part.rpartition("/")[0]
    n = 1
    while f"{folder}/media/harness{n}.{ext}".lstrip("/") in names:
        n += 1
    ed.parts[f"{folder}/media/harness{n}.{ext}".lstrip("/")] = path.read_bytes()
    RELS = rels_name(ed)
    try:
        rels = part_xml(ed, RELS)
    except KeyError:
        rels = etree.Element(f"{{{PKG_REL}}}Relationships", nsmap={None: PKG_REL})
    ids = {r.get("Id") for r in rels}
    k = 1
    while f"rId{k}" in ids:
        k += 1
    etree.SubElement(rels, f"{{{PKG_REL}}}Relationship", Id=f"rId{k}", Type=R_NS + "/image",
                     Target=f"media/harness{n}.{ext}")
    store_xml(ed, RELS, rels)
    types = part_xml(ed, TYPES)
    if not any(d.get("Extension", "").lower() == ext for d in types.findall(f"{{{CT_NS}}}Default")):
        types.insert(0, etree.Element(f"{{{CT_NS}}}Default", Extension=ext, ContentType=IMAGE_TYPES[ext]))
        store_xml(ed, TYPES, types)
    return f"rId{k}", w, h, dpi


def text_area(body, block):
    """(width, height) in EMU available to `block`: its table cell, else its section column."""
    for a in block.iterancestors():
        if a.tag == q("tc"):
            tcw = a.find(f"{q('tcPr')}/{q('tcW')}")
            if tcw is not None and tcw.get(q("type"), "dxa") == "dxa" and tcw.get(q("w"), "0").isdigit():
                return int(tcw.get(q("w"))) * 635 - 2 * 115 * 635, None  # minus default cell margins
    sect = section_for(body, block)
    pg = sect.find(q("pgSz")) if sect is not None else None
    mar = sect.find(q("pgMar")) if sect is not None else None

    def val(el, attr, default):
        try:
            return int(float(el.get(q(attr), default))) if el is not None else int(default)
        except ValueError:
            return int(default)
    width = val(pg, "w", "12240") - val(mar, "left", "1440") - val(mar, "right", "1440")
    height = val(pg, "h", "15840") - val(mar, "top", "1440") - val(mar, "bottom", "1440")
    cols = sect.find(q("cols")) if sect is not None else None
    if cols is not None:
        own = [int(c.get(q("w"))) for c in cols.findall(q("col")) if (c.get(q("w")) or "").isdigit()]
        num = val(cols, "num", "1")
        if own:
            width = min(own)
        elif num > 1:
            width = (width - (num - 1) * val(cols, "space", "720")) // num
    return width * 635, height * 635


def insert_image(root, body, op, ed):
    anchor = find_block(body, op.get("after", op.get("before")), "anchor")
    rid, pw, ph, dpi = add_media(ed, op["file"])
    maxw, maxh = text_area(body, anchor)
    if "width_cm" in op:
        if type(op["width_cm"]) not in (int, float) or not math.isfinite(op["width_cm"]) or op["width_cm"] <= 0:
            raise OpError("width_cm must be finite and positive")
        cx = round(op["width_cm"] * 360000)
        if cx > maxw:
            raise OpError(f"image {op['width_cm']} cm is wider than the text area ({maxw / 360000:.1f} cm)")
    else:
        cx = min(round(pw / dpi * 914400), maxw)  # natural size, capped to the column
    cy = round(cx * ph / pw)
    if maxh and cy > maxh:
        if "width_cm" in op:
            raise OpError(f"image would be {cy / 360000:.1f} cm tall, taller than the page text area; reduce width_cm")
        cy, cx = maxh, round(maxh * pw / ph)
    if cx < 1 or cy < 1:
        raise OpError("image dimensions round to zero")
    align = op.get("align", "center")
    if align not in ALIGN:
        raise OpError(f"align must be one of {sorted(ALIGN)}")
    ids = [int(v) for v in root.xpath("//wp:docPr/@id", namespaces={"wp": WP}) if v.isdigit()]
    pid = max(ids + [0]) + 1
    name = pathlib.Path(op["file"]).name.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
    p = etree.fromstring(
        f'<w:p xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{R_NS}">'
        f'<w:pPr><w:jc w:val="{ALIGN[align]}"/></w:pPr><w:r><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{pid}" name="Picture {pid}" descr="{name}"/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic><a:graphicData uri="{PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="{pid}" name="{name}"/>'
        f'<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        f'</wp:inline></w:drawing></w:r></w:p>')
    (anchor.addnext if "after" in op else anchor.addprevious)(p)
    track_new_block(p, ed)
    print(f"     image {cx / 360000:.2f} x {cy / 360000:.2f} cm")


def replace_image(root, op, ed):
    """Swap the picture shown by drawing `image`, fitted inside its current box (aspect kept),
    so the surrounding layout cannot grow. Position, wrapping and borders stay."""
    if ed.author:
        raise OpError("replace_image: drawing revisions are unsupported; use an untracked copy")
    frames = drawing_frames(root)
    i = op["image"]
    if type(i) is not int or not 0 <= i < len(frames):
        raise OpError(f"image index must be in 0..{len(frames) - 1} (see --images)")
    frame = frames[i]
    blips = frame.xpath(".//a:blip[@r:embed]", namespaces={"a": A, "r": R_NS})
    if len(blips) != 1:
        raise OpError("drawing is not a single embedded picture (chart, shape, group or linked image)")
    ext = frame.find(f"{{{WP}}}extent")
    bw, bh = int(ext.get("cx")), int(ext.get("cy"))
    if bw <= 0 or bh <= 0:
        raise OpError("invalid drawing size")
    rid, pw, ph, _ = add_media(ed, op["file"])
    blip = blips[0]
    blip.set(f"{{{R_NS}}}embed", rid)
    for child in list(blip):  # alternate renditions (SVG) and effects belonged to the old picture
        blip.remove(child)
    for crop in frame.xpath(".//a:srcRect", namespaces={"a": A}):
        crop.getparent().remove(crop)
    scale = min(bw / pw, bh / ph)
    cx, cy = max(1, round(pw * scale)), max(1, round(ph * scale))
    ext.set("cx", str(cx)); ext.set("cy", str(cy))
    for x in frame.xpath(".//a:xfrm/a:ext", namespaces={"a": A}):
        x.set("cx", str(cx)); x.set("cy", str(cy))
    print(f"     image {i}: {bw / 360000:.2f} x {bh / 360000:.2f} -> {cx / 360000:.2f} x {cy / 360000:.2f} cm")


def cell_guard(b, what):
    """A table cell must keep at least one paragraph, and end with one."""
    parent = b.getparent()
    if parent is not None and parent.tag == q("tc"):
        rest = [c for c in parent if c is not b and c.tag != q("tcPr")]
        if not rest or rest[-1].tag != q("p"):
            raise OpError(f"{what}: a table cell must keep a final paragraph; use set_text to empty it instead")


# ---------- op executor ----------
def run_ops(ops, trees, ed, smap, changed):
    global _FIELD_CACHE
    previous = _FIELD_CACHE
    try:
        return _run_ops(ops, trees, ed, smap, changed)
    finally:
        _FIELD_CACHE = previous


def _run_ops(ops, trees, ed, smap, changed):
    global _FIELD_CACHE
    body = trees["word/document.xml"].find(q("body"))
    if not isinstance(ops, list) or not all(isinstance(op, dict) for op in ops):
        raise OpError("ops must be a JSON list of objects")
    for n, op in enumerate(ops, 1):
        # Field nodes are immutable during text edits; rebuild between structural ops.
        _FIELD_CACHE = {}
        kind = op.get("op", "replace")
        tag = f"op {n} ({kind})"
        try:
            validate_op(op)
            if kind == "replace":
                do_replace(op, trees, ed, changed)
                continue
            if kind == "format":
                do_format(op, trees, ed, smap, changed)
                print(f"ok   {tag}")
                continue
            changed.add("word/document.xml")
            if kind == "resize_image":
                resize_image(trees["word/document.xml"], op, ed)
            elif kind == "replace_image":
                replace_image(trees["word/document.xml"], op, ed)
            elif kind == "insert_image":
                insert_image(trees["word/document.xml"], body, op, ed)
            elif kind in ("add_col", "del_col", "col_widths", "merge_cells"):
                {"add_col": add_col, "del_col": del_col, "col_widths": col_widths, "merge_cells": merge_cells}[kind](
                    get_table(trees["word/document.xml"], op["table"]), op, ed)
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
                anchor = find_block(body, op.get("after", op.get("before")), "anchor")
                t = table_like(like, rows) if like is not None else table_plain(body, rows, section_for(body, anchor))
                place(body, t, op)
                track_new_block(t, ed)
            elif kind == "delete":
                seg = block_range(body, op)
                section_guard(seg, op)
                for b in seg:
                    cell_guard(b, "delete")
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
                section_guard(bl[i:j], op)
                for b in bl[i:j]:
                    remove_block(b, ed)
            elif kind == "move":
                seg = block_range(body, op)
                section_guard(seg, op)
                anchor = find_block(body, op.get("after", op.get("before")), "anchor")
                if anchor in seg:
                    raise OpError("move: anchor is inside the moved range")
                for b in seg:
                    cell_guard(b, "move")
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
                tbl = get_table(trees["word/document.xml"], op["table"])
                tr = get_row(tbl, op["row"])
                cells = row_cells(tr)
                c = index_value(op["col"], "col")
                if not 0 <= c < len(cells):
                    raise OpError(f"no cell {c} (row has {len(cells)})")
                set_cell(vertical_source(tbl, tr, cells[c]), op["text"], ed)
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
    ap.add_argument("--full", action="store_true", help="with --blocks: complete text of every block and table")
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

    if pathlib.Path(a.src).suffix.lower() != ".docx":
        sys.exit(f"ERROR {pathlib.Path(a.src).suffix or 'this file'} is not .docx; convert first: "
                 f"python tools/render.py FILE --to-docx")
    zin, trees = open_package(a.src)
    smap = styles_map(zin)
    global _FIELD_CACHE
    _FIELD_CACHE = {}  # listings read every paragraph: scan each part for fields once, not per paragraph
    if a.images:
        list_images(trees["word/document.xml"])
        return
    if a.blocks or a.tables:
        (list_blocks(trees["word/document.xml"].find(q("body")), smap, a.full) if a.blocks
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
    ed.zin, ed.main_part = zin, trees.main_part
    if a.ops:
        ed.base = pathlib.Path(a.ops).resolve().parent
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
            parts = dict(ed.parts)
            for info in zin.infolist():
                data = zin.read(info.filename)
                key = MAIN if info.filename == trees.main_part else info.filename
                if key in changed:
                    data = etree.tostring(trees[key], xml_declaration=True, encoding="UTF-8", standalone=True)
                elif info.filename in parts:
                    data = parts.pop(info.filename)
                zout.writestr(info, data)
            for name, data in parts.items():  # parts that did not exist before (pictures, rels)
                zout.writestr(name, data)
        with zipfile.ZipFile(temp) as chk:
            if chk.testzip():
                raise OpError("output ZIP failed CRC validation")
            for name in list(changed) + [n for n in ed.parts if n.endswith((".xml", ".rels"))]:
                parse(chk.read(trees.main_part if name == MAIN else name))
        zin.close()
    print(f"wrote {dst}" + (f" (tracked as {a.track})" if a.track else ""))


if __name__ == "__main__":
    try:
        main()
    except (OpError, OSError, ValueError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        print(f"ERROR {exc}; nothing published")
        sys.exit(1)
