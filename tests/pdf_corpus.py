"""Universality sweep for in-place PDF edits (tools/pdf_flow.py) over real-world PDFs.

    python tests/pdf_corpus.py [--only SUBSTR] [--pages 3] [DIR ...]

DIR defaults to out/pdfcorpus/* (pdf.js test PDFs, arXiv papers, our own renders — fetch them
yourself, they are not in git). On the first pages of every PDF the sweep tries, each on a fresh
copy: del_row, add_row, cell on a table row; insert after and delete of a text block.

Every attempt must end as ok or refused (a clean error, nothing written). A crash, or an ok
whose result does not hold what the op promised, fails the sweep. pdf_flow verifies by itself
that everything outside the edit is pixel-identical; here the functional outcome is checked:
the deleted row's text is gone, the added row's values are extractable as a row, and so on.
"""
import argparse
import collections
import pathlib
import re
import sys
import traceback

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import pymupdf as fitz  # noqa: E402
import pdf_flow as pf  # noqa: E402

OUT = ROOT / "out" / "pdfcorpus_run"


def reason_key(msg):
    msg = re.sub(r"'[^']*'|\"[^\"]*\"|\[[^\]]*\]|\([^)]*\)", "…", msg)
    return re.sub(r"\d+(\.\d+)?", "N", msg)[:100]


def attempt(data, op, check):
    """Run one op on a fresh document. Returns (verdict, detail, bytes or None)."""
    try:
        flow = pf.Flow(data)
        try:
            flow.run(op)
            out = flow.doc.tobytes(garbage=3, deflate=True)
        finally:
            flow.doc.close()
    except pf.OpError as e:
        return "refused", str(e), None
    except Exception as e:  # anything else is a bug in the tool
        return "crash", f"{type(e).__name__}: {e} @ {traceback.format_exc().strip().splitlines()[-3].strip()[:80]}", None
    try:
        with fitz.open(stream=out, filetype="pdf") as doc:
            problem = check(doc)
    except Exception as e:
        return "corrupt", f"{type(e).__name__}: {e}", out
    return ("wrong", problem, out) if problem else ("ok", "", out)


def row_text(row):
    return pf.norm_ws(" ".join(c[2] for c in row.cells))


def candidates(page):
    """(table row index, block anchor) worth trying on this page."""
    rows = pf.visual_rows(page)
    texts = [row_text(r) for r in rows]
    table = next((i for i in range(1, len(rows) - 1)
                  if len(rows[i].cells) >= 3 and len(rows[i - 1].cells) == len(rows[i].cells) == len(rows[i + 1].cells)
                  and texts.count(texts[i]) == 1 and len(texts[i]) > 5), None)
    anchor = None
    for b in page.get_text("dict")["blocks"]:
        if b["type"] == 0 and len(b.get("lines", [])) >= 2 and b["bbox"][2] - b["bbox"][0] > 180:
            first = pf.norm_ws("".join(s["text"] for s in b["lines"][0]["spans"]))[:45]
            if len(first) > 12:
                anchor = first
                break
    return rows, table, anchor


def sweep(path, max_pages):
    res = []
    try:
        data = path.read_bytes()
        doc = fitz.open(stream=data, filetype="pdf")
        if doc.needs_pass:
            return [("open", "refused", "encrypted")]
    except Exception as e:
        return [("open", "unreadable", f"{type(e).__name__}: {e}"[:100])]
    done = set()
    work = OUT / f"{path.parent.name}__{path.stem}"
    for page in list(doc)[:max_pages]:
        if page.rotation or not pf.page_spans(page):
            continue
        n = page.number + 1
        rows, ti, anchor = candidates(page)
        ops = []
        if ti is not None and "row" not in done:
            done.add("row")
            target = row_text(rows[ti])
            values = [c[2] for c in rows[ti].cells]
            marker = values[0]
            ops += [
                ("del_row", {"op": "del_row", "page": n, "row": ti},
                 lambda d, t=target, k=page.number: f"row text still on page: {t[:40]!r}"
                 if t in pf.norm_ws(" ".join(row_text(r) for r in pf.visual_rows(d[k]))) else None),
                ("add_row", {"op": "add_row", "page": n, "after": ti, "values": values},
                 lambda d, t=target, k=page.number: None
                 if [row_text(r) for r in pf.visual_rows(d[k])].count(t) == 2 else f"no second row {t[:40]!r}"),
                ("cell", {"op": "cell", "page": n, "row": ti, "col": 0, "text": marker[::-1]},
                 lambda d, m=marker[::-1], k=page.number: None if m in d[k].get_text() else f"new cell text {m!r} not found"),
            ]
        if anchor and "block" not in done:
            done.add("block")
            words = [w for w in anchor.split() if w.isalpha()] or ["text"]
            new = " ".join((words * 12)[:18])
            ops += [
                ("insert", {"op": "insert", "page": n, "after": anchor, "text": new},
                 lambda d, m=pf.norm_ws(new)[:25], k=page.number: None
                 if m in pf.norm_ws(d[k].get_text()) else "inserted text not found"),
                ("delete", {"op": "delete", "page": n, "block": anchor},
                 lambda d, a=anchor, k=page.number: f"block still there: {a[:30]!r}"
                 if a in pf.norm_ws(" ".join(r for r in map(row_text, pf.visual_rows(d[k])))) else None),
            ]
        for name, op, check in ops:
            verdict, why, out = attempt(data, op, check)
            res.append((name, verdict, why))
            if out and verdict in ("ok", "wrong"):
                work.mkdir(parents=True, exist_ok=True)
                (work / f"{name}.pdf").write_bytes(out)
    doc.close()
    return res


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("dirs", nargs="*")
    ap.add_argument("--only")
    ap.add_argument("--pages", type=int, default=3)
    a = ap.parse_args()
    dirs = [pathlib.Path(d) for d in a.dirs] or sorted(p for p in (ROOT / "out" / "pdfcorpus").iterdir() if p.is_dir())
    pdfs = sorted(f for d in dirs for f in d.glob("*.pdf") if not a.only or a.only in f.name)
    if not pdfs:
        sys.exit("no PDFs: put files under out/pdfcorpus/<set>/")
    OUT.mkdir(parents=True, exist_ok=True)
    stages = collections.defaultdict(collections.Counter)
    reasons = collections.defaultdict(lambda: collections.defaultdict(list))
    for n, path in enumerate(pdfs, 1):
        res = sweep(path, a.pages)
        bad = [(s, v, w) for s, v, w in res if v in ("crash", "corrupt", "wrong")]
        print(f"[{n}/{len(pdfs)}] {'BAD ' if bad else 'ok  '} {path.parent.name}/{path.name} "
              f"{' '.join(f'{s}={v}' for s, v, _ in res)} {bad or ''}", flush=True)
        for s, v, w in res:
            stages[s][v] += 1
            if v != "ok":
                reasons[v][f"{s}: {reason_key(w)}"].append(path.name)
    print(f"\n{len(pdfs)} PDFs")
    for s, c in stages.items():
        print(f"  {s:<8} " + "  ".join(f"{k}={v}" for k, v in sorted(c.items())))
    for v in ("crash", "corrupt", "wrong", "unreadable", "refused"):
        for why, names in sorted(reasons[v].items(), key=lambda x: -len(x[1])):
            print(f"{v.upper():<10} x{len(names):<3} {why}   e.g. {names[0]}")
    bad = sum(c[v] for c in stages.values() for v in ("crash", "corrupt", "wrong"))
    print(f"\n{'FAILED: ' + str(bad) + ' bad results' if bad else 'ALL GREEN'}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
