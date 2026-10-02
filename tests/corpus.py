"""Universality sweep: run the DOCX editor over a corpus of real-world documents.

    python tests/corpus.py [--render] [--jobs 4] [--only SUBSTR] [DIR ...]

DIR defaults to out/corpus/* (see out/corpus/README for where the files come from: Apache POI
test-data, pandoc test/docx, python-docx fixtures — fetch them yourself, they are not in git).

Per document, every stage must end as ok / refused (clean ERROR, nothing written) — never as a
crash (traceback), a corrupt output (unreadable ZIP/XML) or a silent text change:
  list      --blocks, --tables, --images
  replace   unique word -> other word -> back; all paragraph texts equal to the original
  blocks    insert paragraph after the first text block, set_text, move, delete; texts equal
  table     add_row/del_row and add_col/del_col on the first table; texts equal; format on a side copy
  tracked   the replace + insert as tracked changes; accepted view must still be well-formed
  render    (--render) qa.py round-trip.docx --original: page count and unchanged pages identical
Exit 1 if any crash / corrupt / mismatch. Refusals are listed so new ones can be reviewed.
"""
import argparse
import collections
import concurrent.futures
import json
import os
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import docx_edit as de  # noqa: E402

OUT = ROOT / "out" / "corpus_run"
ENV = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}


def tool(name, *args, timeout=300):
    r = subprocess.run([sys.executable, str(ROOT / "tools" / name), *map(str, args)], capture_output=True,
                       text=True, encoding="utf-8", errors="replace", cwd=ROOT, env=ENV, timeout=timeout)
    return r.returncode, r.stdout + r.stderr


def load(path):
    zin, trees = de.open_package(path)
    zin.close()
    return trees


def texts(path):
    """Every paragraph's visible text in every part — the invariant a round trip must keep."""
    de._FIELD_CACHE = {}  # one field scan per part instead of one per paragraph
    return [de.model(p)[0] for _, root in sorted(load(path).items()) for p in root.iter(de.q("p"))]


def verdict(rc, out):
    if "Traceback" in out:
        return "crash", out.strip().splitlines()[-1][:160]
    if rc == 0:
        return "ok", ""
    m = re.search(r"ERROR (.*)", out)
    return "refused", (m.group(1) if m else out.strip()[-160:])[:160]


def reason_key(msg):
    """Group refusals/crashes: drop quoted document text and numbers."""
    msg = re.sub(r"op \d+ \((\w+)\)", r"\1", msg)
    msg = re.sub(r"'[^']*'|\"[^\"]*\"|\[[^\]]*\]", "…", msg)
    return re.sub(r"\d+", "N", msg)[:110]


def ops_file(work, name, ops):
    f = work / f"{name}.json"
    f.write_text(json.dumps(ops, ensure_ascii=False), encoding="utf-8")
    return f


def edit(work, src, name, ops, *extra):
    dst = work / f"{work.name}__{name}.docx"   # unique name: qa.py keys its render folder by file name
    dst.unlink(missing_ok=True)
    rc, out = tool("docx_edit.py", src, dst, "--ops", ops_file(work, name, ops), *extra)
    v, why = verdict(rc, out)
    if v == "ok":
        try:
            texts(dst)
        except Exception as e:  # unreadable output
            return "corrupt", f"{type(e).__name__}: {e}"[:160], dst
    elif dst.exists():
        return "corrupt", "output written despite failure: " + why, dst
    return v, why, dst


def stage_roundtrip(work, src, base, name, forward, backward, extra=()):
    """forward ops then backward ops; texts must equal `base`. Returns (verdict, why, final path)."""
    v, why, mid = edit(work, src, name + "_fwd", forward, *extra)
    if v != "ok":
        return v, why, src
    if texts(mid) == base and not extra:
        return "mismatch", "forward edit changed nothing", src
    v, why, fin = edit(work, mid, name + "_back", backward, *extra)
    if v != "ok":
        return v, "backward: " + why, src
    if texts(fin) != base:
        a, b = base, texts(fin)
        diff = next(((x, y) for x, y in zip(a, b) if x != y), (len(a), len(b)))
        return "mismatch", f"text differs after round trip: {str(diff)[:120]}", src
    return "ok", "", fin


def check(doc, render):
    rel = f"{doc.parent.name}__{doc.stem}"
    work = OUT / rel
    work.mkdir(parents=True, exist_ok=True)
    res = {}
    try:
        base = texts(doc)
        trees = load(doc)
    except Exception as e:
        return rel, {"open": ("unreadable", f"{type(e).__name__}: {e}"[:120])}
    for flag in ("--blocks", "--tables", "--images"):
        res["list" + flag] = verdict(*tool("docx_edit.py", doc, flag))
    body = trees[de.MAIN].find(de.q("body"))
    if body is None:
        return rel, res
    de._FIELD_CACHE = {}
    blocks = de.blocks(body)
    btext = [de.block_text(b) for b in blocks]
    paras = [(i, btext[i]) for i, b in enumerate(blocks) if b.tag == de.q("p")]
    paras = [(i, t) for i, t in paras if t.strip() and de.OPAQUE not in de.model(blocks[i])[0]][:400]
    cur = doc

    # replace: a word that occurs exactly once in the whole document
    alltext = "\n".join(base)
    words = collections.Counter(re.findall(r"[^\W\d_]{5,}", alltext))
    cand = [w for _, t in paras for w in re.findall(r"[^\W\d_]{5,}", t) if words[w] == 1 and alltext.count(w) == 1][:4]
    v, why = "skipped", "no unique word"
    for w in cand:
        new = w[::-1] + "q"
        if new in alltext:
            continue
        v, why, fin = stage_roundtrip(work, cur, base, "replace", [{"op": "replace", "find": w, "replace": new}],
                                      [{"op": "replace", "find": new, "replace": w}])
        if v != "refused":
            cur = fin
            break
    res["replace"] = (v, why)

    # block ops on the first uniquely addressable paragraph
    uniq = next(([(i, t)] for i, t in paras if len(t) > 3 and sum(t in x for x in btext) == 1), [])
    if uniq:
        i, t = uniq[0]
        mark, mark2 = "Zq corpus inserted paragraph 7141", "Zq corpus rewritten paragraph 9152"
        fwd = [{"op": "insert", "after": f"#{i}", "text": mark},
               {"op": "set_text", "block": mark, "text": mark2},
               {"op": "move", "block": mark2, "before": f"#{i}"}]
        v, why, fin = stage_roundtrip(work, cur, base, "blocks", fwd, [{"op": "delete", "block": mark2}])
        res["blocks"] = (v, why)
        cur = fin if v == "ok" else cur
        # tracked variant: must apply cleanly and stay readable
        v, why, _ = edit(work, cur, "tracked", fwd[:2], "--track", "Corpus")
        res["tracked"] = (v, why)
    else:
        res["blocks"] = ("skipped", "no uniquely addressable paragraph")

    # table ops on the first top-level table
    tbls = [b for b in blocks if b.tag == de.q("tbl")]
    if tbls:
        ti = list(trees[de.MAIN].iter(de.q("tbl"))).index(tbls[0])
        rows = de.table_rows(tbls[0])
        ncell = len(de.row_cells(rows[-1])) if rows else 0
        if ncell:
            fwd = [{"op": "add_row", "table": ti, "after": -1, "values": [f"Zq{c}" for c in range(ncell)]}]
            v, why, fin = stage_roundtrip(work, cur, base, "table_rows", fwd, [{"op": "del_row", "table": ti, "row": -1}])
            res["table_rows"] = (v, why)
            cur = fin if v == "ok" else cur
            fwd = [{"op": "add_col", "table": ti, "after": -1, "values": [f"Zc{r}" for r in range(len(rows))]}]
            v, why, _ = stage_roundtrip(work, cur, base, "table_cols", fwd, [{"op": "del_col", "table": ti, "col": -1}])
            res["table_cols"] = (v, why)
            # not carried into the render check: adding and removing a column rescales the grid
            # twice, which may leave a 1-twip rounding difference (tests/features.py bounds it)
            # formatting changes the look on purpose, so it is checked on a side copy
            v, why, _ = edit(work, cur, "format", [
                {"op": "format", "table": ti, "row": 0, "bold": True, "color": "C00000", "fill": "FFF2CC", "align": "center"},
                {"op": "format", "table": ti, "col": -1, "italic": True, "size": 9}])
            res["format"] = (v, why)
    if uniq:
        v, why, _ = edit(work, cur, "format_text", [
            {"op": "format", "block": f"#{uniq[0][0]}", "underline": True, "highlight": "yellow", "font": "Georgia"}],
            "--track", "Corpus")
        res["format_tracked"] = (v, why)

    if render and cur != doc:
        rc, out = tool("qa.py", cur, "--original", doc, timeout=600)
        if "Traceback" in out:
            # the original itself may be unrenderable (LibreOffice rejects it) — not our regression
            rc0, out0 = tool("render.py", doc, "--out", work / "orig_render", timeout=300)
            res["render"] = ("crash", out.strip().splitlines()[-1][:160]) if rc0 == 0 else ("skipped", "original does not render")
        else:
            errs = [ln for ln in out.splitlines() if ln.startswith("ERROR")]
            res["render"] = ("ok", "") if rc == 0 else ("mismatch", "; ".join(errs)[:200])
    return rel, res


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("dirs", nargs="*")
    ap.add_argument("--render", action="store_true")
    ap.add_argument("--jobs", type=int, default=4)
    ap.add_argument("--only")
    a = ap.parse_args()
    dirs = [pathlib.Path(d) for d in a.dirs] or sorted(p for p in (ROOT / "out" / "corpus").iterdir() if p.is_dir())
    docs = sorted(f for d in dirs for f in d.glob("*.docx") if not a.only or a.only in f.name)
    if not docs:
        sys.exit("no documents: put .docx files under out/corpus/<set>/")
    OUT.mkdir(parents=True, exist_ok=True)
    results = {}
    with concurrent.futures.ThreadPoolExecutor(a.jobs) as ex:
        futs = {ex.submit(check, d, a.render): d for d in docs}
        for n, f in enumerate(concurrent.futures.as_completed(futs), 1):
            try:
                rel, res = f.result()
            except Exception as e:
                rel, res = futs[f].stem, {"harness": ("crash", f"{type(e).__name__}: {e}"[:160])}
            results[rel] = res
            bad = {k: v for k, v in res.items() if v[0] in ("crash", "corrupt", "mismatch")}
            print(f"[{n}/{len(docs)}] {'BAD ' if bad else 'ok  '} {rel} {bad or ''}", flush=True)
    (OUT / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")

    stages = collections.defaultdict(collections.Counter)
    reasons = collections.defaultdict(lambda: collections.defaultdict(list))
    for rel, res in results.items():
        for st, (v, why) in res.items():
            stages[st][v] += 1
            if v not in ("ok", "skipped"):
                reasons[v][f"{st}: {reason_key(why)}"].append(rel)
    print(f"\n{len(docs)} documents")
    for st, c in stages.items():
        print(f"  {st:<14} " + "  ".join(f"{k}={v}" for k, v in sorted(c.items())))
    for v in ("crash", "corrupt", "mismatch", "unreadable", "refused"):
        for why, rels in sorted(reasons[v].items(), key=lambda x: -len(x[1])):
            print(f"{v.upper():<9} x{len(rels):<3} {why}   e.g. {rels[0]}")
    bad = sum(c[v] for c in stages.values() for v in ("crash", "corrupt", "mismatch"))
    print(f"\n{'FAILED: ' + str(bad) + ' bad stage results' if bad else 'ALL GREEN'}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
