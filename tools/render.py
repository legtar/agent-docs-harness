"""Render a document to PDF + PNG pages + one contact sheet for visual QA.

    python tools/render.py FILE [--dpi 110] [--out out/render/<name>]

FILE: .docx .doc .odt .rtf (LibreOffice) | .typ (Typst) | .html (WeasyPrint) | .pdf
Prints the paths of the PDF, the page PNGs and sheet.png (all pages on one image).
"""
import argparse
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile

import fitz  # PyMuPDF

ROOT = pathlib.Path(__file__).resolve().parent.parent
FONTS = ROOT / "fonts"


def soffice() -> str:
    for c in (shutil.which("soffice"),
              r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if c and os.path.exists(c):
            return c
    sys.exit("LibreOffice not found: winget install -e --id TheDocumentFoundation.LibreOffice")


def to_pdf(src: pathlib.Path, out: pathlib.Path) -> pathlib.Path:
    pdf = out / (src.stem + ".pdf")
    ext = src.suffix.lower()
    if ext == ".pdf":
        if src.resolve() != pdf.resolve():
            shutil.copyfile(src, pdf)
    elif ext == ".typ":
        import typst
        typst.compile(str(src), output=str(pdf), root=str(ROOT), font_paths=[str(FONTS)])
    elif ext in (".html", ".htm"):
        from weasyprint import HTML
        HTML(filename=str(src)).write_pdf(str(pdf))
    elif ext in (".docx", ".odt", ".doc", ".rtf") and (lo_py := pathlib.Path(soffice()).parent / "python.exe").exists():
        # UNO path: refreshes TOC/fields before export, like Word does on open.
        with tempfile.TemporaryDirectory(prefix="lo_profile_", ignore_cleanup_errors=True) as prof:
            subprocess.run([str(lo_py), str(pathlib.Path(__file__).with_name("lo_export.py")), str(src), str(pdf), prof],
                           check=True, capture_output=True, timeout=300)
    elif ext in (".docx", ".odt", ".doc", ".rtf"):
        # Isolated profile: parallel soffice runs otherwise fight over the user profile lock.
        with tempfile.TemporaryDirectory(prefix="lo_profile_") as prof:
            subprocess.run([soffice(), f"-env:UserInstallation={pathlib.Path(prof).as_uri()}",
                            "--headless", "--convert-to", "pdf", "--outdir", str(out), str(src)],
                           check=True, capture_output=True, timeout=300)
    else:
        raise ValueError(f"unsupported document extension: {ext}")
    if not pdf.exists():
        sys.exit(f"render failed: {pdf} not produced")
    return pdf


def to_png(pdf: pathlib.Path, out: pathlib.Path, dpi: int) -> list[pathlib.Path]:
    for old in out.glob("page-*.png"):
        old.unlink()
    pages = []
    with fitz.open(pdf) as doc:
        for i, page in enumerate(doc, 1):
            p = out / f"page-{i:03}.png"
            page.get_pixmap(dpi=dpi).save(p)
            pages.append(p)
        sheet(doc, out / "sheet.png")
    return pages


def sheet(doc, path: pathlib.Path, cols: int = 4, max_pages: int = 24) -> None:
    """All pages as thumbnails on one image: fastest way to eyeball rhythm and stray pages."""
    n = min(len(doc), max_pages)
    w, h = doc[0].rect.width, doc[0].rect.height
    gap, rows = 12, (n + cols - 1) // cols
    cw = min(cols, n)
    out = fitz.open()
    pg = out.new_page(width=cw * (w + gap) + gap, height=rows * (h + gap) + gap)
    pg.draw_rect(pg.rect, color=None, fill=(0.85, 0.85, 0.85))
    for i in range(n):
        r, c = divmod(i, cols)
        rect = fitz.Rect(gap + c * (w + gap), gap + r * (h + gap), 0, 0)
        rect.x1, rect.y1 = rect.x0 + w, rect.y0 + h
        pg.draw_rect(rect, color=None, fill=(1, 1, 1))
        if doc[i].get_contents():
            pg.show_pdf_page(rect, doc, i, rotate=-doc[i].rotation)
    pg.get_pixmap(dpi=36 if n > 8 else 50).save(path)


def render(src, out=None, dpi=110):
    src = pathlib.Path(src).resolve()
    out = pathlib.Path(out or ROOT / "out" / "render" / src.stem).resolve()
    out.mkdir(parents=True, exist_ok=True)
    # Convert in a fresh directory: a stale PDF cannot masquerade as successful export.
    if src.suffix.lower() == ".pdf":
        pdf = to_pdf(src, out)
    else:
        with tempfile.TemporaryDirectory(prefix="render-", dir=out) as staging:
            fresh = to_pdf(src, pathlib.Path(staging))
            with fitz.open(fresh) as check:
                if not len(check):
                    raise ValueError("render produced an empty PDF")
            pdf = out / fresh.name
            os.replace(fresh, pdf)
    return pdf, to_png(pdf, out, dpi)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--out")
    a = ap.parse_args()
    pdf, pages = render(a.file, a.out, a.dpi)
    print(f"pdf:   {pdf}\nsheet: {pdf.parent / 'sheet.png'}\npages: {len(pages)} in {pdf.parent}")
