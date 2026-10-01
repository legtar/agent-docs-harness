"""DOCX/ODT -> PDF through LibreOffice UNO, refreshing TOC/indexes and fields first
(plain `soffice --convert-to pdf` leaves a TOC empty). Runs under LibreOffice's own python:

    "C:/Program Files/LibreOffice/program/python.exe" tools/lo_export.py IN OUT.pdf PROFILE_DIR

Used by render.py; not meant to be called directly.
"""
import pathlib
import subprocess
import sys
import time
import uuid

import uno
from com.sun.star.beans import PropertyValue


def prop(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v
    return p


def main(src, dst, profile):
    office = pathlib.Path(sys.executable).parent / "soffice.exe"
    pipe = f"lo_export_{uuid.uuid4().hex}"
    proc = subprocess.Popen([str(office), f"-env:UserInstallation={pathlib.Path(profile).as_uri()}",
                             "--headless", "--invisible", "--norestore", "--nologo",
                             f"--accept=pipe,name={pipe};urp;"])
    desktop, doc = None, None
    try:
        resolver = uno.getComponentContext().ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", uno.getComponentContext())
        for _ in range(120):
            try:
                ctx = resolver.resolve(f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        else:
            sys.exit("LibreOffice did not start")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(pathlib.Path(src).resolve().as_uri(), "_blank", 0,
                                           (prop("Hidden", True), prop("ReadOnly", True)))
        if doc is None:
            raise RuntimeError(f"LibreOffice could not load {src}")
        if hasattr(doc, "getDocumentIndexes"):
            idx = doc.getDocumentIndexes()
            for _ in range(2):  # 2nd pass: TOC page numbers settle after the TOC itself takes space
                for i in range(idx.getCount()):
                    idx.getByIndex(i).update()
                doc.getTextFields().refresh()
        doc.storeToURL(pathlib.Path(dst).resolve().as_uri(), (prop("FilterName", "writer_pdf_Export"),))
        try:
            doc.close(True)
        except Exception:  # LO sometimes drops the bridge on close; the PDF is already written
            pass
    finally:
        if doc is not None:
            try:
                doc.close(True)
            except Exception:
                pass
        try:
            if desktop is not None:
                desktop.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=20)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=10)


if __name__ == "__main__":
    main(*sys.argv[1:4])
