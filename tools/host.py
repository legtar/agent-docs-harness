"""Where things live on this machine — LibreOffice, pandoc, fonts — on Windows, macOS and Linux.
Every lookup can be overridden by an environment variable (SOFFICE, PANDOC)."""
import functools
import os
import pathlib
import re
import shutil
import subprocess
import sys

HOME = pathlib.Path.home()
INSTALL = {
    "soffice": "winget install -e --id TheDocumentFoundation.LibreOffice | brew install --cask libreoffice | "
               "sudo apt install libreoffice   (or set SOFFICE=/path/to/soffice)",
    "pandoc": "winget install -e --id JohnMacFarlane.Pandoc | brew install pandoc | sudo apt install pandoc   "
              "(or set PANDOC=/path/to/pandoc)",
}


def _first(candidates):
    for c in candidates:
        if c and pathlib.Path(c).is_file():
            return str(c)
    return None


@functools.cache
def soffice() -> str:
    found = _first([
        os.environ.get("SOFFICE"), shutil.which("soffice"), shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        HOME / "Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/lib/libreoffice/program/soffice", "/usr/lib64/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice", "/snap/bin/libreoffice",
    ])
    if not found:
        sys.exit("LibreOffice not found: " + INSTALL["soffice"])
    return found


@functools.cache
def lo_python():
    """A Python that can `import uno`: LibreOffice's bundled one (Windows, macOS, tarball installs)
    or the system python3 with python3-uno (Linux packages). None if there is none — callers then
    fall back to plain `soffice --convert-to`."""
    program = pathlib.Path(soffice()).resolve().parent
    bundled = _first([program / "python.exe", program / "python", program.parent / "Resources" / "python"])
    if bundled:
        return bundled
    for py in dict.fromkeys(filter(None, ("/usr/bin/python3", shutil.which("python3")))):
        try:
            if subprocess.run([py, "-c", "import uno"], capture_output=True, timeout=20).returncode == 0:
                return py
        except (OSError, subprocess.SubprocessError):
            pass
    return None


@functools.cache
def pandoc() -> str:
    found = _first([
        os.environ.get("PANDOC"), shutil.which("pandoc"),
        HOME / "AppData/Local/Pandoc/pandoc.exe", r"C:\Program Files\Pandoc\pandoc.exe",
        "/opt/homebrew/bin/pandoc", "/usr/local/bin/pandoc", "/usr/bin/pandoc",
    ])
    if not found:
        sys.exit("pandoc not found: " + INSTALL["pandoc"])
    return found


def font_dirs() -> list[pathlib.Path]:
    if sys.platform == "win32":
        dirs = [pathlib.Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
                HOME / "AppData/Local/Microsoft/Windows/Fonts"]
    elif sys.platform == "darwin":
        dirs = [pathlib.Path("/System/Library/Fonts"), pathlib.Path("/System/Library/Fonts/Supplemental"),
                pathlib.Path("/Library/Fonts"), HOME / "Library/Fonts"]
    else:
        dirs = [pathlib.Path("/usr/share/fonts"), pathlib.Path("/usr/local/share/fonts"),
                HOME / ".fonts", HOME / ".local/share/fonts"]
    return [d for d in dirs if d.is_dir()]


def font_files() -> list[pathlib.Path]:
    """Every installed font file (Linux and macOS keep them in nested folders)."""
    return [f for d in font_dirs() for f in d.rglob("*") if f.suffix.lower() in (".ttf", ".otf", ".ttc")]


def norm_family(name: str) -> str:
    fam = re.sub(r"\s+(Bold|Italic|Oblique|Light|Semibold|SemiBold|Black|Regular|Medium|Thin)\b.*", "", name)
    return re.sub(r"[^a-z0-9]", "", fam.lower())


@functools.cache
def installed_fonts() -> frozenset:
    """Normalised family names available to renderers on this machine (empty = unknown)."""
    names = set()
    if sys.platform == "win32":
        import winreg
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                key = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
            except OSError:
                continue
            for i in range(winreg.QueryInfoKey(key)[1]):
                entry = re.sub(r"\s*\((TrueType|OpenType)\)", "", winreg.EnumValue(key, i)[0])
                names.update(norm_family(part) for part in entry.split(" & "))
        return frozenset(names)
    fc = shutil.which("fc-list")
    if fc:
        try:
            out = subprocess.run([fc, ":", "family"], capture_output=True, text=True, timeout=60).stdout
            names.update(norm_family(f) for line in out.splitlines() for f in line.split(","))
        except (OSError, subprocess.SubprocessError):
            pass
    if not names:  # no fontconfig (stock macOS): read family names from the font files
        import pymupdf as fitz
        for f in font_files():
            try:
                names.add(norm_family(fitz.Font(fontfile=str(f)).name))
            except Exception:
                continue
    names.discard("")
    return frozenset(names)


# ---------------------------------------------------------------- font files by family and style
ROOT = pathlib.Path(__file__).resolve().parent.parent
STYLE = re.compile(r"\s+(Regular|Bold Italic|Bold Oblique|Bold|Italic|Oblique)$", re.I)
ALIASES = {"helvetica": "arial", "times": "timesnewroman", "timesroman": "timesnewroman", "courier": "couriernew"}


def _index(files):
    import pymupdf as fitz
    idx = {}
    for f in files:
        try:
            font = fitz.Font(fontfile=str(f))
        except Exception:
            continue
        family = re.sub(r"[^a-z0-9]", "", STYLE.sub("", font.name).lower())
        idx.setdefault(family, {}).setdefault((bool(font.is_bold), bool(font.is_italic)), f)
    return idx


@functools.cache
def bundled_fonts():
    return _index(sorted(f for f in (ROOT / "fonts").glob("*") if f.suffix.lower() in (".ttf", ".otf", ".ttc")))


@functools.cache
def system_fonts():
    return _index(font_files())


def pdf_family(name: str) -> str:
    """'ABCDEF+TimesNewRomanPS-BoldMT' -> 'timesnewroman'."""
    base = re.split(r"[-,]", name.split("+")[-1])[0]
    family = re.sub(r"(psmt|mt|ps)$", "", re.sub(r"[^a-z0-9]", "", base.lower()))
    return ALIASES.get(family, family)


def find_font(pdf_name: str, bold: bool, italic: bool):
    """File of the same family and style as a PDF font: the fonts shipped in fonts/ first, then
    the system's. None when the family or that style is not available — never another family."""
    family = pdf_family(pdf_name)
    for index in (bundled_fonts, system_fonts):
        styles = index().get(family)
        if styles and (bold, italic) in styles:
            return styles[(bold, italic)]
    return None


if __name__ == "__main__":  # python tools/host.py — what the harness found on this machine
    print("platform :", sys.platform)
    for label, fn in (("soffice", soffice), ("uno py", lo_python), ("pandoc", pandoc)):
        try:
            print(f"{label:<9}:", fn())
        except SystemExit as e:
            print(f"{label:<9}: MISSING — {e}")
    fonts = installed_fonts()
    print("fonts    :", len(fonts), "families;", "house fonts present:" if fonts else "",
          [f for f in ("calibri", "georgia", "carlito", "liberationserif") if f in fonts])
