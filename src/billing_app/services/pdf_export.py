"""Export .xlsx / .docx files to PDF.

Strategy (tried in order):

1. **MS Office via COM** on Windows — the client almost certainly has Word
   and Excel installed; this is the zero-install path we prefer. Requires
   the `pywin32` package, which is bundled into the Windows PyInstaller
   build but lazy-imported so the module still loads on macOS/Linux.
2. **LibreOffice headless** — fallback for machines without MS Office
   (e.g. the developer's Linux box, or a macOS client that uses LibreOffice).
3. **No PDF** — surface a clear, actionable error so the UI can show a
   friendly message instead of failing silently.
"""
from __future__ import annotations

import platform
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path


class PdfExportError(Exception):
    pass


@dataclass
class PdfExportResult:
    generated: list[Path]
    skipped: list[Path]
    backend: str  # "msoffice", "libreoffice", or "" when nothing ran


_SUPPORTED_EXTS = {".xlsx", ".xlsm", ".docx"}


def _find_libreoffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        hit = shutil.which(name)
        if hit:
            return hit
    for candidate in (
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return None


def _have_msoffice_com() -> bool:
    """Probe whether pywin32 is importable and MS Office is installed."""
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    # We don't actually Dispatch here (that would launch Office); instead
    # check the registry for the ProgIDs. pywin32 exposes this via winreg.
    try:
        import winreg
        for progid in ("Word.Application", "Excel.Application"):
            try:
                winreg.CloseKey(winreg.OpenKey(
                    winreg.HKEY_CLASSES_ROOT, progid,
                ))
            except OSError:
                return False
        return True
    except ImportError:
        # winreg exists on all Windows Python builds; absence means we're
        # not on Windows after all.
        return False


def pdf_backend_available() -> str:
    """Return the preferred available backend name, or '' if none."""
    if _have_msoffice_com():
        return "msoffice"
    if _find_libreoffice():
        return "libreoffice"
    return ""


def install_hint() -> str:
    """Return a message suggesting how to enable PDF export on this machine."""
    if platform.system() == "Windows":
        return (
            "PDF export uses Microsoft Word and Excel. Install Microsoft "
            "Office (or LibreOffice) and try again."
        )
    return (
        "PDF export uses LibreOffice on macOS/Linux. Install it from "
        "https://www.libreoffice.org/download/ and try again."
    )


def export_to_pdf(files: list[Path], out_dir: Path) -> PdfExportResult:
    """Convert each .xlsx/.docx file in `files` to PDF in `out_dir`.

    Raises PdfExportError when no PDF backend is available or a conversion
    fails.
    """
    backend = pdf_backend_available()
    if not backend:
        raise PdfExportError(install_hint())

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if backend == "msoffice":
        return _export_msoffice(files, out_dir)
    return _export_libreoffice(files, out_dir)


# ------------------------------------------------------------- MS Office --

def _export_msoffice(files: list[Path], out_dir: Path) -> PdfExportResult:
    # Import here so non-Windows Python doesn't choke at module import time.
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]

    pythoncom.CoInitialize()
    word = None
    excel = None
    generated: list[Path] = []
    skipped: list[Path] = []
    try:
        for f in files:
            p = Path(f)
            if not p.exists():
                skipped.append(p)
                continue
            ext = p.suffix.lower()
            if ext not in _SUPPORTED_EXTS:
                skipped.append(p)
                continue
            pdf = out_dir / (p.stem + ".pdf")
            # Word/Excel COM objects want absolute paths.
            abs_in = str(p.resolve())
            abs_out = str(pdf.resolve())
            if ext == ".docx":
                if word is None:
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0
                doc = word.Documents.Open(abs_in, ReadOnly=True)
                try:
                    # wdFormatPDF = 17
                    doc.SaveAs(abs_out, FileFormat=17)
                finally:
                    doc.Close(SaveChanges=False)
            else:  # .xlsx / .xlsm
                if excel is None:
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(abs_in, ReadOnly=True)
                try:
                    # xlTypePDF = 0
                    wb.ExportAsFixedFormat(0, abs_out)
                finally:
                    wb.Close(SaveChanges=False)
            if pdf.exists():
                generated.append(pdf)
            else:
                raise PdfExportError(
                    f"MS Office reported success but no PDF appeared for {p.name}."
                )
    except PdfExportError:
        raise
    except Exception as e:  # noqa: BLE001
        raise PdfExportError(
            f"MS Office PDF export failed: {type(e).__name__}: {e}"
        ) from e
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:  # noqa: BLE001
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:  # noqa: BLE001
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass

    return PdfExportResult(
        generated=generated, skipped=skipped, backend="msoffice",
    )


# ---------------------------------------------------------- LibreOffice ---

def _export_libreoffice(files: list[Path], out_dir: Path) -> PdfExportResult:
    soffice = _find_libreoffice()
    assert soffice is not None  # caller checked via pdf_backend_available()

    generated: list[Path] = []
    skipped: list[Path] = []
    for f in files:
        p = Path(f)
        if not p.exists():
            skipped.append(p)
            continue
        if p.suffix.lower() not in _SUPPORTED_EXTS:
            skipped.append(p)
            continue
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir), str(p)],
            capture_output=True, text=True, timeout=120,
        )
        pdf = out_dir / (p.stem + ".pdf")
        if result.returncode == 0 and pdf.exists():
            generated.append(pdf)
        else:
            raise PdfExportError(
                f"LibreOffice failed to convert {p.name}: "
                f"{result.stderr.strip() or result.stdout.strip() or 'unknown error'}"
            )

    return PdfExportResult(
        generated=generated, skipped=skipped, backend="libreoffice",
    )


# Backwards-compat alias — callers previously used find_libreoffice().
def find_libreoffice() -> str | None:
    return _find_libreoffice()
