"""Lossless .xls -> .xlsx conversion.

Backends, tried in order:

1. **Microsoft Excel via COM** (Windows + Excel installed) — exact fidelity,
   uses pywin32. The client's Windows machine almost certainly already has
   Excel, so this is the zero-install path.
2. **LibreOffice headless** (any platform) — high fidelity for spreadsheet
   content (formulas, cell formatting, merges, column widths). Requires
   LibreOffice to be installed.

If neither is available we raise XlsConvertError with a clear instruction
so the UI can tell the user how to proceed (open in Excel, Save As .xlsx).
"""
from __future__ import annotations

import platform
import shutil
import subprocess
from pathlib import Path


class XlsConvertError(Exception):
    pass


def find_libreoffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        hit = shutil.which(name)
        if hit:
            return hit
    for candidate in (
        "/snap/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return None


def _have_msexcel_com() -> bool:
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
        import winreg
    except ImportError:
        return False
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
        return True
    except OSError:
        return False


def converter_available() -> str:
    """Return the preferred backend name, or '' if none."""
    if _have_msexcel_com():
        return "msexcel"
    if find_libreoffice():
        return "libreoffice"
    return ""


def install_hint() -> str:
    if platform.system() == "Windows":
        return (
            "To convert legacy .xls files automatically, install Microsoft "
            "Excel (or LibreOffice). Alternatively, open the .xls file in "
            "Excel and use File → Save As → Excel Workbook (.xlsx)."
        )
    return (
        "To convert legacy .xls files automatically, install LibreOffice "
        "from https://www.libreoffice.org/download/. Alternatively, open "
        "the .xls in any spreadsheet app and Save As .xlsx."
    )


def convert_xls_to_xlsx(src: Path, dst: Path) -> str:
    """Convert `src` (.xls) into `dst` (.xlsx) preserving formatting/formulas.

    Returns the backend name used (`"msexcel"` or `"libreoffice"`).
    Raises XlsConvertError if no backend is available or conversion fails.
    """
    src = Path(src)
    dst = Path(dst)
    if not src.exists():
        raise XlsConvertError(f"Source .xls not found: {src}")

    backend = converter_available()
    if not backend:
        raise XlsConvertError(install_hint())

    dst.parent.mkdir(parents=True, exist_ok=True)
    if backend == "msexcel":
        _convert_via_msexcel(src, dst)
    else:
        _convert_via_libreoffice(find_libreoffice(), src, dst)
    if not dst.exists():
        raise XlsConvertError(
            f"Conversion finished but no .xlsx was produced at {dst}."
        )
    return backend


def _convert_via_msexcel(src: Path, dst: Path) -> None:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(src.resolve()), ReadOnly=True)
        try:
            # xlOpenXMLWorkbook = 51 → produces .xlsx
            wb.SaveAs(str(dst.resolve()), FileFormat=51)
        finally:
            wb.Close(SaveChanges=False)
    except Exception as e:  # noqa: BLE001
        raise XlsConvertError(
            f"MS Excel could not convert {src.name}: {type(e).__name__}: {e}"
        ) from e
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:  # noqa: BLE001
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass


def _convert_via_libreoffice(soffice: str, src: Path, dst: Path) -> None:
    """Drive LibreOffice with a *unique* user profile each call so it works
    even when the user already has LibreOffice open with another document
    (the default profile is single-instance and would silently no-op)."""
    import tempfile
    out_dir = dst.parent
    profile_dir = tempfile.mkdtemp(prefix="billingapp-soffice-", dir=str(Path.home()))
    profile_url = "file://" + profile_dir
    try:
        result = subprocess.run(
            [
                soffice,
                f"-env:UserInstallation={profile_url}",
                "--headless",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(out_dir.resolve()),
                str(src.resolve()),
            ],
            capture_output=True,
            text=True,
            timeout=180,
        )
        if result.returncode != 0:
            raise XlsConvertError(
                f"LibreOffice failed to convert {src.name}: "
                f"{(result.stderr or result.stdout).strip() or 'unknown error'}"
            )
        produced = out_dir / (src.stem + ".xlsx")
        if not produced.exists():
            raise XlsConvertError(
                f"LibreOffice exit ok but no .xlsx produced. "
                f"stdout: {result.stdout.strip()!r} "
                f"stderr: {result.stderr.strip()!r}"
            )
        if produced != dst:
            if dst.exists():
                dst.unlink()
            produced.rename(dst)
    finally:
        try:
            shutil.rmtree(profile_dir, ignore_errors=True)
        except OSError:
            pass
