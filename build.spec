# PyInstaller spec — cross-platform one-file build.
# Usage:  pyinstaller --clean --noconfirm build.spec

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(os.path.abspath(SPECPATH))  # type: ignore[name-defined]

block_cipher = None

entry = str(ROOT / "src" / "billing_app" / "main.py")

datas = []
asset_dir = ROOT / "assets"
if asset_dir.exists():
    for item in asset_dir.iterdir():
        datas.append((str(item), "assets"))

hidden_imports = [
    "babel.numbers",  # tkcalendar depends on this
    "customtkinter",
    "openpyxl",
    "docx",
    "xlrd",  # legacy .xls reader
    "sqlalchemy.dialects.sqlite",
]

excludes = ["matplotlib", "numpy", "scipy", "PyQt5", "PyQt6", "PySide2", "PySide6"]

if sys.platform == "win32":
    icon_candidate = asset_dir / "icon.ico"
elif sys.platform == "darwin":
    icon_candidate = asset_dir / "icon.icns"
else:
    icon_candidate = asset_dir / "icon.png"
icon_arg = str(icon_candidate) if icon_candidate.exists() else None

a = Analysis(  # noqa: F821
    [entry],
    pathex=[str(ROOT / "src")],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)  # noqa: F821

exe = EXE(  # noqa: F821
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="BillingApp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_arg,
)

if sys.platform == "darwin":
    app = BUNDLE(  # noqa: F821
        exe,
        name="BillingApp.app",
        icon=icon_arg,
        bundle_identifier="com.billingapp.invoice",
    )
