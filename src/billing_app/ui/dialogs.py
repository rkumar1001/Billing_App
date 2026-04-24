from __future__ import annotations

import platform
import shutil
import subprocess
import time
from pathlib import Path
from tkinter import filedialog


def ask_directory(parent, initialdir: str | None = None, title: str = "Choose directory") -> str:
    """Open a folder chooser with deterministic Linux-native preference."""
    normalized = _normalize_initialdir(initialdir)

    # Prefer native Linux pickers first because Tk's askdirectory can be flaky
    # on some desktop/session combinations.
    if platform.system() == "Linux":
        folder = _ask_directory_zenity(title=title, initialdir=normalized)
        if folder:
            return folder
        folder = _ask_directory_kdialog(title=title, initialdir=normalized)
        if folder:
            return folder

    return filedialog.askdirectory(
        parent=parent,
        initialdir=normalized,
        title=title,
        mustexist=True,
    )


def _normalize_initialdir(initialdir: str | None) -> str:
    if not initialdir:
        return str(Path.home())
    p = Path(initialdir).expanduser()
    return str(p if p.is_dir() else Path.home())


def _ask_directory_zenity(title: str, initialdir: str) -> str:
    if not shutil.which("zenity"):
        return ""
    cmd = [
        "zenity",
        "--file-selection",
        "--directory",
        "--title",
        title,
        "--filename",
        f"{initialdir.rstrip('/')}/",
    ]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, check=False)
    except OSError:
        return ""
    if res.returncode != 0:
        return ""
    return (res.stdout or "").strip()


def _ask_directory_kdialog(title: str, initialdir: str) -> str:
    if not shutil.which("kdialog"):
        return ""
    cmd = ["kdialog", "--getexistingdirectory", initialdir, "--title", title]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, check=False)
    except OSError:
        return ""
    if res.returncode != 0:
        return ""
    return (res.stdout or "").strip()
