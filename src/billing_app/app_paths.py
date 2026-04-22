from __future__ import annotations

import os
import sys
from pathlib import Path

from platformdirs import PlatformDirs

APP_NAME = "BillingApp"
APP_AUTHOR = "BillingApp"

_dirs = PlatformDirs(APP_NAME, APP_AUTHOR)


def user_data_dir() -> Path:
    path = Path(_dirs.user_data_dir)
    path.mkdir(parents=True, exist_ok=True)
    return path


def user_config_dir() -> Path:
    path = Path(_dirs.user_config_dir)
    path.mkdir(parents=True, exist_ok=True)
    return path


def db_path() -> Path:
    return user_data_dir() / "billing.db"


def config_path() -> Path:
    return user_config_dir() / "config.json"


def default_output_root() -> Path:
    path = user_data_dir() / "Invoices"
    path.mkdir(parents=True, exist_ok=True)
    return path


def bundle_root() -> Path:
    # When frozen by PyInstaller, sys._MEIPASS points at the extracted bundle.
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parents[2]


def asset_path(name: str) -> Path:
    return bundle_root() / "assets" / name
