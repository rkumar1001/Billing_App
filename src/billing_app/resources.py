from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from .app_paths import config_path, default_output_root


@dataclass
class AppConfig:
    theme: str = "System"
    color_theme: str = "blue"
    last_source_folder: str = ""
    default_output_root: str = ""
    window_geometry: str = "1200x760"
    date_format: str = "%B %d, %Y"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AppConfig":
        allowed = {f for f in cls.__dataclass_fields__}
        return cls(**{k: v for k, v in data.items() if k in allowed})


def load_config() -> AppConfig:
    path = config_path()
    if not path.exists():
        cfg = AppConfig(default_output_root=str(default_output_root()))
        save_config(cfg)
        return cfg
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return AppConfig(default_output_root=str(default_output_root()))
    cfg = AppConfig.from_dict(data)
    if not cfg.default_output_root:
        cfg.default_output_root = str(default_output_root())
    return cfg


def save_config(cfg: AppConfig) -> None:
    path = config_path()
    path.write_text(json.dumps(cfg.to_dict(), indent=2), encoding="utf-8")
