from __future__ import annotations

import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CONFIG_DIR = Path.home() / ".config" / "outlook-organizer"


def config_dir() -> Path:
    value = os.environ.get("OUTLOOK_ORGANIZER_CONFIG") or os.environ.get(
        "OUTLOOK_DISTILLER_CONFIG"
    )
    return Path(value).expanduser().resolve() if value else DEFAULT_CONFIG_DIR


def state_dir() -> Path:
    value = os.environ.get("OUTLOOK_ORGANIZER_STATE") or os.environ.get(
        "OUTLOOK_DISTILLER_STATE"
    )
    path = Path(value).expanduser().resolve() if value else PROJECT_ROOT / "state"
    path.mkdir(parents=True, exist_ok=True)
    return path
