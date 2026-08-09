from __future__ import annotations

from pathlib import Path

from outlook_organizer.calendar.models import CalendarConfig
from outlook_organizer.paths import config_dir
from outlook_organizer.yaml_config import load_yaml_model


def load_calendar_config(directory: Path | None = None) -> CalendarConfig:
    root = (directory or config_dir()).resolve()
    return load_yaml_model(root / "calendar.yaml", CalendarConfig)
