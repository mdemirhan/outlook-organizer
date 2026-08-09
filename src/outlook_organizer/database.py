from __future__ import annotations

import sqlite3
from pathlib import Path

from outlook_organizer.paths import state_dir


class SqliteDatabase:
    """Side-effect-free SQLite connection factory."""

    def __init__(self, path: Path | None = None) -> None:
        self.path = path or (state_dir() / "outlook-organizer.sqlite")

    @property
    def exists(self) -> bool:
        return self.path.is_file()

    def connect_existing(self) -> sqlite3.Connection | None:
        if not self.exists:
            return None
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        return connection

    def connect_for_write(self) -> sqlite3.Connection:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        return connection
