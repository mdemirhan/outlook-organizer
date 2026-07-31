from __future__ import annotations

import json
import sqlite3
import uuid
from collections.abc import Iterable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from outlook_organizer.models import TriagePlan
from outlook_organizer.paths import state_dir
from outlook_organizer.serialization import plan_from_dict, plan_to_dict

SCHEMA = """
CREATE TABLE IF NOT EXISTS plans (
    plan_id TEXT PRIMARY KEY,
    created_at TEXT NOT NULL,
    config_fingerprint TEXT NOT NULL,
    status TEXT NOT NULL,
    payload TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS runs (
    run_id TEXT PRIMARY KEY,
    plan_id TEXT NOT NULL,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    status TEXT NOT NULL,
    error TEXT,
    FOREIGN KEY(plan_id) REFERENCES plans(plan_id)
);
CREATE TABLE IF NOT EXISTS run_actions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id TEXT NOT NULL,
    sequence INTEGER NOT NULL,
    outlook_id INTEGER NOT NULL,
    message_id TEXT NOT NULL,
    subject TEXT NOT NULL,
    status TEXT NOT NULL,
    before_state TEXT NOT NULL,
    after_state TEXT NOT NULL,
    error TEXT,
    FOREIGN KEY(run_id) REFERENCES runs(run_id)
);
CREATE INDEX IF NOT EXISTS idx_run_actions_run ON run_actions(run_id, sequence);
"""


class StateStore:
    def __init__(self, path: Path | None = None) -> None:
        if path is None:
            directory = state_dir()
            organizer_path = directory / "outlook-organizer.sqlite"
            legacy_path = directory / "outlook-distiller.sqlite"
            if not organizer_path.exists() and legacy_path.exists():
                legacy_path.replace(organizer_path)
            self.path = organizer_path
        else:
            self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as connection:
            connection.executescript(SCHEMA)

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        return connection

    def save_plan(self, plan: TriagePlan) -> None:
        payload = json.dumps(plan_to_dict(plan), ensure_ascii=False)
        with self.connect() as connection:
            connection.execute(
                """
                INSERT OR REPLACE INTO plans
                    (plan_id, created_at, config_fingerprint, status, payload)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    plan.plan_id,
                    plan.created_at.isoformat(),
                    plan.config_fingerprint,
                    "previewed",
                    payload,
                ),
            )

    def load_plan(self, plan_id: str) -> TriagePlan:
        with self.connect() as connection:
            row = connection.execute(
                "SELECT payload FROM plans WHERE plan_id = ?", (plan_id,)
            ).fetchone()
        if row is None:
            raise KeyError(f"Unknown plan: {plan_id}")
        return plan_from_dict(json.loads(row["payload"]))

    def plan_status(self, plan_id: str) -> str:
        with self.connect() as connection:
            row = connection.execute(
                "SELECT status FROM plans WHERE plan_id = ?", (plan_id,)
            ).fetchone()
        if row is None:
            raise KeyError(f"Unknown plan: {plan_id}")
        return str(row["status"])

    def start_run(self, plan_id: str) -> str:
        run_id = f"run-{uuid.uuid4().hex[:12]}"
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO runs (run_id, plan_id, started_at, status)
                VALUES (?, ?, ?, ?)
                """,
                (run_id, plan_id, datetime.now(UTC).isoformat(), "running"),
            )
            connection.execute(
                "UPDATE plans SET status = ? WHERE plan_id = ?", ("applying", plan_id)
            )
        return run_id

    def record_action(
        self,
        *,
        run_id: str,
        sequence: int,
        outlook_id: int,
        message_id: str,
        subject: str,
        status: str,
        before_state: dict[str, Any],
        after_state: dict[str, Any],
        error: str | None = None,
    ) -> None:
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO run_actions
                    (run_id, sequence, outlook_id, message_id, subject, status,
                     before_state, after_state, error)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    sequence,
                    outlook_id,
                    message_id,
                    subject,
                    status,
                    json.dumps(before_state, ensure_ascii=False),
                    json.dumps(after_state, ensure_ascii=False),
                    error,
                ),
            )

    def finish_run(self, run_id: str, status: str, error: str | None = None) -> None:
        with self.connect() as connection:
            row = connection.execute(
                "SELECT plan_id FROM runs WHERE run_id = ?", (run_id,)
            ).fetchone()
            connection.execute(
                """
                UPDATE runs
                SET completed_at = ?, status = ?, error = ?
                WHERE run_id = ?
                """,
                (datetime.now(UTC).isoformat(), status, error, run_id),
            )
            if row:
                connection.execute(
                    "UPDATE plans SET status = ? WHERE plan_id = ?",
                    ("applied" if status == "completed" else status, row["plan_id"]),
                )

    def run_actions(self, run_id: str, *, reverse: bool = False) -> list[sqlite3.Row]:
        order = "DESC" if reverse else "ASC"
        with self.connect() as connection:
            return list(
                connection.execute(
                    f"SELECT * FROM run_actions WHERE run_id = ? ORDER BY sequence {order}",
                    (run_id,),
                )
            )

    def list_runs(self, limit: int = 20) -> list[dict[str, Any]]:
        with self.connect() as connection:
            rows = connection.execute(
                """
                SELECT run_id, plan_id, started_at, completed_at, status, error
                FROM runs ORDER BY started_at DESC LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [dict(row) for row in rows]

    def mark_run_undone(self, run_id: str, action_ids: Iterable[int]) -> None:
        ids = list(action_ids)
        with self.connect() as connection:
            if ids:
                placeholders = ",".join("?" for _ in ids)
                connection.execute(
                    f"UPDATE run_actions SET status = 'undone' WHERE id IN ({placeholders})",
                    ids,
                )
            row = connection.execute(
                "SELECT plan_id FROM runs WHERE run_id = ?", (run_id,)
            ).fetchone()
            connection.execute(
                """
                UPDATE runs
                SET completed_at = ?, status = 'undone', error = NULL
                WHERE run_id = ?
                """,
                (datetime.now(UTC).isoformat(), run_id),
            )
            if row:
                connection.execute(
                    "UPDATE plans SET status = 'undone' WHERE plan_id = ?",
                    (row["plan_id"],),
                )
