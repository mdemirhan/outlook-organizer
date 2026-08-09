from __future__ import annotations

import json
import sqlite3
import uuid
from collections.abc import Iterable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from outlook_organizer.database import SqliteDatabase

AUDIT_SCHEMA = """
CREATE TABLE IF NOT EXISTS audit_runs (
    run_id TEXT PRIMARY KEY,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    config_fingerprint TEXT NOT NULL,
    parameters TEXT NOT NULL,
    status TEXT NOT NULL,
    error TEXT
);
CREATE TABLE IF NOT EXISTS audit_actions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id TEXT NOT NULL,
    sequence INTEGER NOT NULL,
    outlook_id INTEGER NOT NULL,
    message_id TEXT NOT NULL,
    subject TEXT NOT NULL,
    decision TEXT NOT NULL,
    before_state TEXT NOT NULL,
    intended_state TEXT NOT NULL,
    actual_state TEXT,
    status TEXT NOT NULL,
    error TEXT,
    FOREIGN KEY(run_id) REFERENCES audit_runs(run_id)
);
CREATE INDEX IF NOT EXISTS idx_audit_actions_run
    ON audit_actions(run_id, sequence);
"""


class SqliteAuditRepository:
    def __init__(self, path: Path | None = None, *, database: SqliteDatabase | None = None) -> None:
        self.database = database or SqliteDatabase(path)

    @property
    def path(self) -> Path:
        return self.database.path

    def _write_connection(self) -> sqlite3.Connection:
        connection = self.database.connect_for_write()
        connection.executescript(AUDIT_SCHEMA)
        return connection

    def begin_run(self, *, config_fingerprint: str, parameters: dict[str, Any]) -> str:
        run_id = f"run-{uuid.uuid4().hex[:12]}"
        with self._write_connection() as connection:
            connection.execute(
                """
                INSERT INTO audit_runs
                    (run_id, started_at, config_fingerprint, parameters, status)
                VALUES (?, ?, ?, ?, 'running')
                """,
                (
                    run_id,
                    datetime.now(UTC).isoformat(),
                    config_fingerprint,
                    json.dumps(parameters, ensure_ascii=False),
                ),
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
        decision: dict[str, Any],
        before_state: dict[str, Any],
        intended_state: dict[str, Any],
        actual_state: dict[str, Any] | None,
        status: str,
        error: str | None = None,
    ) -> None:
        with self._write_connection() as connection:
            connection.execute(
                """
                INSERT INTO audit_actions
                    (run_id, sequence, outlook_id, message_id, subject, decision,
                     before_state, intended_state, actual_state, status, error)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    sequence,
                    outlook_id,
                    message_id,
                    subject,
                    json.dumps(decision, ensure_ascii=False),
                    json.dumps(before_state, ensure_ascii=False),
                    json.dumps(intended_state, ensure_ascii=False),
                    json.dumps(actual_state, ensure_ascii=False)
                    if actual_state is not None
                    else None,
                    status,
                    error,
                ),
            )

    def finish_run(self, run_id: str, status: str, error: str | None = None) -> None:
        with self._write_connection() as connection:
            connection.execute(
                """
                UPDATE audit_runs SET completed_at = ?, status = ?, error = ?
                WHERE run_id = ?
                """,
                (datetime.now(UTC).isoformat(), status, error, run_id),
            )

    def list_runs(self, limit: int = 20) -> list[dict[str, Any]]:
        connection = self.database.connect_existing()
        if connection is None:
            return []
        try:
            rows = connection.execute(
                """
                SELECT run_id, started_at, completed_at, config_fingerprint,
                       parameters, status, error
                FROM audit_runs ORDER BY started_at DESC LIMIT ?
                """,
                (limit,),
            ).fetchall()
        except sqlite3.OperationalError:
            return []
        finally:
            connection.close()
        return [
            {
                **dict(row),
                "parameters": json.loads(row["parameters"]),
            }
            for row in rows
        ]

    def run_actions(self, run_id: str, *, reverse: bool = False) -> list[sqlite3.Row]:
        connection = self.database.connect_existing()
        if connection is None:
            return []
        order = "DESC" if reverse else "ASC"
        try:
            return list(
                connection.execute(
                    f"SELECT * FROM audit_actions WHERE run_id = ? ORDER BY sequence {order}",
                    (run_id,),
                )
            )
        except sqlite3.OperationalError:
            return []
        finally:
            connection.close()

    def mark_run_undone(self, run_id: str, action_ids: Iterable[int]) -> None:
        ids = list(action_ids)
        with self._write_connection() as connection:
            if ids:
                placeholders = ",".join("?" for _ in ids)
                connection.execute(
                    f"UPDATE audit_actions SET status = 'undone' WHERE id IN ({placeholders})",
                    ids,
                )
            connection.execute(
                """
                UPDATE audit_runs
                SET completed_at = ?, status = 'undone', error = NULL
                WHERE run_id = ?
                """,
                (datetime.now(UTC).isoformat(), run_id),
            )
