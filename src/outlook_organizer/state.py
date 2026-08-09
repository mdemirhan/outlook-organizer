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
CREATE TABLE IF NOT EXISTS thread_routes (
    scope TEXT NOT NULL,
    thread_guid TEXT NOT NULL,
    folder_key TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    PRIMARY KEY(scope, thread_guid)
);
CREATE TABLE IF NOT EXISTS thread_members (
    scope TEXT NOT NULL,
    thread_guid TEXT NOT NULL,
    outlook_id INTEGER NOT NULL,
    message_id TEXT NOT NULL,
    last_folder_id INTEGER NOT NULL,
    last_folder_key TEXT,
    detached INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL,
    PRIMARY KEY(scope, thread_guid, outlook_id)
);
CREATE INDEX IF NOT EXISTS idx_thread_members_lookup
    ON thread_members(scope, thread_guid, detached);
CREATE INDEX IF NOT EXISTS idx_thread_members_outlook
    ON thread_members(scope, outlook_id);
"""


class StateStore:
    def __init__(self, path: Path | None = None) -> None:
        if path is None:
            directory = state_dir()
            self.path = directory / "outlook-organizer.sqlite"
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

    def thread_index_status(
        self,
        scope: str,
    ) -> dict[str, Any]:
        with self.connect() as connection:
            route_count = connection.execute(
                "SELECT count(*) FROM thread_routes WHERE scope = ?",
                (scope,),
            ).fetchone()[0]
            member_count = connection.execute(
                "SELECT count(*) FROM thread_members WHERE scope = ?",
                (scope,),
            ).fetchone()[0]
        return {
            "threads": int(route_count),
            "members": int(member_count),
        }

    def thread_contexts(
        self,
        scope: str,
        thread_guids: Iterable[str],
    ) -> dict[str, dict[str, Any]]:
        guids = sorted({value for value in thread_guids if value})
        if not guids:
            return {}
        placeholders = ",".join("?" for _ in guids)
        parameters = [scope, *guids]
        with self.connect() as connection:
            route_rows = connection.execute(
                f"""
                SELECT thread_guid, folder_key
                FROM thread_routes
                WHERE scope = ? AND thread_guid IN ({placeholders})
                """,
                parameters,
            ).fetchall()
            member_rows = connection.execute(
                f"""
                SELECT thread_guid, outlook_id, message_id, last_folder_id,
                       last_folder_key, detached
                FROM thread_members
                WHERE scope = ? AND thread_guid IN ({placeholders})
                ORDER BY outlook_id
                """,
                parameters,
            ).fetchall()
        contexts = {
            str(row["thread_guid"]): {
                "thread_guid": str(row["thread_guid"]),
                "folder_key": str(row["folder_key"]),
                "members": [],
            }
            for row in route_rows
        }
        for row in member_rows:
            context = contexts.get(str(row["thread_guid"]))
            if context is None:
                continue
            context["members"].append(
                {
                    "thread_guid": str(row["thread_guid"]),
                    "outlook_id": int(row["outlook_id"]),
                    "message_id": str(row["message_id"]),
                    "folder_id": int(row["last_folder_id"]),
                    "folder_key": row["last_folder_key"],
                    "detached": bool(row["detached"]),
                }
            )
        return contexts

    def update_thread_index(
        self,
        *,
        scope: str,
        routes: dict[str, str] | None = None,
        members: list[dict[str, Any]] | None = None,
    ) -> None:
        now = datetime.now(UTC).isoformat()
        routes = routes or {}
        members = members or []
        with self.connect() as connection:
            for thread_guid, folder_key in routes.items():
                connection.execute(
                    """
                    INSERT INTO thread_routes
                        (scope, thread_guid, folder_key, updated_at)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(scope, thread_guid) DO UPDATE SET
                        folder_key = excluded.folder_key,
                        updated_at = excluded.updated_at
                    """,
                    (scope, thread_guid, folder_key, now),
                )
            for member in members:
                message_id = str(member.get("message_id", ""))
                if message_id:
                    connection.execute(
                        """
                        DELETE FROM thread_members
                        WHERE scope = ? AND message_id = ? AND outlook_id != ?
                        """,
                        (scope, message_id, int(member["outlook_id"])),
                    )
                connection.execute(
                    """
                    INSERT INTO thread_members
                        (scope, thread_guid, outlook_id, message_id, last_folder_id,
                         last_folder_key, detached, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(scope, thread_guid, outlook_id) DO UPDATE SET
                        message_id = excluded.message_id,
                        last_folder_id = excluded.last_folder_id,
                        last_folder_key = excluded.last_folder_key,
                        detached = excluded.detached,
                        updated_at = excluded.updated_at
                    """,
                    (
                        scope,
                        str(member["thread_guid"]),
                        int(member["outlook_id"]),
                        message_id,
                        int(member["folder_id"]),
                        member.get("folder_key"),
                        int(bool(member.get("detached", False))),
                        now,
                    ),
                )

    def delete_thread_members(
        self,
        scope: str,
        outlook_ids: Iterable[int],
    ) -> set[str]:
        ids = sorted({int(value) for value in outlook_ids})
        if not ids:
            return set()
        placeholders = ",".join("?" for _ in ids)
        with self.connect() as connection:
            rows = connection.execute(
                f"""
                SELECT DISTINCT thread_guid FROM thread_members
                WHERE scope = ? AND outlook_id IN ({placeholders})
                """,
                [scope, *ids],
            ).fetchall()
            connection.execute(
                f"""
                DELETE FROM thread_members
                WHERE scope = ? AND outlook_id IN ({placeholders})
                """,
                [scope, *ids],
            )
        return {str(row["thread_guid"]) for row in rows}

    def delete_thread_routes(self, scope: str, thread_guids: Iterable[str]) -> None:
        guids = sorted({value for value in thread_guids if value})
        if not guids:
            return
        placeholders = ",".join("?" for _ in guids)
        with self.connect() as connection:
            connection.execute(
                f"""
                DELETE FROM thread_routes
                WHERE scope = ? AND thread_guid IN ({placeholders})
                """,
                [scope, *guids],
            )
