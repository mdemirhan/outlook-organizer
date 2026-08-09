from __future__ import annotations

import sqlite3
from collections.abc import Iterable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Protocol

from outlook_organizer.database import SqliteDatabase
from outlook_organizer.mail import DomainClass, MailMessage
from outlook_organizer.triage.config import TriageContext
from outlook_organizer.triage.models import RuleMatch, ThreadResolution, TriageDecision

THREAD_SCHEMA = """
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
"""


class ThreadIndexRepository(Protocol):
    def contexts(self, scope: str, thread_guids: Iterable[str]) -> dict[str, dict[str, Any]]: ...

    def update(
        self,
        *,
        scope: str,
        routes: dict[str, str],
        members: list[dict[str, Any]],
    ) -> None: ...

    def forget_members(self, *, scope: str, outlook_ids: Iterable[int]) -> None: ...


class SqliteThreadIndexRepository:
    """Thread cache whose reads never create or mutate the database."""

    def __init__(self, path: Path | None = None, *, database: SqliteDatabase | None = None) -> None:
        self.database = database or SqliteDatabase(path)

    def contexts(self, scope: str, thread_guids: Iterable[str]) -> dict[str, dict[str, Any]]:
        guids = sorted({value for value in thread_guids if value})
        if not guids:
            return {}
        connection = self.database.connect_existing()
        if connection is None:
            return {}
        placeholders = ",".join("?" for _ in guids)
        try:
            routes = connection.execute(
                f"""
                SELECT thread_guid, folder_key FROM thread_routes
                WHERE scope = ? AND thread_guid IN ({placeholders})
                """,
                [scope, *guids],
            ).fetchall()
            members = connection.execute(
                f"""
                SELECT thread_guid, outlook_id, message_id, last_folder_id,
                       last_folder_key, detached
                FROM thread_members
                WHERE scope = ? AND thread_guid IN ({placeholders})
                """,
                [scope, *guids],
            ).fetchall()
        except sqlite3.OperationalError:
            return {}
        finally:
            connection.close()
        contexts = {
            str(row["thread_guid"]): {
                "thread_guid": str(row["thread_guid"]),
                "folder_key": str(row["folder_key"]),
                "members": [],
            }
            for row in routes
        }
        for row in members:
            context = contexts.get(str(row["thread_guid"]))
            if context is not None:
                context["members"].append(
                    {
                        "outlook_id": int(row["outlook_id"]),
                        "message_id": str(row["message_id"]),
                        "folder_id": int(row["last_folder_id"]),
                        "folder_key": row["last_folder_key"],
                        "detached": bool(row["detached"]),
                    }
                )
        return contexts

    def update(
        self,
        *,
        scope: str,
        routes: dict[str, str],
        members: list[dict[str, Any]],
    ) -> None:
        if not routes and not members:
            return
        now = datetime.now(UTC).isoformat()
        with self.database.connect_for_write() as connection:
            connection.executescript(THREAD_SCHEMA)
            for thread_guid, folder_key in routes.items():
                connection.execute(
                    """
                    INSERT INTO thread_routes (scope, thread_guid, folder_key, updated_at)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(scope, thread_guid) DO UPDATE SET
                        folder_key = excluded.folder_key,
                        updated_at = excluded.updated_at
                    """,
                    (scope, thread_guid, folder_key, now),
                )
            for member in members:
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
                        member["thread_guid"],
                        member["outlook_id"],
                        member["message_id"],
                        member["folder_id"],
                        member.get("folder_key"),
                        int(bool(member.get("detached", False))),
                        now,
                    ),
                )

    def forget_members(self, *, scope: str, outlook_ids: Iterable[int]) -> None:
        ids = sorted({int(value) for value in outlook_ids})
        if not ids:
            return
        connection = self.database.connect_existing()
        if connection is None:
            return
        placeholders = ",".join("?" for _ in ids)
        try:
            connection.execute(
                f"DELETE FROM thread_members WHERE scope = ? AND outlook_id IN ({placeholders})",
                [scope, *ids],
            )
            connection.execute(
                """
                DELETE FROM thread_routes
                WHERE scope = ?
                  AND NOT EXISTS (
                      SELECT 1 FROM thread_members
                      WHERE thread_members.scope = thread_routes.scope
                        AND thread_members.thread_guid = thread_routes.thread_guid
                  )
                """,
                (scope,),
            )
            connection.commit()
        except sqlite3.OperationalError:
            connection.rollback()
        finally:
            connection.close()


class ThreadAffinityResolver:
    def __init__(self, context: TriageContext, repository: ThreadIndexRepository) -> None:
        self.context = context
        self.repository = repository
        self.enabled = context.config.threading.enabled
        self.scope = f"inbox:{context.mail.folders.folders['inbox'].id}"
        self.route_priority = {
            route.move_to: priority for priority, route in enumerate(context.config.routes)
        }
        self.managed = set(self.route_priority)
        self.safety = {
            route.move_to
            for route in context.config.routes
            if route.when.sender_type in {"junk_external", "unclassified_external"}
        }
        self.affinity = self.managed - self.safety

    def resolve(
        self,
        messages: list[MailMessage],
        decisions: list[TriageDecision],
    ) -> ThreadResolution:
        if not self.enabled:
            return ThreadResolution(decisions)
        contexts = self.repository.contexts(
            self.scope, (message.thread_guid for message in messages)
        )
        grouped: dict[str, list[TriageDecision]] = {}
        for message, decision in zip(messages, decisions, strict=True):
            if message.thread_guid:
                grouped.setdefault(message.thread_guid, []).append(decision)
        routes: dict[str, str] = {}
        for thread_guid, group in grouped.items():
            candidates = {
                decision.move_to
                for decision in group
                if self.uses_affinity(decision) and decision.move_to
            }
            context = contexts.get(thread_guid)
            if context and context["folder_key"] in self.affinity:
                candidates.add(str(context["folder_key"]))
            if not candidates:
                continue
            destination = min(
                candidates,
                key=lambda key: (self.route_priority.get(key, 1_000_000), key),
            )
            routes[thread_guid] = destination
            for decision in group:
                if not self.uses_affinity(decision) or decision.move_to == destination:
                    continue
                decision.move_to = destination
                decision.report_section = self.context.mail.folders.folders[destination].name
                decision.matches.append(
                    RuleMatch(
                        "thread-affinity",
                        "Conversation inherited its indexed destination",
                        -1,
                        [f"thread destination is {destination}"],
                    )
                )
        return ThreadResolution(decisions, canonical_routes=routes)

    def persist_successes(
        self,
        *,
        messages: list[MailMessage],
        decisions_by_id: dict[int, TriageDecision],
        successful_ids: set[int],
        canonical_routes: dict[str, str],
    ) -> None:
        if not self.enabled or not successful_ids:
            return
        routes: dict[str, str] = {}
        members: list[dict[str, Any]] = []
        for message in messages:
            if message.outlook_id not in successful_ids or not message.thread_guid:
                continue
            decision = decisions_by_id[message.outlook_id]
            folder_key = decision.move_to if self.uses_affinity(decision) else None
            if folder_key is None:
                continue
            routes[message.thread_guid] = canonical_routes.get(message.thread_guid, folder_key)
            members.append(
                {
                    "thread_guid": message.thread_guid,
                    "outlook_id": message.outlook_id,
                    "message_id": message.stable_id,
                    "folder_id": self.context.mail.folders.folders[folder_key].id,
                    "folder_key": folder_key,
                    "detached": False,
                }
            )
        self.repository.update(scope=self.scope, routes=routes, members=members)

    def uses_affinity(self, decision: TriageDecision) -> bool:
        return bool(
            decision.move_to in self.affinity
            and not decision.keep_in_inbox
            and decision.domain_class
            not in {DomainClass.JUNK_EXTERNAL, DomainClass.UNCLASSIFIED_EXTERNAL}
        )
