from __future__ import annotations

import sqlite3

import pytest
from fakes import FakeMailGateway

from outlook_organizer.audit import HistoryService, SqliteAuditRepository
from outlook_organizer.database import SqliteDatabase
from outlook_organizer.triage.apply import ApplyTriageService
from outlook_organizer.triage.thread_index import (
    SqliteThreadIndexRepository,
    ThreadAffinityResolver,
)


def services(tmp_path, triage_context, message):
    database = SqliteDatabase(tmp_path / "state.sqlite")
    gateway = FakeMailGateway([message])
    thread_repository = SqliteThreadIndexRepository(database=database)
    resolver = ThreadAffinityResolver(triage_context, thread_repository)
    audit = SqliteAuditRepository(database=database)
    apply = ApplyTriageService(triage_context, gateway, gateway, audit, resolver)
    return database, gateway, audit, apply


def test_confirmed_change_creates_audit_without_plans(
    tmp_path, triage_context, direct_message
) -> None:
    database, gateway, audit, service = services(tmp_path, triage_context, direct_message)

    report = service.apply(limit=1)

    assert report["execution"]["status"] == "completed"
    assert gateway.applied[0]["target_folder_id"] == 110
    assert len(audit.list_runs()) == 1
    with sqlite3.connect(database.path) as connection:
        tables = {
            row[0]
            for row in connection.execute("SELECT name FROM sqlite_master WHERE type = 'table'")
        }
    assert "audit_runs" in tables
    assert "audit_actions" in tables
    assert "plans" not in tables


def test_noop_apply_does_not_create_database(tmp_path, triage_context, direct_message) -> None:
    direct_message.folder_id = 110
    direct_message.folder_name = "Internal General"
    direct_message.categories = ["@Internal General", "@Only Me"]
    database, _, _, service = services(tmp_path, triage_context, direct_message)

    report = service.apply(limit=1)

    assert report["execution"]["status"] == "no_changes"
    assert not database.path.exists()


def test_undo_uses_audit_actions_only(tmp_path, triage_context, direct_message) -> None:
    database, gateway, audit, service = services(tmp_path, triage_context, direct_message)
    report = service.apply(limit=1)

    thread_repository = SqliteThreadIndexRepository(database=database)

    result = HistoryService(
        audit,
        gateway,
        gateway,
        lambda outlook_ids: thread_repository.forget_members(
            scope="inbox:101", outlook_ids=outlook_ids
        ),
    ).undo(report["execution"]["run_id"], confirm=True)

    assert result["status"] == "undone"
    assert gateway.applied[-1]["target_folder_id"] == 101


def test_apply_stops_if_a_message_disappears_before_verification(
    tmp_path, triage_context, direct_message
) -> None:
    database, gateway, audit, service = services(tmp_path, triage_context, direct_message)
    gateway.get_messages = lambda outlook_ids, body_limit=0: []

    with pytest.raises(RuntimeError, match="messages disappeared"):
        service.apply(limit=1)

    assert gateway.applied == []
    assert audit.list_runs() == []
    assert not database.path.exists()
