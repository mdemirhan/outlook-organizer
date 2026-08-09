from __future__ import annotations

from typing import Any

from outlook_organizer.audit import HistoryService, SqliteAuditRepository
from outlook_organizer.brief import load_brief_context
from outlook_organizer.calendar import load_calendar_config
from outlook_organizer.database import SqliteDatabase
from outlook_organizer.mail import load_mail_context
from outlook_organizer.mail.service import FolderAdminService
from outlook_organizer.outlook import OutlookAdapter, OutlookReader
from outlook_organizer.triage import (
    ApplyTriageService,
    TriagePreviewService,
    load_triage_context,
)
from outlook_organizer.triage.thread_index import (
    SqliteThreadIndexRepository,
    ThreadAffinityResolver,
)


def folder_admin_service() -> FolderAdminService:
    return FolderAdminService(load_mail_context(), OutlookAdapter())


def triage_preview_service() -> TriagePreviewService:
    context = load_triage_context()
    database = SqliteDatabase()
    reader = OutlookReader()
    resolver = ThreadAffinityResolver(context, SqliteThreadIndexRepository(database=database))
    return TriagePreviewService(context, reader, resolver)


def triage_apply_service() -> ApplyTriageService:
    context = load_triage_context()
    database = SqliteDatabase()
    outlook = OutlookAdapter()
    resolver = ThreadAffinityResolver(context, SqliteThreadIndexRepository(database=database))
    return ApplyTriageService(
        context,
        outlook,
        outlook,
        SqliteAuditRepository(database=database),
        resolver,
    )


def history_service() -> HistoryService:
    context = load_triage_context()
    database = SqliteDatabase()
    outlook = OutlookAdapter()
    thread_repository = SqliteThreadIndexRepository(database=database)
    scope = f"inbox:{context.mail.folders.folders['inbox'].id}"
    return HistoryService(
        SqliteAuditRepository(database=database),
        outlook,
        outlook,
        lambda outlook_ids: thread_repository.forget_members(scope=scope, outlook_ids=outlook_ids),
    )


def validate_configuration() -> dict[str, Any]:
    mail = load_mail_context()
    brief = load_brief_context()
    triage = load_triage_context()
    calendar = load_calendar_config()
    return {
        "valid": True,
        "mail_fingerprint": mail.fingerprint,
        "triage_fingerprint": triage.fingerprint,
        "brief_fingerprint": brief.fingerprint,
        "brief": {
            "default_profile": brief.config.default_profile,
            "profiles": sorted(brief.config.profiles),
        },
        "threading_enabled": triage.config.threading.enabled,
        "folders": {
            key: {"id": folder.id, "name": folder.name, "parent": folder.parent}
            for key, folder in mail.folders.folders.items()
        },
        "calendar_timezone": calendar.timezone,
    }
