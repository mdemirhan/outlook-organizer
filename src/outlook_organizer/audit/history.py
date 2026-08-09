from __future__ import annotations

import json
from collections.abc import Callable
from typing import Any

from outlook_organizer.audit.repository import AuditRepository
from outlook_organizer.mail import FlagStatus
from outlook_organizer.mail.ports import MailReader, MailWriter


class HistoryService:
    def __init__(
        self,
        audit: AuditRepository,
        reader: MailReader,
        writer: MailWriter,
        forget_indexed_messages: Callable[[set[int]], None] | None = None,
    ) -> None:
        self.audit = audit
        self.reader = reader
        self.writer = writer
        self.forget_indexed_messages = forget_indexed_messages

    def list_runs(self, limit: int = 20) -> list[dict[str, Any]]:
        return self.audit.list_runs(limit)

    def undo(self, run_id: str, *, confirm: bool) -> dict[str, Any]:
        if not confirm:
            raise ValueError("confirm=true is required for Outlook writes")
        rows = self.audit.run_actions(run_id, reverse=True)
        if not rows:
            raise KeyError(f"Unknown or empty audit run: {run_id}")
        restored = 0
        action_ids: list[int] = []
        try:
            for row in rows:
                if row["status"] not in {"applied", "failed"}:
                    continue
                before = json.loads(row["before_state"])
                current = self.reader.get_message(int(row["outlook_id"]), body_limit=0)
                original_folder = int(before["folder_id"])
                self.writer.apply_mail_state(
                    int(row["outlook_id"]),
                    categories=list(before["categories"]),
                    flag_status=FlagStatus(before["flag_status"]),
                    target_folder_id=(
                        original_folder if current.folder_id != original_folder else None
                    ),
                )
                restored += 1
                action_ids.append(int(row["id"]))
        except Exception as exc:
            return {
                "run_id": run_id,
                "applied": restored,
                "status": "undo_failed",
                "error": str(exc),
            }
        if self.forget_indexed_messages is not None:
            self.forget_indexed_messages(
                {int(row["outlook_id"]) for row in rows if row["status"] == "applied"}
            )
        self.audit.mark_run_undone(run_id, action_ids)
        return {
            "run_id": run_id,
            "applied": restored,
            "status": "undone",
            "error": None,
        }
