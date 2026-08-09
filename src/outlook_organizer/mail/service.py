from __future__ import annotations

from dataclasses import asdict
from typing import Any

from outlook_organizer.mail.config import MailContext
from outlook_organizer.mail.models import MailMessage
from outlook_organizer.mail.ports import MailReader, MailWriter


class MailReadService:
    def __init__(self, context: MailContext, reader: MailReader) -> None:
        self.context = context
        self.reader = reader

    def folders(self) -> list[dict[str, Any]]:
        return [
            asdict(folder) for folder in self.reader.list_folders(self.context.folders.scan_limit)
        ]

    def configured_folder_status(self) -> dict[str, Any]:
        actual_by_id = {
            folder.outlook_id: folder
            for folder in self.reader.list_folders(self.context.folders.scan_limit)
        }
        result: dict[str, dict[str, Any]] = {}
        for key, expected in self.context.folders.folders.items():
            actual = actual_by_id.get(expected.id)
            parent = self.context.folders.folders[expected.parent].name if expected.parent else None
            valid = bool(
                actual
                and actual.name.casefold() in {name.casefold() for name in expected.names}
                and (parent is None or actual.parent_name.casefold() == parent.casefold())
            )
            result[key] = {
                "valid": valid,
                "expected": {"id": expected.id, "name": expected.name, "parent": parent},
                "actual": (
                    {
                        "id": actual.outlook_id,
                        "name": actual.name,
                        "parent": actual.parent_name,
                    }
                    if actual
                    else None
                ),
            }
        return {"valid": all(item["valid"] for item in result.values()), "folders": result}

    def get_message(self, outlook_id: int, include_body: bool = False) -> dict[str, Any]:
        message = self.reader.get_message(outlook_id, body_limit=20_000 if include_body else 0)
        return self._message_dict(message, include_body=include_body)

    def search_messages(
        self,
        query: str,
        *,
        limit: int = 20,
        scan_limit: int = 250,
        include_body: bool = False,
    ) -> list[dict[str, Any]]:
        inbox = self.context.folders.folders["inbox"]
        messages = self.reader.latest_messages(
            inbox.id, limit=scan_limit, body_limit=5000 if query or include_body else 0
        )
        needle = query.casefold()
        matches = [
            message
            for message in messages
            if needle in message.subject.casefold()
            or needle in message.sender_name.casefold()
            or needle in message.sender_address.casefold()
            or needle in message.body.casefold()
        ]
        return [
            self._message_dict(message, include_body=include_body) for message in matches[:limit]
        ]

    @staticmethod
    def _message_dict(message: MailMessage, *, include_body: bool) -> dict[str, Any]:
        result = asdict(message)
        result["flag_status"] = message.flag_status.value
        result["stable_id"] = message.stable_id
        if not include_body:
            result.pop("body", None)
        return result


class FolderAdminService:
    def __init__(self, context: MailContext, writer: MailWriter) -> None:
        self.context = context
        self.writer = writer

    def setup(self, *, confirm: bool) -> dict[str, Any]:
        if not confirm:
            raise ValueError("confirm=true is required to create Outlook folders")
        inbox = self.context.folders.folders["inbox"]
        roots = ("organized_primary", "organized_secondary")
        return {
            "folders": {
                key: self.writer.ensure_mail_folder(
                    inbox.id, self.context.folders.folders[key].name
                )
                for key in roots
            }
        }
