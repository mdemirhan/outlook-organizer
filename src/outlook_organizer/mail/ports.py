from __future__ import annotations

from collections.abc import Iterable
from typing import Any, Protocol

from outlook_organizer.mail.models import (
    FlagStatus,
    MailMessage,
    OutlookFolder,
)


class MailReader(Protocol):
    def list_folders(
        self, maximum_id: int = 5000, *, refresh: bool = False
    ) -> list[OutlookFolder]: ...

    def latest_messages(
        self, folder_id: int, limit: int = 20, body_limit: int = 2000
    ) -> list[MailMessage]: ...

    def get_message(self, outlook_id: int, body_limit: int = 20_000) -> MailMessage: ...

    def get_messages(
        self, outlook_ids: Iterable[int], *, body_limit: int = 0
    ) -> list[MailMessage]: ...

    def messages_in_window(
        self,
        folder_id: int,
        *,
        start_offset_seconds: int,
        end_offset_seconds: int,
        read_state: str,
        limit: int,
        body_limit: int = 0,
    ) -> list[MailMessage]: ...


class MailWriter(Protocol):
    def apply_mail_state(
        self,
        outlook_id: int,
        *,
        categories: list[str],
        flag_status: FlagStatus,
        target_folder_id: int | None = None,
    ) -> None: ...

    def apply_mail_states(self, updates: list[dict[str, Any]]) -> list[dict[str, str | int]]: ...

    def ensure_mail_folder(self, inbox_id: int, folder_name: str) -> dict[str, str | int]: ...
