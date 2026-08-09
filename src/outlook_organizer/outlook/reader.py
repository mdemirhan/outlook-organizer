from __future__ import annotations

from collections.abc import Iterable

from outlook_organizer.calendar import CalendarEvent, CalendarInfo
from outlook_organizer.mail import MailMessage, OutlookFolder
from outlook_organizer.outlook.adapter import OutlookAdapter


class OutlookReader:
    """Read-only Outlook capability exposed to query-side services."""

    def __init__(self, adapter: OutlookAdapter | None = None) -> None:
        self._adapter = adapter or OutlookAdapter()

    def list_folders(self, maximum_id: int = 5000, *, refresh: bool = False) -> list[OutlookFolder]:
        return self._adapter.list_folders(maximum_id, refresh=refresh)

    def find_folder(self, names: Iterable[str], maximum_id: int = 5000) -> OutlookFolder:
        return self._adapter.find_folder(names, maximum_id)

    def latest_messages(
        self, folder_id: int, limit: int = 20, body_limit: int = 2000
    ) -> list[MailMessage]:
        return self._adapter.latest_messages(folder_id, limit=limit, body_limit=body_limit)

    def get_message(self, outlook_id: int, body_limit: int = 20_000) -> MailMessage:
        return self._adapter.get_message(outlook_id, body_limit=body_limit)

    def get_messages(self, outlook_ids: Iterable[int], *, body_limit: int = 0) -> list[MailMessage]:
        return self._adapter.get_messages(outlook_ids, body_limit=body_limit)

    def messages_in_window(
        self,
        folder_id: int,
        *,
        start_offset_seconds: int,
        end_offset_seconds: int,
        read_state: str,
        limit: int,
        body_limit: int = 0,
    ) -> list[MailMessage]:
        return self._adapter.messages_in_window(
            folder_id,
            start_offset_seconds=start_offset_seconds,
            end_offset_seconds=end_offset_seconds,
            read_state=read_state,
            limit=limit,
            body_limit=body_limit,
        )

    def list_calendars(
        self, maximum_id: int = 5000, *, refresh: bool = False
    ) -> list[CalendarInfo]:
        return self._adapter.list_calendars(maximum_id, refresh=refresh)

    def find_calendar(self, names: Iterable[str], maximum_id: int = 5000) -> CalendarInfo:
        return self._adapter.find_calendar(names, maximum_id)

    def calendar_events(
        self,
        calendar_id: int,
        *,
        days_behind: int = 0,
        days_ahead: int = 7,
        body_limit: int = 0,
    ) -> list[CalendarEvent]:
        return self._adapter.calendar_events(
            calendar_id,
            days_behind=days_behind,
            days_ahead=days_ahead,
            body_limit=body_limit,
        )
