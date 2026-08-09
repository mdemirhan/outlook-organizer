from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from outlook_organizer.calendar.models import (
    CalendarAttendee,
    CalendarEvent,
    CalendarInfo,
)
from outlook_organizer.mail.models import (
    FlagStatus,
    MailMessage,
    OutlookFolder,
    Recipient,
)
from outlook_organizer.outlook.applescript import AppleScriptRunner, OutlookError
from outlook_organizer.outlook.scripts import (
    APPLY_MAIL_STATE,
    APPLY_MAIL_STATES,
    ENSURE_MAIL_FOLDER,
    LIST_CALENDARS,
    LIST_FOLDERS,
    READ_CALENDAR_EVENTS,
    READ_MESSAGE_BY_ID,
    READ_MESSAGES_BY_IDS,
    READ_MESSAGES_IN_FOLDER_ORDER,
    READ_MESSAGES_IN_FOLDER_WINDOW,
)

RS = "\x1e"
US = "\x1f"
GS = "\x1d"
FS = "\x1c"


def _bool(value: str) -> bool:
    return value.strip().lower() == "true"


def _flag(value: str) -> FlagStatus:
    normalized = value.strip().lower()
    if normalized == "not completed":
        return FlagStatus.FLAGGED
    if normalized == "completed":
        return FlagStatus.COMPLETED
    return FlagStatus.NOT_FLAGGED


def _recipient_list(value: str) -> list[Recipient]:
    recipients: list[Recipient] = []
    for encoded in filter(None, value.split(GS)):
        parts = encoded.split(FS)
        parts.extend([""] * (3 - len(parts)))
        recipients.append(Recipient(name=parts[0], address=parts[1].lower(), kind=parts[2]))
    return recipients


class OutlookAdapter:
    def __init__(self, runner: AppleScriptRunner | None = None) -> None:
        self.runner = runner or AppleScriptRunner()
        self._folder_cache: dict[int, list[OutlookFolder]] = {}
        self._calendar_cache: dict[int, list[CalendarInfo]] = {}

    def list_folders(self, maximum_id: int = 5000, *, refresh: bool = False) -> list[OutlookFolder]:
        if not refresh and maximum_id in self._folder_cache:
            return list(self._folder_cache[maximum_id])
        result = self.runner.run(LIST_FOLDERS, str(maximum_id))
        folders: list[OutlookFolder] = []
        for line in result.stdout.splitlines():
            parts = line.split("\t")
            if len(parts) != 4:
                continue
            folders.append(
                OutlookFolder(
                    outlook_id=int(parts[0]),
                    name=parts[1],
                    parent_name=parts[2],
                    message_count=int(parts[3]),
                )
            )
        self._folder_cache[maximum_id] = folders
        return list(folders)

    def find_folder(self, names: Iterable[str], maximum_id: int = 5000) -> OutlookFolder:
        wanted = {name.casefold() for name in names}
        matches = [
            folder for folder in self.list_folders(maximum_id) if folder.name.casefold() in wanted
        ]
        if not matches:
            raise OutlookError(f"No Outlook folder found for: {', '.join(names)}")
        return max(matches, key=lambda folder: folder.message_count)

    def list_calendars(
        self, maximum_id: int = 5000, *, refresh: bool = False
    ) -> list[CalendarInfo]:
        if not refresh and maximum_id in self._calendar_cache:
            return list(self._calendar_cache[maximum_id])
        result = self.runner.run(LIST_CALENDARS, str(maximum_id))
        calendars: list[CalendarInfo] = []
        for line in result.stdout.splitlines():
            parts = line.split("\t")
            if len(parts) != 3:
                continue
            calendars.append(
                CalendarInfo(
                    outlook_id=int(parts[0]),
                    name=parts[1],
                    event_count=int(parts[2]),
                )
            )
        self._calendar_cache[maximum_id] = calendars
        return list(calendars)

    def find_calendar(self, names: Iterable[str], maximum_id: int = 5000) -> CalendarInfo:
        wanted = {name.casefold() for name in names}
        matches = [
            calendar
            for calendar in self.list_calendars(maximum_id)
            if calendar.name.casefold() in wanted
        ]
        if not matches:
            raise OutlookError(f"No Outlook calendar found for: {', '.join(names)}")
        return max(matches, key=lambda calendar: calendar.event_count)

    def latest_messages(
        self,
        folder_id: int,
        limit: int = 20,
        body_limit: int = 2000,
    ) -> list[MailMessage]:
        result = self.runner.run(
            READ_MESSAGES_IN_FOLDER_ORDER,
            str(folder_id),
            str(limit),
            str(body_limit),
        )
        messages = self._parse_messages(result.stdout)
        return sorted(
            messages,
            key=lambda message: (message.received_at, message.outlook_id),
            reverse=True,
        )

    def get_message(self, outlook_id: int, body_limit: int = 20_000) -> MailMessage:
        result = self.runner.run(READ_MESSAGE_BY_ID, str(outlook_id), str(body_limit))
        messages = self._parse_messages(result.stdout)
        if not messages:
            raise OutlookError(f"Outlook message {outlook_id} was not found")
        return messages[0]

    def get_messages(self, outlook_ids: Iterable[int], *, body_limit: int = 0) -> list[MailMessage]:
        ids = [int(value) for value in outlook_ids]
        if not ids:
            return []
        result = self.runner.run(
            READ_MESSAGES_BY_IDS,
            str(body_limit),
            *(str(value) for value in ids),
        )
        return self._parse_messages(result.stdout)

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
        if read_state not in {"unread", "read", "all"}:
            raise ValueError(f"Unsupported read state: {read_state}")
        result = self.runner.run(
            READ_MESSAGES_IN_FOLDER_WINDOW,
            str(folder_id),
            str(start_offset_seconds),
            str(end_offset_seconds),
            read_state,
            str(limit),
            str(body_limit),
        )
        return sorted(
            self._parse_messages(result.stdout),
            key=lambda message: (message.received_at, message.outlook_id),
            reverse=True,
        )

    def _parse_messages(self, output: str) -> list[MailMessage]:
        if not output.strip():
            return []
        messages: list[MailMessage] = []
        for row in filter(None, output.split(RS)):
            fields = row.split(US)
            if len(fields) < 14:
                raise OutlookError(f"Unexpected Outlook message response with {len(fields)} fields")
            messages.append(
                MailMessage(
                    outlook_id=int(fields[0]),
                    exchange_id=fields[1],
                    folder_id=int(fields[2]),
                    folder_name=fields[3],
                    subject=fields[4],
                    sender_name=fields[5],
                    sender_address=fields[6].lower(),
                    to=_recipient_list(fields[7]),
                    cc=_recipient_list(fields[8]),
                    received_at=fields[9],
                    flag_status=_flag(fields[10]),
                    categories=[value for value in fields[11].split(GS) if value],
                    body=fields[12],
                    has_attachments=_bool(fields[13]),
                    thread_guid=fields[14].strip() if len(fields) > 14 else "",
                    is_read=_bool(fields[15]) if len(fields) > 15 else True,
                    replied_to=_bool(fields[16]) if len(fields) > 16 else False,
                )
            )
        return messages

    def calendar_events(
        self,
        calendar_id: int,
        *,
        days_behind: int = 0,
        days_ahead: int = 7,
        body_limit: int = 0,
    ) -> list[CalendarEvent]:
        result = self.runner.run(
            READ_CALENDAR_EVENTS,
            str(calendar_id),
            str(days_behind),
            str(days_ahead),
            str(body_limit),
        )
        events: list[CalendarEvent] = []
        for row in filter(None, result.stdout.split(RS)):
            fields = row.split(US)
            if len(fields) < 14:
                raise OutlookError(
                    f"Unexpected Outlook calendar response with {len(fields)} fields"
                )
            attendees: list[CalendarAttendee] = []
            for encoded in filter(None, fields[12].split(GS)):
                attendee_fields = encoded.split(FS)
                attendee_fields.extend([""] * (4 - len(attendee_fields)))
                attendees.append(
                    CalendarAttendee(
                        name=attendee_fields[0],
                        address=attendee_fields[1].lower(),
                        attendee_type=attendee_fields[2],
                        status=attendee_fields[3],
                    )
                )
            events.append(
                CalendarEvent(
                    outlook_id=int(fields[0]),
                    exchange_id=fields[1],
                    calendar_id=int(fields[2]),
                    subject=fields[3],
                    start_at=fields[4],
                    end_at=fields[5],
                    location=fields[6],
                    organizer=fields[7],
                    all_day=_bool(fields[8]),
                    free_busy_status=fields[9],
                    is_private=_bool(fields[10]),
                    categories=[value for value in fields[11].split(GS) if value],
                    attendees=attendees,
                    body=fields[13],
                )
            )
        return events

    def apply_mail_state(
        self,
        outlook_id: int,
        *,
        categories: list[str],
        flag_status: FlagStatus,
        target_folder_id: int | None = None,
    ) -> None:
        self.runner.run(
            APPLY_MAIL_STATE,
            str(outlook_id),
            GS.join(categories),
            flag_status.value,
            str(target_folder_id or 0),
        )

    def apply_mail_states(self, updates: list[dict[str, Any]]) -> list[dict[str, str | int]]:
        if not updates:
            return []
        arguments: list[str] = []
        for update in updates:
            flag_status = update["flag_status"]
            if not isinstance(flag_status, FlagStatus):
                flag_status = FlagStatus(str(flag_status))
            arguments.extend(
                [
                    str(update["outlook_id"]),
                    GS.join(str(value) for value in update["categories"]),
                    flag_status.value,
                    str(update.get("target_folder_id") or 0),
                ]
            )
        result = self.runner.run(APPLY_MAIL_STATES, *arguments)
        parsed: list[dict[str, str | int]] = []
        for row in filter(None, result.stdout.split(RS)):
            fields = row.split(US)
            fields.extend([""] * (3 - len(fields)))
            parsed.append(
                {
                    "outlook_id": int(fields[0]),
                    "status": fields[1],
                    "error": fields[2],
                }
            )
        return parsed

    def ensure_mail_folder(self, inbox_id: int, folder_name: str) -> dict[str, str | int]:
        result = self.runner.run(ENSURE_MAIL_FOLDER, str(inbox_id), folder_name)
        fields = result.stdout.strip().split("\t")
        if len(fields) != 3:
            raise OutlookError("Unexpected Outlook folder-creation response")
        self._folder_cache.clear()
        return {
            "outlook_id": int(fields[0]),
            "name": fields[1],
            "status": fields[2],
        }
