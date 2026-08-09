from __future__ import annotations

from dataclasses import asdict
from datetime import date, time
from typing import Any, Protocol

from outlook_organizer.calendar.models import CalendarConfig, CalendarEvent, CalendarInfo
from outlook_organizer.calendar_analysis import analyze_calendar, find_free_slots


class CalendarReader(Protocol):
    def list_calendars(
        self, maximum_id: int = 5000, *, refresh: bool = False
    ) -> list[CalendarInfo]: ...

    def find_calendar(self, names: list[str], maximum_id: int = 5000) -> CalendarInfo: ...

    def calendar_events(
        self,
        calendar_id: int,
        *,
        days_behind: int = 0,
        days_ahead: int = 7,
        body_limit: int = 0,
    ) -> list[CalendarEvent]: ...


class CalendarService:
    def __init__(self, config: CalendarConfig, reader: CalendarReader) -> None:
        self.config = config
        self.reader = reader

    def calendars(self) -> list[dict[str, Any]]:
        return [
            asdict(value) for value in self.reader.list_calendars(self.config.maximum_calendar_id)
        ]

    def events(
        self, *, days_behind: int = 0, days_ahead: int = 7, include_body: bool = False
    ) -> list[dict[str, Any]]:
        calendar = self.reader.find_calendar(
            self.config.calendar_names, self.config.maximum_calendar_id
        )
        events = self.reader.calendar_events(
            calendar.outlook_id,
            days_behind=days_behind,
            days_ahead=days_ahead,
            body_limit=5000 if include_body else 0,
        )
        return [self._event_dict(event, include_body=include_body) for event in events]

    def analyze(self, *, days_behind: int = 0, days_ahead: int = 7) -> dict[str, Any]:
        calendar = self.reader.find_calendar(
            self.config.calendar_names, self.config.maximum_calendar_id
        )
        result = analyze_calendar(
            self.reader.calendar_events(
                calendar.outlook_id, days_behind=days_behind, days_ahead=days_ahead
            )
        )
        result["calendar"] = calendar.name
        return result

    def free_slots(
        self, target_date: date, *, minimum_minutes: int | None = None
    ) -> list[dict[str, str | int]]:
        calendar = self.reader.find_calendar(
            self.config.calendar_names, self.config.maximum_calendar_id
        )
        events = self.reader.calendar_events(calendar.outlook_id, days_behind=1, days_ahead=14)
        hours = self.config.working_hours.get(target_date.strftime("%A").lower())
        if not hours:
            return []
        return find_free_slots(
            events,
            target_date,
            work_start=time.fromisoformat(hours[0]),
            work_end=time.fromisoformat(hours[1]),
            minimum_minutes=minimum_minutes or self.config.preferences.minimum_focus_block_minutes,
            buffer_minutes=self.config.preferences.meeting_buffer_minutes,
            blocked_windows=[
                (
                    time.fromisoformat(self.config.preferences.lunch_window[0]),
                    time.fromisoformat(self.config.preferences.lunch_window[1]),
                )
            ],
        )

    @staticmethod
    def _event_dict(event: CalendarEvent, *, include_body: bool) -> dict[str, Any]:
        result = asdict(event)
        if not include_body or event.is_private:
            result.pop("body", None)
        if event.is_private:
            result.update(subject="Private appointment", location="", attendees=[])
        return result
