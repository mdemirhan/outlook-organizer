from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, time, timedelta
from typing import Any

from outlook_organizer.calendar.models import CalendarEvent

OUTLOOK_DATE_FORMATS = (
    "%Y-%m-%dT%H:%M:%S",
    "%A, %B %d, %Y at %I:%M:%S %p",
    "%A, %B %d, %Y at %H:%M:%S",
    "%A, %d %B %Y at %H:%M:%S",
)


def parse_outlook_date(value: str) -> datetime | None:
    for date_format in OUTLOOK_DATE_FORMATS:
        try:
            return datetime.strptime(value, date_format)
        except ValueError:
            continue
    return None


def analyze_calendar(events: list[CalendarEvent]) -> dict[str, Any]:
    normalized = [
        (event, parse_outlook_date(event.start_at), parse_outlook_date(event.end_at))
        for event in events
    ]
    normalized = [item for item in normalized if item[1] is not None and item[2] is not None]
    normalized.sort(key=lambda item: item[1])

    conflicts: list[dict[str, str]] = []
    back_to_back: list[dict[str, str]] = []
    meeting_minutes_by_day: dict[str, int] = defaultdict(int)

    for index, (event, start, end) in enumerate(normalized):
        assert start is not None and end is not None
        if not event.all_day and event.free_busy_status.casefold() != "free":
            meeting_minutes_by_day[start.date().isoformat()] += max(
                0, int((end - start).total_seconds() // 60)
            )
        for previous, previous_start, previous_end in normalized[:index]:
            assert previous_start is not None and previous_end is not None
            if start < previous_end and end > previous_start:
                conflicts.append(
                    {
                        "first": previous.subject,
                        "second": event.subject,
                        "overlap_start": max(start, previous_start).isoformat(),
                    }
                )
            elif start == previous_end:
                back_to_back.append(
                    {
                        "first": previous.subject,
                        "second": event.subject,
                        "at": start.isoformat(),
                    }
                )

    return {
        "event_count": len(events),
        "meeting_hours_by_day": {
            key: round(minutes / 60, 2) for key, minutes in meeting_minutes_by_day.items()
        },
        "conflicts": conflicts,
        "back_to_back": back_to_back,
    }


def find_free_slots(
    events: list[CalendarEvent],
    target_date: date,
    *,
    work_start: time,
    work_end: time,
    minimum_minutes: int,
    buffer_minutes: int = 0,
    blocked_windows: list[tuple[time, time]] | None = None,
) -> list[dict[str, str | int]]:
    day_start = datetime.combine(target_date, work_start)
    day_end = datetime.combine(target_date, work_end)
    busy: list[tuple[datetime, datetime]] = []
    for event in events:
        start = parse_outlook_date(event.start_at)
        end = parse_outlook_date(event.end_at)
        if start is None or end is None or event.free_busy_status.casefold() == "free":
            continue
        if end <= day_start or start >= day_end:
            continue
        buffered_start = start - timedelta(minutes=buffer_minutes)
        buffered_end = end + timedelta(minutes=buffer_minutes)
        busy.append((max(buffered_start, day_start), min(buffered_end, day_end)))
    for blocked_start, blocked_end in blocked_windows or []:
        busy.append(
            (
                datetime.combine(target_date, blocked_start),
                datetime.combine(target_date, blocked_end),
            )
        )
    busy.sort()

    slots: list[dict[str, str | int]] = []
    cursor = day_start
    for start, end in busy:
        if start > cursor:
            minutes = int((start - cursor).total_seconds() // 60)
            if minutes >= minimum_minutes:
                slots.append(
                    {"start": cursor.isoformat(), "end": start.isoformat(), "minutes": minutes}
                )
        cursor = max(cursor, end)
    if cursor < day_end:
        minutes = int((day_end - cursor).total_seconds() // 60)
        if minutes >= minimum_minutes:
            slots.append(
                {"start": cursor.isoformat(), "end": day_end.isoformat(), "minutes": minutes}
            )
    return slots
