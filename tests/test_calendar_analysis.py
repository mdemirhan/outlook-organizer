from __future__ import annotations

from datetime import date, time

from outlook_organizer.calendar_analysis import analyze_calendar, find_free_slots
from outlook_organizer.models import CalendarEvent


def event(subject: str, start: str, end: str) -> CalendarEvent:
    return CalendarEvent(
        outlook_id=1,
        exchange_id=subject,
        calendar_id=200,
        subject=subject,
        start_at=start,
        end_at=end,
        location="",
        organizer="",
        all_day=False,
        free_busy_status="busy",
        is_private=False,
        categories=[],
        attendees=[],
    )


def test_calendar_conflicts_and_back_to_back() -> None:
    events = [
        event(
            "First",
            "Tuesday, July 28, 2026 at 09:00:00 AM",
            "Tuesday, July 28, 2026 at 10:00:00 AM",
        ),
        event(
            "Second",
            "Tuesday, July 28, 2026 at 09:30:00 AM",
            "Tuesday, July 28, 2026 at 10:30:00 AM",
        ),
        event(
            "Third",
            "Tuesday, July 28, 2026 at 10:30:00 AM",
            "Tuesday, July 28, 2026 at 11:00:00 AM",
        ),
    ]
    result = analyze_calendar(events)
    assert len(result["conflicts"]) == 1
    assert len(result["back_to_back"]) == 1


def test_find_focus_slots() -> None:
    events = [
        event(
            "Meeting",
            "Tuesday, July 28, 2026 at 11:00:00 AM",
            "Tuesday, July 28, 2026 at 12:00:00 PM",
        )
    ]
    slots = find_free_slots(
        events,
        date(2026, 7, 28),
        work_start=time(9),
        work_end=time(18),
        minimum_minutes=90,
    )
    assert slots[0]["minutes"] == 120
    assert slots[1]["minutes"] == 360


def test_focus_slots_respect_buffers_and_blocked_windows() -> None:
    events = [
        event(
            "Meeting",
            "2026-07-28T10:00:00",
            "2026-07-28T11:00:00",
        )
    ]
    slots = find_free_slots(
        events,
        date(2026, 7, 28),
        work_start=time(9),
        work_end=time(14),
        minimum_minutes=30,
        buffer_minutes=10,
        blocked_windows=[(time(12), time(13))],
    )
    assert slots == [
        {
            "start": "2026-07-28T09:00:00",
            "end": "2026-07-28T09:50:00",
            "minutes": 50,
        },
        {
            "start": "2026-07-28T11:10:00",
            "end": "2026-07-28T12:00:00",
            "minutes": 50,
        },
        {
            "start": "2026-07-28T13:00:00",
            "end": "2026-07-28T14:00:00",
            "minutes": 60,
        },
    ]


def test_nested_overlaps_are_all_reported() -> None:
    result = analyze_calendar(
        [
            event("Long", "2026-07-28T09:00:00", "2026-07-28T12:00:00"),
            event("Short", "2026-07-28T10:00:00", "2026-07-28T10:30:00"),
            event("Later", "2026-07-28T11:00:00", "2026-07-28T11:30:00"),
        ]
    )
    assert {(item["first"], item["second"]) for item in result["conflicts"]} == {
        ("Long", "Short"),
        ("Long", "Later"),
    }
