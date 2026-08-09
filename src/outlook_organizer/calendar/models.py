from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from pydantic import Field

from outlook_organizer.mail.models import StrictModel


class CalendarPreferences(StrictModel):
    lunch_window: tuple[str, str]
    minimum_focus_block_minutes: int = Field(ge=15, le=480)
    meeting_buffer_minutes: int = Field(ge=0, le=120)
    maximum_meeting_hours_per_day: float = Field(ge=0, le=24)
    avoid_back_to_back_meetings: bool = True
    preferred_focus_windows: list[tuple[str, str]] = Field(default_factory=list)


class ProtectedRelationships(StrictModel):
    high_priority: list[str] = Field(default_factory=list)


class CalendarConfig(StrictModel):
    version: Literal[1]
    timezone: str
    calendar_names: list[str]
    maximum_calendar_id: int = Field(default=5000, ge=10, le=100_000)
    working_hours: dict[str, tuple[str, str]]
    preferences: CalendarPreferences
    protected_relationships: ProtectedRelationships


@dataclass(slots=True)
class CalendarInfo:
    outlook_id: int
    name: str
    event_count: int


@dataclass(slots=True)
class CalendarAttendee:
    name: str
    address: str
    attendee_type: str
    status: str


@dataclass(slots=True)
class CalendarEvent:
    outlook_id: int
    exchange_id: str
    calendar_id: int
    subject: str
    start_at: str
    end_at: str
    location: str
    organizer: str
    all_day: bool
    free_busy_status: str
    is_private: bool
    categories: list[str]
    attendees: list[CalendarAttendee]
    body: str = ""
