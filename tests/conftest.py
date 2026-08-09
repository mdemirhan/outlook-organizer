from __future__ import annotations

from pathlib import Path

import pytest

from outlook_organizer.brief import BriefContext, load_brief_context
from outlook_organizer.mail import (
    FlagStatus,
    MailContext,
    MailMessage,
    Recipient,
    load_mail_context,
)
from outlook_organizer.triage import TriageContext, load_triage_context

CONFIG = Path(__file__).resolve().parents[1] / "config"


@pytest.fixture
def mail_context() -> MailContext:
    return load_mail_context(CONFIG)


@pytest.fixture
def triage_context() -> TriageContext:
    return load_triage_context(CONFIG)


@pytest.fixture
def brief_context() -> BriefContext:
    return load_brief_context(CONFIG)


@pytest.fixture
def direct_message() -> MailMessage:
    return MailMessage(
        outlook_id=42,
        exchange_id="exchange-42",
        folder_id=101,
        folder_name="Inbox",
        subject="Please review",
        sender_name="Team Member",
        sender_address="member@corp.example",
        to=[Recipient("Example User", "example.user@corp.example")],
        cc=[],
        received_at="Tuesday, July 28, 2026 at 10:00:00 AM",
        flag_status=FlagStatus.NOT_FLAGGED,
        categories=[],
        body="Please review the proposal.",
    )
