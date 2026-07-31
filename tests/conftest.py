from __future__ import annotations

from pathlib import Path

import pytest

from outlook_organizer.config import AppConfig, load_config
from outlook_organizer.models import FlagStatus, MailMessage, Recipient


@pytest.fixture
def app_config() -> AppConfig:
    return load_config(Path(__file__).resolve().parents[1] / "config")


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
