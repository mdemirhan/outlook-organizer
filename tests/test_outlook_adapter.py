from __future__ import annotations

from outlook_organizer.outlook import OutlookAdapter
from outlook_organizer.outlook.scripts import (
    READ_MESSAGES_IN_FOLDER_ORDER,
    READ_MESSAGES_IN_FOLDER_WINDOW,
)


def test_empty_outlook_output_returns_no_messages() -> None:
    adapter = OutlookAdapter()

    assert adapter._parse_messages("") == []
    assert adapter._parse_messages("\n") == []
    assert adapter._parse_messages("  \n") == []


def test_folder_order_reader_skips_messages_without_received_date() -> None:
    date_guard = """\
            set receivedDate to missing value
            try
                set receivedDate to time received of messageRef
            end try
            if receivedDate is not missing value then
                set end of rows to my messageRow(messageRef, folderRef, bodyLimit)
            end if"""

    assert date_guard in READ_MESSAGES_IN_FOLDER_ORDER


def test_message_reader_includes_hidden_outlook_thread_guid() -> None:
    assert "«class lOTd» of messageRef" in READ_MESSAGES_IN_FOLDER_ORDER


def test_message_thread_guid_is_stripped_before_indexing() -> None:
    adapter = OutlookAdapter()
    row = "\x1f".join(
        [
            "42",
            "exchange-42",
            "101",
            "Inbox",
            "Subject",
            "Sender",
            "sender@example.com",
            "",
            "",
            "2026-08-09T00:00:00Z",
            "not flagged",
            "",
            "",
            "false",
            "thread-guid\n",
            "false",
            "true",
        ]
    )

    messages = adapter._parse_messages(row)

    assert messages[0].thread_guid == "thread-guid"
    assert not messages[0].is_read
    assert messages[0].replied_to


def test_window_reader_filters_by_period_and_read_state() -> None:
    assert "time received >= windowStart" in READ_MESSAGES_IN_FOLDER_WINDOW
    assert "time received < windowEnd" in READ_MESSAGES_IN_FOLDER_WINDOW
    assert 'desiredReadState is "unread"' in READ_MESSAGES_IN_FOLDER_WINDOW
    assert "get is read of messageRef" in READ_MESSAGES_IN_FOLDER_WINDOW
