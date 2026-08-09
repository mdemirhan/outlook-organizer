from __future__ import annotations

from outlook_organizer.outlook import OutlookAdapter
from outlook_organizer.outlook.scripts import (
    READ_MESSAGES_IN_FOLDER_ORDER,
    READ_THREAD_STATES_BY_IDS,
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
        ]
    )

    messages = adapter._parse_messages(row)

    assert messages[0].thread_guid == "thread-guid"


def test_thread_state_reader_is_targeted_by_message_id() -> None:
    assert "repeat with messageIDText in argv" in READ_THREAD_STATES_BY_IDS
    assert "every message of folderRef" not in READ_THREAD_STATES_BY_IDS


def test_thread_state_response_is_parsed() -> None:
    adapter = OutlookAdapter()

    states = adapter._parse_thread_states(
        "42\x1fexchange-42\x1f110\x1fInternal General\x1fthread-guid\n"
    )

    assert len(states) == 1
    assert states[0].outlook_id == 42
    assert states[0].folder_id == 110
    assert states[0].thread_guid == "thread-guid"
