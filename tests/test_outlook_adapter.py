from __future__ import annotations

from outlook_organizer.outlook import OutlookAdapter
from outlook_organizer.outlook.scripts import READ_MESSAGES_IN_FOLDER_ORDER


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
