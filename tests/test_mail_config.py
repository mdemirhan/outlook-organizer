from __future__ import annotations

from dataclasses import replace

from outlook_organizer.mail.models import FolderCatalogConfig
from outlook_organizer.mail.service import FolderAdminService


class RecordingFolderWriter:
    def __init__(self) -> None:
        self.created: list[tuple[int, str]] = []

    def ensure_mail_folder(self, inbox_id: int, folder_name: str) -> dict[str, str | int]:
        self.created.append((inbox_id, folder_name))
        return {"id": len(self.created), "name": folder_name, "status": "existing"}


def test_rootless_folder_catalog_and_setup(mail_context) -> None:
    catalog = FolderCatalogConfig.model_validate(
        {
            "version": 1,
            "folders": {
                "inbox": {"name": "Inbox", "id": 101},
                "top_level": {"name": "Top Level", "id": 102},
                "nested": {"name": "Nested", "id": 103, "parent": "top_level"},
            },
        }
    )
    writer = RecordingFolderWriter()

    result = FolderAdminService(replace(mail_context, folders=catalog), writer).setup(confirm=True)

    assert list(result["folders"]) == ["top_level"]
    assert writer.created == [(101, "Top Level")]
