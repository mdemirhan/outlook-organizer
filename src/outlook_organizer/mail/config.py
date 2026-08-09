from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path

from outlook_organizer.mail.models import FolderCatalogConfig, MailDefinitionsConfig
from outlook_organizer.paths import config_dir
from outlook_organizer.yaml_config import load_yaml_model


@dataclass(frozen=True)
class MailContext:
    definitions: MailDefinitionsConfig
    folders: FolderCatalogConfig
    fingerprint: str
    directory: Path


def load_mail_context(directory: Path | None = None) -> MailContext:
    root = (directory or config_dir()).resolve()
    definitions = load_yaml_model(root / "mail-definitions.yaml", MailDefinitionsConfig)
    folders = load_yaml_model(root / "mail-folders.yaml", FolderCatalogConfig)
    canonical = json.dumps(
        {
            "definitions": definitions.model_dump(mode="json"),
            "folders": folders.model_dump(mode="json"),
        },
        sort_keys=True,
        separators=(",", ":"),
    )
    return MailContext(
        definitions=definitions,
        folders=folders,
        fingerprint=hashlib.sha256(canonical.encode()).hexdigest()[:16],
        directory=root,
    )
