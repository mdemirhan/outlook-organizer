from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml
from pydantic import BaseModel

from outlook_organizer.models import CalendarConfig, MailDefinitionsConfig, MailRulesConfig
from outlook_organizer.paths import config_dir


@dataclass(frozen=True)
class AppConfig:
    definitions: MailDefinitionsConfig
    mail: MailRulesConfig
    calendar: CalendarConfig
    fingerprint: str
    directory: Path


def _read_yaml(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"{path} must contain a YAML mapping")
    return data


def _load_model[ModelT: BaseModel](path: Path, model: type[ModelT]) -> ModelT:
    return model.model_validate(_read_yaml(path))


def load_config(directory: Path | None = None) -> AppConfig:
    root = (directory or config_dir()).resolve()
    definitions = _load_model(root / "mail-definitions.yaml", MailDefinitionsConfig)
    mail = _load_model(root / "mail-rules.yaml", MailRulesConfig)
    calendar = _load_model(root / "calendar.yaml", CalendarConfig)

    canonical = json.dumps(
        {
            "definitions": definitions.model_dump(mode="json"),
            "mail": mail.model_dump(mode="json"),
            "calendar": calendar.model_dump(mode="json"),
        },
        sort_keys=True,
        separators=(",", ":"),
    )
    fingerprint = hashlib.sha256(canonical.encode()).hexdigest()[:16]
    return AppConfig(
        definitions=definitions,
        mail=mail,
        calendar=calendar,
        fingerprint=fingerprint,
        directory=root,
    )
