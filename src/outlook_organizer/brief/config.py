from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from pydantic import Field, field_validator, model_validator

from outlook_organizer.mail.config import MailContext, load_mail_context
from outlook_organizer.mail.models import StrictModel
from outlook_organizer.paths import config_dir
from outlook_organizer.yaml_config import load_yaml_model

BriefPeriod = Literal[
    "last_hour",
    "last_24_hours",
    "today",
    "yesterday",
    "since_yesterday",
    "since_last_workday",
]
BriefReadState = Literal["unread", "read", "all"]
BriefGroupBy = Literal["none", "conversation"]
BriefContentMode = Literal["detailed", "concise", "rollup", "metadata_only"]


class BriefScopeConfig(StrictModel):
    folder: str
    recursive: bool = False


class BriefFolderPolicyConfig(StrictModel):
    content: BriefContentMode = "concise"
    snippet_chars: int = Field(default=1200, ge=0, le=20_000)

    @model_validator(mode="after")
    def validate_content_limit(self) -> BriefFolderPolicyConfig:
        if self.content in {"rollup", "metadata_only"} and self.snippet_chars != 0:
            raise ValueError(f"{self.content} policies must use snippet_chars: 0")
        if self.content in {"detailed", "concise"} and self.snippet_chars == 0:
            raise ValueError(f"{self.content} policies must use a positive snippet_chars value")
        return self


class BriefDefaultsConfig(StrictModel):
    scopes: list[BriefScopeConfig] = Field(
        default_factory=lambda: [BriefScopeConfig(folder="organized_primary", recursive=True)]
    )
    period: BriefPeriod = "today"
    read_state: BriefReadState = "all"
    include_attention_debt: bool = False
    attention_debt_days: int = Field(default=7, ge=1, le=90)
    attention_debt_folders: list[str] = Field(default_factory=list)
    group_by: BriefGroupBy = "none"
    max_messages: int = Field(default=75, ge=1, le=500)
    folder_policies: dict[str, BriefFolderPolicyConfig] = Field(default_factory=dict)


class BriefProfileConfig(StrictModel):
    name: str
    description: str = ""
    aliases: list[str] = Field(default_factory=list)
    scopes: list[BriefScopeConfig] | None = None
    period: BriefPeriod | None = None
    read_state: BriefReadState | None = None
    include_attention_debt: bool | None = None
    attention_debt_days: int | None = Field(default=None, ge=1, le=90)
    attention_debt_folders: list[str] | None = None
    group_by: BriefGroupBy | None = None
    max_messages: int | None = Field(default=None, ge=1, le=500)

    @field_validator("aliases")
    @classmethod
    def normalize_aliases(cls, values: list[str]) -> list[str]:
        return sorted({value.strip() for value in values if value.strip()}, key=str.casefold)


class BriefConfig(StrictModel):
    version: Literal[1]
    timezone: str
    default_profile: str | None = None
    defaults: BriefDefaultsConfig = Field(default_factory=BriefDefaultsConfig)
    profiles: dict[str, BriefProfileConfig] = Field(default_factory=dict)


@dataclass(frozen=True)
class BriefContext:
    mail: MailContext
    config: BriefConfig
    fingerprint: str


def load_brief_context(directory: Path | None = None) -> BriefContext:
    root = (directory or config_dir()).resolve()
    mail = load_mail_context(root)
    brief = load_yaml_model(root / "brief.yaml", BriefConfig)
    folder_keys = set(mail.folders.folders)

    def require(values: set[str], context: str) -> None:
        if missing := sorted(values - folder_keys):
            raise ValueError(f"{context} references undefined folders: {missing}")

    require({scope.folder for scope in brief.defaults.scopes}, "brief defaults")
    require(set(brief.defaults.folder_policies), "brief folder policies")
    require(set(brief.defaults.attention_debt_folders), "brief attention debt")
    for key, profile in brief.profiles.items():
        require({scope.folder for scope in profile.scopes or []}, f"brief profile {key!r}")
        require(set(profile.attention_debt_folders or []), f"brief profile {key!r} attention debt")
    if brief.default_profile and brief.default_profile not in brief.profiles:
        raise ValueError(
            f"brief default_profile references undefined profile {brief.default_profile!r}"
        )
    aliases: dict[str, str] = {}
    for key, profile in brief.profiles.items():
        for value in [key, profile.name, *profile.aliases]:
            normalized = value.strip().casefold()
            if not normalized:
                raise ValueError(f"brief profile {key!r} has an empty name or alias")
            if previous := aliases.get(normalized):
                if previous != key:
                    raise ValueError(
                        f"brief profile name or alias {value!r} is shared by "
                        f"{previous!r} and {key!r}"
                    )
            aliases[normalized] = key
    canonical = json.dumps(
        {
            "brief": brief.model_dump(mode="json"),
            "mail_fingerprint": mail.fingerprint,
        },
        sort_keys=True,
        separators=(",", ":"),
    )
    return BriefContext(
        mail=mail,
        config=brief,
        fingerprint=hashlib.sha256(canonical.encode()).hexdigest()[:16],
    )
