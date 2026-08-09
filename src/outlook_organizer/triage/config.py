from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from pydantic import Field, model_validator

from outlook_organizer.mail.config import MailContext, load_mail_context
from outlook_organizer.mail.models import StrictModel
from outlook_organizer.paths import config_dir
from outlook_organizer.yaml_config import load_yaml_model


class MatchConfig(StrictModel):
    flagged: bool | None = None
    recipient: (
        Literal["only_me", "direct_to_me", "multi_recipient", "not_to_me", "unknown"]
        | list[Literal["only_me", "direct_to_me", "multi_recipient", "not_to_me", "unknown"]]
        | None
    ) = None
    sender_group: str | None = None
    sender_type: (
        Literal[
            "internal",
            "junk_external",
            "safe_external",
            "unclassified_external",
            "unknown",
        ]
        | None
    ) = None
    distribution_list_group: str | None = None
    distribution_list: bool | None = None
    distribution_delivery: bool | None = None


class AnnotationRule(StrictModel):
    id: str
    description: str = ""
    when: MatchConfig
    add_category: str | None = None
    section: str | None = None
    keep_in_inbox: bool = False


class RouteRule(StrictModel):
    id: str
    description: str = ""
    when: MatchConfig
    move_to: str
    category: str | None = None


class DefaultRouting(StrictModel):
    keep_in_inbox: bool = True
    category: str | None = None
    section: str = "Others"


class ThreadingConfig(StrictModel):
    enabled: bool = False


class TriageConfig(StrictModel):
    version: Literal[1]
    threading: ThreadingConfig = Field(default_factory=ThreadingConfig)
    annotations: list[AnnotationRule] = Field(default_factory=list)
    routes: list[RouteRule]
    default: DefaultRouting

    @model_validator(mode="after")
    def validate_rule_ids(self) -> TriageConfig:
        ids = [rule.id for rule in [*self.annotations, *self.routes]]
        if len(ids) != len(set(ids)):
            raise ValueError("triage rule IDs must be unique")
        return self


@dataclass(frozen=True)
class TriageContext:
    mail: MailContext
    config: TriageConfig
    fingerprint: str


def load_triage_context(directory: Path | None = None) -> TriageContext:
    root = (directory or config_dir()).resolve()
    mail = load_mail_context(root)
    triage = load_yaml_model(root / "triage.yaml", TriageConfig)
    folders = set(mail.folders.folders)
    if missing := sorted({route.move_to for route in triage.routes} - folders):
        raise ValueError(f"triage routes reference undefined folders: {missing}")
    referenced_groups = {
        rule.when.sender_group
        for rule in [*triage.annotations, *triage.routes]
        if rule.when.sender_group
    }
    if missing := sorted(referenced_groups - set(mail.definitions.groups)):
        raise ValueError(f"triage rules reference undefined people groups: {missing}")
    referenced_lists = {
        rule.when.distribution_list_group
        for rule in [*triage.annotations, *triage.routes]
        if rule.when.distribution_list_group
    }
    if missing := sorted(referenced_lists - set(mail.definitions.distribution_list_groups)):
        raise ValueError(f"triage rules reference undefined distribution groups: {missing}")
    canonical = json.dumps(
        {
            "triage": triage.model_dump(mode="json"),
            "mail_fingerprint": mail.fingerprint,
        },
        sort_keys=True,
        separators=(",", ":"),
    )
    return TriageContext(
        mail=mail,
        config=triage,
        fingerprint=hashlib.sha256(canonical.encode()).hexdigest()[:16],
    )
