from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

from outlook_organizer.mail.models import DomainClass, FlagStatus


@dataclass(slots=True)
class RuleMatch:
    rule_id: str
    description: str
    priority: int
    reasons: list[str]


@dataclass(slots=True)
class TriageDecision:
    message_id: str
    outlook_id: int
    subject: str
    sender_name: str
    sender_address: str
    received_at: str
    add_categories: list[str]
    remove_categories: list[str]
    move_to: str | None
    set_flag: FlagStatus | None
    report_section: str
    keep_in_inbox: bool
    matches: list[RuleMatch]
    domain_class: DomainClass


@dataclass(slots=True)
class TriageAssessment:
    created_at: datetime
    config_fingerprint: str
    folder_id: int
    decisions: list[TriageDecision]


@dataclass(slots=True)
class ExecutionResult:
    run_id: str | None
    applied: int
    status: str
    error: str | None = None
    thread_routed: int = 0


@dataclass(slots=True)
class ThreadResolution:
    decisions: list[TriageDecision]
    canonical_routes: dict[str, str] = field(default_factory=dict)
