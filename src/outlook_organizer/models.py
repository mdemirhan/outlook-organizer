from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class IdentityConfig(StrictModel):
    name: str
    addresses: list[str]

    @field_validator("addresses")
    @classmethod
    def normalize_addresses(cls, values: list[str]) -> list[str]:
        return sorted({value.strip().lower() for value in values})


def _normalize_domains(values: list[str]) -> list[str]:
    normalized: set[str] = set()
    for value in values:
        domain = value.strip().lower().rstrip(".")
        if "@" in domain or not domain:
            raise ValueError("domains must be bare DNS names")
        normalized.add(domain.encode("idna").decode("ascii"))
    return sorted(normalized)


def _normalize_addresses(values: list[str]) -> list[str]:
    normalized = sorted({value.strip().lower() for value in values})
    if any("@" not in value for value in normalized):
        raise ValueError("addresses must contain complete email addresses")
    return normalized


class ExternalSendersConfig(StrictModel):
    domains: list[str] = Field(default_factory=list)
    addresses: list[str] = Field(default_factory=list)

    @field_validator("domains")
    @classmethod
    def normalize_domains(cls, values: list[str]) -> list[str]:
        return _normalize_domains(values)

    @field_validator("addresses")
    @classmethod
    def normalize_addresses(cls, values: list[str]) -> list[str]:
        return _normalize_addresses(values)


class JunkExternalConfig(ExternalSendersConfig):
    keywords: list[str] = Field(default_factory=list)

    @field_validator("keywords")
    @classmethod
    def normalize_keywords(cls, values: list[str]) -> list[str]:
        return sorted({value.strip().casefold() for value in values if value.strip()})


class MailDefinitionsConfig(StrictModel):
    version: Literal[2]
    identity: IdentityConfig
    internal_domains: list[str]
    groups: dict[str, list[str]] = Field(default_factory=dict)
    safe_external: ExternalSendersConfig = Field(default_factory=ExternalSendersConfig)
    junk_external: JunkExternalConfig = Field(default_factory=JunkExternalConfig)
    distribution_list_groups: dict[str, list[str]] = Field(default_factory=dict)

    @field_validator("internal_domains")
    @classmethod
    def normalize_internal_domains(cls, values: list[str]) -> list[str]:
        return _normalize_domains(values)

    @field_validator("groups", "distribution_list_groups")
    @classmethod
    def normalize_group_addresses(
        cls, values: dict[str, list[str]]
    ) -> dict[str, list[str]]:
        return {
            group_name: _normalize_addresses(addresses)
            for group_name, addresses in values.items()
        }

    @model_validator(mode="after")
    def validate_unique_classifications(self) -> MailDefinitionsConfig:
        internal_domains = set(self.internal_domains)
        conflicting_safe_domains = internal_domains & set(self.safe_external.domains)
        if conflicting_safe_domains:
            raise ValueError(
                "domains cannot be both internal and safe external: "
                f"{sorted(conflicting_safe_domains)}"
            )
        conflicting_junk_domains = internal_domains & set(self.junk_external.domains)
        if conflicting_junk_domains:
            raise ValueError(
                "domains cannot be both internal and junk external: "
                f"{sorted(conflicting_junk_domains)}"
            )
        overlapping_domains = set(self.safe_external.domains) & set(self.junk_external.domains)
        if overlapping_domains:
            raise ValueError(
                f"domains cannot be both safe and junk: {sorted(overlapping_domains)}"
            )
        overlapping_addresses = set(self.safe_external.addresses) & set(
            self.junk_external.addresses
        )
        if overlapping_addresses:
            raise ValueError(
                f"addresses cannot be both safe and junk: {sorted(overlapping_addresses)}"
            )

        group_addresses: dict[str, str] = {}
        for group_name, addresses in self.groups.items():
            for address in addresses:
                previous = group_addresses.get(address)
                if previous and previous != group_name:
                    raise ValueError(
                        f"{address} belongs to multiple groups: {previous}, {group_name}"
                    )
                group_addresses[address] = group_name

        distribution_addresses: dict[str, str] = {}
        for group_name, addresses in self.distribution_list_groups.items():
            for address in addresses:
                previous = distribution_addresses.get(address)
                if previous and previous != group_name:
                    raise ValueError(
                        f"{address} belongs to multiple distribution-list "
                        f"groups: {previous}, {group_name}"
                    )
                distribution_addresses[address] = group_name
        return self


class FolderConfig(StrictModel):
    name: str
    id: int = Field(ge=1)
    aliases: list[str] = Field(default_factory=list)
    parent: str | None = None

    @property
    def names(self) -> list[str]:
        return [self.name, *self.aliases]


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



class MailRulesConfig(StrictModel):
    version: Literal[2]
    folder_scan_limit: int = Field(default=1000, ge=10, le=100_000)
    threading: ThreadingConfig = Field(default_factory=ThreadingConfig)
    folders: dict[str, FolderConfig]
    annotations: list[AnnotationRule] = Field(default_factory=list)
    routes: list[RouteRule]
    default: DefaultRouting

    @model_validator(mode="after")
    def validate_references(self) -> MailRulesConfig:
        if "inbox" not in self.folders:
            raise ValueError("folders must define inbox")
        if "organized_primary" not in self.folders:
            raise ValueError("folders must define organized_primary")
        if "organized_secondary" not in self.folders:
            raise ValueError("folders must define organized_secondary")

        ids = [rule.id for rule in [*self.annotations, *self.routes]]
        if len(ids) != len(set(ids)):
            raise ValueError("rule IDs must be unique")

        duplicate_folder_ids = sorted(
            folder_id
            for folder_id in {folder.id for folder in self.folders.values()}
            if sum(folder.id == folder_id for folder in self.folders.values()) > 1
        )
        if duplicate_folder_ids:
            raise ValueError(f"folder IDs must be unique: {duplicate_folder_ids}")

        for key, folder in self.folders.items():
            if folder.parent == key:
                raise ValueError(f"folder {key!r} cannot be its own parent")
            if folder.parent and folder.parent not in self.folders:
                raise ValueError(
                    f"folder {key!r} references undefined parent {folder.parent!r}"
                )
            ancestors: set[str] = {key}
            parent = folder.parent
            while parent:
                if parent in ancestors:
                    raise ValueError(f"folder parent cycle includes {parent!r}")
                ancestors.add(parent)
                parent = self.folders[parent].parent
        for route in self.routes:
            if route.move_to not in self.folders:
                raise ValueError(
                    f"route {route.id!r} references undefined folder {route.move_to!r}"
                )
        return self


class CalendarPreferences(StrictModel):
    lunch_window: tuple[str, str]
    minimum_focus_block_minutes: int = Field(ge=15, le=480)
    meeting_buffer_minutes: int = Field(ge=0, le=120)
    maximum_meeting_hours_per_day: float = Field(ge=0, le=24)
    avoid_back_to_back_meetings: bool = True
    preferred_focus_windows: list[tuple[str, str]] = Field(default_factory=list)


class ProtectedRelationships(StrictModel):
    high_priority: list[str] = Field(default_factory=list)


class CalendarConfig(StrictModel):
    version: Literal[1]
    timezone: str
    calendar_names: list[str]
    maximum_calendar_id: int = Field(default=5000, ge=10, le=100_000)
    working_hours: dict[str, tuple[str, str]]
    preferences: CalendarPreferences
    protected_relationships: ProtectedRelationships


class DomainClass(StrEnum):
    INTERNAL = "internal"
    JUNK_EXTERNAL = "junk_external"
    SAFE_EXTERNAL = "safe_external"
    UNCLASSIFIED_EXTERNAL = "unclassified_external"
    UNKNOWN = "unknown"


class Relationship(StrEnum):
    MANAGER = "manager"
    TEAM = "team"
    PEER = "peer"
    LEADERSHIP = "leadership"
    KNOWN_EXTERNAL = "known_external"
    UNKNOWN = "unknown"


class Directness(StrEnum):
    ONLY_ME = "only_me"
    DIRECT_TO_ME = "direct_to_me"
    MULTI_RECIPIENT = "multi_recipient"
    NOT_TO_ME = "not_to_me"
    UNKNOWN = "unknown"


class FlagStatus(StrEnum):
    NOT_FLAGGED = "not_flagged"
    FLAGGED = "flagged"
    COMPLETED = "completed"


@dataclass(slots=True)
class Recipient:
    name: str
    address: str
    kind: str = "unknown"


@dataclass(slots=True)
class MailMessage:
    outlook_id: int
    exchange_id: str
    folder_id: int
    folder_name: str
    subject: str
    sender_name: str
    sender_address: str
    to: list[Recipient]
    cc: list[Recipient]
    received_at: str
    flag_status: FlagStatus
    categories: list[str]
    body: str = ""
    has_attachments: bool = False
    thread_guid: str = ""

    @property
    def stable_id(self) -> str:
        return self.exchange_id or f"outlook:{self.outlook_id}"


@dataclass(slots=True)
class MessageFacts:
    message: MailMessage
    sender_domain: str | None
    domain_class: DomainClass
    sender_relationship: str
    directness: Directness
    visible_recipient_count: int
    includes_me: bool
    only_me: bool
    direct_to_me: bool
    has_distribution_list: bool
    delivered_via_distribution_list: bool
    distribution_lists: list[str]
    distribution_list_groups: list[str] = field(default_factory=list)


@dataclass(slots=True)
class RuleMatch:
    rule_id: str
    description: str
    priority: int
    reasons: list[str]


@dataclass(slots=True)
class PlannedMessageAction:
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
class TriagePlan:
    plan_id: str
    created_at: datetime
    config_fingerprint: str
    folder_id: int
    actions: list[PlannedMessageAction]
    dry_run: bool = True
    thread_promotions: int = 0


@dataclass(slots=True)
class ThreadMessageState:
    outlook_id: int
    exchange_id: str
    folder_id: int
    folder_name: str
    thread_guid: str


@dataclass(slots=True)
class OutlookFolder:
    outlook_id: int
    name: str
    parent_name: str
    message_count: int


@dataclass(slots=True)
class CalendarInfo:
    outlook_id: int
    name: str
    event_count: int


@dataclass(slots=True)
class CalendarAttendee:
    name: str
    address: str
    attendee_type: str
    status: str


@dataclass(slots=True)
class CalendarEvent:
    outlook_id: int
    exchange_id: str
    calendar_id: int
    subject: str
    start_at: str
    end_at: str
    location: str
    organizer: str
    all_day: bool
    free_busy_status: str
    is_private: bool
    categories: list[str]
    attendees: list[CalendarAttendee]
    body: str = ""
