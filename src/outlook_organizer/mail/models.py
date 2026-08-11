from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


def normalize_domains(values: list[str]) -> list[str]:
    normalized: set[str] = set()
    for value in values:
        domain = value.strip().lower().rstrip(".")
        if "@" in domain or not domain:
            raise ValueError("domains must be bare DNS names")
        normalized.add(domain.encode("idna").decode("ascii"))
    return sorted(normalized)


def normalize_addresses(values: list[str]) -> list[str]:
    normalized = sorted({value.strip().lower() for value in values})
    if any("@" not in value for value in normalized):
        raise ValueError("addresses must contain complete email addresses")
    return normalized


class IdentityConfig(StrictModel):
    name: str
    addresses: list[str]

    @field_validator("addresses")
    @classmethod
    def normalize_identity_addresses(cls, values: list[str]) -> list[str]:
        return normalize_addresses(values)


class ExternalSendersConfig(StrictModel):
    domains: list[str] = Field(default_factory=list)
    addresses: list[str] = Field(default_factory=list)

    @field_validator("domains")
    @classmethod
    def normalize_external_domains(cls, values: list[str]) -> list[str]:
        return normalize_domains(values)

    @field_validator("addresses")
    @classmethod
    def normalize_external_addresses(cls, values: list[str]) -> list[str]:
        return normalize_addresses(values)


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
        return normalize_domains(values)

    @field_validator("groups", "distribution_list_groups")
    @classmethod
    def normalize_group_addresses(cls, values: dict[str, list[str]]) -> dict[str, list[str]]:
        return {name: normalize_addresses(addresses) for name, addresses in values.items()}

    @model_validator(mode="after")
    def validate_unique_classifications(self) -> MailDefinitionsConfig:
        internal = set(self.internal_domains)
        if overlap := internal & set(self.safe_external.domains):
            raise ValueError(
                f"domains cannot be both internal and safe external: {sorted(overlap)}"
            )
        if overlap := internal & set(self.junk_external.domains):
            raise ValueError(
                f"domains cannot be both internal and junk external: {sorted(overlap)}"
            )
        if overlap := set(self.safe_external.domains) & set(self.junk_external.domains):
            raise ValueError(f"domains cannot be both safe and junk: {sorted(overlap)}")
        if overlap := set(self.safe_external.addresses) & set(self.junk_external.addresses):
            raise ValueError(f"addresses cannot be both safe and junk: {sorted(overlap)}")
        for collection_name, groups in (
            ("people", self.groups),
            ("distribution", self.distribution_list_groups),
        ):
            owners: dict[str, str] = {}
            for group_name, addresses in groups.items():
                for address in addresses:
                    previous = owners.get(address)
                    if previous and previous != group_name:
                        raise ValueError(
                            f"{collection_name} address {address} belongs to multiple groups: "
                            f"{previous}, {group_name}"
                        )
                    owners[address] = group_name
        return self


class FolderConfig(StrictModel):
    name: str
    id: int = Field(ge=1)
    aliases: list[str] = Field(default_factory=list)
    parent: str | None = None

    @property
    def names(self) -> list[str]:
        return [self.name, *self.aliases]


class FolderCatalogConfig(StrictModel):
    version: Literal[1]
    scan_limit: int = Field(default=1000, ge=10, le=100_000)
    folders: dict[str, FolderConfig]

    @model_validator(mode="after")
    def validate_catalog(self) -> FolderCatalogConfig:
        required = {"inbox"}
        if missing := sorted(required - set(self.folders)):
            raise ValueError(f"mail-folders.yaml is missing required folders: {missing}")
        ids = [folder.id for folder in self.folders.values()]
        if len(ids) != len(set(ids)):
            raise ValueError("folder IDs must be unique")
        for key, folder in self.folders.items():
            if folder.parent == key:
                raise ValueError(f"folder {key!r} cannot be its own parent")
            if folder.parent and folder.parent not in self.folders:
                raise ValueError(f"folder {key!r} references undefined parent {folder.parent!r}")
            ancestors = {key}
            parent = folder.parent
            while parent:
                if parent in ancestors:
                    raise ValueError(f"folder parent cycle includes {parent!r}")
                ancestors.add(parent)
                parent = self.folders[parent].parent
        return self


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
    is_read: bool = True
    replied_to: bool = False

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
class OutlookFolder:
    outlook_id: int
    name: str
    parent_name: str
    message_count: int
