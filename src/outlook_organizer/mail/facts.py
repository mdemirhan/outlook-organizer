from __future__ import annotations

from dataclasses import dataclass

from outlook_organizer.mail.models import (
    Directness,
    DomainClass,
    MailDefinitionsConfig,
    MailMessage,
    MessageFacts,
    Relationship,
)


def normalize_email(value: str) -> str:
    return value.strip().lower()


def extract_domain(address: str) -> str | None:
    normalized = normalize_email(address)
    if "@" not in normalized:
        return None
    local, domain = normalized.rsplit("@", 1)
    domain = domain.rstrip(".")
    if not local or not domain:
        return None
    try:
        return domain.encode("idna").decode("ascii")
    except UnicodeError:
        return None


def domain_matches(candidate: str, configured_domain: str) -> bool:
    return candidate == configured_domain or candidate.endswith("." + configured_domain)


@dataclass(frozen=True)
class DomainClassification:
    domain: str | None
    domain_class: DomainClass
    matched_domain: str | None = None


class DomainClassifier:
    def __init__(self, config: MailDefinitionsConfig) -> None:
        self.config = config

    def classify(self, address: str, subject: str = "") -> DomainClassification:
        normalized_address = normalize_email(address)
        domain = extract_domain(address)
        if domain is None:
            return DomainClassification(None, DomainClass.UNKNOWN)
        for configured in self.config.internal_domains:
            if domain_matches(domain, configured):
                return DomainClassification(domain, DomainClass.INTERNAL, configured)
        if normalized_address in self.config.junk_external.addresses:
            return DomainClassification(domain, DomainClass.JUNK_EXTERNAL)
        for configured in self.config.junk_external.domains:
            if domain_matches(domain, configured):
                return DomainClassification(domain, DomainClass.JUNK_EXTERNAL, configured)
        if normalized_address in self.config.safe_external.addresses:
            return DomainClassification(domain, DomainClass.SAFE_EXTERNAL)
        for configured in self.config.safe_external.domains:
            if domain_matches(domain, configured):
                return DomainClassification(domain, DomainClass.SAFE_EXTERNAL, configured)
        searchable = f"{subject}\n{normalized_address}".casefold()
        if any(keyword in searchable for keyword in self.config.junk_external.keywords):
            return DomainClassification(domain, DomainClass.JUNK_EXTERNAL)
        return DomainClassification(domain, DomainClass.UNCLASSIFIED_EXTERNAL)


class RelationshipResolver:
    def __init__(self, config: MailDefinitionsConfig) -> None:
        self._address_to_group = {
            normalize_email(address): group
            for group, addresses in config.groups.items()
            for address in addresses
        }

    def resolve(self, address: str, domain_class: DomainClass) -> str:
        normalized = normalize_email(address)
        if normalized in self._address_to_group:
            return self._address_to_group[normalized]
        if domain_class is DomainClass.SAFE_EXTERNAL:
            return Relationship.KNOWN_EXTERNAL.value
        return Relationship.UNKNOWN.value


class FactBuilder:
    def __init__(self, definitions: MailDefinitionsConfig) -> None:
        self.domain_classifier = DomainClassifier(definitions)
        self.relationships = RelationshipResolver(definitions)
        self.me = {normalize_email(value) for value in definitions.identity.addresses}
        self.known_lists = {
            normalize_email(address): group
            for group, addresses in definitions.distribution_list_groups.items()
            for address in addresses
        }

    def build(self, message: MailMessage) -> MessageFacts:
        classification = self.domain_classifier.classify(message.sender_address, message.subject)
        relationship = self.relationships.resolve(
            message.sender_address, classification.domain_class
        )
        sender_address = normalize_email(message.sender_address)
        visible = [*message.to, *message.cc]
        to_addresses = {
            normalized
            for recipient in message.to
            if (normalized := normalize_email(recipient.address))
        }
        cc_addresses = {
            normalized
            for recipient in message.cc
            if (normalized := normalize_email(recipient.address))
        }
        visible_addresses = to_addresses | cc_addresses
        includes_me = bool(visible_addresses & self.me)
        only_me = bool(to_addresses) and to_addresses <= self.me and not cc_addresses
        direct_to_me = bool(to_addresses) and to_addresses <= self.me and bool(cc_addresses)
        if only_me:
            directness = Directness.ONLY_ME
        elif direct_to_me:
            directness = Directness.DIRECT_TO_ME
        elif len(visible_addresses) > 1:
            directness = Directness.MULTI_RECIPIENT
        elif includes_me:
            directness = Directness.DIRECT_TO_ME
        elif visible_addresses:
            directness = Directness.NOT_TO_ME
        else:
            directness = Directness.UNKNOWN

        distribution_lists: list[str] = []
        distribution_groups: list[str] = []
        configured_sender_list = self.known_lists.get(sender_address)
        delivered_via_list = configured_sender_list is not None
        if configured_sender_list:
            distribution_lists.append(message.sender_name or sender_address)
            distribution_groups.append(configured_sender_list)
        for recipient in visible:
            address = normalize_email(recipient.address)
            configured_list = self.known_lists.get(address)
            if "group" in recipient.kind.casefold() or configured_list is not None:
                distribution_lists.append(recipient.name or address)
                if configured_list:
                    distribution_groups.append(configured_list)
        identity_in_to = bool(to_addresses & self.me)
        group_in_to = any(
            "group" in recipient.kind.casefold()
            or normalize_email(recipient.address) in self.known_lists
            for recipient in message.to
        )
        delivered_via_list = delivered_via_list or (group_in_to and not identity_in_to)
        return MessageFacts(
            message=message,
            sender_domain=classification.domain,
            domain_class=classification.domain_class,
            sender_relationship=relationship,
            directness=directness,
            visible_recipient_count=len(visible_addresses),
            includes_me=includes_me,
            only_me=only_me,
            direct_to_me=direct_to_me,
            has_distribution_list=bool(distribution_lists),
            delivered_via_distribution_list=delivered_via_list,
            distribution_lists=distribution_lists,
            distribution_list_groups=sorted(set(distribution_groups)),
        )
