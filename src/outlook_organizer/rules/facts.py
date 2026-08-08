from __future__ import annotations

from outlook_organizer.models import (
    Directness,
    MailDefinitionsConfig,
    MailMessage,
    MessageFacts,
)
from outlook_organizer.rules.domains import DomainClassifier, normalize_email
from outlook_organizer.rules.relationships import RelationshipResolver


class FactBuilder:
    def __init__(self, definitions: MailDefinitionsConfig) -> None:
        self.definitions = definitions
        self.domain_classifier = DomainClassifier(definitions)
        self.relationships = RelationshipResolver(definitions)
        self.me = {normalize_email(value) for value in definitions.identity.addresses}
        self.known_lists: dict[str, str] = {}
        for group_name, addresses in definitions.distribution_list_groups.items():
            for address in addresses:
                self.known_lists[normalize_email(address)] = group_name

    def build(self, message: MailMessage) -> MessageFacts:
        classification = self.domain_classifier.classify(
            message.sender_address, message.subject
        )
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
        direct_to_me = (
            bool(to_addresses) and to_addresses <= self.me and bool(cc_addresses)
        )

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
        distribution_list_groups: list[str] = []
        configured_sender_list = self.known_lists.get(sender_address)
        delivered_via_distribution_list = configured_sender_list is not None
        if configured_sender_list:
            distribution_lists.append(message.sender_name or sender_address)
            distribution_list_groups.append(configured_sender_list)

        for recipient in visible:
            address = normalize_email(recipient.address)
            type_text = recipient.kind.casefold()
            configured_list = self.known_lists.get(address)
            if (
                "public group" in type_text
                or "private group" in type_text
                or configured_list is not None
            ):
                distribution_lists.append(recipient.name or address)
                if configured_list:
                    distribution_list_groups.append(configured_list)

        # Outlook's group type tells us that an address is a group, but not
        # whether the message reached the owner through that group. Treat an
        # Outlook-detected group as the delivery path only when it is in To and
        # no configured identity is directly in To. A group copied on CC is
        # context on a direct message, not evidence of distribution delivery.
        identity_in_to = bool(to_addresses & self.me)
        group_in_to = any(
            "public group" in recipient.kind.casefold()
            or "private group" in recipient.kind.casefold()
            or normalize_email(recipient.address) in self.known_lists
            for recipient in message.to
        )
        delivered_via_distribution_list = delivered_via_distribution_list or (
            group_in_to and not identity_in_to
        )

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
            delivered_via_distribution_list=delivered_via_distribution_list,
            distribution_lists=distribution_lists,
            distribution_list_groups=sorted(set(distribution_list_groups)),
        )
