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
        to_addresses = [normalize_email(recipient.address) for recipient in message.to]
        cc_addresses = [normalize_email(recipient.address) for recipient in message.cc]
        includes_me = any(address in self.me for address in [*to_addresses, *cc_addresses])
        only_me = len(to_addresses) == 1 and to_addresses[0] in self.me and not cc_addresses
        direct_to_me = len(to_addresses) == 1 and to_addresses[0] in self.me and bool(cc_addresses)

        if only_me:
            directness = Directness.ONLY_ME
        elif direct_to_me:
            directness = Directness.DIRECT_TO_ME
        elif len(visible) > 1:
            directness = Directness.MULTI_RECIPIENT
        elif includes_me:
            directness = Directness.DIRECT_TO_ME
        elif visible:
            directness = Directness.NOT_TO_ME
        else:
            directness = Directness.UNKNOWN

        distribution_lists: list[str] = []
        distribution_list_groups: list[str] = []
        configured_sender_list = self.known_lists.get(sender_address)
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

        return MessageFacts(
            message=message,
            sender_domain=classification.domain,
            domain_class=classification.domain_class,
            sender_relationship=relationship,
            directness=directness,
            visible_recipient_count=len(visible),
            includes_me=includes_me,
            only_me=only_me,
            direct_to_me=direct_to_me,
            has_distribution_list=bool(distribution_lists),
            distribution_lists=distribution_lists,
            distribution_list_groups=sorted(set(distribution_list_groups)),
        )
