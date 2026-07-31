from __future__ import annotations

from outlook_organizer.models import DomainClass, MailDefinitionsConfig, Relationship
from outlook_organizer.rules.domains import normalize_email


class RelationshipResolver:
    def __init__(self, config: MailDefinitionsConfig) -> None:
        self._address_to_group: dict[str, str] = {}
        for group_name, addresses in config.groups.items():
            for address in addresses:
                self._address_to_group[normalize_email(address)] = group_name

    def resolve(self, address: str, domain_class: DomainClass) -> str:
        normalized = normalize_email(address)
        if normalized in self._address_to_group:
            return self._address_to_group[normalized]
        if domain_class is DomainClass.SAFE_EXTERNAL:
            return Relationship.KNOWN_EXTERNAL.value
        return Relationship.UNKNOWN.value

    def address_in_group(self, address: str, group: str) -> bool:
        return self._address_to_group.get(normalize_email(address)) == group
