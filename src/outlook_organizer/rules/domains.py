from __future__ import annotations

from dataclasses import dataclass

from outlook_organizer.models import DomainClass, MailDefinitionsConfig


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
    if candidate == configured_domain:
        return True
    return candidate.endswith("." + configured_domain)


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

        for configured_domain in self.config.internal_domains:
            if domain_matches(domain, configured_domain):
                return DomainClassification(
                    domain, DomainClass.INTERNAL, configured_domain
                )

        if normalized_address in self.config.junk_external.addresses:
            return DomainClassification(domain, DomainClass.JUNK_EXTERNAL)
        for configured_domain in self.config.junk_external.domains:
            if domain_matches(domain, configured_domain):
                return DomainClassification(
                    domain, DomainClass.JUNK_EXTERNAL, configured_domain
                )

        if normalized_address in self.config.safe_external.addresses:
            return DomainClassification(domain, DomainClass.SAFE_EXTERNAL)
        for configured_domain in self.config.safe_external.domains:
            if domain_matches(domain, configured_domain):
                return DomainClassification(
                    domain, DomainClass.SAFE_EXTERNAL, configured_domain
                )

        searchable_header = f"{subject}\n{normalized_address}".casefold()
        if any(
            keyword in searchable_header for keyword in self.config.junk_external.keywords
        ):
            return DomainClassification(domain, DomainClass.JUNK_EXTERNAL)

        return DomainClassification(domain, DomainClass.UNCLASSIFIED_EXTERNAL)
