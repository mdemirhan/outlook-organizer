from __future__ import annotations

import pytest

from outlook_organizer.models import DomainClass
from outlook_organizer.rules.domains import DomainClassifier, domain_matches, extract_domain


def test_configured_root_and_subdomains_are_internal(app_config) -> None:
    classifier = DomainClassifier(app_config.definitions)
    assert classifier.classify("person@corp.example").domain_class is DomainClass.INTERNAL
    assert (
        classifier.classify("person@division.corp.example").domain_class is DomainClass.INTERNAL
    )
    assert (
        classifier.classify("group@sub.corp.invalid").domain_class
        is DomainClass.UNKNOWN_EXTERNAL
    )


def test_lookalike_domains_are_not_internal(app_config) -> None:
    classifier = DomainClassifier(app_config.definitions)
    for address in (
        "person@evil-corp.example",
        "person@corp.example.attacker.example",
        "person@notcorp.example",
    ):
        assert classifier.classify(address).domain_class is DomainClass.UNKNOWN_EXTERNAL


def test_boundary_aware_safe_external_match() -> None:
    assert domain_matches("partner.example", "partner.example")
    assert domain_matches("mail.partner.example", "partner.example")
    assert not domain_matches("evil-partner.example", "partner.example")


def test_exact_safe_external_sender_does_not_trust_whole_domain(app_config) -> None:
    classifier = DomainClassifier(app_config.definitions)
    assert (
        classifier.classify("trusted.sender@public.example").domain_class
        is DomainClass.SAFE_EXTERNAL
    )
    assert (
        classifier.classify("someone-else@public.example").domain_class
        is DomainClass.UNKNOWN_EXTERNAL
    )


def test_safe_external_newsletter_is_not_reclassified_as_junk(app_config) -> None:
    classifier = DomainClassifier(app_config.definitions)
    classification = classifier.classify(
        "newsletter@trusted-partner.example", "Monthly newsletter"
    )
    assert classification.domain_class is DomainClass.SAFE_EXTERNAL


@pytest.mark.parametrize(
    "address",
    [
        "sender@unwanted.example",
        "sender@updates.unwanted.example",
        "marketing@otherwise-valid.example",
    ],
)
def test_configured_junk_domains_and_addresses_are_known_junk(app_config, address) -> None:
    classifier = DomainClassifier(app_config.definitions)
    assert classifier.classify(address).domain_class is DomainClass.JUNK_EXTERNAL


@pytest.mark.parametrize(
    "address",
    [
        "someone-else@otherwise-valid.example",
    ],
)
def test_exact_junk_sender_does_not_mark_whole_domain_as_junk(app_config, address) -> None:
    classifier = DomainClassifier(app_config.definitions)
    assert classifier.classify(address).domain_class is DomainClass.UNKNOWN_EXTERNAL


def test_extract_domain_normalizes_case_and_invalid_values() -> None:
    assert extract_domain("Person@CORP.EXAMPLE.") == "corp.example"
    assert extract_domain("not-an-address") is None
