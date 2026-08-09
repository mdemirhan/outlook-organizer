from __future__ import annotations

from outlook_organizer.models import Directness, DomainClass, Recipient
from outlook_organizer.rules.facts import FactBuilder


def test_only_me_directness_and_internal_sender(app_config, direct_message) -> None:
    facts = FactBuilder(app_config.definitions).build(direct_message)
    assert facts.directness is Directness.ONLY_ME
    assert facts.only_me
    assert facts.domain_class is DomainClass.INTERNAL
    assert not facts.has_distribution_list


def test_only_me_ignores_duplicate_identity_recipients(
    app_config, direct_message
) -> None:
    direct_message.to = [
        Recipient("Example User", "example.user@corp.example"),
        Recipient("Example User", "EXAMPLE.USER@CORP.EXAMPLE"),
        Recipient("Example Alias", "example.alias@corp.example"),
    ]

    facts = FactBuilder(app_config.definitions).build(direct_message)

    assert facts.directness is Directness.ONLY_ME
    assert facts.only_me
    assert facts.visible_recipient_count == 2


def test_distribution_list_detection(app_config, direct_message) -> None:
    direct_message.to = [
        Recipient(
            "Company Announcements",
            "announcements@corp.example",
            "public group address",
        )
    ]
    facts = FactBuilder(app_config.definitions).build(direct_message)
    assert facts.has_distribution_list
    assert facts.delivered_via_distribution_list
    assert facts.distribution_lists == ["Company Announcements"]
    assert facts.distribution_list_groups == ["company_announcements"]


def test_configured_distribution_list_sender_detection(
    app_config, direct_message
) -> None:
    direct_message.sender_name = "Performance Monitoring"
    direct_message.sender_address = "monitoring@corp.example"

    facts = FactBuilder(app_config.definitions).build(direct_message)

    assert facts.has_distribution_list
    assert facts.delivered_via_distribution_list
    assert facts.distribution_lists == ["Performance Monitoring"]
    assert facts.distribution_list_groups == ["company_announcements"]


def test_outlook_detected_unconfigured_distribution_list_has_no_named_group(
    app_config, direct_message
) -> None:
    direct_message.to = [
        Recipient(
            "Unconfigured list",
            "unconfigured-list@corp.example",
            "public group address",
        )
    ]
    facts = FactBuilder(app_config.definitions).build(direct_message)
    assert facts.has_distribution_list
    assert facts.delivered_via_distribution_list
    assert facts.distribution_list_groups == []


def test_group_copied_on_direct_mail_is_context_not_delivery(
    app_config, direct_message
) -> None:
    direct_message.cc = [
        Recipient(
            "Project team",
            "project-team@corp.example",
            "public group address",
        )
    ]

    facts = FactBuilder(app_config.definitions).build(direct_message)

    assert facts.has_distribution_list
    assert facts.distribution_lists == ["Project team"]
    assert not facts.delivered_via_distribution_list


def test_group_in_to_is_not_delivery_when_owner_is_also_directly_in_to(
    app_config, direct_message
) -> None:
    direct_message.to.append(
        Recipient(
            "Project team",
            "project-team@corp.example",
            "public group address",
        )
    )

    facts = FactBuilder(app_config.definitions).build(direct_message)

    assert facts.has_distribution_list
    assert not facts.delivered_via_distribution_list


def test_unclassified_external_is_classified_for_routing(
    app_config, direct_message
) -> None:
    direct_message.sender_address = "sender@unclassified-external.example"
    facts = FactBuilder(app_config.definitions).build(direct_message)
    assert facts.domain_class is DomainClass.UNCLASSIFIED_EXTERNAL
