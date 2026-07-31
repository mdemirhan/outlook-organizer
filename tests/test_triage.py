from __future__ import annotations

import pytest

from outlook_organizer.models import FlagStatus, Recipient
from outlook_organizer.rules.triage import MailTriagePlanner
from outlook_organizer.serialization import plan_from_dict, plan_to_dict


def test_flagged_rule_creates_action(app_config, direct_message) -> None:
    direct_message.flag_status = FlagStatus.FLAGGED
    planner = MailTriagePlanner(app_config)
    plan = planner.create_plan([direct_message], direct_message.folder_id)
    action = plan.actions[0]
    assert "@Action" in action.add_categories
    assert action.keep_in_inbox
    assert action.matches[0].rule_id == "flagged-needs-action"


def test_unknown_external_policy_survives_default_rule(app_config, direct_message) -> None:
    direct_message.sender_address = "sender@untrusted.example"
    planner = MailTriagePlanner(app_config)
    action = planner.create_plan([direct_message], direct_message.folder_id).actions[0]
    assert "@Untrusted External" in action.add_categories
    assert "@Only Me" in action.add_categories
    assert action.report_section == "Untrusted External"
    assert not action.keep_in_inbox
    assert action.move_to == "untrusted_external"


@pytest.mark.parametrize(
    ("sender_address", "subject"),
    [
        ("sender@unknown-external.example", "Monthly Newsletter"),
        ("product-newsletter@unknown-external.example", "Product update"),
    ],
)
def test_unknown_external_junk_keywords_route_to_junk_external(
    app_config, direct_message, sender_address, subject
) -> None:
    direct_message.sender_address = sender_address
    direct_message.subject = subject
    action = (
        MailTriagePlanner(app_config)
        .create_plan([direct_message], direct_message.folder_id)
        .actions[0]
    )
    assert action.move_to == "junk_external"
    assert action.report_section == "Junk External"
    assert "@Junk External" in action.add_categories
    assert "route-junk-external" in [match.rule_id for match in action.matches]


def test_known_junk_sender_routes_to_configured_folder(app_config, direct_message) -> None:
    direct_message.sender_address = "sender@unwanted.example"
    action = (
        MailTriagePlanner(app_config)
        .create_plan([direct_message], direct_message.folder_id)
        .actions[0]
    )
    assert action.move_to == "junk_external"
    assert action.report_section == "Junk External"
    assert action.domain_class.value == "junk_external"
    assert "@Junk External" in action.add_categories
    assert app_config.mail.folders[action.move_to].id == 111


def test_routine_internal_mail_moves_to_internal_general(app_config, direct_message) -> None:
    planner = MailTriagePlanner(app_config)
    action = planner.create_plan([direct_message], direct_message.folder_id).actions[0]
    assert action.move_to == "internal_general"
    assert "@Internal General" in action.add_categories
    assert not action.keep_in_inbox


@pytest.mark.parametrize(
    "address",
    [
        "announcements@corp.example",
        "monitoring@corp.example",
    ],
)
def test_known_company_distribution_lists_are_routed(
    app_config, direct_message, address
) -> None:
    direct_message.to = [Recipient(address, address)]
    action = (
        MailTriagePlanner(app_config)
        .create_plan([direct_message], direct_message.folder_id)
        .actions[0]
    )
    assert action.move_to == "company_announcements"
    assert "@Company Announcements" in action.add_categories


@pytest.mark.parametrize(
    "address",
    [
        "announcements@corp.example",
        "monitoring@corp.example",
    ],
)
def test_known_company_distribution_list_senders_are_routed(
    app_config, direct_message, address
) -> None:
    direct_message.sender_name = address
    direct_message.sender_address = address

    action = (
        MailTriagePlanner(app_config)
        .create_plan([direct_message], direct_message.folder_id)
        .actions[0]
    )

    assert action.move_to == "company_announcements"
    assert "@Company Announcements" in action.add_categories
    assert [match.rule_id for match in action.matches][-1] == (
        "route-company-announcements"
    )


def test_unconfigured_internal_distribution_list_uses_fallback_route(
    app_config, direct_message
) -> None:
    direct_message.to = [
        Recipient(
            "Unconfigured list",
            "unconfigured-list@corp.example",
            "public group address",
        )
    ]
    action = (
        MailTriagePlanner(app_config)
        .create_plan([direct_message], direct_message.folder_id)
        .actions[0]
    )
    assert action.move_to == "company_announcements"
    assert [match.rule_id for match in action.matches][-1] == (
        "route-other-internal-distribution"
    )


def test_unmatched_mail_uses_others_category(app_config, direct_message) -> None:
    direct_message.sender_address = "not-an-email-address"
    action = (
        MailTriagePlanner(app_config)
        .create_plan([direct_message], direct_message.folder_id)
        .actions[0]
    )
    assert action.move_to is None
    assert action.keep_in_inbox
    assert action.report_section == "Others"
    assert "@Others" in action.add_categories


def test_leadership_takes_precedence_over_other_internal(app_config, direct_message) -> None:
    direct_message.sender_address = "leader@corp.example"
    action = (
        MailTriagePlanner(app_config)
        .create_plan([direct_message], direct_message.folder_id)
        .actions[0]
    )
    assert action.move_to == "leadership"
    assert "@Leadership" in action.add_categories
    assert [match.rule_id for match in action.matches][-1] == "route-leadership"


def test_team_member_routes_to_my_team(app_config, direct_message) -> None:
    direct_message.sender_address = "teammate@corp.example"
    action = (
        MailTriagePlanner(app_config)
        .create_plan([direct_message], direct_message.folder_id)
        .actions[0]
    )
    assert action.move_to == "my_team"
    assert "@My Team" in action.add_categories


def test_plan_round_trip(app_config, direct_message) -> None:
    plan = MailTriagePlanner(app_config).create_plan([direct_message], direct_message.folder_id)
    restored = plan_from_dict(plan_to_dict(plan))
    assert restored.plan_id == plan.plan_id
    assert restored.actions[0].outlook_id == direct_message.outlook_id


def test_legacy_digest_section_is_still_readable(app_config, direct_message) -> None:
    plan = MailTriagePlanner(app_config).create_plan([direct_message], direct_message.folder_id)
    payload = plan_to_dict(plan)
    action = payload["actions"][0]
    action["digest_section"] = action.pop("report_section")

    restored = plan_from_dict(payload)

    assert restored.actions[0].report_section == plan.actions[0].report_section
