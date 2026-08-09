from __future__ import annotations

import pytest

from outlook_organizer.mail import FlagStatus, Recipient
from outlook_organizer.triage.classifier import TriageClassifier


def decision(triage_context, message):
    return TriageClassifier(triage_context).classify(message)


def test_flagged_mail_stays_in_inbox(triage_context, direct_message) -> None:
    direct_message.flag_status = FlagStatus.FLAGGED
    result = decision(triage_context, direct_message)
    assert result.keep_in_inbox
    assert result.move_to is None
    assert "@Action" in result.add_categories


def test_internal_mail_routes_to_general(triage_context, direct_message) -> None:
    result = decision(triage_context, direct_message)
    assert result.move_to == "internal_general"
    assert result.add_categories == ["@Internal General", "@Only Me"]


def test_leadership_route_precedes_general(triage_context, direct_message) -> None:
    direct_message.sender_address = "leader@corp.example"
    result = decision(triage_context, direct_message)
    assert result.move_to == "leadership"
    assert result.matches[-1].rule_id == "route-leadership"


@pytest.mark.parametrize("address", ["announcements@corp.example", "monitoring@corp.example"])
def test_known_distribution_lists_route_explicitly(triage_context, direct_message, address) -> None:
    direct_message.to = [Recipient(address, address)]
    result = decision(triage_context, direct_message)
    assert result.move_to == "company_announcements"


def test_unclassified_sender_uses_safety_folder(triage_context, direct_message) -> None:
    direct_message.sender_address = "sender@unknown.example"
    result = decision(triage_context, direct_message)
    assert result.move_to == "unclassified_external"
