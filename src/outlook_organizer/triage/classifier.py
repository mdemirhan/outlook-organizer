from __future__ import annotations

from datetime import UTC, datetime

from outlook_organizer.mail import FactBuilder, MailMessage, MessageFacts
from outlook_organizer.triage.config import MatchConfig, TriageContext
from outlook_organizer.triage.models import (
    RuleMatch,
    TriageAssessment,
    TriageDecision,
)


class RuleEngine:
    @staticmethod
    def match(
        facts: MessageFacts,
        *,
        rule_id: str,
        description: str,
        condition: MatchConfig,
        order: int,
    ) -> RuleMatch | None:
        reasons: list[str] = []
        if condition.flagged is not None:
            actual = facts.message.flag_status.value == "flagged"
            if actual != condition.flagged:
                return None
            reasons.append(f"flagged={actual}")
        if condition.recipient is not None:
            expected = (
                {condition.recipient}
                if isinstance(condition.recipient, str)
                else set(condition.recipient)
            )
            if facts.directness.value not in expected:
                return None
            reasons.append(f"recipient={facts.directness.value}")
        if condition.sender_group is not None:
            if facts.sender_relationship != condition.sender_group:
                return None
            reasons.append(f"sender_group={condition.sender_group}")
        if condition.sender_type is not None:
            if facts.domain_class.value != condition.sender_type:
                return None
            reasons.append(f"sender_type={condition.sender_type}")
        if condition.distribution_list_group is not None:
            if condition.distribution_list_group not in facts.distribution_list_groups:
                return None
            reasons.append(f"distribution_list_group={condition.distribution_list_group}")
        if condition.distribution_list is not None:
            if facts.has_distribution_list != condition.distribution_list:
                return None
            reasons.append(f"distribution_list={condition.distribution_list}")
        if condition.distribution_delivery is not None:
            if facts.delivered_via_distribution_list != condition.distribution_delivery:
                return None
            reasons.append(f"distribution_delivery={condition.distribution_delivery}")
        return RuleMatch(rule_id, description, order, reasons)


class TriageClassifier:
    def __init__(self, context: TriageContext) -> None:
        self.context = context
        self.fact_builder = FactBuilder(context.mail.definitions)
        self.engine = RuleEngine()

    def assess(self, messages: list[MailMessage], folder_id: int) -> TriageAssessment:
        return TriageAssessment(
            created_at=datetime.now(UTC),
            config_fingerprint=self.context.fingerprint,
            folder_id=folder_id,
            decisions=[self.classify(message) for message in messages],
        )

    def classify(self, message: MailMessage) -> TriageDecision:
        facts = self.fact_builder.build(message)
        config = self.context.config
        categories: list[str] = []
        keep_in_inbox = False
        section: str | None = None
        matches: list[RuleMatch] = []
        for order, annotation in enumerate(config.annotations):
            match = self.engine.match(
                facts,
                rule_id=annotation.id,
                description=annotation.description,
                condition=annotation.when,
                order=order,
            )
            if match is None:
                continue
            matches.append(match)
            if annotation.add_category:
                categories.append(annotation.add_category)
            section = section or annotation.section
            keep_in_inbox = keep_in_inbox or annotation.keep_in_inbox
        move_to: str | None = None
        route_offset = len(config.annotations)
        for order, route in enumerate(config.routes, start=route_offset):
            match = self.engine.match(
                facts,
                rule_id=route.id,
                description=route.description,
                condition=route.when,
                order=order,
            )
            if match is None:
                continue
            matches.append(match)
            move_to = route.move_to
            if route.category:
                categories.append(route.category)
            section = section or self.context.mail.folders.folders[route.move_to].name
            break
        else:
            default = config.default
            keep_in_inbox = keep_in_inbox or default.keep_in_inbox
            if default.category:
                categories.append(default.category)
            section = section or default.section
        if keep_in_inbox:
            move_to = None
        return TriageDecision(
            message_id=message.stable_id,
            outlook_id=message.outlook_id,
            subject=message.subject,
            sender_name=message.sender_name,
            sender_address=message.sender_address,
            received_at=message.received_at,
            add_categories=sorted(set(categories)),
            remove_categories=[],
            move_to=move_to,
            set_flag=None,
            report_section=section or "Needs review",
            keep_in_inbox=keep_in_inbox,
            matches=matches,
            domain_class=facts.domain_class,
        )
