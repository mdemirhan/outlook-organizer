from __future__ import annotations

import uuid
from datetime import UTC, datetime

from outlook_organizer.config import AppConfig
from outlook_organizer.models import (
    MailMessage,
    MessageFacts,
    PlannedMessageAction,
    TriagePlan,
)
from outlook_organizer.rules.engine import RuleEngine
from outlook_organizer.rules.facts import FactBuilder


class MailTriagePlanner:
    def __init__(self, config: AppConfig) -> None:
        self.config = config
        self.fact_builder = FactBuilder(config.definitions)
        self.engine = RuleEngine()

    def validate(self) -> None:
        # Pydantic validates match fields and local folder references. Cross-file
        # group references are validated by OutlookOrganizerService.validate().
        return None

    def create_plan(self, messages: list, folder_id: int) -> TriagePlan:
        actions = [self.plan_message(message) for message in messages]
        return TriagePlan(
            plan_id=f"plan-{uuid.uuid4().hex[:12]}",
            created_at=datetime.now(UTC),
            config_fingerprint=self.config.fingerprint,
            folder_id=folder_id,
            actions=actions,
            dry_run=True,
        )

    def plan_message(self, message: MailMessage) -> PlannedMessageAction:
        """Classify one current Outlook message into its complete desired action."""
        return self._plan_message(self.fact_builder.build(message))

    def _plan_message(self, facts: MessageFacts) -> PlannedMessageAction:
        add_categories: list[str] = []
        keep_in_inbox = False
        report_section: str | None = None
        matches = []

        for order, annotation in enumerate(self.config.mail.annotations):
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
                add_categories.append(annotation.add_category)
            if report_section is None and annotation.section:
                report_section = annotation.section
            keep_in_inbox = keep_in_inbox or annotation.keep_in_inbox

        move_to: str | None = None
        matched_route = False
        route_order_offset = len(self.config.mail.annotations)
        for order, route in enumerate(self.config.mail.routes, start=route_order_offset):
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
            matched_route = True
            move_to = route.move_to
            if route.category:
                add_categories.append(route.category)
            if report_section is None:
                report_section = self.config.mail.folders[route.move_to].name
            break

        if not matched_route:
            default = self.config.mail.default
            keep_in_inbox = keep_in_inbox or default.keep_in_inbox
            if default.category:
                add_categories.append(default.category)
            if report_section is None:
                report_section = default.section

        if keep_in_inbox:
            move_to = None

        return PlannedMessageAction(
            message_id=facts.message.stable_id,
            outlook_id=facts.message.outlook_id,
            subject=facts.message.subject,
            sender_name=facts.message.sender_name,
            sender_address=facts.message.sender_address,
            received_at=facts.message.received_at,
            add_categories=sorted(set(add_categories)),
            remove_categories=[],
            move_to=move_to,
            set_flag=None,
            report_section=report_section or "Needs review",
            keep_in_inbox=keep_in_inbox,
            matches=matches,
            domain_class=facts.domain_class,
        )
