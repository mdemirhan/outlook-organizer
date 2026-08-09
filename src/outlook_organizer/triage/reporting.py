from __future__ import annotations

from collections import Counter
from dataclasses import asdict
from typing import Any

from outlook_organizer.mail import DomainClass
from outlook_organizer.triage.config import TriageContext
from outlook_organizer.triage.models import ExecutionResult, TriageAssessment


def build_triage_report(
    assessment: TriageAssessment,
    context: TriageContext,
    *,
    execution: ExecutionResult | None = None,
) -> dict[str, Any]:
    sections: dict[str, list[dict[str, Any]]] = {}
    for index, decision in enumerate(assessment.decisions, start=1):
        destination = (
            context.mail.folders.folders[decision.move_to].name if decision.move_to else None
        )
        sections.setdefault(decision.report_section, []).append(
            {
                "index": index,
                "outlook_id": decision.outlook_id,
                "subject": decision.subject,
                "sender_name": decision.sender_name,
                "sender_address": decision.sender_address,
                "received_at": decision.received_at,
                "domain_class": decision.domain_class.value,
                "categories_to_add": decision.add_categories,
                "move_to": destination,
                "keep_in_inbox": decision.keep_in_inbox,
                "matched_rules": [match.rule_id for match in decision.matches],
            }
        )
    routes = Counter(
        context.mail.folders.folders[decision.move_to].name if decision.move_to else "Inbox"
        for decision in assessment.decisions
    )
    categories = Counter(
        category for decision in assessment.decisions for category in decision.add_categories
    )
    thread_routed = sum(
        any(match.rule_id == "thread-affinity" for match in decision.matches)
        for decision in assessment.decisions
    )
    return {
        "created_at": assessment.created_at.isoformat(),
        "dry_run": execution is None,
        "summary": {
            "messages": len(assessment.decisions),
            "proposed_moves": sum(
                decision.move_to is not None for decision in assessment.decisions
            ),
            "kept_in_inbox": sum(decision.keep_in_inbox for decision in assessment.decisions),
            "possible_spam": sum(
                decision.domain_class
                in {DomainClass.JUNK_EXTERNAL, DomainClass.UNCLASSIFIED_EXTERNAL}
                for decision in assessment.decisions
            ),
            "thread_routed": thread_routed,
        },
        "action_summary": {
            "routes": dict(routes),
            "categories": dict(categories),
        },
        "sections": sections,
        "execution": asdict(execution) if execution else None,
    }
