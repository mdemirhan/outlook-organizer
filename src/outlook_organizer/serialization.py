from __future__ import annotations

from datetime import datetime
from typing import Any

from outlook_organizer.models import (
    DomainClass,
    FlagStatus,
    PlannedMessageAction,
    RuleMatch,
    TriagePlan,
)


def _domain_class(value: str) -> DomainClass:
    legacy = {
        "known_junk": "junk_external",
        "untrusted_external": "unknown_external",
    }
    return DomainClass(legacy.get(value, value))


def plan_to_dict(plan: TriagePlan) -> dict[str, Any]:
    return {
        "plan_id": plan.plan_id,
        "created_at": plan.created_at.isoformat(),
        "config_fingerprint": plan.config_fingerprint,
        "folder_id": plan.folder_id,
        "dry_run": plan.dry_run,
        "actions": [
            {
                "message_id": action.message_id,
                "outlook_id": action.outlook_id,
                "subject": action.subject,
                "sender_name": action.sender_name,
                "sender_address": action.sender_address,
                "received_at": action.received_at,
                "add_categories": action.add_categories,
                "remove_categories": action.remove_categories,
                "move_to": action.move_to,
                "set_flag": action.set_flag.value if action.set_flag else None,
                "report_section": action.report_section,
                "keep_in_inbox": action.keep_in_inbox,
                "domain_class": action.domain_class.value,
                "matches": [
                    {
                        "rule_id": match.rule_id,
                        "description": match.description,
                        "priority": match.priority,
                        "reasons": match.reasons,
                    }
                    for match in action.matches
                ],
            }
            for action in plan.actions
        ],
    }


def plan_from_dict(value: dict[str, Any]) -> TriagePlan:
    return TriagePlan(
        plan_id=value["plan_id"],
        created_at=datetime.fromisoformat(value["created_at"]),
        config_fingerprint=value["config_fingerprint"],
        folder_id=int(value["folder_id"]),
        dry_run=bool(value.get("dry_run", True)),
        actions=[
            PlannedMessageAction(
                message_id=action["message_id"],
                outlook_id=int(action["outlook_id"]),
                subject=action["subject"],
                sender_name=action.get("sender_name", ""),
                sender_address=action.get("sender_address", ""),
                received_at=action.get("received_at", ""),
                add_categories=list(action["add_categories"]),
                remove_categories=list(action["remove_categories"]),
                move_to=action.get("move_to"),
                set_flag=FlagStatus(action["set_flag"]) if action.get("set_flag") else None,
                report_section=action.get(
                    "report_section",
                    action.get("digest_section", "Needs review"),
                ),
                keep_in_inbox=bool(action["keep_in_inbox"]),
                domain_class=_domain_class(action["domain_class"]),
                matches=[
                    RuleMatch(
                        rule_id=match["rule_id"],
                        description=match["description"],
                        priority=int(match["priority"]),
                        reasons=list(match["reasons"]),
                    )
                    for match in action["matches"]
                ],
            )
            for action in value["actions"]
        ],
    )
