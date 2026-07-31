from __future__ import annotations

from outlook_organizer.models import MatchConfig, MessageFacts, RuleMatch


class RuleEngine:
    def match(
        self,
        facts: MessageFacts,
        *,
        rule_id: str,
        description: str,
        condition: MatchConfig,
        order: int,
    ) -> RuleMatch | None:
        reasons: list[str] = []
        for key, expected in condition.model_dump(exclude_none=True).items():
            matched, reason = self._predicate(facts, key, expected)
            if not matched:
                return None
            reasons.append(reason)
        return RuleMatch(
            rule_id=rule_id,
            description=description,
            priority=-order,
            reasons=reasons,
        )

    def _predicate(
        self, facts: MessageFacts, key: str, expected: object
    ) -> tuple[bool, str]:
        if key == "flagged":
            actual = facts.message.flag_status.value == "flagged"
            return actual is bool(expected), f"flagged is {actual}"
        if key == "recipient":
            values = (
                {str(value) for value in expected}
                if isinstance(expected, list)
                else {str(expected)}
            )
            actual = facts.directness.value
            return actual in values, f"recipient directness is {actual}"
        if key == "sender_group":
            actual = facts.sender_relationship
            return actual == str(expected), f"sender group is {actual}"
        if key == "sender_type":
            actual = facts.domain_class.value
            return actual == str(expected), f"sender type is {actual}"
        if key == "distribution_list_group":
            group = str(expected)
            return (
                group in facts.distribution_list_groups,
                f"distribution-list groups include {facts.distribution_list_groups}",
            )
        if key == "distribution_list":
            actual = facts.has_distribution_list
            return actual is bool(expected), f"distribution-list presence is {actual}"
        raise ValueError(f"Unknown match field: {key}")
