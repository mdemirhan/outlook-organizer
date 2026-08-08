from __future__ import annotations

from collections import Counter
from dataclasses import asdict
from datetime import date, time
from typing import Any

from outlook_organizer.actions import ActionExecutor
from outlook_organizer.calendar_analysis import analyze_calendar, find_free_slots
from outlook_organizer.config import AppConfig, load_config
from outlook_organizer.models import CalendarEvent, MailMessage, TriagePlan
from outlook_organizer.outlook import OutlookAdapter
from outlook_organizer.progress import ProgressCallback, format_count, report_progress
from outlook_organizer.rules import MailTriagePlanner
from outlook_organizer.serialization import plan_to_dict
from outlook_organizer.state import StateStore


class OutlookOrganizerService:
    def __init__(
        self,
        config: AppConfig | None = None,
        adapter: OutlookAdapter | None = None,
        store: StateStore | None = None,
    ) -> None:
        self.config = config or load_config()
        self.adapter = adapter or OutlookAdapter()
        self.store = store or StateStore()
        self.planner = MailTriagePlanner(self.config)
        self.executor = ActionExecutor(self.config, self.adapter, self.store)

    def validate(self) -> dict[str, Any]:
        self.planner.validate()
        configured_groups = set(self.config.definitions.groups)
        configured_distribution_groups = set(
            self.config.definitions.distribution_list_groups
        )
        conditions = [
            rule.when for rule in [*self.config.mail.annotations, *self.config.mail.routes]
        ]
        referenced_groups = {
            condition.sender_group for condition in conditions if condition.sender_group
        }
        referenced_distribution_groups = {
            condition.distribution_list_group
            for condition in conditions
            if condition.distribution_list_group
        }
        missing_groups = sorted(referenced_groups - configured_groups)
        if missing_groups:
            raise ValueError(f"Rules reference undefined people groups: {missing_groups}")
        missing_distribution_groups = sorted(
            referenced_distribution_groups - configured_distribution_groups
        )
        if missing_distribution_groups:
            raise ValueError(
                "Rules reference undefined distribution-list groups: "
                f"{missing_distribution_groups}"
            )
        return {
            "valid": True,
            "fingerprint": self.config.fingerprint,
            "annotations": len(self.config.mail.annotations),
            "routes": len(self.config.mail.routes),
            "groups": sorted(configured_groups),
            "distribution_list_groups": sorted(configured_distribution_groups),
            "folders": {
                key: {
                    "id": folder.id,
                    "name": folder.name,
                    "parent": folder.parent,
                }
                for key, folder in self.config.mail.folders.items()
            },
            "internal_domains": list(self.config.definitions.internal_domains),
            "junk_external_domains": list(self.config.definitions.junk_external.domains),
            "junk_external_addresses": list(
                self.config.definitions.junk_external.addresses
            ),
            "junk_external_keywords": list(
                self.config.definitions.junk_external.keywords
            ),
            "safe_external_domains": list(self.config.definitions.safe_external.domains),
            "safe_external_addresses": list(
                self.config.definitions.safe_external.addresses
            ),
        }

    def folders(self) -> list[dict[str, Any]]:
        return [
            asdict(folder)
            for folder in self.adapter.list_folders(self.config.mail.folder_scan_limit)
        ]

    def configured_folder_status(self) -> dict[str, Any]:
        actual_by_id = {
            folder.outlook_id: folder
            for folder in self.adapter.list_folders(self.config.mail.folder_scan_limit)
        }
        folders: dict[str, dict[str, Any]] = {}
        for key, expected in self.config.mail.folders.items():
            actual = actual_by_id.get(expected.id)
            expected_parent = (
                self.config.mail.folders[expected.parent].name
                if expected.parent
                else None
            )
            name_matches = bool(
                actual
                and actual.name.casefold()
                in {name.casefold() for name in expected.names}
            )
            parent_matches = bool(
                actual
                and (
                    expected_parent is None
                    or actual.parent_name.casefold() == expected_parent.casefold()
                )
            )
            folders[key] = {
                "valid": name_matches and parent_matches,
                "expected": {
                    "id": expected.id,
                    "name": expected.name,
                    "parent": expected_parent,
                },
                "actual": (
                    {
                        "id": actual.outlook_id,
                        "name": actual.name,
                        "parent": actual.parent_name,
                    }
                    if actual
                    else None
                ),
            }
        return {
            "valid": all(folder["valid"] for folder in folders.values()),
            "folders": folders,
        }

    def setup_mail_folders(self, *, confirm: bool) -> dict[str, Any]:
        if not confirm:
            raise ValueError("confirm=true is required to create Outlook folders")
        inbox = self.config.mail.folders["inbox"]
        roots = ("organized_primary", "organized_secondary")
        return {
            "folders": {
                key: self.adapter.ensure_mail_folder(
                    inbox.id,
                    self.config.mail.folders[key].name,
                )
                for key in roots
            }
        }

    def calendars(self) -> list[dict[str, Any]]:
        return [
            asdict(calendar)
            for calendar in self.adapter.list_calendars(self.config.calendar.maximum_calendar_id)
        ]

    def preview_triage(
        self,
        *,
        limit: int = 25,
        body_limit: int = 0,
    ) -> dict[str, Any]:
        plan = self._build_plan(limit=limit, body_limit=body_limit)
        self.store.save_plan(plan)
        return plan_to_dict(plan)

    def _build_plan(
        self,
        *,
        limit: int,
        body_limit: int,
        progress: ProgressCallback | None = None,
    ) -> TriagePlan:
        inbox_id = self.config.mail.folders["inbox"].id
        report_progress(
            progress,
            f"Reading up to {format_count(limit, 'message')} from Outlook",
        )
        messages = self.adapter.latest_messages(inbox_id, limit=limit, body_limit=body_limit)
        report_progress(
            progress,
            f"Classifying {format_count(len(messages), 'message')}",
        )
        return self.planner.create_plan(messages, inbox_id)

    def triage_mail(
        self,
        *,
        limit: int = 50,
        body_limit: int = 0,
        confirm: bool = False,
        progress: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        plan_object = self._build_plan(
            limit=limit,
            body_limit=body_limit,
            progress=progress,
        )
        execution: dict[str, Any] | None = None
        if confirm:
            report_progress(progress, "Saving the confirmed triage plan")
            self.store.save_plan(plan_object)
            execution = asdict(self.executor.apply_plan(plan_object, progress=progress))
        report_progress(progress, "Building the mail triage report")
        plan = plan_to_dict(plan_object)
        sections: dict[str, list[dict[str, Any]]] = {}
        for index, action in enumerate(plan["actions"], start=1):
            destination = (
                self.config.mail.folders[action["move_to"]].name
                if action["move_to"]
                else None
            )
            sections.setdefault(action["report_section"], []).append(
                {
                    "index": index,
                    "outlook_id": action["outlook_id"],
                    "subject": action["subject"],
                    "sender_name": action["sender_name"],
                    "sender_address": action["sender_address"],
                    "received_at": action["received_at"],
                    "domain_class": action["domain_class"],
                    "categories_to_add": action["add_categories"],
                    "move_to": destination,
                    "keep_in_inbox": action["keep_in_inbox"],
                    "matched_rules": [match["rule_id"] for match in action["matches"]],
                }
            )
        routes = Counter(
            (
                self.config.mail.folders[action["move_to"]].name
                if action["move_to"]
                else "Inbox"
            )
            for action in plan["actions"]
        )
        categories = Counter(
            category for action in plan["actions"] for category in action["add_categories"]
        )
        return {
            "created_at": plan["created_at"],
            "dry_run": not confirm,
            "summary": {
                "messages": len(plan["actions"]),
                "proposed_moves": sum(action["move_to"] is not None for action in plan["actions"]),
                "kept_in_inbox": sum(action["keep_in_inbox"] for action in plan["actions"]),
                "possible_spam": sum(
                    action["domain_class"] in {"junk_external", "unknown_external"}
                    for action in plan["actions"]
                ),
            },
            "action_summary": {
                "routes": dict(routes),
                "categories": dict(categories),
            },
            "sections": sections,
            "execution": execution,
        }

    def get_plan(self, plan_id: str) -> dict[str, Any]:
        return plan_to_dict(self.store.load_plan(plan_id))

    def get_message(self, outlook_id: int, include_body: bool = False) -> dict[str, Any]:
        message = self.adapter.get_message(outlook_id, body_limit=20_000 if include_body else 0)
        return self._message_dict(message, include_body=include_body)

    def search_messages(
        self,
        query: str,
        *,
        limit: int = 20,
        scan_limit: int = 250,
        include_body: bool = False,
    ) -> list[dict[str, Any]]:
        inbox = self.config.mail.folders["inbox"]
        body_limit = 5000 if include_body or query else 0
        messages = self.adapter.latest_messages(
            inbox.id, limit=scan_limit, body_limit=body_limit
        )
        needle = query.casefold()
        matches = [
            message
            for message in messages
            if needle in message.subject.casefold()
            or needle in message.sender_name.casefold()
            or needle in message.sender_address.casefold()
            or needle in message.body.casefold()
        ]
        return [
            self._message_dict(message, include_body=include_body) for message in matches[:limit]
        ]

    def apply_plan(
        self,
        plan_id: str,
        *,
        confirm: bool,
        selected_indexes: list[int] | None = None,
    ) -> dict[str, Any]:
        if not confirm:
            raise ValueError("confirm=true is required for Outlook writes")
        status = self.store.plan_status(plan_id)
        if status not in {"previewed", "failed"}:
            raise ValueError(
                f"Plan {plan_id} has status {status!r}; create a fresh preview before applying"
            )
        plan = self.store.load_plan(plan_id)
        if selected_indexes is not None:
            invalid = sorted(
                index for index in selected_indexes if index < 1 or index > len(plan.actions)
            )
            if invalid:
                raise ValueError(f"Selected action indexes are out of range: {invalid}")
        result = self.executor.apply_plan(
            plan,
            selected_indexes=(set(selected_indexes) if selected_indexes is not None else None),
        )
        return asdict(result)

    def undo_run(self, run_id: str, *, confirm: bool) -> dict[str, Any]:
        if not confirm:
            raise ValueError("confirm=true is required for Outlook writes")
        return asdict(self.executor.undo_run(run_id))

    def recent_runs(self, limit: int = 20) -> list[dict[str, Any]]:
        return self.store.list_runs(limit)

    def calendar_events(
        self,
        *,
        days_behind: int = 0,
        days_ahead: int = 7,
        include_body: bool = False,
    ) -> list[dict[str, Any]]:
        calendar = self.adapter.find_calendar(
            self.config.calendar.calendar_names,
            self.config.calendar.maximum_calendar_id,
        )
        events = self.adapter.calendar_events(
            calendar.outlook_id,
            days_behind=days_behind,
            days_ahead=days_ahead,
            body_limit=5000 if include_body else 0,
        )
        return [self._calendar_event_dict(event, include_body=include_body) for event in events]

    def analyze_calendar(
        self,
        *,
        days_behind: int = 0,
        days_ahead: int = 7,
    ) -> dict[str, Any]:
        calendar = self.adapter.find_calendar(
            self.config.calendar.calendar_names,
            self.config.calendar.maximum_calendar_id,
        )
        events = self.adapter.calendar_events(
            calendar.outlook_id,
            days_behind=days_behind,
            days_ahead=days_ahead,
        )
        result = analyze_calendar(events)
        result["calendar"] = calendar.name
        return result

    def find_free_slots(
        self,
        target_date: date,
        *,
        minimum_minutes: int | None = None,
    ) -> list[dict[str, str | int]]:
        calendar = self.adapter.find_calendar(
            self.config.calendar.calendar_names,
            self.config.calendar.maximum_calendar_id,
        )
        events = self.adapter.calendar_events(calendar.outlook_id, days_behind=1, days_ahead=14)
        day_name = target_date.strftime("%A").lower()
        hours = self.config.calendar.working_hours.get(day_name)
        if not hours:
            return []
        work_start = time.fromisoformat(hours[0])
        work_end = time.fromisoformat(hours[1])
        lunch_start = time.fromisoformat(self.config.calendar.preferences.lunch_window[0])
        lunch_end = time.fromisoformat(self.config.calendar.preferences.lunch_window[1])
        minimum = minimum_minutes or self.config.calendar.preferences.minimum_focus_block_minutes
        return find_free_slots(
            events,
            target_date,
            work_start=work_start,
            work_end=work_end,
            minimum_minutes=minimum,
            buffer_minutes=self.config.calendar.preferences.meeting_buffer_minutes,
            blocked_windows=[(lunch_start, lunch_end)],
        )

    @staticmethod
    def _message_dict(message: MailMessage, *, include_body: bool) -> dict[str, Any]:
        result = asdict(message)
        result["flag_status"] = message.flag_status.value
        result["stable_id"] = message.stable_id
        if not include_body:
            result.pop("body", None)
        return result

    @staticmethod
    def _calendar_event_dict(event: CalendarEvent, *, include_body: bool) -> dict[str, Any]:
        result = asdict(event)
        if not include_body or event.is_private:
            result.pop("body", None)
        if event.is_private:
            result["subject"] = "Private appointment"
            result["location"] = ""
            result["attendees"] = []
        return result
