from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

from outlook_organizer.config import AppConfig
from outlook_organizer.models import (
    DomainClass,
    MailMessage,
    PlannedMessageAction,
    RuleMatch,
    ThreadMessageState,
)
from outlook_organizer.outlook import OutlookAdapter
from outlook_organizer.progress import ProgressCallback, format_count, report_progress
from outlook_organizer.state import StateStore


@dataclass(slots=True)
class ThreadPromotion:
    outlook_id: int
    thread_guid: str
    destination_key: str


@dataclass(slots=True)
class ThreadResolution:
    actions: list[PlannedMessageAction]
    promotions: list[ThreadPromotion] = field(default_factory=list)
    canonical_routes: dict[str, str] = field(default_factory=dict)


class ThreadRouter:
    def __init__(
        self,
        config: AppConfig,
        adapter: OutlookAdapter,
        store: StateStore,
    ) -> None:
        self.config = config
        self.adapter = adapter
        self.store = store
        self.enabled = config.mail.threading.enabled
        inbox_id = config.mail.folders["inbox"].id
        self.scope = f"inbox:{inbox_id}"
        self.folder_key_by_id = {
            folder.id: key for key, folder in config.mail.folders.items()
        }
        self.route_priority: dict[str, int] = {}
        self.managed_folder_keys: set[str] = set()
        self.safety_folder_keys: set[str] = set()
        for priority, route in enumerate(config.mail.routes):
            self.route_priority.setdefault(route.move_to, priority)
            self.managed_folder_keys.add(route.move_to)
            if route.when.sender_type in {"junk_external", "unclassified_external"}:
                self.safety_folder_keys.add(route.move_to)
        self.affinity_folder_keys = self.managed_folder_keys - self.safety_folder_keys
    def status(self) -> dict[str, Any]:
        status = self.store.thread_index_status(self.scope)
        return {
            "enabled": self.enabled,
            "ready": self.enabled,
            "mode": "prospective",
            "scope": self.scope,
            "managed_folders": sorted(self.managed_folder_keys),
            **status,
        }

    def resolve(
        self,
        messages: list[MailMessage],
        actions: list[PlannedMessageAction],
        *,
        progress: ProgressCallback | None = None,
    ) -> ThreadResolution:
        if len(messages) != len(actions):
            raise ValueError("Thread routing requires one action per message")
        if not self.enabled or not any(message.thread_guid for message in messages):
            return ThreadResolution(actions=actions)

        guids = {message.thread_guid for message in messages if message.thread_guid}
        contexts = self.store.thread_contexts(self.scope, guids)
        current_ids = {message.outlook_id for message in messages}
        known_ids = sorted(
            {
                int(member["outlook_id"])
                for context in contexts.values()
                for member in context["members"]
                if int(member["outlook_id"]) not in current_ids
            }
        )
        if known_ids:
            report_progress(
                progress,
                f"Checking {format_count(len(known_ids), 'known thread message')}",
            )
            states = self.adapter.thread_states_by_ids(known_ids)
            self._reconcile_contexts(contexts, states, current_ids=current_ids)

        grouped: dict[str, list[tuple[MailMessage, PlannedMessageAction]]] = defaultdict(list)
        for message, action in zip(messages, actions, strict=True):
            if message.thread_guid:
                grouped[message.thread_guid].append((message, action))

        promotions: dict[int, ThreadPromotion] = {}
        canonical_routes: dict[str, str] = {}
        for thread_guid, pairs in grouped.items():
            context = contexts.get(thread_guid)
            candidates: set[str] = set()
            if context and context["folder_key"] in self.affinity_folder_keys:
                candidates.add(str(context["folder_key"]))
            candidates.update(
                action.move_to
                for _, action in pairs
                if self.action_uses_affinity(action)
            )
            if not candidates:
                continue
            destination = self._highest_priority(candidates)
            canonical_routes[thread_guid] = destination

            for _, action in pairs:
                if not self.action_uses_affinity(action):
                    continue
                original = action.move_to
                action.move_to = destination
                action.report_section = self.config.mail.folders[destination].name
                if original != destination:
                    promoted = (
                        original is not None
                        and self.route_priority[destination]
                        < self.route_priority.get(original, 1_000_000)
                    )
                    action.matches.append(
                        RuleMatch(
                            rule_id=(
                                "thread-priority-promotion"
                                if promoted
                                else "thread-affinity"
                            ),
                            description=(
                                "Conversation promoted to its highest-priority folder"
                                if promoted
                                else "Conversation inherited its established folder"
                            ),
                            priority=-1,
                            reasons=[f"thread destination is {destination}"],
                        )
                    )

            if context is None:
                continue
            for member in context["members"]:
                outlook_id = int(member["outlook_id"])
                if (
                    outlook_id in current_ids
                    or member["detached"]
                    or member["folder_key"] not in self.affinity_folder_keys
                    or member["folder_key"] == destination
                ):
                    continue
                promotions[outlook_id] = ThreadPromotion(
                    outlook_id=outlook_id,
                    thread_guid=thread_guid,
                    destination_key=destination,
                )

        return ThreadResolution(
            actions=actions,
            promotions=list(promotions.values()),
            canonical_routes=canonical_routes,
        )

    def action_uses_affinity(self, action: PlannedMessageAction) -> bool:
        return bool(
            action.move_to in self.affinity_folder_keys
            and not action.keep_in_inbox
            and action.domain_class
            not in {DomainClass.JUNK_EXTERNAL, DomainClass.UNCLASSIFIED_EXTERNAL}
        )

    @staticmethod
    def action_was_thread_routed(action: PlannedMessageAction) -> bool:
        return any(
            match.rule_id in {"thread-affinity", "thread-priority-promotion"}
            for match in action.matches
        )

    def persist_applied(
        self,
        *,
        messages: list[MailMessage],
        actions_by_id: dict[int, PlannedMessageAction],
        successful_ids: set[int],
        canonical_routes: dict[str, str],
        promoted_messages: dict[int, tuple[MailMessage, ThreadPromotion]],
    ) -> None:
        if not self.enabled:
            return
        routes: dict[str, str] = {}
        members: list[dict[str, Any]] = []
        for message in messages:
            if message.outlook_id not in successful_ids or not message.thread_guid:
                continue
            action = actions_by_id[message.outlook_id]
            destination = (
                action.move_to
                if action.move_to and not action.keep_in_inbox
                else self.folder_key_by_id.get(message.folder_id)
            )
            if destination is None:
                continue
            members.append(
                self._member_record(
                    message,
                    folder_key=destination,
                    detached=not self.action_uses_affinity(action),
                )
            )
            routes[message.thread_guid] = canonical_routes.get(
                message.thread_guid, destination
            )

        for outlook_id, (message, promotion) in promoted_messages.items():
            if outlook_id not in successful_ids:
                continue
            members.append(
                self._member_record(
                    message,
                    folder_key=promotion.destination_key,
                    detached=False,
                )
            )
            routes[promotion.thread_guid] = promotion.destination_key

        self.store.update_thread_index(
            scope=self.scope,
            routes=routes,
            members=members,
        )

    def forget_messages(self, outlook_ids: set[int]) -> None:
        if not self.enabled or not outlook_ids:
            return
        affected = self.store.delete_thread_members(self.scope, outlook_ids)
        if not affected:
            return
        self._recompute_routes(affected)

    def restore_members(self, restored: list[tuple[MailMessage, int]]) -> None:
        if not self.enabled or not restored:
            return
        members: list[dict[str, Any]] = []
        affected: set[str] = set()
        for message, folder_id in restored:
            if not message.thread_guid:
                continue
            folder_key = self.folder_key_by_id.get(folder_id)
            if folder_key is None:
                continue
            affected.add(message.thread_guid)
            members.append(
                {
                    "thread_guid": message.thread_guid,
                    "outlook_id": message.outlook_id,
                    "message_id": message.stable_id,
                    "folder_id": folder_id,
                    "folder_key": folder_key,
                    "detached": folder_key not in self.affinity_folder_keys,
                }
            )
        if members:
            self.store.update_thread_index(scope=self.scope, members=members)
            self._recompute_routes(affected)

    def _recompute_routes(self, affected: set[str]) -> None:
        contexts = self.store.thread_contexts(self.scope, affected)
        route_updates: dict[str, str] = {}
        empty: set[str] = set()
        for thread_guid in affected:
            context = contexts.get(thread_guid)
            if context is None or not context["members"]:
                empty.add(thread_guid)
                continue
            active_folders = {
                str(member["folder_key"])
                for member in context["members"]
                if not member["detached"]
                and member["folder_key"] in self.affinity_folder_keys
            }
            all_folders = {
                str(member["folder_key"])
                for member in context["members"]
                if member["folder_key"]
            }
            candidates = active_folders or all_folders
            if candidates:
                route_updates[thread_guid] = self._highest_priority(candidates)
            else:
                empty.add(thread_guid)
        if route_updates:
            self.store.update_thread_index(scope=self.scope, routes=route_updates)
        self.store.delete_thread_routes(self.scope, empty)

    def _reconcile_contexts(
        self,
        contexts: dict[str, dict[str, Any]],
        states: list[ThreadMessageState],
        *,
        current_ids: set[int],
    ) -> None:
        state_by_id = {state.outlook_id: state for state in states}
        route_updates: dict[str, str] = {}
        member_updates: dict[tuple[str, int], dict[str, Any]] = {}
        for thread_guid, context in contexts.items():
            active = [
                member
                for member in context["members"]
                if not member["detached"] and member["outlook_id"] not in current_ids
            ]
            if not active:
                continue
            valid_state_by_id = {
                member["outlook_id"]: state_by_id[member["outlook_id"]]
                for member in active
                if member["outlook_id"] in state_by_id
                and self._state_matches_member(
                    state_by_id[member["outlook_id"]],
                    thread_guid=thread_guid,
                )
            }
            for member in active:
                state = valid_state_by_id.get(member["outlook_id"])
                if (
                    state is not None
                    and state.exchange_id
                    and state.exchange_id != member["message_id"]
                ):
                    member["message_id"] = state.exchange_id
                    member_updates[(thread_guid, member["outlook_id"])] = member
            found = list(valid_state_by_id.values())
            changed = [
                member
                for member in active
                if member["outlook_id"] not in valid_state_by_id
                or valid_state_by_id[member["outlook_id"]].folder_id
                != member["folder_id"]
            ]
            if not changed:
                continue
            actual_keys = {
                self.folder_key_by_id.get(state.folder_id)
                for state in found
            }
            whole_thread_move = bool(
                len(found) == len(active)
                and len(actual_keys) == 1
                and next(iter(actual_keys)) in self.managed_folder_keys
            )
            if whole_thread_move:
                destination = str(next(iter(actual_keys)))
                context["folder_key"] = destination
                route_updates[thread_guid] = destination
                for member in active:
                    state = valid_state_by_id[member["outlook_id"]]
                    member.update(
                        message_id=state.exchange_id or member["message_id"],
                        folder_id=state.folder_id,
                        folder_key=destination,
                        detached=destination in self.safety_folder_keys,
                    )
                    member_updates[(thread_guid, member["outlook_id"])] = member
                continue

            for member in changed:
                state = valid_state_by_id.get(member["outlook_id"])
                if state is not None:
                    member.update(
                        message_id=state.exchange_id or member["message_id"],
                        folder_id=state.folder_id,
                        folder_key=self.folder_key_by_id.get(state.folder_id),
                    )
                member["detached"] = True
                member_updates[(thread_guid, member["outlook_id"])] = member

        if route_updates or member_updates:
            self.store.update_thread_index(
                scope=self.scope,
                routes=route_updates,
                members=member_updates.values(),
            )

    @staticmethod
    def _state_matches_member(
        state: ThreadMessageState,
        *,
        thread_guid: str,
    ) -> bool:
        return state.thread_guid == thread_guid

    def _highest_priority(self, folders: set[str]) -> str:
        return min(
            folders,
            key=lambda key: (self.route_priority.get(key, 1_000_000), key),
        )

    def _member_record(
        self,
        message: MailMessage,
        *,
        folder_key: str,
        detached: bool,
    ) -> dict[str, Any]:
        return {
            "thread_guid": message.thread_guid,
            "outlook_id": message.outlook_id,
            "message_id": message.stable_id,
            "folder_id": self.config.mail.folders[folder_key].id,
            "folder_key": folder_key,
            "detached": detached,
        }
