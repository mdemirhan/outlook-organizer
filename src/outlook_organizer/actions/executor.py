from __future__ import annotations

import json
from dataclasses import dataclass

from outlook_organizer.config import AppConfig
from outlook_organizer.models import FlagStatus, MailMessage, TriagePlan
from outlook_organizer.outlook import OutlookAdapter
from outlook_organizer.progress import ProgressCallback, format_count, report_progress
from outlook_organizer.rules.triage import MailTriagePlanner
from outlook_organizer.state import StateStore
from outlook_organizer.threads import ThreadPromotion, ThreadRouter


@dataclass(slots=True)
class ExecutionResult:
    run_id: str
    applied: int
    status: str
    error: str | None = None
    promoted: int = 0
    thread_routed: int = 0


class ActionExecutor:
    def __init__(
        self,
        config: AppConfig,
        adapter: OutlookAdapter,
        store: StateStore,
    ) -> None:
        self.config = config
        self.adapter = adapter
        self.store = store
        self.planner = MailTriagePlanner(config)
        self.thread_router = ThreadRouter(config, adapter, store)

    def apply_plan(
        self,
        plan: TriagePlan,
        *,
        selected_indexes: set[int] | None = None,
        progress: ProgressCallback | None = None,
    ) -> ExecutionResult:
        if plan.config_fingerprint != self.config.fingerprint:
            raise ValueError("Plan configuration fingerprint no longer matches current config")

        selected_actions = [
            (sequence, action)
            for sequence, action in enumerate(plan.actions, start=1)
            if selected_indexes is None or sequence in selected_indexes
        ]
        report_progress(
            progress,
            f"Verifying {format_count(len(selected_actions), 'message')} before making changes",
        )
        current_messages = {
            message.outlook_id: message
            for message in self.adapter.get_messages(
                [action.outlook_id for _, action in selected_actions],
                body_limit=0,
            )
        }
        missing_ids = [
            action.outlook_id
            for _, action in selected_actions
            if action.outlook_id not in current_messages
        ]
        if missing_ids:
            raise ValueError(f"Outlook messages disappeared before apply: {missing_ids}")

        report_progress(
            progress,
            f"Preparing {format_count(len(selected_actions), 'Outlook update')}",
        )
        ordered_messages = [
            current_messages[action.outlook_id] for _, action in selected_actions
        ]
        refreshed_actions = [
            self.planner.plan_message(message) for message in ordered_messages
        ]
        resolution = self.thread_router.resolve(
            ordered_messages,
            refreshed_actions,
            progress=progress,
        )
        actions_by_id = {
            action.outlook_id: action for action in resolution.actions
        }
        prepared: list[dict] = []
        for sequence, previewed_action in selected_actions:
            current = current_messages[previewed_action.outlook_id]
            # The folder preview and this verification read are separate Outlook
            # snapshots. Reclassify the fresh message so recipient, flag, and
            # category changes cannot leave us applying a stale rule decision.
            action = actions_by_id[current.outlook_id]
            plan.actions[sequence - 1] = action
            before = {
                "folder_id": current.folder_id,
                "folder_name": current.folder_name,
                "categories": current.categories,
                "flag_status": current.flag_status.value,
                "thread_guid": current.thread_guid,
            }
            desired_categories = sorted(
                (set(current.categories) - set(action.remove_categories))
                | set(action.add_categories)
            )
            desired_flag = action.set_flag or current.flag_status
            desired_folder_id = current.folder_id
            target_folder_id: int | None = None
            if action.move_to and not action.keep_in_inbox:
                desired_folder_id = self.config.mail.folders[action.move_to].id
                if desired_folder_id != current.folder_id:
                    target_folder_id = desired_folder_id
            after = {
                "folder_id": desired_folder_id,
                "folder_name": (
                    self.config.mail.folders[action.move_to].name
                    if desired_folder_id != current.folder_id
                    else current.folder_name
                ),
                "categories": desired_categories,
                "flag_status": desired_flag.value,
                "thread_guid": current.thread_guid,
            }
            prepared.append(
                {
                    "sequence": sequence,
                    "action": action,
                    "message": current,
                    "promotion": None,
                    "before": before,
                    "after": after,
                    "update": {
                        "outlook_id": action.outlook_id,
                        "categories": desired_categories,
                        "flag_status": desired_flag,
                        "target_folder_id": target_folder_id,
                    },
                }
            )

        promotion_by_id = {
            promotion.outlook_id: promotion
            for promotion in resolution.promotions
        }
        promoted_messages = {
            message.outlook_id: message
            for message in self.adapter.get_messages(
                promotion_by_id,
                body_limit=0,
            )
        }
        applied_promotions: dict[int, tuple[MailMessage, ThreadPromotion]] = {}
        next_sequence = len(plan.actions) + 1
        for outlook_id, promotion in promotion_by_id.items():
            current = promoted_messages.get(outlook_id)
            if current is None:
                continue
            base_action = self.planner.plan_message(current)
            if not self.thread_router.action_uses_affinity(base_action):
                continue
            destination = self.config.mail.folders[promotion.destination_key]
            if current.folder_id == destination.id:
                continue
            before = {
                "folder_id": current.folder_id,
                "folder_name": current.folder_name,
                "categories": current.categories,
                "flag_status": current.flag_status.value,
                "thread_guid": current.thread_guid,
            }
            after = {
                "folder_id": destination.id,
                "folder_name": destination.name,
                "categories": current.categories,
                "flag_status": current.flag_status.value,
                "thread_guid": current.thread_guid,
            }
            base_action.move_to = promotion.destination_key
            base_action.report_section = destination.name
            prepared.append(
                {
                    "sequence": next_sequence,
                    "action": base_action,
                    "message": current,
                    "promotion": promotion,
                    "before": before,
                    "after": after,
                    "update": {
                        "outlook_id": outlook_id,
                        "categories": current.categories,
                        "flag_status": current.flag_status,
                        "target_folder_id": destination.id,
                    },
                }
            )
            applied_promotions[outlook_id] = (current, promotion)
            next_sequence += 1
        plan.thread_promotions = len(applied_promotions)

        # Keep the persisted plan aligned with the actions actually sent to
        # Outlook. This also creates the plan record for direct executor use.
        self.store.save_plan(plan)
        run_id = self.store.start_run(plan.plan_id)
        report_progress(
            progress,
            f"Applying {format_count(len(prepared), 'update')} in Outlook",
        )
        try:
            results = self.adapter.apply_mail_states([item["update"] for item in prepared])
        except Exception as exc:
            report_progress(progress, "Recording Outlook failure details")
            for item in prepared:
                action = item["action"]
                self.store.record_action(
                    run_id=run_id,
                    sequence=item["sequence"],
                    outlook_id=action.outlook_id,
                    message_id=action.message_id,
                    subject=action.subject,
                    status="failed",
                    before_state=item["before"],
                    after_state=item["after"],
                    error=str(exc),
                )
            self.store.finish_run(run_id, "failed", str(exc))
            return ExecutionResult(run_id, 0, "failed", str(exc))

        report_progress(progress, "Recording the audit trail")
        results_by_id = {int(result["outlook_id"]): result for result in results}
        applied = 0
        promoted = 0
        thread_routed = 0
        successful_ids: set[int] = set()
        errors: list[str] = []
        for item in prepared:
            action = item["action"]
            result = results_by_id.get(action.outlook_id)
            succeeded = result is not None and result["status"] == "applied"
            error = (
                ""
                if succeeded
                else str(result["error"] if result is not None else "No result returned by Outlook")
            )
            if succeeded:
                applied += 1
                successful_ids.add(action.outlook_id)
                if item["promotion"] is not None:
                    promoted += 1
                elif self.thread_router.action_was_thread_routed(action):
                    thread_routed += 1
            else:
                errors.append(f"{action.subject}: {error}")
            self.store.record_action(
                run_id=run_id,
                sequence=item["sequence"],
                outlook_id=action.outlook_id,
                message_id=action.message_id,
                subject=action.subject,
                status="applied" if succeeded else "failed",
                before_state=item["before"],
                after_state=item["after"],
                error=error or None,
            )

        self.thread_router.persist_applied(
            messages=ordered_messages,
            actions_by_id=actions_by_id,
            successful_ids=successful_ids,
            canonical_routes=resolution.canonical_routes,
            promoted_messages=applied_promotions,
        )

        if errors:
            status = "partial" if applied else "failed"
            error_text = f"{len(errors)} action(s) failed: " + "; ".join(errors[:3])
            self.store.finish_run(run_id, status, error_text)
            return ExecutionResult(
                run_id,
                applied,
                status,
                error_text,
                promoted=promoted,
                thread_routed=thread_routed,
            )
        self.store.finish_run(run_id, "completed")
        return ExecutionResult(
            run_id,
            applied,
            "completed",
            promoted=promoted,
            thread_routed=thread_routed,
        )

    def undo_run(self, run_id: str) -> ExecutionResult:
        rows = self.store.run_actions(run_id, reverse=True)
        restored = 0
        undone_ids: list[int] = []
        restored_thread_members: list[tuple[MailMessage, int]] = []
        try:
            for row in rows:
                # Outlook mutations are not transactional. A failed action may
                # have changed categories or flags before a later operation
                # failed, so restoring its captured before-state is safe.
                if row["status"] not in {"applied", "failed"}:
                    continue
                before = json.loads(row["before_state"])
                current = self.adapter.get_message(int(row["outlook_id"]), body_limit=0)
                original_folder_id = int(before["folder_id"])
                self.adapter.apply_mail_state(
                    int(row["outlook_id"]),
                    categories=list(before["categories"]),
                    flag_status=FlagStatus(before["flag_status"]),
                    target_folder_id=(
                        original_folder_id if current.folder_id != original_folder_id else None
                    ),
                )
                restored += 1
                undone_ids.append(int(row["id"]))
                restored_thread_members.append((current, original_folder_id))
        except Exception as exc:
            self.thread_router.restore_members(restored_thread_members)
            return ExecutionResult(run_id, restored, "undo_failed", str(exc))
        self.store.mark_run_undone(run_id, undone_ids)
        self.thread_router.restore_members(restored_thread_members)
        return ExecutionResult(run_id, restored, "undone")
