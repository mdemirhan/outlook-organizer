from __future__ import annotations

import json
from dataclasses import dataclass

from outlook_organizer.config import AppConfig
from outlook_organizer.models import FlagStatus, TriagePlan
from outlook_organizer.outlook import OutlookAdapter
from outlook_organizer.progress import ProgressCallback, format_count, report_progress
from outlook_organizer.rules.triage import MailTriagePlanner
from outlook_organizer.state import StateStore


@dataclass(slots=True)
class ExecutionResult:
    run_id: str
    applied: int
    status: str
    error: str | None = None


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
        prepared: list[dict] = []
        for sequence, previewed_action in selected_actions:
            current = current_messages[previewed_action.outlook_id]
            # The folder preview and this verification read are separate Outlook
            # snapshots. Reclassify the fresh message so recipient, flag, and
            # category changes cannot leave us applying a stale rule decision.
            action = self.planner.plan_message(current)
            plan.actions[sequence - 1] = action
            before = {
                "folder_id": current.folder_id,
                "folder_name": current.folder_name,
                "categories": current.categories,
                "flag_status": current.flag_status.value,
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
            }
            prepared.append(
                {
                    "sequence": sequence,
                    "action": action,
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

        if errors:
            status = "partial" if applied else "failed"
            error_text = f"{len(errors)} action(s) failed: " + "; ".join(errors[:3])
            self.store.finish_run(run_id, status, error_text)
            return ExecutionResult(run_id, applied, status, error_text)
        self.store.finish_run(run_id, "completed")
        return ExecutionResult(run_id, applied, "completed")

    def undo_run(self, run_id: str) -> ExecutionResult:
        rows = self.store.run_actions(run_id, reverse=True)
        restored = 0
        undone_ids: list[int] = []
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
        except Exception as exc:
            return ExecutionResult(run_id, restored, "undo_failed", str(exc))
        self.store.mark_run_undone(run_id, undone_ids)
        return ExecutionResult(run_id, restored, "undone")
