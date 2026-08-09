from __future__ import annotations

from dataclasses import asdict
from typing import Any

from outlook_organizer.audit.repository import AuditRepository
from outlook_organizer.mail import MailMessage
from outlook_organizer.mail.ports import MailReader, MailWriter
from outlook_organizer.progress import ProgressCallback, format_count, report_progress
from outlook_organizer.triage.classifier import TriageClassifier
from outlook_organizer.triage.config import TriageContext
from outlook_organizer.triage.models import ExecutionResult
from outlook_organizer.triage.reporting import build_triage_report
from outlook_organizer.triage.thread_index import ThreadAffinityResolver


class ApplyTriageService:
    def __init__(
        self,
        context: TriageContext,
        reader: MailReader,
        writer: MailWriter,
        audit: AuditRepository,
        thread_resolver: ThreadAffinityResolver,
    ) -> None:
        self.context = context
        self.reader = reader
        self.writer = writer
        self.audit = audit
        self.classifier = TriageClassifier(context)
        self.thread_resolver = thread_resolver

    def apply(
        self,
        *,
        limit: int = 50,
        body_limit: int = 0,
        progress: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        inbox_id = self.context.mail.folders.folders["inbox"].id
        report_progress(progress, f"Reading up to {format_count(limit, 'message')} from Outlook")
        initial = self.reader.latest_messages(inbox_id, limit=limit, body_limit=body_limit)
        report_progress(
            progress, f"Verifying {format_count(len(initial), 'message')} before changes"
        )
        initial_ids = [message.outlook_id for message in initial]
        refreshed = self.reader.get_messages(initial_ids, body_limit=body_limit)
        refreshed_by_id = {message.outlook_id: message for message in refreshed}
        missing_ids = [
            outlook_id for outlook_id in initial_ids if outlook_id not in refreshed_by_id
        ]
        if missing_ids:
            rendered = ", ".join(str(outlook_id) for outlook_id in missing_ids)
            raise RuntimeError(
                f"Triage stopped before making changes because Outlook messages disappeared: "
                f"{rendered}"
            )
        messages = [refreshed_by_id[outlook_id] for outlook_id in initial_ids]
        assessment = self.classifier.assess(messages, inbox_id)
        resolution = self.thread_resolver.resolve(messages, assessment.decisions)
        assessment.decisions = resolution.decisions
        decisions_by_id = {decision.outlook_id: decision for decision in assessment.decisions}
        messages_by_id = {message.outlook_id: message for message in messages}
        prepared: list[dict[str, Any]] = []
        for sequence, decision in enumerate(assessment.decisions, start=1):
            message = messages_by_id[decision.outlook_id]
            categories = sorted(
                (set(message.categories) - set(decision.remove_categories))
                | set(decision.add_categories)
            )
            flag = decision.set_flag or message.flag_status
            target_id = None
            intended_folder_id = message.folder_id
            intended_folder_name = message.folder_name
            if decision.move_to and not decision.keep_in_inbox:
                destination = self.context.mail.folders.folders[decision.move_to]
                intended_folder_id = destination.id
                intended_folder_name = destination.name
                if destination.id != message.folder_id:
                    target_id = destination.id
            before = self._state(message)
            intended = {
                "folder_id": intended_folder_id,
                "folder_name": intended_folder_name,
                "categories": categories,
                "flag_status": flag.value,
                "thread_guid": message.thread_guid,
            }
            if before == intended:
                continue
            prepared.append(
                {
                    "sequence": sequence,
                    "message": message,
                    "decision": decision,
                    "before": before,
                    "intended": intended,
                    "update": {
                        "outlook_id": message.outlook_id,
                        "categories": categories,
                        "flag_status": flag,
                        "target_folder_id": target_id,
                    },
                }
            )

        if not prepared:
            result = ExecutionResult(None, 0, "no_changes")
            report_progress(progress, "No Outlook changes are needed")
            return build_triage_report(assessment, self.context, execution=result)

        run_id = self.audit.begin_run(
            config_fingerprint=self.context.fingerprint,
            parameters={"limit": limit, "body_limit": body_limit},
        )
        report_progress(progress, f"Applying {format_count(len(prepared), 'update')} in Outlook")
        try:
            results = self.writer.apply_mail_states([item["update"] for item in prepared])
        except Exception as exc:
            for item in prepared:
                self._record(item, run_id=run_id, status="failed", error=str(exc))
            self.audit.finish_run(run_id, "failed", str(exc))
            execution = ExecutionResult(run_id, 0, "failed", str(exc))
            return build_triage_report(assessment, self.context, execution=execution)

        results_by_id = {int(result["outlook_id"]): result for result in results}
        successful: set[int] = set()
        errors: list[str] = []
        for item in prepared:
            outlook_id = item["message"].outlook_id
            result = results_by_id.get(outlook_id)
            succeeded = bool(result and result["status"] == "applied")
            error = "" if succeeded else str(result["error"] if result else "No Outlook result")
            if succeeded:
                successful.add(outlook_id)
            else:
                errors.append(f"{item['message'].subject}: {error}")
            self._record(
                item,
                run_id=run_id,
                status="applied" if succeeded else "failed",
                error=error or None,
                actual=item["intended"] if succeeded else None,
            )
        self.thread_resolver.persist_successes(
            messages=messages,
            decisions_by_id=decisions_by_id,
            successful_ids=successful,
            canonical_routes=resolution.canonical_routes,
        )
        status = "completed" if not errors else ("partial" if successful else "failed")
        error_text = "; ".join(errors[:3]) or None
        self.audit.finish_run(run_id, status, error_text)
        execution = ExecutionResult(
            run_id,
            len(successful),
            status,
            error_text,
            thread_routed=sum(
                any(match.rule_id == "thread-affinity" for match in decision.matches)
                and decision.outlook_id in successful
                for decision in assessment.decisions
            ),
        )
        report_progress(progress, "Building the mail triage report")
        return build_triage_report(assessment, self.context, execution=execution)

    def _record(
        self,
        item: dict[str, Any],
        *,
        run_id: str,
        status: str,
        error: str | None,
        actual: dict[str, Any] | None = None,
    ) -> None:
        decision = item["decision"]
        self.audit.record_action(
            run_id=run_id,
            sequence=item["sequence"],
            outlook_id=decision.outlook_id,
            message_id=decision.message_id,
            subject=decision.subject,
            decision=asdict(decision),
            before_state=item["before"],
            intended_state=item["intended"],
            actual_state=actual,
            status=status,
            error=error,
        )

    @staticmethod
    def _state(message: MailMessage) -> dict[str, Any]:
        return {
            "folder_id": message.folder_id,
            "folder_name": message.folder_name,
            "categories": message.categories,
            "flag_status": message.flag_status.value,
            "thread_guid": message.thread_guid,
        }
