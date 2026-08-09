from __future__ import annotations

from typing import Any

from outlook_organizer.mail.ports import MailReader
from outlook_organizer.progress import ProgressCallback, format_count, report_progress
from outlook_organizer.triage.classifier import TriageClassifier
from outlook_organizer.triage.config import TriageContext
from outlook_organizer.triage.models import TriageAssessment
from outlook_organizer.triage.reporting import build_triage_report
from outlook_organizer.triage.thread_index import ThreadAffinityResolver


class TriagePreviewService:
    def __init__(
        self,
        context: TriageContext,
        reader: MailReader,
        thread_resolver: ThreadAffinityResolver,
    ) -> None:
        self.context = context
        self.reader = reader
        self.classifier = TriageClassifier(context)
        self.thread_resolver = thread_resolver

    def assess(
        self,
        *,
        limit: int = 50,
        body_limit: int = 0,
        progress: ProgressCallback | None = None,
    ) -> TriageAssessment:
        inbox_id = self.context.mail.folders.folders["inbox"].id
        report_progress(progress, f"Reading up to {format_count(limit, 'message')} from Outlook")
        messages = self.reader.latest_messages(inbox_id, limit=limit, body_limit=body_limit)
        report_progress(progress, f"Classifying {format_count(len(messages), 'message')}")
        assessment = self.classifier.assess(messages, inbox_id)
        assessment.decisions = self.thread_resolver.resolve(
            messages, assessment.decisions
        ).decisions
        return assessment

    def preview(
        self,
        *,
        limit: int = 50,
        body_limit: int = 0,
        progress: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        assessment = self.assess(limit=limit, body_limit=body_limit, progress=progress)
        report_progress(progress, "Building the mail triage report")
        return build_triage_report(assessment, self.context)
