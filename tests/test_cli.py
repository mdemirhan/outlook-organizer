from __future__ import annotations

from typer.testing import CliRunner

import outlook_organizer.cli as cli


class FakePreviewService:
    def preview(self, *, limit, body_limit, progress):
        progress(f"Reading up to {limit} messages from Outlook")
        progress("Classifying 0 messages")
        progress("Building the mail triage report")
        return {
            "dry_run": True,
            "summary": {
                "messages": 0,
                "proposed_moves": 0,
                "kept_in_inbox": 0,
                "possible_spam": 0,
                "thread_routed": 0,
            },
            "sections": {},
            "action_summary": {"routes": {}, "categories": {}},
            "execution": None,
        }


def test_mail_triage_preview_is_cli_only_and_shows_progress(monkeypatch) -> None:
    monkeypatch.setattr(cli, "_triage_preview_service", FakePreviewService)

    result = CliRunner().invoke(cli.app, ["mail", "triage", "--limit", "2"])

    assert result.exit_code == 0
    assert "Reading up to 2 messages from Outlook" in result.output
    assert "Triage ready — no Outlook changes" in result.output
    assert "DRY RUN" in result.output


def test_mail_help_exposes_brief_and_triage_without_thread_admin() -> None:
    result = CliRunner().invoke(cli.app, ["mail", "--help"])
    assert result.exit_code == 0
    assert "brief" in result.output
    assert "triage" in result.output
    assert "threads" not in result.output


def test_root_help_exposes_mcp_and_cli_history() -> None:
    result = CliRunner().invoke(cli.app, ["--help"])
    assert result.exit_code == 0
    assert "mcp" in result.output
    assert "history" in result.output
