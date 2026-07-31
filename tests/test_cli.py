from __future__ import annotations

from typer.testing import CliRunner

import outlook_organizer.cli as cli


class FakeTriageService:
    def triage_mail(self, *, limit, body_limit, confirm, progress):
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
            },
            "sections": {},
            "action_summary": {"routes": {}, "categories": {}},
            "execution": None,
        }


def test_mail_triage_shows_phase_progress(monkeypatch) -> None:
    monkeypatch.setattr(cli, "_service", FakeTriageService)

    result = CliRunner().invoke(cli.app, ["mail", "triage", "--limit", "2"])

    assert result.exit_code == 0
    assert "→ Loading configuration" in result.output
    assert "→ Reading up to 2 messages from Outlook" in result.output
    assert "→ Classifying 0 messages" in result.output
    assert "✓ Triage ready — no Outlook changes" in result.output


def test_mail_triage_progress_can_be_disabled(monkeypatch) -> None:
    monkeypatch.setattr(cli, "_service", FakeTriageService)

    result = CliRunner().invoke(
        cli.app,
        ["mail", "triage", "--limit", "2", "--no-progress"],
    )

    assert result.exit_code == 0
    assert "→" not in result.output
    assert "✓ Triage ready" not in result.output
    assert "Outlook Organizer" in result.output
    assert "Outlook Organizer · Triage summary" in result.output
    assert "Triage summary" in result.output
    assert "Overview" not in result.output
    assert "Action summary" not in result.output
    assert "DRY RUN" in result.output
    assert "Add --apply" in result.output


def test_triage_report_renders_message_content_as_literal_text(capsys) -> None:
    cli._print_triage_report(
        {
            "created_at": "2026-07-30T12:00:00+00:00",
            "dry_run": True,
            "summary": {
                "messages": 1,
                "proposed_moves": 1,
                "kept_in_inbox": 0,
                "possible_spam": 0,
            },
            "sections": {
                "Company Announcements": [
                    {
                        "index": 1,
                        "sender_name": "Monitoring",
                        "sender_address": "monitoring@corp.example",
                        "subject": "[red]Alert[/red]",
                        "keep_in_inbox": False,
                        "move_to": "Company Announcements",
                        "categories_to_add": ["@Company Announcements"],
                    }
                ],
                "Assistant": [
                    {
                        "index": 2,
                        "sender_name": "Assistant",
                        "sender_address": "assistant@corp.example",
                        "subject": "Calendar coordination",
                        "keep_in_inbox": False,
                        "move_to": "Assistant",
                        "categories_to_add": ["@Assistant"],
                    }
                ],
            },
            "action_summary": {
                "routes": {"Company Announcements": 1},
                "categories": {"@Company Announcements": 1},
            },
            "execution": None,
        }
    )

    output = capsys.readouterr().out
    assert "Company Announcements" in output
    assert "[red]Alert[/red]" in output
    assert "Monitoring" not in output
    message_line = next(line for line in output.splitlines() if "monitoring@" in line)
    assert "[red]Alert[/red]" in message_line
    assert "Move to" in output
    lines = output.splitlines()
    section_line = next(line for line in lines if "Company Announcements  1 message" in line)
    assert "─" in section_line
    assert not any("Email" in line and "Subject" in line and "Action" in line for line in lines)
    message_index = next(index for index, line in enumerate(lines) if "monitoring@" in line)
    assert lines[message_index + 1] == ""
    assert "Assistant  1 message" in lines[message_index + 2]
    assistant_index = next(index for index, line in enumerate(lines) if "assistant@" in line)
    assert lines[assistant_index + 1] == ""
    assert lines[assistant_index + 2].startswith("╭─")
    assert "Triage summary" in output
    assert output.index("monitoring@") < output.index("Outlook Organizer")
    assert output.index("Outlook Organizer") < output.index("Triage summary")
    summary_route_line = [
        line for line in output.splitlines() if "Move to Company Announcements" in line
    ][-1]
    route_suffix = summary_route_line.split("Move to Company Announcements", 1)[1]
    assert route_suffix.index("1") <= 4


def test_subjects_are_condensed_and_truncated() -> None:
    assert cli._condense_subject("First line\nSecond line") == "First line Second line"

    condensed = cli._condense_subject("x" * 100)

    assert condensed == ("x" * 71) + "…"
    assert len(condensed) == 72


def test_confirmed_execution_status_is_combined_in_footer(capsys) -> None:
    cli._print_triage_report(
        {
            "created_at": "2026-07-30T11:42:05.965028+00:00",
            "dry_run": False,
            "summary": {
                "messages": 10,
                "proposed_moves": 5,
                "kept_in_inbox": 5,
                "possible_spam": 0,
            },
            "sections": {},
            "action_summary": {
                "routes": {"Inbox": 5, "CXO": 5},
                "categories": {},
            },
            "execution": {
                "status": "completed",
                "applied": 10,
                "run_id": "run-33ed9007c89c",
                "error": None,
            },
        }
    )

    output = capsys.readouterr().out
    assert output.count("Outlook Organizer") == 1
    assert output.count("COMPLETED") == 1
    assert output.count("10 actions applied") == 1
    assert output.count("run-33ed9007c89c") == 1
    assert output.index("No messages found") < output.index("Outlook Organizer")
    assert output.index("Triage summary") < output.index("COMPLETED")
    status_lines = [line for line in output.splitlines() if "COMPLETED" in line]
    assert len(status_lines) == 1
    assert "run-33ed9007c89c" in status_lines[0]


def test_help_uses_organizer_command_names() -> None:
    runner = CliRunner()

    root_help = runner.invoke(cli.app, ["--help"])
    mail_help = runner.invoke(cli.app, ["mail", "--help"])
    calendar_help = runner.invoke(cli.app, ["calendar", "--help"])

    assert root_help.exit_code == 0
    assert "check" in root_help.output
    assert "history" in root_help.output
    assert "doctor" not in root_help.output
    assert "runs" not in root_help.output
    assert "triage" in mail_help.output
    assert "setup" in mail_help.output
    assert "show" in mail_help.output
    assert "folders" in mail_help.output
    assert "digest" not in mail_help.output
    assert "workload" in calendar_help.output
    assert "│ analyze " not in calendar_help.output


def test_legacy_mail_digest_alias_still_works(monkeypatch) -> None:
    monkeypatch.setattr(cli, "_service", FakeTriageService)

    result = CliRunner().invoke(
        cli.app,
        ["mail", "digest", "--limit", "2", "--no-progress"],
    )

    assert result.exit_code == 0
    assert "Outlook Organizer · Triage summary" in result.output
