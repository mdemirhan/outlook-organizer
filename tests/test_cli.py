from __future__ import annotations

from rich.console import Console as RichConsole
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


def test_triage_report_renders_message_content_as_literal_text(
    capsys, monkeypatch
) -> None:
    monkeypatch.setattr(cli, "Console", lambda: RichConsole(width=200))
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
                        "matched_rules": [
                            "route-company-announcements",
                            "thread-priority-promotion",
                        ],
                    }
                ],
                "Assistant": [
                    {
                        "index": 2,
                        "sender_name": "Assistant",
                        "sender_address": "assistant@corp.example",
                        "subject": "Calendar coordination",
                        "keep_in_inbox": True,
                        "move_to": None,
                        "categories_to_add": ["@Assistant"],
                        "matched_rules": ["route-assistant"],
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
    assert "→ Company Announcements" in output
    assert "Keep in Inbox" in output
    assert "Threading" in output
    assert "1 message  ·  0 possible spam" in output
    lines = output.splitlines()
    normalized_output = " ".join(output.split())
    assert "→ Company Announcements · Threading" in normalized_output
    section_line = next(line for line in lines if "Company Announcements  1 message" in line)
    assert "─" in section_line
    assert not any("Email" in line and "Subject" in line and "Action" in line for line in lines)
    message_index = next(index for index, line in enumerate(lines) if "monitoring@" in line)
    assistant_section_index = next(
        index for index, line in enumerate(lines) if "Assistant  1 message" in line
    )
    assert assistant_section_index > message_index
    assert "" in lines[message_index:assistant_section_index]
    assistant_index = next(index for index, line in enumerate(lines) if "assistant@" in line)
    assert lines[assistant_index + 1] == ""
    assert lines[assistant_index + 2].startswith("╭─")
    assert "Triage summary" in output
    assert output.index("monitoring@") < output.index("Outlook Organizer")
    assert output.index("Outlook Organizer") < output.index("Triage summary")
    route_index = next(
        index for index, line in enumerate(lines) if line.strip("│ ") == "Routes"
    )
    assert "Company Announcements" in lines[route_index + 1]
    assert "1" in lines[route_index + 2]


def test_subjects_are_condensed_and_truncated() -> None:
    assert cli._condense_subject("First line\nSecond line") == "First line Second line"

    condensed = cli._condense_subject("x" * 100)

    assert condensed == ("x" * 71) + "…"
    assert len(condensed) == 72


def test_confirmed_execution_status_is_combined_in_footer(capsys, monkeypatch) -> None:
    monkeypatch.setattr(cli, "Console", lambda: RichConsole(width=200))
    cli._print_triage_report(
        {
            "created_at": "2026-07-30T11:42:05.965028+00:00",
            "dry_run": False,
            "summary": {
                "messages": 10,
                "proposed_moves": 5,
                "kept_in_inbox": 5,
                "possible_spam": 0,
                "thread_routed": 99,
                "thread_promotions": 99,
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
                "thread_routed": 2,
                "promoted": 1,
            },
        }
    )

    output = capsys.readouterr().out
    normalized_output = " ".join(output.split())
    assert output.count("Outlook Organizer") == 1
    assert output.count("COMPLETED") == 1
    assert output.count("10 actions applied") == 1
    assert output.count("run-33ed9007c89c") == 1
    assert "2 routed by threading" in normalized_output
    assert "1 earlier message promoted" in normalized_output
    metric_lines = [line for line in output.splitlines() if "10 messages" in line]
    assert len(metric_lines) == 1
    assert "0 possible spam" in metric_lines[0]
    assert "2 routed by threading" in metric_lines[0]
    assert "99" not in output
    assert output.index("No messages found") < output.index("Outlook Organizer")
    assert output.index("Triage summary") < output.index("COMPLETED")
    status_lines = [line for line in output.splitlines() if "COMPLETED" in line]
    assert len(status_lines) == 1
    assert "run-33ed9007c89c" in status_lines[0]
    lines = output.splitlines()
    route_index = next(
        index for index, line in enumerate(lines) if line.strip("│ ") == "Routes"
    )
    assert "Inbox" in lines[route_index + 1]
    assert "CXO" in lines[route_index + 1]
    assert lines[route_index + 2].count("5") == 2


def test_successful_historical_promotions_render_in_a_separate_responsive_section(
    capsys, monkeypatch
) -> None:
    monkeypatch.setattr(cli, "Console", lambda: RichConsole(width=88))
    cli._print_triage_report(
        {
            "created_at": "2026-07-30T11:42:05+00:00",
            "dry_run": False,
            "summary": {
                "messages": 1,
                "proposed_moves": 1,
                "kept_in_inbox": 0,
                "possible_spam": 0,
                "thread_routed": 0,
                "thread_promotions": 2,
            },
            "sections": {
                "CXO": [
                    {
                        "index": 1,
                        "sender_address": "executive@corp.example",
                        "subject": "New reply",
                        "keep_in_inbox": False,
                        "move_to": "CXO",
                        "categories_to_add": [],
                        "matched_rules": [],
                    }
                ]
            },
            "promoted_messages": [
                {
                    "outlook_id": 41,
                    "sender_name": "Earlier Sender",
                    "sender_address": "earlier.sender@corp.example",
                    "subject": "[red]Earlier message[/red]",
                    "source_folder": "Turkcell General",
                    "destination_folder": "CXO",
                },
                {
                    "outlook_id": 40,
                    "sender_name": "Another Sender",
                    "sender_address": "another.sender@corp.example",
                    "subject": (
                        "A long historical conversation subject that still fits responsively"
                    ),
                    "source_folder": "My Directs",
                    "destination_folder": "CXO",
                },
            ],
            "action_summary": {"routes": {"CXO": 1}, "categories": {}},
            "execution": {
                "status": "completed",
                "applied": 3,
                "run_id": "run-promotions",
                "error": None,
                "thread_routed": 0,
                "promoted": 2,
            },
        }
    )

    output = capsys.readouterr().out
    normalized_output = " ".join(output.split())
    assert "Earlier messages promoted by threading 2 messages" in normalized_output
    assert "earlier.sender@" in output
    assert "[red]Earlier message[/red]" in normalized_output
    assert "Turkcell General → CXO" in normalized_output
    assert "My Directs → CXO" in normalized_output
    assert "2 earlier messages promoted" in normalized_output
    assert "No messages found" not in output
    assert output.index("executive@corp") < output.index("earlier.sender@")
    assert output.index("earlier.sender@") < output.index("Triage summary")


def test_route_summary_wraps_only_between_columns() -> None:
    rows = cli._route_summary_rows(
        {
            "Inbox": 10,
            "Turkcell General": 20,
            "My Directs": 5,
        },
        maximum_width=25,
    )

    assert len(rows) == 4
    assert "Inbox" in rows[0].plain
    assert "Turkcell General" in rows[0].plain
    assert "My Directs" in rows[2].plain
    assert all(len(row.plain) <= 25 for row in rows)


def test_help_uses_organizer_command_names() -> None:
    runner = CliRunner()

    root_help = runner.invoke(cli.app, ["--help"])
    mail_help = runner.invoke(cli.app, ["mail", "--help"])
    calendar_help = runner.invoke(cli.app, ["calendar", "--help"])

    assert root_help.exit_code == 0
    assert "check" in root_help.output
    assert "history" in root_help.output
    assert "triage" in mail_help.output
    assert "setup" in mail_help.output
    assert "show" in mail_help.output
    assert "folders" in mail_help.output
    assert "threads" in mail_help.output
    assert "workload" in calendar_help.output


def test_mail_threads_status_is_exposed(monkeypatch) -> None:
    class FakeThreadService:
        def thread_index_status(self):
            return {
                "enabled": True,
                "ready": True,
                "mode": "prospective",
                "threads": 3,
                "members": 5,
            }

    monkeypatch.setattr(cli, "_service", FakeThreadService)

    result = CliRunner().invoke(cli.app, ["mail", "threads", "status"])

    assert result.exit_code == 0
    assert '"mode": "prospective"' in result.output
    assert '"threads": 3' in result.output
