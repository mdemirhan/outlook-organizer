from __future__ import annotations

import json
import platform
from datetime import date
from pathlib import Path
from time import monotonic
from typing import Annotated

import typer
from rich.console import Console, Group
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
from rich.table import Table
from rich.text import Text

from outlook_organizer.mcp_server import run_mcp
from outlook_organizer.outlook import OutlookError
from outlook_organizer.service import OutlookOrganizerService

app = typer.Typer(help="Local-first Outlook email organization and calendar analysis")
config_app = typer.Typer(help="Validate and inspect configuration")
mail_app = typer.Typer(help="Review, organize, and search Outlook email")
threads_app = typer.Typer(help="Inspect prospective conversation-thread indexing")
calendar_app = typer.Typer(help="Inspect and analyze the Outlook calendar")
history_app = typer.Typer(help="Inspect or undo recorded change history")
app.add_typer(config_app, name="config")
app.add_typer(mail_app, name="mail")
mail_app.add_typer(threads_app, name="threads")
app.add_typer(calendar_app, name="calendar")
app.add_typer(history_app, name="history")


class _ConsoleProgress:
    def __init__(self, *, enabled: bool) -> None:
        self.enabled = enabled
        self.console = Console(stderr=True)
        self.started_at = 0.0
        self.description = ""
        self.task_id: int | None = None
        self.finished = False
        self.live = self.console.is_terminal
        self.progress = Progress(
            SpinnerColumn("dots", style="cyan"),
            TextColumn("[bold]{task.description}"),
            TimeElapsedColumn(),
            console=self.console,
            transient=True,
        )

    def __enter__(self) -> _ConsoleProgress:
        self.started_at = monotonic()
        if self.enabled and self.live:
            self.progress.start()
            self.task_id = self.progress.add_task("Starting Outlook Organizer", total=None)
        return self

    def update(self, description: str) -> None:
        if not self.enabled or description == self.description:
            return
        self.description = description
        if self.live and self.task_id is not None:
            self.progress.update(self.task_id, description=description, refresh=True)
        else:
            self.console.print(f"[cyan]→[/cyan] {description}")

    def finish(self, description: str, *, success: bool) -> None:
        if self.finished:
            return
        self.finished = True
        if self.enabled and self.live:
            self.progress.stop()
        if self.enabled:
            marker = "[green]✓[/green]" if success else "[red]×[/red]"
            elapsed = monotonic() - self.started_at
            self.console.print(f"{marker} {description} [dim]({elapsed:.1f}s)[/dim]")

    def __exit__(self, exc_type: object, exc: object, traceback: object) -> bool:
        if not self.finished:
            self.finish("Triage stopped" if exc else "Triage finished", success=exc is None)
        return False


def _print(value: object) -> None:
    typer.echo(json.dumps(value, ensure_ascii=False, indent=2, default=str))


def _service() -> OutlookOrganizerService:
    try:
        return OutlookOrganizerService()
    except Exception as exc:
        typer.echo(f"Configuration error: {exc}", err=True)
        raise typer.Exit(2) from exc


def _condense_subject(value: object, maximum_length: int = 72) -> str:
    subject = " ".join(str(value or "(no subject)").split())
    if len(subject) <= maximum_length:
        return subject
    return subject[: maximum_length - 1].rstrip() + "…"


def _route_summary_rows(
    routes: dict[str, int],
    maximum_width: int,
) -> list[Text]:
    entries = [
        (str(destination), str(count), max(len(str(destination)), len(str(count))))
        for destination, count in routes.items()
    ]
    groups: list[list[tuple[str, str, int]]] = []
    current: list[tuple[str, str, int]] = []
    current_width = 0
    for entry in entries:
        separator_width = 3 if current else 0
        projected_width = current_width + separator_width + entry[2]
        if current and projected_width > maximum_width:
            groups.append(current)
            current = []
            current_width = 0
            separator_width = 0
        current.append(entry)
        current_width += separator_width + entry[2]
    if current:
        groups.append(current)

    rows: list[Text] = []
    for group in groups:
        labels = Text()
        counts = Text()
        for index, (label, count, width) in enumerate(group):
            if index:
                labels.append(" │ ", style="bright_black")
                counts.append(" │ ", style="bright_black")
            labels.append(label.center(width), style="cyan")
            counts.append(count.center(width), style="bold")
        rows.extend([labels, counts])
    return rows


def _print_triage_report(report: dict) -> None:
    console = Console()
    dry_run = bool(report["dry_run"])
    summary = report["summary"]
    execution = report.get("execution")
    promoted_messages = report.get("promoted_messages", [])

    status_line = Text(overflow="ellipsis", no_wrap=True)
    if dry_run:
        status_line.append("DRY RUN", style="bold yellow")
        status_line.append("  ·  No Outlook changes")
        status_line.append("  ·  Add --apply to organize these messages.", style="dim")
    elif execution:
        status = execution["status"]
        completed = status == "completed"
        status_style = "bold green" if completed else "bold red"
        applied = execution["applied"]
        action_label = "action" if applied == 1 else "actions"
        status_line.append(status.upper(), style=status_style)
        status_line.append(f"  ·  {applied} {action_label} applied")
        status_line.append(f"  ·  Run ID {execution['run_id']}", style="dim")
        if execution.get("error"):
            status_line.append(f"  ·  {execution['error']}", style="red")
    else:
        status_line.append("CONFIRMED RUN", style="bold green")
    if report.get("created_at"):
        raw_timestamp = str(report["created_at"])
        compact_timestamp = raw_timestamp[:16].replace("T", " ")
        if raw_timestamp.endswith(("Z", "+00:00")):
            compact_timestamp += " UTC"
        elif len(raw_timestamp) >= 6 and raw_timestamp[-6] in "+-":
            compact_timestamp += raw_timestamp[-6:]
        status_line.append(f"  ·  {compact_timestamp}", style="dim")

    summary_line = Text()
    message_count = int(summary["messages"])
    summary_line.append(str(message_count), style="bold cyan")
    summary_line.append(" message" if message_count == 1 else " messages", style="dim")
    summary_line.append("  ·  ", style="bright_black")
    spam_count = summary["possible_spam"]
    summary_line.append(
        str(spam_count),
        style="bold red" if spam_count else "bold green",
    )
    summary_line.append(" possible spam", style="dim")
    thread_routed = (
        int(execution.get("thread_routed", 0))
        if execution
        else int(summary.get("thread_routed", 0))
    )
    thread_promotions = (
        int(execution.get("promoted", 0))
        if execution
        else int(summary.get("thread_promotions", 0))
    )
    if thread_routed:
        summary_line.append("  ·  ", style="bright_black")
        summary_line.append(str(thread_routed), style="bold magenta")
        summary_line.append(" routed by threading", style="dim")
    if thread_promotions:
        summary_line.append("  ·  ", style="bright_black")
        summary_line.append(str(thread_promotions), style="bold magenta")
        summary_line.append(
            (
                " earlier message promoted"
                if thread_promotions == 1
                else " earlier messages promoted"
            ),
            style="dim",
        )

    route_rows = _route_summary_rows(
        report["action_summary"]["routes"],
        maximum_width=max(20, console.width - 6),
    )

    summary_parts: list[object] = [
        status_line,
        summary_line,
    ]
    summary_parts.append(Text("\nRoutes", style="bold"))
    summary_parts.extend(route_rows)
    category_counts = report["action_summary"]["categories"]
    if category_counts:
        category_text = Text("\nCategories  ", style="bold")
        for index, (category, count) in enumerate(category_counts.items()):
            if index:
                category_text.append("  ·  ", style="dim")
            category_text.append(str(category), style="magenta")
            category_text.append(f" {count}", style="bold")
        summary_parts.append(category_text)

    summary_panel = Panel(
        Group(*summary_parts),
        title=Text(" Outlook Organizer · Triage summary ", style="bold cyan"),
        title_align="left",
        border_style="bright_black",
        padding=(0, 1),
    )

    if not report["sections"] and not promoted_messages:
        console.print(
            Panel(
                Text("No messages found for this triage.", style="dim", justify="center"),
                border_style="bright_black",
            )
        )

    section_entries = list(report["sections"].items())
    for section_index, (section_name, items) in enumerate(section_entries):
        section_title = Text(str(section_name), style="bold cyan")
        section_title.append(f"  {len(items)}", style="bold")
        section_title.append(" message" if len(items) == 1 else " messages", style="dim")
        console.rule(section_title, align="left", style="bright_black")
        table = Table.grid(expand=True, padding=(0, 1))
        table.add_column(justify="right", style="dim", no_wrap=True, width=3)
        table.add_column(ratio=3, max_width=38, overflow="ellipsis", no_wrap=True)
        table.add_column(ratio=5, overflow="ellipsis", no_wrap=True)
        table.add_column(ratio=3, max_width=38, overflow="fold")
        table.add_column(
            ratio=2,
            max_width=28,
            overflow="ellipsis",
            no_wrap=True,
        )

        for item in items:
            sender_address = item["sender_address"] or "Unknown sender"
            destination = (
                "Keep in Inbox"
                if item["keep_in_inbox"] or not item["move_to"]
                else f"→ {item['move_to']}"
            )
            action_style = "cyan" if destination == "Keep in Inbox" else "green"
            destination_text = Text(destination, style=action_style)
            if any(
                rule_id in {"thread-affinity", "thread-priority-promotion"}
                for rule_id in item.get("matched_rules", [])
            ):
                destination_text.append(" · Threading", style="bold magenta")
            categories = Text()
            if item["categories_to_add"]:
                for category_index, category in enumerate(item["categories_to_add"]):
                    if category_index:
                        categories.append(", ", style="dim")
                    categories.append(str(category), style="magenta")
            else:
                categories.append("None", style="dim")
            table.add_row(
                str(item["index"]),
                Text(str(sender_address), style="bold"),
                Text(_condense_subject(item["subject"])),
                destination_text,
                categories,
            )
        console.print(table)
        if section_index + 1 < len(section_entries):
            console.print()

    if promoted_messages:
        if section_entries:
            console.print()
        promotion_title = Text("Earlier messages promoted by threading", style="bold magenta")
        promotion_title.append(f"  {len(promoted_messages)}", style="bold")
        promotion_title.append(
            " message" if len(promoted_messages) == 1 else " messages",
            style="dim",
        )
        console.rule(promotion_title, align="left", style="bright_black")
        promotion_table = Table.grid(expand=True, padding=(0, 1))
        promotion_table.add_column(justify="right", style="dim", no_wrap=True, width=3)
        promotion_table.add_column(ratio=4, max_width=44, overflow="fold")
        promotion_table.add_column(
            ratio=3,
            max_width=38,
            overflow="ellipsis",
            no_wrap=True,
        )
        promotion_table.add_column(ratio=5, overflow="ellipsis", no_wrap=True)
        for index, item in enumerate(promoted_messages, start=1):
            route = Text(str(item["source_folder"]), style="dim")
            route.append(" → ", style="bright_black")
            route.append(str(item["destination_folder"]), style="green")
            promotion_table.add_row(
                str(index),
                route,
                Text(str(item.get("sender_address") or "Unknown sender"), style="bold"),
                Text(_condense_subject(item.get("subject"))),
            )
        console.print(promotion_table)

    console.print()
    console.print(summary_panel)


@app.command("check")
def check() -> None:
    """Check configuration and local Outlook visibility."""
    service = _service()
    result: dict[str, object] = {
        "platform": platform.platform(),
        "config": service.validate(),
        "outlook_app_exists": Path("/Applications/Microsoft Outlook.app").exists(),
    }
    try:
        inbox_config = service.config.mail.folders["inbox"]
        inbox = service.adapter.find_folder(
            inbox_config.names,
            service.config.mail.folder_scan_limit,
        )
        calendar = service.adapter.find_calendar(
            service.config.calendar.calendar_names,
            service.config.calendar.maximum_calendar_id,
        )
        result["inbox"] = {
            "id": inbox.outlook_id,
            "name": inbox.name,
            "message_count": inbox.message_count,
        }
        result["configured_folders"] = service.configured_folder_status()
        result["calendar"] = {
            "id": calendar.outlook_id,
            "name": calendar.name,
            "event_count": calendar.event_count,
        }
    except OutlookError as exc:
        result["outlook_error"] = str(exc)
    _print(result)


@config_app.command("validate")
def config_validate() -> None:
    """Validate YAML configuration and rule predicates."""
    _print(_service().validate())


@mail_app.command("folders")
def mail_folders() -> None:
    """List scriptable Outlook mail folders."""
    _print(_service().folders())


@mail_app.command("setup")
def mail_setup(
    confirm: Annotated[bool, typer.Option("--confirm")] = False,
) -> None:
    """Ensure the configured mail root folders exist. Requires --confirm."""
    _print(_service().setup_mail_folders(confirm=confirm))


@mail_app.command("triage")
def mail_triage(
    limit: Annotated[int, typer.Option(min=1, max=500)] = 50,
    body_limit: Annotated[int, typer.Option(min=0, max=50_000)] = 0,
    apply: Annotated[
        bool,
        typer.Option("--apply", "--confirm", help="Apply the proposed Outlook changes."),
    ] = False,
    progress: Annotated[
        bool,
        typer.Option("--progress/--no-progress", help="Show live progress on stderr."),
    ] = True,
) -> None:
    """Preview mail triage, or apply it immediately with --apply."""
    with _ConsoleProgress(enabled=progress) as progress_view:
        progress_view.update("Loading configuration")
        report = _service().triage_mail(
            limit=limit,
            body_limit=body_limit,
            confirm=apply,
            progress=progress_view.update,
        )
        execution = report.get("execution")
        succeeded = execution is None or execution["status"] == "completed"
        if execution is None:
            completion = "Triage ready — no Outlook changes"
        else:
            action_count = execution["applied"]
            action_label = "action" if action_count == 1 else "actions"
            completion = f"Triage {execution['status']} — {action_count} {action_label} applied"
        progress_view.finish(completion, success=succeeded)

    _print_triage_report(report)
    if execution and execution["status"] != "completed":
        raise typer.Exit(1)


@mail_app.command("search")
def mail_search(
    query: str,
    limit: Annotated[int, typer.Option(min=1, max=100)] = 20,
    include_body: bool = False,
) -> None:
    """Search recent Inbox messages."""
    _print(_service().search_messages(query, limit=limit, include_body=include_body))


@mail_app.command("show")
def mail_show(outlook_id: int, include_body: bool = False) -> None:
    """Read one Outlook message by its actionable Outlook ID."""
    _print(_service().get_message(outlook_id, include_body=include_body))


@threads_app.command("status")
def mail_threads_status() -> None:
    """Show whether the authoritative local thread index is ready."""
    _print(_service().thread_index_status())


@history_app.command("list")
def history_list(limit: Annotated[int, typer.Option(min=1, max=100)] = 20) -> None:
    """List recent applied runs and their status."""
    _print(_service().recent_runs(limit))


@history_app.command("undo")
def history_undo(
    run_id: str,
    confirm: Annotated[bool, typer.Option("--confirm")] = False,
) -> None:
    """Undo a completed email run. Requires --confirm."""
    _print(_service().undo_run(run_id, confirm=confirm))


@calendar_app.command("list")
def calendar_list() -> None:
    """List scriptable Outlook calendars."""
    _print(_service().calendars())


@calendar_app.command("agenda")
def calendar_agenda(
    days_ahead: Annotated[int, typer.Option(min=1, max=90)] = 7,
    days_behind: Annotated[int, typer.Option(min=0, max=30)] = 0,
    include_body: bool = False,
) -> None:
    """Show calendar events around today."""
    _print(
        _service().calendar_events(
            days_ahead=days_ahead,
            days_behind=days_behind,
            include_body=include_body,
        )
    )


@calendar_app.command("workload")
def calendar_workload(
    days_ahead: Annotated[int, typer.Option(min=1, max=90)] = 7,
    days_behind: Annotated[int, typer.Option(min=0, max=30)] = 0,
) -> None:
    """Summarize meeting hours, conflicts, and back-to-back meetings."""
    _print(_service().analyze_calendar(days_ahead=days_ahead, days_behind=days_behind))


@calendar_app.command("free-slots")
def calendar_free_slots(
    target_date: str,
    minimum_minutes: Annotated[int | None, typer.Option(min=15, max=480)] = None,
) -> None:
    """Find available focus blocks during configured working hours."""
    try:
        parsed_date = date.fromisoformat(target_date)
    except ValueError as exc:
        raise typer.BadParameter("target_date must use YYYY-MM-DD") from exc
    _print(_service().find_free_slots(parsed_date, minimum_minutes=minimum_minutes))


@app.command("mcp")
def mcp_command() -> None:
    """Run the local MCP server over STDIO."""
    run_mcp()


if __name__ == "__main__":
    app()
