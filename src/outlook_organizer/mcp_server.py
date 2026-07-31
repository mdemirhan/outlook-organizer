from __future__ import annotations

from datetime import date
from typing import Any

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from outlook_organizer.service import OutlookOrganizerService

mcp = FastMCP(
    "Outlook Organizer",
    instructions=(
        "Local Outlook for Mac tools. mail_triage is read-only unless "
        "confirm=true; with confirmation it computes and applies one triage plan. "
        "Never claim a write succeeded unless the tool result reports completed."
    ),
)


def _service() -> OutlookOrganizerService:
    return OutlookOrganizerService()


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def outlook_validate_config() -> dict[str, Any]:
    """Validate mail definitions, ordered rules, folder references, and calendar config."""
    return _service().validate()


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def mail_list_folders() -> list[dict[str, Any]]:
    """List actionable Outlook mail folders and message counts."""
    return _service().folders()


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=False,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def mail_setup(confirm: bool = False) -> dict[str, Any]:
    """Ensure the configured Exchange mail roots exist. Requires confirmation."""
    return _service().setup_mail_folders(confirm=confirm)


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=False,
        destructiveHint=False,
        idempotentHint=False,
        openWorldHint=False,
    )
)
def mail_triage(
    limit: int = 50,
    body_limit: int = 0,
    confirm: bool = False,
) -> dict[str, Any]:
    """Preview mail triage, or compute and apply it immediately with confirmation."""
    return _service().triage_mail(
        limit=limit,
        body_limit=body_limit,
        confirm=confirm,
    )


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def mail_get_message(outlook_id: int, include_body: bool = False) -> dict[str, Any]:
    """Read a message by Outlook ID. Body text is omitted unless explicitly requested."""
    return _service().get_message(outlook_id, include_body=include_body)


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def mail_search(
    query: str,
    limit: int = 20,
    include_body: bool = False,
) -> list[dict[str, Any]]:
    """Search recent Inbox messages by sender, subject, or body text."""
    return _service().search_messages(query, limit=limit, include_body=include_body)


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=False,
        destructiveHint=False,
        idempotentHint=False,
        openWorldHint=False,
    )
)
def history_undo_run(run_id: str, confirm: bool = False) -> dict[str, Any]:
    """Restore folder, categories, and flag state recorded before a run."""
    return _service().undo_run(run_id, confirm=confirm)


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def history_list_runs(limit: int = 20) -> list[dict[str, Any]]:
    """List recent apply runs and their status."""
    return _service().recent_runs(limit)


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def calendar_list_calendars() -> list[dict[str, Any]]:
    """List actionable Outlook calendars."""
    return _service().calendars()


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def calendar_get_agenda(
    days_ahead: int = 7,
    days_behind: int = 0,
    include_body: bool = False,
) -> list[dict[str, Any]]:
    """Read calendar events around today. Private event details are redacted."""
    return _service().calendar_events(
        days_ahead=days_ahead,
        days_behind=days_behind,
        include_body=include_body,
    )


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def calendar_analyze_workload(
    days_ahead: int = 7,
    days_behind: int = 0,
) -> dict[str, Any]:
    """Calculate meeting hours, conflicts, and back-to-back meetings."""
    return _service().analyze_calendar(days_ahead=days_ahead, days_behind=days_behind)


@mcp.tool(
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    )
)
def calendar_find_free_slots(
    target_date: str,
    minimum_minutes: int | None = None,
) -> list[dict[str, str | int]]:
    """Find focus-time slots on an ISO date using configured working hours."""
    return _service().find_free_slots(
        date.fromisoformat(target_date), minimum_minutes=minimum_minutes
    )


def run_mcp() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    run_mcp()
