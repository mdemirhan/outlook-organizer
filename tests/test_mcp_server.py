from __future__ import annotations

import asyncio

import outlook_organizer.mcp_server as mcp_server


def test_mcp_surface_is_read_only_and_has_no_triage_or_history_tools() -> None:
    tools = asyncio.run(mcp_server.mcp.list_tools())
    names = {tool.name for tool in tools}
    assert "mail_brief" in names
    assert "mail_list_brief_profiles" in names
    assert "mail_search" in names
    assert "mail_triage" not in names
    assert "mail_setup" not in names
    assert "history_undo_run" not in names
    assert "history_list_runs" not in names
    assert all(tool.annotations.readOnlyHint for tool in tools)


def test_mcp_functions_are_callable() -> None:
    assert callable(mcp_server.mail_brief)
    assert callable(mcp_server.mail_list_brief_profiles)
    assert callable(mcp_server.calendar_get_agenda)
