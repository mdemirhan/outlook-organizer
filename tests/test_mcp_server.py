from __future__ import annotations

import outlook_organizer.mcp_server as mcp_server


def test_mcp_exports_organizer_tool_names() -> None:
    assert callable(mcp_server.mail_setup)
    assert callable(mcp_server.mail_triage)
    assert callable(mcp_server.mail_thread_index_status)
    assert callable(mcp_server.history_list_runs)
    assert callable(mcp_server.history_undo_run)
