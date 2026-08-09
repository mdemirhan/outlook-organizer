from __future__ import annotations

from collections.abc import Iterable
from typing import Any, Protocol


class AuditRepository(Protocol):
    def begin_run(self, *, config_fingerprint: str, parameters: dict[str, Any]) -> str: ...

    def record_action(
        self,
        *,
        run_id: str,
        sequence: int,
        outlook_id: int,
        message_id: str,
        subject: str,
        decision: dict[str, Any],
        before_state: dict[str, Any],
        intended_state: dict[str, Any],
        actual_state: dict[str, Any] | None,
        status: str,
        error: str | None = None,
    ) -> None: ...

    def finish_run(self, run_id: str, status: str, error: str | None = None) -> None: ...

    def list_runs(self, limit: int = 20) -> list[dict[str, Any]]: ...

    def run_actions(self, run_id: str, *, reverse: bool = False) -> list[Any]: ...

    def mark_run_undone(self, run_id: str, action_ids: Iterable[int]) -> None: ...
