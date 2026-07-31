from __future__ import annotations

from collections.abc import Callable

ProgressCallback = Callable[[str], None]


def format_count(count: int, noun: str) -> str:
    suffix = "" if count == 1 else "s"
    return f"{count} {noun}{suffix}"


def report_progress(callback: ProgressCallback | None, description: str) -> None:
    if callback is not None:
        callback(description)
