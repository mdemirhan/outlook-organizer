from __future__ import annotations

import subprocess
import threading
from dataclasses import dataclass


class OutlookError(RuntimeError):
    """Raised when Outlook AppleScript automation fails."""


@dataclass(slots=True)
class AppleScriptResult:
    stdout: str
    stderr: str


class AppleScriptRunner:
    """Serialize AppleScript calls because Outlook automation is single-threaded."""

    def __init__(self, timeout_seconds: int = 90) -> None:
        self.timeout_seconds = timeout_seconds
        self._lock = threading.Lock()

    def run(self, source: str, *args: str) -> AppleScriptResult:
        command = ["/usr/bin/osascript", "-e", source, *args]
        with self._lock:
            try:
                completed = subprocess.run(
                    command,
                    check=False,
                    capture_output=True,
                    text=True,
                    timeout=self.timeout_seconds,
                )
            except subprocess.TimeoutExpired as exc:
                raise OutlookError(
                    f"Outlook scripting timed out after {self.timeout_seconds}s"
                ) from exc

        if completed.returncode != 0:
            message = completed.stderr.strip() or completed.stdout.strip()
            raise OutlookError(message or f"osascript exited with {completed.returncode}")
        return AppleScriptResult(stdout=completed.stdout, stderr=completed.stderr)
