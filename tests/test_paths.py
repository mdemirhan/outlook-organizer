from __future__ import annotations

from pathlib import Path

from outlook_organizer.paths import config_dir


def test_default_config_directory_is_outside_repository(monkeypatch) -> None:
    monkeypatch.delenv("OUTLOOK_ORGANIZER_CONFIG", raising=False)

    assert config_dir() == Path.home() / ".config" / "outlook-organizer"


def test_config_environment_variable_overrides_default(monkeypatch, tmp_path) -> None:
    configured = tmp_path / "configured"
    monkeypatch.setenv("OUTLOOK_ORGANIZER_CONFIG", str(configured))

    assert config_dir() == configured.resolve()
