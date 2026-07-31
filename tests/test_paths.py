from __future__ import annotations

from pathlib import Path

from outlook_organizer.paths import config_dir


def test_default_config_directory_is_outside_repository(monkeypatch) -> None:
    monkeypatch.delenv("OUTLOOK_ORGANIZER_CONFIG", raising=False)
    monkeypatch.delenv("OUTLOOK_DISTILLER_CONFIG", raising=False)

    assert config_dir() == Path.home() / ".config" / "outlook-organizer"


def test_primary_config_environment_variable_takes_precedence(
    monkeypatch, tmp_path
) -> None:
    primary = tmp_path / "primary"
    legacy = tmp_path / "legacy"
    monkeypatch.setenv("OUTLOOK_ORGANIZER_CONFIG", str(primary))
    monkeypatch.setenv("OUTLOOK_DISTILLER_CONFIG", str(legacy))

    assert config_dir() == primary.resolve()
