from __future__ import annotations

import shutil
from pathlib import Path

import pytest
import yaml

from outlook_organizer.brief import load_brief_context


def copied_config(tmp_path: Path) -> Path:
    source = Path(__file__).resolve().parents[1] / "config"
    destination = tmp_path / "config"
    shutil.copytree(source, destination)
    return destination


def test_brief_profile_folder_references_are_validated(tmp_path) -> None:
    directory = copied_config(tmp_path)
    path = directory / "brief.yaml"
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    data["profiles"]["morning"]["scopes"][0]["folder"] = "missing-folder"
    path.write_text(yaml.safe_dump(data, sort_keys=False), encoding="utf-8")

    with pytest.raises(ValueError, match="missing-folder"):
        load_brief_context(directory)


def test_brief_profile_aliases_must_be_unambiguous(tmp_path) -> None:
    directory = copied_config(tmp_path)
    path = directory / "brief.yaml"
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    data["profiles"]["quick-check"]["aliases"] = ["Morning brief"]
    path.write_text(yaml.safe_dump(data, sort_keys=False), encoding="utf-8")

    with pytest.raises(ValueError, match="shared by"):
        load_brief_context(directory)
