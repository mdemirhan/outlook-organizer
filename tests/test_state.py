from __future__ import annotations

import outlook_organizer.state as state_module
from outlook_organizer.rules.triage import MailTriagePlanner
from outlook_organizer.state import StateStore


def test_plan_persistence(tmp_path, app_config, direct_message) -> None:
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(app_config).create_plan([direct_message], direct_message.folder_id)
    store.save_plan(plan)
    restored = store.load_plan(plan.plan_id)
    assert restored.config_fingerprint == plan.config_fingerprint
    assert restored.actions[0].subject == direct_message.subject
    assert store.plan_status(plan.plan_id) == "previewed"


def test_legacy_database_filename_is_migrated(
    tmp_path, monkeypatch, app_config, direct_message
) -> None:
    legacy_path = tmp_path / "outlook-distiller.sqlite"
    legacy_store = StateStore(legacy_path)
    plan = MailTriagePlanner(app_config).create_plan(
        [direct_message],
        direct_message.folder_id,
    )
    legacy_store.save_plan(plan)
    monkeypatch.setattr(state_module, "state_dir", lambda: tmp_path)

    migrated_store = StateStore()

    assert migrated_store.path == tmp_path / "outlook-organizer.sqlite"
    assert migrated_store.load_plan(plan.plan_id).plan_id == plan.plan_id
    assert not legacy_path.exists()
