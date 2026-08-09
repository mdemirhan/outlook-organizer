from __future__ import annotations

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
