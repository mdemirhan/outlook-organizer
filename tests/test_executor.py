from __future__ import annotations

from outlook_organizer.actions import ActionExecutor
from outlook_organizer.models import FlagStatus, OutlookFolder, Recipient
from outlook_organizer.rules import MailTriagePlanner
from outlook_organizer.service import OutlookOrganizerService
from outlook_organizer.state import StateStore


class FakeOutlookAdapter:
    def __init__(self, message) -> None:
        self.message = message
        self.applied: list[dict] = []
        self.ensured_folders: list[str] = []
        self.folder_lookups = 0

    def find_folder(self, names, maximum_id):
        self.folder_lookups += 1
        return OutlookFolder(500, names[0], "", 0)

    def get_message(self, outlook_id, body_limit=0):
        assert outlook_id == self.message.outlook_id
        return self.message

    def latest_messages(self, folder_id, limit=20, body_limit=2000):
        return [self.message][:limit]

    def get_messages(self, outlook_ids, body_limit=0):
        wanted = set(outlook_ids)
        return [self.message] if self.message.outlook_id in wanted else []

    def apply_mail_state(
        self,
        outlook_id,
        *,
        categories,
        flag_status,
        target_folder_id=None,
    ):
        self.applied.append(
            {
                "outlook_id": outlook_id,
                "categories": categories,
                "flag_status": flag_status,
                "target_folder_id": target_folder_id,
            }
        )
        self.message.categories = list(categories)
        self.message.flag_status = flag_status
        if target_folder_id is not None:
            self.message.folder_id = target_folder_id

    def apply_mail_states(self, updates):
        results = []
        for update in updates:
            self.apply_mail_state(
                update["outlook_id"],
                categories=update["categories"],
                flag_status=update["flag_status"],
                target_folder_id=update["target_folder_id"],
            )
            results.append(
                {
                    "outlook_id": update["outlook_id"],
                    "status": "applied",
                    "error": "",
                }
            )
        return results

    def ensure_mail_folder(self, inbox_id, folder_name):
        self.ensured_folders.append(folder_name)
        return {
            "outlook_id": 900 + len(self.ensured_folders),
            "name": folder_name,
            "status": "existing",
        }


def test_apply_is_audited_and_undo_restores_state(tmp_path, app_config, direct_message) -> None:
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(app_config).create_plan([direct_message], direct_message.folder_id)
    store.save_plan(plan)
    executor = ActionExecutor(app_config, adapter, store)

    result = executor.apply_plan(plan)

    assert result.status == "completed"
    assert result.applied == 1
    assert adapter.folder_lookups == 0
    assert adapter.applied[0]["target_folder_id"] == 110
    assert adapter.applied[0]["categories"] == ["@Internal General", "@Only Me"]

    undo = executor.undo_run(result.run_id)

    assert undo.status == "undone"
    assert undo.applied == 1
    assert adapter.applied[1] == {
        "outlook_id": direct_message.outlook_id,
        "categories": [],
        "flag_status": FlagStatus.NOT_FLAGGED,
        "target_folder_id": direct_message.folder_id,
    }
    assert store.plan_status(plan.plan_id) == "undone"


def test_configured_junk_uses_cached_junk_external_folder_id(
    tmp_path, app_config, direct_message
) -> None:
    direct_message.sender_address = "sender@unwanted.example"
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(app_config).create_plan([direct_message], direct_message.folder_id)

    result = ActionExecutor(app_config, adapter, store).apply_plan(plan)

    assert result.status == "completed"
    assert adapter.folder_lookups == 0
    assert adapter.applied[0]["target_folder_id"] == 111


def test_apply_reclassifies_message_from_fresh_reread(
    tmp_path, app_config, direct_message
) -> None:
    direct_message.to = []
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(app_config).create_plan(
        [direct_message], direct_message.folder_id
    )
    store.save_plan(plan)
    assert plan.actions[0].add_categories == ["@Internal General"]

    direct_message.to = [
        Recipient("Example User", "example.user@corp.example")
    ]

    result = ActionExecutor(app_config, adapter, store).apply_plan(plan)

    assert result.status == "completed"
    assert adapter.applied[0]["categories"] == ["@Internal General", "@Only Me"]
    refreshed = store.load_plan(plan.plan_id).actions[0]
    assert refreshed.add_categories == ["@Internal General", "@Only Me"]
    assert [match.rule_id for match in refreshed.matches] == [
        "sent-only-to-me",
        "route-internal-general",
    ]


def test_apply_reclassifies_flag_change_before_moving(
    tmp_path, app_config, direct_message
) -> None:
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(app_config).create_plan(
        [direct_message], direct_message.folder_id
    )
    store.save_plan(plan)
    assert plan.actions[0].move_to == "internal_general"

    direct_message.flag_status = FlagStatus.FLAGGED

    result = ActionExecutor(app_config, adapter, store).apply_plan(plan)

    assert result.status == "completed"
    assert adapter.applied[0]["target_folder_id"] is None
    assert adapter.applied[0]["categories"] == [
        "@Action",
        "@Internal General",
        "@Only Me",
    ]
    refreshed = store.load_plan(plan.plan_id).actions[0]
    assert refreshed.keep_in_inbox
    assert refreshed.move_to is None


def test_apply_preserves_an_existing_rule_category(
    tmp_path, app_config, direct_message
) -> None:
    direct_message.categories = ["@Only Me"]
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(app_config).create_plan(
        [direct_message], direct_message.folder_id
    )

    result = ActionExecutor(app_config, adapter, store).apply_plan(plan)

    assert result.status == "completed"
    assert adapter.applied[0]["categories"] == ["@Internal General", "@Only Me"]


def test_explicit_empty_selection_applies_nothing(tmp_path, app_config, direct_message) -> None:
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(app_config).create_plan([direct_message], direct_message.folder_id)
    store.save_plan(plan)
    service = OutlookOrganizerService(app_config, adapter, store)

    result = service.apply_plan(plan.plan_id, confirm=True, selected_indexes=[])

    assert result["status"] == "completed"
    assert result["applied"] == 0
    assert adapter.applied == []
    assert adapter.folder_lookups == 0


def test_mail_setup_ensures_both_organized_roots(
    tmp_path, app_config, direct_message
) -> None:
    adapter = FakeOutlookAdapter(direct_message)
    service = OutlookOrganizerService(
        app_config,
        adapter,
        StateStore(tmp_path / "state.sqlite"),
    )

    result = service.setup_mail_folders(confirm=True)

    assert adapter.ensured_folders == ["aOrganized", "bOrganized"]
    assert list(result["folders"]) == ["organized_primary", "organized_secondary"]


def test_undo_restores_a_failed_partial_action(tmp_path, app_config, direct_message) -> None:
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(app_config).create_plan([direct_message], direct_message.folder_id)
    store.save_plan(plan)
    run_id = store.start_run(plan.plan_id)
    store.record_action(
        run_id=run_id,
        sequence=1,
        outlook_id=direct_message.outlook_id,
        message_id=direct_message.stable_id,
        subject=direct_message.subject,
        status="failed",
        before_state={
            "folder_id": direct_message.folder_id,
            "folder_name": direct_message.folder_name,
            "categories": [],
            "flag_status": "not_flagged",
        },
        after_state={
            "folder_id": 110,
            "folder_name": "Internal General",
            "categories": ["@Internal General", "@Only Me"],
            "flag_status": "not_flagged",
        },
        error="partial Outlook mutation",
    )
    store.finish_run(run_id, "failed", "partial Outlook mutation")
    direct_message.categories = ["@Internal General", "@Only Me"]

    result = ActionExecutor(app_config, adapter, store).undo_run(run_id)

    assert result.status == "undone"
    assert result.applied == 1
    assert adapter.applied == [
        {
            "outlook_id": direct_message.outlook_id,
            "categories": [],
            "flag_status": FlagStatus.NOT_FLAGGED,
            "target_folder_id": None,
        }
    ]
    assert store.plan_status(plan.plan_id) == "undone"


def test_triage_mail_dry_run_does_not_persist_or_apply(
    tmp_path, app_config, direct_message
) -> None:
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    service = OutlookOrganizerService(app_config, adapter, store)
    progress: list[str] = []

    result = service.triage_mail(limit=1, body_limit=0, progress=progress.append)

    with store.connect() as connection:
        assert connection.execute("SELECT count(*) FROM plans").fetchone()[0] == 0
        assert connection.execute("SELECT count(*) FROM runs").fetchone()[0] == 0
    assert result["dry_run"]
    assert result["execution"] is None
    assert adapter.applied == []
    assert result["sections"]["Internal General"][0]["move_to"] == "Internal General"
    assert result["action_summary"]["routes"] == {"Internal General": 1}
    assert progress == [
        "Reading up to 1 message from Outlook",
        "Classifying 1 message",
        "Building the mail triage report",
    ]


def test_configured_folder_status_checks_names_and_parents(
    tmp_path, app_config, direct_message
) -> None:
    adapter = FakeOutlookAdapter(direct_message)
    actual_folders = []
    for folder in app_config.mail.folders.values():
        parent_name = (
            app_config.mail.folders[folder.parent].name if folder.parent else ""
        )
        actual_folders.append(OutlookFolder(folder.id, folder.name, parent_name, 0))
    adapter.list_folders = lambda maximum_id: actual_folders
    service = OutlookOrganizerService(
        app_config, adapter, StateStore(tmp_path / "state.sqlite")
    )

    status = service.configured_folder_status()

    assert status["valid"]
    assert status["folders"]["junk_external"]["actual"]["parent"] == "bOrganized"


def test_confirmed_triage_mail_persists_and_applies_in_one_call(
    tmp_path, app_config, direct_message
) -> None:
    adapter = FakeOutlookAdapter(direct_message)
    store = StateStore(tmp_path / "state.sqlite")
    service = OutlookOrganizerService(app_config, adapter, store)
    progress: list[str] = []

    result = service.triage_mail(
        limit=1,
        body_limit=0,
        confirm=True,
        progress=progress.append,
    )

    with store.connect() as connection:
        assert connection.execute("SELECT count(*) FROM plans").fetchone()[0] == 1
        assert connection.execute("SELECT count(*) FROM runs").fetchone()[0] == 1
    assert not result["dry_run"]
    assert result["execution"]["status"] == "completed"
    assert result["execution"]["applied"] == 1
    assert len(adapter.applied) == 1
    assert progress == [
        "Reading up to 1 message from Outlook",
        "Classifying 1 message",
        "Saving the confirmed triage plan",
        "Verifying 1 message before making changes",
        "Preparing 1 Outlook update",
        "Applying 1 update in Outlook",
        "Recording the audit trail",
        "Building the mail triage report",
    ]
