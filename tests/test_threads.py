from __future__ import annotations

from dataclasses import replace

from outlook_organizer.actions import ActionExecutor
from outlook_organizer.config import AppConfig
from outlook_organizer.models import (
    FlagStatus,
    MailMessage,
    ThreadMessageState,
)
from outlook_organizer.rules import MailTriagePlanner
from outlook_organizer.state import StateStore
from outlook_organizer.threads import ThreadRouter


def threaded_config(app_config) -> AppConfig:
    mail = app_config.mail.model_copy(deep=True)
    mail.threading.enabled = True
    return replace(app_config, mail=mail, fingerprint="threaded-config")


def message_copy(message, **changes) -> MailMessage:
    return replace(message, **changes)


class ThreadStateAdapter:
    def __init__(self, states: list[ThreadMessageState] | None = None) -> None:
        self.states = {state.outlook_id: state for state in states or []}
        self.requested_ids: list[list[int]] = []

    def thread_states_by_ids(self, outlook_ids):
        ids = list(outlook_ids)
        self.requested_ids.append(ids)
        return [self.states[value] for value in ids if value in self.states]


def seed_thread(
    store: StateStore,
    *,
    thread_guid: str,
    folder_key: str,
    members: list[tuple[int, int, str, bool]],
) -> None:
    store.update_thread_index(
        scope="inbox:101",
        routes={thread_guid: folder_key},
        members=[
            {
                "thread_guid": thread_guid,
                "outlook_id": outlook_id,
                "message_id": f"exchange-{outlook_id}",
                "folder_id": folder_id,
                "folder_key": member_folder_key,
                "detached": detached,
            }
            for outlook_id, folder_id, member_folder_key, detached in members
        ],
    )


def plan_actions(config, messages):
    planner = MailTriagePlanner(config)
    return [planner.plan_message(message) for message in messages]


def test_threading_is_disabled_by_default(app_config, direct_message, tmp_path) -> None:
    adapter = ThreadStateAdapter()
    router = ThreadRouter(app_config, adapter, StateStore(tmp_path / "state.sqlite"))
    direct_message.thread_guid = "thread-1"
    action = MailTriagePlanner(app_config).plan_message(direct_message)

    resolution = router.resolve([direct_message], [action])

    assert not router.status()["enabled"]
    assert not router.status()["ready"]
    assert resolution.actions[0].move_to == "internal_general"
    assert adapter.requested_ids == []


def test_sql_miss_is_authoritative_and_never_queries_outlook(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    adapter = ThreadStateAdapter()
    store = StateStore(tmp_path / "state.sqlite")
    router = ThreadRouter(config, adapter, store)
    direct_message.thread_guid = "unknown-thread"

    resolution = router.resolve(
        [direct_message],
        plan_actions(config, [direct_message]),
    )

    assert resolution.actions[0].move_to == "internal_general"
    assert resolution.promotions == []
    assert adapter.requested_ids == []
    assert store.thread_contexts(router.scope, ["unknown-thread"]) == {}


def test_new_messages_in_the_same_batch_choose_the_highest_priority_route(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    general = message_copy(
        direct_message,
        outlook_id=43,
        exchange_id="exchange-43",
        thread_guid="batch-thread",
    )
    leadership = message_copy(
        direct_message,
        outlook_id=44,
        exchange_id="exchange-44",
        sender_address="leader@corp.example",
        thread_guid="batch-thread",
    )
    router = ThreadRouter(
        config,
        ThreadStateAdapter(),
        StateStore(tmp_path / "state.sqlite"),
    )

    resolution = router.resolve(
        [general, leadership],
        plan_actions(config, [general, leadership]),
    )

    assert [action.move_to for action in resolution.actions] == [
        "leadership",
        "leadership",
    ]
    assert resolution.actions[0].matches[-1].rule_id == "thread-priority-promotion"
    assert resolution.promotions == []


def test_known_thread_inherits_existing_higher_priority_folder(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="known-thread",
        folder_key="leadership",
        members=[(41, 106, "leadership", False)],
    )
    adapter = ThreadStateAdapter(
        [ThreadMessageState(41, "exchange-41", 106, "Leadership", "known-thread")]
    )
    direct_message.thread_guid = "known-thread"
    router = ThreadRouter(config, adapter, store)

    resolution = router.resolve(
        [direct_message],
        plan_actions(config, [direct_message]),
    )

    assert resolution.actions[0].move_to == "leadership"
    assert resolution.actions[0].matches[-1].rule_id == "thread-priority-promotion"
    assert resolution.promotions == []
    assert adapter.requested_ids == [[41]]


def test_higher_priority_reply_promotes_known_thread_members(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="promoted-thread",
        folder_key="internal_general",
        members=[(41, 110, "internal_general", False)],
    )
    adapter = ThreadStateAdapter(
        [
            ThreadMessageState(
                41,
                "exchange-41",
                110,
                "Internal General",
                "promoted-thread",
            )
        ]
    )
    direct_message.sender_address = "leader@corp.example"
    direct_message.thread_guid = "promoted-thread"

    resolution = ThreadRouter(config, adapter, store).resolve(
        [direct_message],
        plan_actions(config, [direct_message]),
    )

    assert resolution.actions[0].move_to == "leadership"
    assert [(item.outlook_id, item.destination_key) for item in resolution.promotions] == [
        (41, "leadership")
    ]


def test_manual_whole_thread_move_updates_canonical_destination(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="manual-thread",
        folder_key="internal_general",
        members=[
            (41, 110, "internal_general", False),
            (42, 110, "internal_general", False),
        ],
    )
    adapter = ThreadStateAdapter(
        [
            ThreadMessageState(41, "exchange-41", 106, "Leadership", "manual-thread"),
            ThreadMessageState(42, "exchange-42", 106, "Leadership", "manual-thread"),
        ]
    )
    incoming = message_copy(
        direct_message,
        outlook_id=43,
        exchange_id="exchange-43",
        thread_guid="manual-thread",
    )
    router = ThreadRouter(config, adapter, store)

    resolution = router.resolve([incoming], plan_actions(config, [incoming]))
    context = store.thread_contexts(router.scope, ["manual-thread"])["manual-thread"]

    assert context["folder_key"] == "leadership"
    assert {member["folder_key"] for member in context["members"]} == {"leadership"}
    assert resolution.actions[0].move_to == "leadership"
    assert resolution.promotions == []


def test_manual_single_message_move_becomes_detached_exception(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="split-thread",
        folder_key="internal_general",
        members=[
            (41, 110, "internal_general", False),
            (42, 110, "internal_general", False),
        ],
    )
    adapter = ThreadStateAdapter(
        [
            ThreadMessageState(41, "exchange-41", 106, "Leadership", "split-thread"),
            ThreadMessageState(
                42,
                "exchange-42",
                110,
                "Internal General",
                "split-thread",
            ),
        ]
    )
    incoming = message_copy(
        direct_message,
        outlook_id=43,
        exchange_id="exchange-43",
        thread_guid="split-thread",
    )
    router = ThreadRouter(config, adapter, store)

    resolution = router.resolve([incoming], plan_actions(config, [incoming]))
    context = store.thread_contexts(router.scope, ["split-thread"])["split-thread"]
    moved_member = next(member for member in context["members"] if member["outlook_id"] == 41)

    assert context["folder_key"] == "internal_general"
    assert moved_member["folder_key"] == "leadership"
    assert moved_member["detached"]
    assert resolution.actions[0].move_to == "internal_general"
    assert resolution.promotions == []


def test_missing_known_message_is_detached_without_failing(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="missing-thread",
        folder_key="internal_general",
        members=[(41, 110, "internal_general", False)],
    )
    direct_message.thread_guid = "missing-thread"
    router = ThreadRouter(config, ThreadStateAdapter(), store)

    resolution = router.resolve(
        [direct_message],
        plan_actions(config, [direct_message]),
    )
    member = store.thread_contexts(router.scope, ["missing-thread"])["missing-thread"][
        "members"
    ][0]

    assert member["detached"]
    assert resolution.actions[0].move_to == "internal_general"
    assert resolution.promotions == []


def test_message_manually_moved_outside_managed_tree_is_not_pulled_back(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="archived-thread",
        folder_key="internal_general",
        members=[(41, 110, "internal_general", False)],
    )
    adapter = ThreadStateAdapter(
        [ThreadMessageState(41, "exchange-41", 999, "Archive", "archived-thread")]
    )
    direct_message.thread_guid = "archived-thread"
    router = ThreadRouter(config, adapter, store)

    resolution = router.resolve(
        [direct_message],
        plan_actions(config, [direct_message]),
    )
    member = store.thread_contexts(router.scope, ["archived-thread"])[
        "archived-thread"
    ]["members"][0]

    assert member["folder_id"] == 999
    assert member["folder_key"] is None
    assert member["detached"]
    assert resolution.promotions == []


def test_outlook_id_resolving_to_another_thread_is_detached(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="original-thread",
        folder_key="internal_general",
        members=[(41, 110, "internal_general", False)],
    )
    adapter = ThreadStateAdapter(
        [ThreadMessageState(41, "different-exchange-id", 110, "Internal General", "other-thread")]
    )
    direct_message.thread_guid = "original-thread"
    router = ThreadRouter(config, adapter, store)

    router.resolve([direct_message], plan_actions(config, [direct_message]))
    member = store.thread_contexts(router.scope, ["original-thread"])[
        "original-thread"
    ]["members"][0]

    assert member["detached"]


def test_changed_exchange_id_in_same_thread_refreshes_without_detaching(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="stable-thread",
        folder_key="internal_general",
        members=[(41, 110, "internal_general", False)],
    )
    adapter = ThreadStateAdapter(
        [
            ThreadMessageState(
                41,
                "exchange-after-folder-move",
                110,
                "Internal General",
                "stable-thread",
            )
        ]
    )
    direct_message.thread_guid = "stable-thread"
    router = ThreadRouter(config, adapter, store)

    router.resolve([direct_message], plan_actions(config, [direct_message]))
    member = store.thread_contexts(router.scope, ["stable-thread"])["stable-thread"][
        "members"
    ][0]

    assert member["message_id"] == "exchange-after-folder-move"
    assert not member["detached"]


def test_higher_priority_reply_repromotes_manually_moved_thread_with_new_exchange_ids(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="revisit-thread",
        folder_key="leadership",
        members=[
            (41, 106, "leadership", False),
            (42, 106, "leadership", False),
            (43, 106, "leadership", False),
        ],
    )
    adapter = ThreadStateAdapter(
        [
            ThreadMessageState(
                outlook_id,
                f"moved-exchange-{outlook_id}",
                110,
                "Internal General",
                "revisit-thread",
            )
            for outlook_id in (41, 42, 43)
        ]
    )
    incoming = message_copy(
        direct_message,
        outlook_id=44,
        exchange_id="exchange-44",
        sender_address="leader@corp.example",
        thread_guid="revisit-thread",
    )

    resolution = ThreadRouter(config, adapter, store).resolve(
        [incoming],
        plan_actions(config, [incoming]),
    )
    context = store.thread_contexts("inbox:101", ["revisit-thread"])[
        "revisit-thread"
    ]

    assert resolution.actions[0].move_to == "leadership"
    assert {
        (promotion.outlook_id, promotion.destination_key)
        for promotion in resolution.promotions
    } == {
        (41, "leadership"),
        (42, "leadership"),
        (43, "leadership"),
    }
    assert {member["message_id"] for member in context["members"]} == {
        "moved-exchange-41",
        "moved-exchange-42",
        "moved-exchange-43",
    }


def test_safety_and_flagged_messages_do_not_inherit_thread_destination(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="protected-thread",
        folder_key="leadership",
        members=[(41, 106, "leadership", False)],
    )
    adapter = ThreadStateAdapter(
        [ThreadMessageState(41, "exchange-41", 106, "Leadership", "protected-thread")]
    )
    unsafe = message_copy(
        direct_message,
        outlook_id=42,
        exchange_id="exchange-42",
        sender_address="unknown@external.example",
        thread_guid="protected-thread",
    )
    flagged = message_copy(
        direct_message,
        outlook_id=43,
        exchange_id="exchange-43",
        flag_status=FlagStatus.FLAGGED,
        thread_guid="protected-thread",
    )
    junk = message_copy(
        direct_message,
        outlook_id=44,
        exchange_id="exchange-44",
        sender_address="sender@unwanted.example",
        thread_guid="protected-thread",
    )
    router = ThreadRouter(config, adapter, store)

    resolution = router.resolve(
        [unsafe, flagged, junk],
        plan_actions(config, [unsafe, flagged, junk]),
    )

    assert resolution.actions[0].move_to == "unclassified_external"
    assert resolution.actions[1].move_to is None
    assert resolution.actions[1].keep_in_inbox
    assert resolution.actions[2].move_to == "junk_external"
    assert resolution.canonical_routes == {"protected-thread": "leadership"}
    assert resolution.promotions == []


class ApplyingThreadAdapter(ThreadStateAdapter):
    def __init__(
        self,
        messages: list[MailMessage],
        folder_names: dict[int, str],
        *,
        failed_ids: set[int] | None = None,
    ) -> None:
        super().__init__()
        self.messages = {message.outlook_id: message for message in messages}
        self.folder_names = folder_names
        self.failed_ids = failed_ids or set()
        self.applied: list[dict] = []

    def get_messages(self, outlook_ids, body_limit=0):
        return [self.messages[value] for value in outlook_ids if value in self.messages]

    def get_message(self, outlook_id, body_limit=0):
        return self.messages[outlook_id]

    def thread_states_by_ids(self, outlook_ids):
        ids = list(outlook_ids)
        self.requested_ids.append(ids)
        return [
            ThreadMessageState(
                message.outlook_id,
                message.exchange_id,
                message.folder_id,
                message.folder_name,
                message.thread_guid,
            )
            for value in ids
            if (message := self.messages.get(value)) is not None
        ]

    def apply_mail_states(self, updates):
        results = []
        for update in updates:
            if update["outlook_id"] in self.failed_ids:
                results.append(
                    {
                        "outlook_id": update["outlook_id"],
                        "status": "failed",
                        "error": "simulated failure",
                    }
                )
                continue
            self.apply_mail_state(
                update["outlook_id"],
                categories=update["categories"],
                flag_status=update["flag_status"],
                target_folder_id=update["target_folder_id"],
            )
            results.append(
                {"outlook_id": update["outlook_id"], "status": "applied", "error": ""}
            )
        return results

    def apply_mail_state(
        self,
        outlook_id,
        *,
        categories,
        flag_status,
        target_folder_id=None,
    ):
        message = self.messages[outlook_id]
        self.applied.append(
            {
                "outlook_id": outlook_id,
                "target_folder_id": target_folder_id,
            }
        )
        message.categories = list(categories)
        message.flag_status = flag_status
        if target_folder_id is not None:
            message.folder_id = target_folder_id
            message.folder_name = self.folder_names[target_folder_id]


def test_executor_moves_new_reply_and_promotes_historical_member(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    historical = message_copy(
        direct_message,
        outlook_id=41,
        exchange_id="exchange-41",
        folder_id=110,
        folder_name="Internal General",
        thread_guid="execution-thread",
    )
    incoming = message_copy(
        direct_message,
        outlook_id=42,
        exchange_id="exchange-42",
        sender_address="leader@corp.example",
        thread_guid="execution-thread",
    )
    folder_names = {folder.id: folder.name for folder in config.mail.folders.values()}
    adapter = ApplyingThreadAdapter([historical, incoming], folder_names)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="execution-thread",
        folder_key="internal_general",
        members=[(41, 110, "internal_general", False)],
    )
    plan = MailTriagePlanner(config).create_plan([incoming], incoming.folder_id)

    result = ActionExecutor(config, adapter, store).apply_plan(plan)

    assert result.status == "completed"
    assert result.applied == 2
    assert result.promoted == 1
    assert result.thread_routed == 0
    assert incoming.folder_id == 106
    assert historical.folder_id == 106
    context = store.thread_contexts("inbox:101", ["execution-thread"])[
        "execution-thread"
    ]
    assert context["folder_key"] == "leadership"
    assert {member["folder_key"] for member in context["members"]} == {"leadership"}

    undo = ActionExecutor(config, adapter, store).undo_run(result.run_id)

    assert undo.status == "undone"
    assert incoming.folder_id == 101
    assert historical.folder_id == 110
    restored = store.thread_contexts("inbox:101", ["execution-thread"])[
        "execution-thread"
    ]
    assert restored["folder_key"] == "internal_general"


def test_executor_reports_successful_current_messages_routed_by_thread(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    leadership = message_copy(
        direct_message,
        outlook_id=41,
        exchange_id="exchange-41",
        sender_address="leader@corp.example",
        thread_guid="current-thread",
    )
    general = message_copy(
        direct_message,
        outlook_id=42,
        exchange_id="exchange-42",
        thread_guid="current-thread",
    )
    folder_names = {folder.id: folder.name for folder in config.mail.folders.values()}
    adapter = ApplyingThreadAdapter([leadership, general], folder_names)
    store = StateStore(tmp_path / "state.sqlite")
    plan = MailTriagePlanner(config).create_plan(
        [leadership, general],
        leadership.folder_id,
    )

    result = ActionExecutor(config, adapter, store).apply_plan(plan)

    assert result.status == "completed"
    assert result.applied == 2
    assert result.thread_routed == 1
    assert result.promoted == 0
    assert leadership.folder_id == 106
    assert general.folder_id == 106


def test_executor_does_not_promote_flagged_historical_member(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    historical = message_copy(
        direct_message,
        outlook_id=41,
        exchange_id="exchange-41",
        folder_id=110,
        folder_name="Internal General",
        flag_status=FlagStatus.FLAGGED,
        thread_guid="flagged-history",
    )
    incoming = message_copy(
        direct_message,
        outlook_id=42,
        exchange_id="exchange-42",
        sender_address="leader@corp.example",
        thread_guid="flagged-history",
    )
    folder_names = {folder.id: folder.name for folder in config.mail.folders.values()}
    adapter = ApplyingThreadAdapter([historical, incoming], folder_names)
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="flagged-history",
        folder_key="internal_general",
        members=[(41, 110, "internal_general", False)],
    )
    plan = MailTriagePlanner(config).create_plan([incoming], incoming.folder_id)

    result = ActionExecutor(config, adapter, store).apply_plan(plan)

    assert result.status == "completed"
    assert result.applied == 1
    assert result.promoted == 0
    assert incoming.folder_id == 106
    assert historical.folder_id == 110


def test_partial_promotion_failure_keeps_member_retryable(
    app_config, direct_message, tmp_path
) -> None:
    config = threaded_config(app_config)
    historical = message_copy(
        direct_message,
        outlook_id=41,
        exchange_id="exchange-41",
        folder_id=110,
        folder_name="Internal General",
        thread_guid="partial-thread",
    )
    incoming = message_copy(
        direct_message,
        outlook_id=42,
        exchange_id="exchange-42",
        sender_address="leader@corp.example",
        thread_guid="partial-thread",
    )
    folder_names = {folder.id: folder.name for folder in config.mail.folders.values()}
    adapter = ApplyingThreadAdapter(
        [historical, incoming],
        folder_names,
        failed_ids={41},
    )
    store = StateStore(tmp_path / "state.sqlite")
    seed_thread(
        store,
        thread_guid="partial-thread",
        folder_key="internal_general",
        members=[(41, 110, "internal_general", False)],
    )
    plan = MailTriagePlanner(config).create_plan([incoming], incoming.folder_id)

    result = ActionExecutor(config, adapter, store).apply_plan(plan)
    context = store.thread_contexts("inbox:101", ["partial-thread"])["partial-thread"]
    historical_member = next(
        member for member in context["members"] if member["outlook_id"] == 41
    )

    assert result.status == "partial"
    assert result.applied == 1
    assert result.promoted == 0
    assert context["folder_key"] == "leadership"
    assert historical_member["folder_key"] == "internal_general"
    assert not historical_member["detached"]
