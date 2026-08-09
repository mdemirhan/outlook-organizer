from __future__ import annotations

from dataclasses import replace

from outlook_organizer.database import SqliteDatabase
from outlook_organizer.triage.classifier import TriageClassifier
from outlook_organizer.triage.config import TriageContext
from outlook_organizer.triage.thread_index import (
    SqliteThreadIndexRepository,
    ThreadAffinityResolver,
)


def threaded(context: TriageContext) -> TriageContext:
    config = context.config.model_copy(
        update={"threading": context.config.threading.model_copy(update={"enabled": True})}
    )
    return replace(context, config=config)


def test_missing_thread_index_is_read_without_creation(
    tmp_path, triage_context, direct_message
) -> None:
    context = threaded(triage_context)
    direct_message.thread_guid = "thread-1"
    database = SqliteDatabase(tmp_path / "state.sqlite")
    resolver = ThreadAffinityResolver(context, SqliteThreadIndexRepository(database=database))
    result = resolver.resolve(
        [direct_message], [TriageClassifier(context).classify(direct_message)]
    )
    assert result.decisions[0].move_to == "internal_general"
    assert not database.path.exists()


def test_preview_reads_existing_thread_route_without_writing(
    tmp_path, triage_context, direct_message
) -> None:
    context = threaded(triage_context)
    direct_message.thread_guid = "thread-1"
    database = SqliteDatabase(tmp_path / "state.sqlite")
    repository = SqliteThreadIndexRepository(database=database)
    repository.update(
        scope="inbox:101",
        routes={"thread-1": "leadership"},
        members=[],
    )
    before = database.path.stat().st_mtime_ns
    resolver = ThreadAffinityResolver(context, repository)

    result = resolver.resolve(
        [direct_message], [TriageClassifier(context).classify(direct_message)]
    )

    assert result.decisions[0].move_to == "leadership"
    assert result.decisions[0].matches[-1].rule_id == "thread-affinity"
    assert database.path.stat().st_mtime_ns == before


def test_forgetting_last_member_removes_orphaned_route(
    tmp_path, triage_context, direct_message
) -> None:
    database = SqliteDatabase(tmp_path / "state.sqlite")
    repository = SqliteThreadIndexRepository(database=database)
    repository.update(
        scope="inbox:101",
        routes={"thread-1": "leadership"},
        members=[
            {
                "thread_guid": "thread-1",
                "outlook_id": direct_message.outlook_id,
                "message_id": direct_message.stable_id,
                "folder_id": 111,
                "folder_key": "leadership",
            }
        ],
    )

    repository.forget_members(scope="inbox:101", outlook_ids={direct_message.outlook_id})

    assert repository.contexts("inbox:101", ["thread-1"]) == {}
