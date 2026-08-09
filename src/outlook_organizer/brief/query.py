from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

from outlook_organizer.brief.config import BriefContext, BriefScopeConfig

PERIODS = {
    "last_hour",
    "last_24_hours",
    "today",
    "yesterday",
    "since_yesterday",
    "since_last_workday",
}
READ_STATES = {"unread", "read", "all"}
GROUP_MODES = {"none", "conversation"}


class BriefQueryResolver:
    def __init__(self, context: BriefContext) -> None:
        self.context = context
        self.config = context.config
        self.timezone = ZoneInfo(context.config.timezone)

    def resolve(
        self,
        *,
        profile: str | None,
        folder_keys: list[str] | None,
        additional_folder_keys: list[str] | None,
        exclude_folder_keys: list[str] | None,
        include_subfolders: bool | None,
        period: str | None,
        since: str | None,
        until: str | None,
        read_state: str | None,
        include_attention_debt: bool | None,
        attention_debt_days: int | None,
        group_by: str | None,
        max_messages: int | None,
    ) -> tuple[dict[str, Any], str | None, dict[str, Any]]:
        query = self.config.defaults.model_dump(mode="python")
        query.update(exclude_folder_keys=[], since=None, until=None)
        profile_key = self._profile_key(profile)
        if profile_key:
            query.update(
                self.config.profiles[profile_key].model_dump(
                    mode="python",
                    exclude={"name", "description", "aliases"},
                    exclude_none=True,
                )
            )
        overrides: dict[str, Any] = {}
        if folder_keys is not None:
            if not folder_keys:
                raise ValueError("folder_keys cannot be empty")
            recursive = include_subfolders if include_subfolders is not None else False
            query["scopes"] = [
                BriefScopeConfig(folder=value, recursive=recursive) for value in folder_keys
            ]
            overrides["folder_keys"] = folder_keys
        elif include_subfolders is not None:
            query["scopes"] = [
                BriefScopeConfig(folder=scope.folder, recursive=include_subfolders)
                for scope in query["scopes"]
            ]
            overrides["include_subfolders"] = include_subfolders
        if additional_folder_keys:
            recursive = include_subfolders if include_subfolders is not None else False
            query["scopes"] = [
                *query["scopes"],
                *[
                    BriefScopeConfig(folder=value, recursive=recursive)
                    for value in additional_folder_keys
                ],
            ]
            overrides["additional_folder_keys"] = additional_folder_keys
        if exclude_folder_keys is not None:
            query["exclude_folder_keys"] = list(exclude_folder_keys)
            overrides["exclude_folder_keys"] = exclude_folder_keys
        for key, value in {
            "read_state": read_state,
            "include_attention_debt": include_attention_debt,
            "attention_debt_days": attention_debt_days,
            "group_by": group_by,
            "max_messages": max_messages,
        }.items():
            if value is not None:
                query[key] = value
                overrides[key] = value
        if period is not None:
            query.update(period=period, since=None, until=None)
            overrides["period"] = period
        if since is not None or until is not None:
            if since is None:
                raise ValueError("until requires an explicit since value")
            query.update(period=None, since=since, until=until)
            overrides["since"] = since
            if until is not None:
                overrides["until"] = until
        if query["period"] is not None and query["period"] not in PERIODS:
            raise ValueError(f"Unsupported brief period: {query['period']}")
        if query["read_state"] not in READ_STATES:
            raise ValueError(f"Unsupported read state: {query['read_state']}")
        if query["group_by"] not in GROUP_MODES:
            raise ValueError(f"Unsupported group mode: {query['group_by']}")
        if not 1 <= int(query["attention_debt_days"]) <= 90:
            raise ValueError("attention_debt_days must be between 1 and 90")
        if not 1 <= int(query["max_messages"]) <= 500:
            raise ValueError("max_messages must be between 1 and 500")
        query["scopes"] = [
            scope if isinstance(scope, BriefScopeConfig) else BriefScopeConfig.model_validate(scope)
            for scope in query["scopes"]
        ]
        return query, profile_key, overrides

    def window(
        self,
        *,
        now: datetime,
        period: str | None,
        since: str | None,
        until: str | None,
    ) -> tuple[datetime, datetime]:
        if period == "last_hour":
            start, end = now - timedelta(hours=1), now
        elif period == "last_24_hours":
            start, end = now - timedelta(hours=24), now
        elif period == "today":
            start, end = now.replace(hour=0, minute=0, second=0, microsecond=0), now
        elif period in {"yesterday", "since_yesterday"}:
            today = now.replace(hour=0, minute=0, second=0, microsecond=0)
            start = today - timedelta(days=1)
            end = today if period == "yesterday" else now
        elif period == "since_last_workday":
            start = now.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1)
            while start.weekday() >= 5:
                start -= timedelta(days=1)
            end = now
        elif period is None and since:
            start = self._timestamp(since)
            end = self._timestamp(until) if until else now
        else:
            raise ValueError("A supported period or explicit since value is required")
        if start >= end:
            raise ValueError("brief start must be earlier than brief end")
        return start, end

    def folders(self, scopes: list[BriefScopeConfig], *, excluded: set[str]) -> list[str]:
        configured = self.context.mail.folders.folders
        unknown = sorted(({scope.folder for scope in scopes} | excluded) - set(configured))
        if unknown:
            raise ValueError(f"Unknown brief folders: {unknown}")
        selected: list[str] = []
        for scope in scopes:
            if scope.folder not in selected:
                selected.append(scope.folder)
            if scope.recursive:
                for key in configured:
                    parent = configured[key].parent
                    while parent:
                        if parent == scope.folder:
                            if key not in selected:
                                selected.append(key)
                            break
                        parent = configured[parent].parent
        result = [key for key in selected if key not in excluded]
        if not result:
            raise ValueError("Brief folder selection is empty after exclusions")
        return result

    def _profile_key(self, requested: str | None) -> str | None:
        value = requested or self.config.default_profile
        if not value:
            return None
        normalized = value.strip().casefold()
        for key, profile in self.config.profiles.items():
            if normalized in {
                candidate.strip().casefold() for candidate in [key, profile.name, *profile.aliases]
            }:
                return key
        available = ", ".join(sorted(self.config.profiles)) or "(none)"
        raise ValueError(f"Unknown brief profile {value!r}. Available profiles: {available}")

    def _timestamp(self, value: str) -> datetime:
        parsed = datetime.fromisoformat(value)
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=self.timezone)
        return parsed.astimezone(self.timezone)
