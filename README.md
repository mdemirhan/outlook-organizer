# Outlook Organizer

Local-first email organization and calendar analysis for Outlook for Mac with
on-premises Exchange. Outlook is accessed only through its AppleScript object
model; the project never reads or writes Outlook's internal database.

## Safety model

- `mail triage` is read-only by default.
- The same command applies its computed actions only with `--apply`.
- Every applied action records original folder, categories, and flag state.
- Completed runs can be undone.
- Unclassified external domains are routed for review. The current rules propose
  moving them to `Unclassified External`, but only through an explicitly applied
  triage; deletion is never exposed.
- Email bodies are not persisted by default.
- Private calendar event content is redacted from tool results.
- Sending, replying, deleting mail, and responding to invitations are not exposed.

## Setup

```bash
uv sync
install -d -m 700 ~/.config/outlook-organizer
install -m 600 config/*.yaml ~/.config/outlook-organizer/
uv run outlook-organizer config validate
uv run outlook-organizer check
```

macOS may ask for permission to let the invoking application control Microsoft
Outlook. Outlook must be installed, signed in, and synchronized.

## Configuration

Personal configuration lives outside the repository in
`~/.config/outlook-organizer/`. The tracked `config/` directory contains only
sanitized samples. Copy all three samples during setup, then replace the example
addresses, domains, calendar settings, folder names, and placeholder folder IDs
in your private copies:

- `~/.config/outlook-organizer/mail-definitions.yaml`: identity, address groups, domains, trusted
  senders, junk detection, and distribution lists.
- `~/.config/outlook-organizer/mail-rules.yaml`: Outlook folders, annotations, ordered routing
  rules, and fallback behavior.
- `~/.config/outlook-organizer/calendar.yaml`: calendar discovery, working hours, and focus-time
  preferences.

Set `OUTLOOK_ORGANIZER_CONFIG` to use a different directory. Do not put private
configuration back under `config/` or commit it elsewhere in the repository.

See the [configuration reference](docs/configuration.md) for every option,
allowed value, default, validation rule, and matching behavior.

If the Outlook profile is recreated or folders are deleted and recreated, run
`mail folders` and update the affected definitions before applying triage.

## Daily commands

```bash
# Idempotent root-folder setup for aOrganized and bOrganized
uv run outlook-organizer mail setup --confirm

# Dry-run daily review
uv run outlook-organizer check
uv run outlook-organizer mail threads status
uv run outlook-organizer mail triage --limit 50
uv run outlook-organizer calendar agenda --days-ahead 7
uv run outlook-organizer calendar workload --days-ahead 7
uv run outlook-organizer calendar free-slots 2026-07-29

# Compute the triage report and immediately apply all displayed actions
uv run outlook-organizer mail triage --limit 50 --apply

# Suppress the live phase/elapsed-time display when scripting
uv run outlook-organizer mail triage --limit 50 --no-progress

# Revert a completed run
uv run outlook-organizer history undo RUN_ID --confirm
```

The triage command uses a terminal-friendly Rich report with a compact summary,
routing tables, and a clearly marked preview or execution result. Optional
conversation-aware filing is controlled by `threading.enabled` in
`mail-rules.yaml`; see the configuration reference for its prospective SQL
index, priority-promotion behavior, and manual-move semantics. When threading
affects a run, the main metrics line reports current Inbox messages routed by
threading and earlier filed messages promoted to the thread's new route. An
individual destination receives a `· Threading` suffix when threading changed
that message's ordinary rule destination.

Outlook changes are not transactional. If a run fails after partially changing
a message, undo that run before running triage again:

```bash
uv run outlook-organizer history undo FAILED_RUN_ID --confirm
```

The undo path restores the recorded folder, categories, and flag state for both
completed and partially failed actions.

## MCP

The project contains a project-scoped Codex MCP configuration. After dependency
installation, restart Codex so it discovers the server.

The server can preview or immediately apply mail triage, read/search mail,
undo runs, read the calendar, analyze workload, and find focus slots. Email
writes still require `confirm=true`, even after client approval.

Example requests in Codex:

- “Triage my newest 50 messages. Do not apply anything.”
- “Triage my newest 50 messages and apply the proposed changes.”
- “Show calendar conflicts this week and find a 60-minute focus block tomorrow.”

## Local audit database

SQLite records confirmed Outlook actions, run status, and each message's
folder/category/flag state so `history undo` can safely restore it. When
threading is enabled, it also stores prospective conversation destinations and
known message IDs. Dry-run triage reports are not stored, although targeted
manual-move reconciliation may refresh already-known thread metadata. Message
bodies are never written to this database.
