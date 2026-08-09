# Configuration reference

Outlook Organizer loads five strict YAML files from the configured directory.
Unknown keys and invalid cross-file references fail validation.

```bash
uv run outlook-organizer config validate
```

## `mail-definitions.yaml`

Contains shared mail identity and classification data.

```yaml
version: 2
identity:
  name: Example User
  addresses: [example.user@corp.example]
internal_domains: [corp.example]
groups:
  leadership: [leader@corp.example]
safe_external:
  domains: [trusted.example]
  addresses: []
junk_external:
  domains: [unwanted.example]
  addresses: []
  keywords: [newsletter]
distribution_list_groups:
  company_announcements: [announcements@corp.example]
```

Domains are normalized and match their subdomains on DNS boundaries. Exact
addresses are case-insensitive. A domain cannot be both internal, safe, and
junk. An address may belong to only one people group and one distribution-list
group.

Sender classification precedence is internal, explicit junk, safe external,
junk keyword, then unclassified external. Invalid addresses are `unknown`.

## `mail-folders.yaml`

The shared folder catalog is independent from triage rules and is used by both
triage and brief.

```yaml
version: 1
scan_limit: 1000
folders:
  inbox:
    name: Inbox
    aliases: [Gelen Kutusu]
    id: 101
  organized_primary:
    name: aOrganized
    id: 102
  organized_secondary:
    name: bOrganized
    id: 103
  leadership:
    name: Leadership
    id: 106
    parent: organized_primary
```

`inbox`, `organized_primary`, and `organized_secondary` are required. IDs must
be unique. Parent keys must exist and cannot form cycles. Discover current IDs
with `outlook-organizer mail folders`.

## `triage.yaml`

Triage is deterministic and CLI-only.

```yaml
version: 1
threading:
  enabled: true
annotations:
  - id: flagged-needs-action
    when: {flagged: true}
    keep_in_inbox: true
    section: Needs attention
routes:
  - id: route-leadership
    when: {sender_group: leadership}
    move_to: leadership
default:
  keep_in_inbox: true
  section: Others
```

Every matching annotation applies. Routes are evaluated in order and the first
match wins. `keep_in_inbox` suppresses a route move. Existing categories are
preserved and configured categories are added.

Supported `when` predicates:

- `flagged`: boolean
- `recipient`: `only_me`, `direct_to_me`, `multi_recipient`, `not_to_me`, or `unknown`
- `sender_group`: a group from `mail-definitions.yaml`
- `sender_type`: `internal`, `junk_external`, `safe_external`,
  `unclassified_external`, or `unknown`
- `distribution_list_group`: configured distribution group
- `distribution_list`: boolean presence signal
- `distribution_delivery`: conservative delivery-path signal

Multiple predicates are combined with AND. Annotation and route IDs must be
unique. Every route destination must exist in `mail-folders.yaml`.

### Conversation-aware triage

With `threading.enabled: true`, triage uses Outlook conversation GUIDs and a
prospective SQLite index:

- preview may read an existing indexed destination;
- a missing database or missing conversation is treated as a cold index;
- preview never creates, reconciles, or writes the database;
- only successfully applied Outlook changes update routes and members;
- junk, unclassified external, and keep-in-Inbox decisions never inherit an
  ordinary conversation destination;
- route priority follows the order in `routes`.

Conversation identity is an association, not evidence that a sender is safe.

## `brief.yaml`

Brief is a stateless reading feature designed for MCP/LLM clients.

```yaml
version: 1
timezone: Europe/Istanbul
default_profile: morning
defaults:
  scopes:
    - folder: organized_primary
      recursive: true
  period: today
  read_state: unread
  include_attention_debt: false
  attention_debt_days: 7
  attention_debt_folders: [leadership]
  group_by: none
  max_messages: 75
  folder_policies:
    leadership:
      content: detailed
      snippet_chars: 2000
profiles:
  morning:
    name: Morning brief
    aliases: [morning mail]
    period: since_yesterday
    read_state: unread
    include_attention_debt: true
```

Resolution order is built-in safety limits, defaults, selected profile, then
explicit MCP/CLI arguments.

Periods: `last_hour`, `last_24_hours`, `today`, `yesterday`,
`since_yesterday`, and `since_last_workday`. Callers may instead supply an ISO
`since` and optional `until`; offset-free timestamps use `brief.yaml` timezone.

Read state is `unread`, `read`, or `all`. Scopes reference folder keys and may
include descendants recursively. Explicit request folders replace profile
scopes; additional and excluded folders modify them.

Content policies:

- `detailed` and `concise` return capped, untrusted snippets;
- `rollup` and `metadata_only` return no body and require `snippet_chars: 0`;
- junk and unclassified-external bodies are always suppressed by code.

Attention debt evaluates older flagged, `@Action`, unreplied-direct, and unread
priority-folder messages. `group_by: conversation` groups only messages in the
current packet and never uses the persisted triage thread index.

## `calendar.yaml`

```yaml
version: 1
timezone: Europe/Istanbul
calendar_names: [Calendar]
maximum_calendar_id: 1000
working_hours:
  monday: ["09:00", "18:00"]
preferences:
  lunch_window: ["12:00", "13:30"]
  minimum_focus_block_minutes: 90
  meeting_buffer_minutes: 10
  maximum_meeting_hours_per_day: 5
  avoid_back_to_back_meetings: true
  preferred_focus_windows: []
protected_relationships:
  high_priority: [leadership]
```

Calendar configuration is loaded only by calendar services. Brief carries its
own reporting timezone and does not depend on this file.
