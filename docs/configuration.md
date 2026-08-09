# Configuration reference

Outlook Organizer loads three YAML files from
`~/.config/outlook-organizer/` by default:

- `mail-definitions.yaml` describes people, domains, and distribution lists.
- `mail-rules.yaml` describes Outlook folders and mail-routing behavior.
- `calendar.yaml` describes calendar discovery and focus-time preferences.

All three files are required. Their schemas are strict: unknown keys cause
validation to fail instead of being silently ignored.

The repository's `config/` directory is a sanitized, valid sample—not the
runtime configuration. Create a private copy with owner-only permissions:

```bash
install -d -m 700 ~/.config/outlook-organizer
install -m 600 config/*.yaml ~/.config/outlook-organizer/
```

Replace every example value in that private copy. In particular, the sample
folder IDs are placeholders and must be replaced with IDs discovered from your
Outlook profile before applying triage.

Run validation after every change:

```bash
uv run outlook-organizer config validate
```

To use another private configuration directory, set
`OUTLOOK_ORGANIZER_CONFIG` to a directory containing all three files.

Changing any configuration value changes the configuration fingerprint. A
previously saved preview cannot be applied after that change; create a fresh
triage preview instead.

## Mail definitions

`mail-definitions.yaml` contains frequently edited identity and sender
classification data. Edit the private copy under
`~/.config/outlook-organizer/`; the repository file is only a sample.

### Top-level options

| Option | Type | Required | Description |
| --- | --- | --- | --- |
| `version` | integer | yes | Must be `2`. |
| `identity` | mapping | yes | The mailbox owner's name and email addresses. |
| `internal_domains` | list of domains | yes | Domains classified as internal. |
| `groups` | mapping of lists | no | Named groups of individual sender addresses. Defaults to `{}`. |
| `safe_external` | mapping | no | Trusted external domains and addresses. Defaults to empty lists. |
| `junk_external` | mapping | no | Junk domains, addresses, and keywords. Defaults to empty lists. |
| `distribution_list_groups` | mapping of lists | no | Named distribution-list address groups. Defaults to `{}`. |

### `identity`

```yaml
identity:
  name: Example User
  addresses:
    - example.user@example.com
    - example.alias@example.com
```

| Option | Type | Required | Description |
| --- | --- | --- | --- |
| `name` | string | yes | Display name of the mailbox owner. |
| `addresses` | list of strings | yes | Addresses used to determine whether mail is sent directly to the owner. |

Identity addresses are trimmed, lowercased, deduplicated, and sorted.

### Domains and exact addresses

```yaml
internal_domains:
  - example.com

safe_external:
  domains:
    - trusted-partner.example
  addresses:
    - trusted.sender@public.example

junk_external:
  domains:
    - unwanted.example
  addresses:
    - marketing@otherwise-valid.example
  keywords:
    - newsletter
```

Domain entries must be bare DNS names such as `example.com`, not email
addresses or URL patterns. Matching is boundary-aware and includes subdomains:
`mail.example.com` matches `example.com`, while `notexample.com` does not.
Internationalized domains are normalized to IDNA form.

Exact address entries must contain `@`. Addresses, domains, and keywords are
trimmed, normalized case-insensitively, deduplicated, and sorted.

Sender classification uses this precedence:

1. `internal_domains`
2. `junk_external.addresses`
3. `junk_external.domains`
4. `safe_external.addresses`
5. `safe_external.domains`
6. `junk_external.keywords`
7. otherwise `unclassified_external`

Junk keywords are substring-matched against the subject and sender address
only. They are evaluated only after internal, explicit junk, and safe-external
classification. Message bodies are not searched.

A domain cannot appear in more than one of `internal_domains`,
`safe_external.domains`, and `junk_external.domains`. An exact address cannot
be both safe and junk.

### People groups

`groups` maps a name used by `sender_group` rules to exact sender addresses:

```yaml
groups:
  leadership:
    - leader@example.com
  my_team:
    - teammate@example.com
```

An address may belong to only one people group. Group names are free-form, but
every group referenced by a mail rule must exist here.

### Distribution-list groups

```yaml
distribution_list_groups:
  company_announcements:
    - announcements@example.com
```

An address may belong to only one distribution-list group. Configured list
addresses are recognized when they appear as the sender or as a visible To/CC
recipient. Outlook public and private groups in To/CC are also recognized as
distribution lists, even when they are not assigned to a configured group.

Defining an address here identifies it and assigns it to a named group; it does
not by itself choose a destination. Mail rules decide whether to match the named
group (`distribution_list_group`), mere group presence (`distribution_list`), or
the more conservative delivery-intent signal (`distribution_delivery`).

## Mail rules

`mail-rules.yaml` contains stable Outlook folder references and ordered rules.
Edit the private copy under `~/.config/outlook-organizer/`.

### Top-level options

| Option | Type | Required | Description |
| --- | --- | --- | --- |
| `version` | integer | yes | Must be `2`. |
| `folder_scan_limit` | integer | no | Highest Outlook folder ID inspected during discovery. Default `1000`; allowed range `10`–`100000`. |
| `threading` | mapping | no | Optional prospective conversation-aware filing. Disabled by default. |
| `folders` | mapping | yes | Named Outlook folder definitions. |
| `annotations` | list | no | Non-routing rules; every match applies. Defaults to `[]`. |
| `routes` | list | yes | Routing rules; first match wins. |
| `default` | mapping | yes | Behavior when no route matches. |

### Folder definitions

```yaml
folders:
  inbox:
    name: Inbox
    aliases:
      - Gelen Kutusu
    id: 184

  organized_primary:
    name: aOrganized
    id: 813

  organized_secondary:
    name: bOrganized
    id: 823

  my_team:
    name: My Team
    id: 817
    parent: organized_primary
```

Each folder key is a stable configuration identifier used by routes.

| Option | Type | Required | Description |
| --- | --- | --- | --- |
| `name` | string | yes | Outlook folder name. |
| `id` | positive integer | yes | Scriptable Outlook folder ID. |
| `aliases` | list of strings | no | Alternative names accepted during folder discovery. Defaults to `[]`. |
| `parent` | folder key or `null` | no | Expected parent folder used to validate the hierarchy. |

The keys `inbox`, `organized_primary`, and `organized_secondary` are required.
Folder IDs must be unique. Parent keys must exist, and parent relationships
cannot contain self-references or cycles.

Use the following command to discover current Outlook IDs:

```bash
uv run outlook-organizer mail folders
```

Folder IDs can change when an Outlook profile is recreated or a folder is
deleted and recreated.

### Conversation-aware threading

Threading is opt-in and disabled by default:

```yaml
threading:
  enabled: false
```

Set `enabled: true` to keep SQL-known Outlook conversations together. Each
message still runs through the ordinary annotation and route rules. For an
eligible conversation, the highest-priority destination is the route appearing
earliest in `routes`. A later higher-priority reply promotes both that reply and
previously indexed managed messages. Threading changes folders only; categories
remain based on each individual message.

The following remain per-message exceptions and are never pulled into an
ordinary conversation destination:

- messages retained by `keep_in_inbox`, including flagged mail;
- `junk_external` and `unclassified_external` routes;
- messages manually moved outside configured route folders;
- individual messages manually moved away from the rest of a known thread.

The index is deliberately prospective. Enabling threading does not scan or
backfill existing Outlook folders. A missing thread GUID in SQLite is
authoritative and is treated as a new conversation; a full Outlook enumeration
is never attempted. The thread becomes known only after a confirmed triage run
successfully processes one of its messages.

For SQL-known conversations, the organizer validates only the recorded Outlook
message IDs. If every known member was manually moved to the same configured
route folder, that folder becomes the new conversation destination. A single
manually moved member becomes a detached exception and is not moved back.
Messages no longer found in Outlook, Outlook IDs that now resolve to another
thread, and messages moved outside the configured tree are also detached
safely. Exchange IDs are refreshed rather than treated as immutable because
Outlook changes them when messages move between folders.

#### Junk and unclassified messages in known threads

A conversation GUID establishes association, not sender trust. Consequently,
thread affinity never overrides a `junk_external` or `unclassified_external`
classification:

- the junk or unclassified message follows its own configured safety route;
- it does not inherit the conversation's ordinary destination;
- it does not change the conversation's canonical destination; and
- previously filed ordinary members are not moved because of that message.

For example, a junk-classified reply in a conversation routed to CXO moves only
to `Junk External`; the existing CXO messages remain in CXO, and the next
ordinary reply still inherits CXO. This is deliberately conservative because
the explicit junk address/domain or keyword evidence is stronger than thread
association. If that classification is a false positive, correct
`junk_external` in `mail-definitions.yaml` before applying the dry run. If it
was already applied, undo the run or move the message manually after correcting
the definition. The organizer does not silently move suspected junk into a
trusted folder.

Inspect the prospective index without reading Outlook folders:

```bash
uv run outlook-organizer mail threads status
```

Disabling threading stops all thread lookups and promotions but leaves the
prospective index intact in case it is enabled again later.

The triage summary's main metrics line shows two threading effects when they
are nonzero: current Inbox messages whose destination was changed by thread
affinity, and earlier filed thread members promoted to a higher-priority
destination. Confirmed-run counts include only successful Outlook updates. In
the individual message table, `· Threading` is appended to the destination only
when threading changed that message's ordinary rule destination. Confirmed runs
also list each successfully promoted earlier message in a separate section with
its original folder, destination, sender, and subject. Dry runs show only the
promotion count, avoiding an additional full-message read from Outlook.

### Match conditions

The `when` mapping on an annotation or route accepts the following predicates:

| Predicate | Allowed value | Matches when |
| --- | --- | --- |
| `flagged` | boolean | The Outlook message is currently flagged. |
| `recipient` | one value or a list | Recipient directness is one of the configured values. |
| `sender_group` | group name | The exact sender address belongs to that people group. |
| `sender_type` | classification value | The sender has that domain classification. |
| `distribution_list_group` | group name | A recognized sender or visible recipient belongs to that configured list group. |
| `distribution_list` | boolean | A configured or Outlook-detected distribution list is present. |
| `distribution_delivery` | boolean | The message appears to have been delivered through a distribution list. |

Multiple predicates in one `when` mapping are combined with AND:

```yaml
when:
  sender_type: internal
  distribution_list_group: company_announcements
```

`recipient` accepts:

| Value | Meaning |
| --- | --- |
| `only_me` | To contains exactly one configured identity address and CC is empty. |
| `direct_to_me` | The owner is the sole visible recipient via CC, or is the only To recipient with at least one CC recipient. |
| `multi_recipient` | Outlook exposes at least two visible To/CC recipients, except for the `direct_to_me` case above. |
| `not_to_me` | Outlook exposes one visible recipient that is not a configured identity address. This commonly occurs with distribution lists. |
| `unknown` | Outlook exposes no visible To/CC recipients. |

`sender_type` accepts `internal`, `junk_external`, `safe_external`,
`unclassified_external`, or `unknown`. `unclassified_external` means the
sender has a valid external address but matches neither the safe nor junk
configuration. `unknown` means the sender address could not be parsed as a
complete email address.

#### Distribution-list predicates

These fields are predicates used inside a rule's `when` mapping, not global
feature switches. Like other boolean predicates, they may be set to `false` to
match the inverse condition.

Distribution-list identity, presence, and delivery intent are deliberately
separate:

- `distribution_list: true` matches when a configured list is the sender or
  when a configured or Outlook-recognized public/private group appears anywhere
  in To or CC. It is useful for annotation and diagnostics, but is usually too
  broad for routing.
- `distribution_delivery: true` is the conservative routing predicate. It is
  true only when either (a) the sender is a configured list, or (b) a recognized
  group appears in To and none of the configured identity addresses appears
  directly in To. A group copied only on CC does not count as the delivery path.
- `distribution_list_group: <name>` matches a configured list in sender, To, or
  CC regardless of delivery intent. It acts as an explicit override when placed
  on a route and is appropriate for known announcement lists that should always
  have the same destination.

A "recognized group" is either an address in `distribution_list_groups` or a
To/CC recipient Outlook describes as a `public group address` or
`private group address`.

| Visible message shape | `distribution_list` | `distribution_delivery` | Named `distribution_list_group` |
| --- | --- | --- | --- |
| Configured list is the sender | `true` | `true` | Matches that list's group |
| Owner is in To; unconfigured Outlook group is only in CC | `true` | `false` | Does not match |
| Outlook group is in To; owner is not in To | `true` | `true` | Matches only if configured |
| Owner and Outlook group are both in To | `true` | `false` | Matches only if configured |
| Owner is in To; configured list is in CC | `true` | `false` | Matches that list's group |
| No recognized group is visible and sender is not a configured list | `false` | `false` | Does not match |

`distribution_delivery` does not attempt directory membership expansion. It
uses the sender and visible To/CC metadata exposed by Outlook. A list hidden by
BCC or returned by Outlook as an unresolved individual address cannot be
recognized unless its address is explicitly configured.

For a conservative internal distribution fallback, combine sender type with
delivery intent:

```yaml
when:
  sender_type: internal
  distribution_delivery: true
```

For a known list that should always win even when it is copied on CC, use the
named group predicate instead:

```yaml
when:
  sender_type: internal
  distribution_list_group: company_announcements
```

### Annotations

Every matching annotation applies. An annotation can add context or keep a
message in the Inbox without selecting its destination:

```yaml
annotations:
  - id: flagged-needs-action
    description: Keep flagged messages visible in the Inbox
    when:
      flagged: true
    section: Needs attention
    keep_in_inbox: true

  - id: sent-only-to-me
    description: Message is addressed only to me
    when:
      recipient: only_me
    add_category: "@Only Me"
```

| Option | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `id` | string | yes | — | Unique rule identifier used in reports and audit data. |
| `description` | string | no | `""` | Human-readable explanation. |
| `when` | mapping | yes | — | Match predicates. |
| `add_category` | string or `null` | no | `null` | Outlook category to add. Existing categories are preserved. |
| `section` | string or `null` | no | `null` | Preferred triage-report section. |
| `keep_in_inbox` | boolean | no | `false` | Suppress the move selected by a matching route. |

The first matching annotation with `section` set chooses the report section.
`keep_in_inbox` is cumulative: if any matching annotation enables it, the
message remains in the Inbox even when a route matches.

Categories are additive. Removing `add_category` from the configuration stops
future additions but does not remove that category from existing Outlook mail.

### Routes

Routes are evaluated from top to bottom, and the first matching route wins:

```yaml
routes:
  - id: route-my-team
    description: Mail from my team
    when:
      sender_group: my_team
    move_to: my_team
```

| Option | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `id` | string | yes | — | Unique rule identifier. IDs are shared with annotations and cannot repeat. |
| `description` | string | no | `""` | Human-readable explanation. |
| `when` | mapping | yes | — | Match predicates. |
| `move_to` | folder key | yes | — | Destination from the `folders` mapping. |
| `category` | string or `null` | no | `null` | Outlook category to add when the route matches. |

Route order is significant. Put specific relationship and distribution-list
rules before broad sender-type rules.

### Default routing

`default` applies only when no route matches:

```yaml
default:
  keep_in_inbox: true
  section: Others
```

| Option | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `keep_in_inbox` | boolean | no | `true` | Keep unmatched mail in the Inbox. |
| `category` | string or `null` | no | `null` | Category to add to unmatched mail. |
| `section` | string | no | `"Others"` | Triage-report section for unmatched mail. |

## Calendar

`calendar.yaml` controls calendar discovery and free-slot calculations. Edit
the private copy under `~/.config/outlook-organizer/`.

### Top-level options

| Option | Type | Required | Description |
| --- | --- | --- | --- |
| `version` | integer | yes | Must be `1`. |
| `timezone` | string | yes | Intended calendar timezone, normally an IANA name such as `Europe/Istanbul`. Currently informational; Outlook supplies local timestamps. |
| `calendar_names` | list of strings | yes | Case-insensitive Outlook calendar names accepted during discovery. If several match, the one with the most events is selected. |
| `maximum_calendar_id` | integer | no | Highest Outlook calendar ID inspected. Default `5000`; allowed range `10`–`100000`. |
| `working_hours` | mapping | yes | Lowercase weekday names mapped to start/end times. Missing days produce no free slots. |
| `preferences` | mapping | yes | Focus-time and meeting preferences. |
| `protected_relationships` | mapping | yes | Relationship groups intended to receive scheduling priority. |

Times should use ISO local-time strings such as `"09:00"` or `"13:30"`.

### Working hours

```yaml
working_hours:
  monday: ["09:00", "18:00"]
  tuesday: ["09:00", "18:00"]
```

`calendar free-slots` considers only the configured interval for the requested
weekday. Days omitted from the mapping return no slots.

### Preferences

```yaml
preferences:
  lunch_window: ["12:00", "13:30"]
  minimum_focus_block_minutes: 90
  meeting_buffer_minutes: 10
  maximum_meeting_hours_per_day: 5
  avoid_back_to_back_meetings: true
  preferred_focus_windows:
    - ["09:00", "11:00"]
```

| Option | Type | Required | Constraints | Current behavior |
| --- | --- | --- | --- | --- |
| `lunch_window` | two times | yes | Start and end | Always blocked during free-slot searches. |
| `minimum_focus_block_minutes` | integer | yes | `15`–`480` | Minimum returned free-slot length unless overridden on the command line. |
| `meeting_buffer_minutes` | integer | yes | `0`–`120` | Added before and after busy events during free-slot searches. |
| `maximum_meeting_hours_per_day` | number | yes | `0`–`24` | Accepted and validated; not currently used to flag workload. |
| `avoid_back_to_back_meetings` | boolean | no | Default `true` | Accepted and validated; workload analysis reports back-to-back meetings regardless of this value. |
| `preferred_focus_windows` | list of time pairs | no | Default `[]` | Accepted and validated; not currently used to rank or restrict free slots. |

### Protected relationships

```yaml
protected_relationships:
  high_priority:
    - leadership
    - my_team
```

`high_priority` is an optional list defaulting to `[]`. It is accepted and
validated but is not currently used by calendar analysis or free-slot
selection. Names are not currently cross-validated against mail people groups.

## Validation summary

Configuration validation catches:

- unknown YAML options;
- unsupported schema versions;
- malformed domain and address entries;
- conflicting sender classifications;
- duplicate membership across people or distribution-list groups;
- missing rule group references;
- duplicate rule IDs and folder IDs;
- missing route destinations; and
- invalid folder parent relationships.

Validation does not contact Outlook. Use `uv run outlook-organizer check` after
validation to confirm that configured folders and calendars are visible in the
current Outlook profile.
