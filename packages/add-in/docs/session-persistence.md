# Session persistence and workbook identity

> Status: implemented (2026-11-10). This document records the invariants the
> session-history feature depends on; changing any of them needs a matching
> change here and in the tests listed at the end.

## Why this exists

Two user-visible bugs drove this design:

1. **Session history showed only old sessions.** Sessions that were saved while
   the workbook identity was unavailable, or saved under an identity the current
   document no longer reports, were *hidden* by the resume overlay's
   workbook filter. The overlay also failed silently when the metadata index was
   missing.
2. **The latest session was lost after restarting Excel.** Sessions were only
   written after the *first completed assistant message*, there was no flush on
   teardown, and the restore path could point `workbook.latestSession` at a
   blank record that restore then skipped.

## Storage layout

| Key / store | Owner | Contents |
|---|---|---|
| `sessions` (IndexedDB store) | `SessionsStore` | full `SessionData` (messages, model, thinking level) |
| `sessions-metadata` (store) | `SessionsStore` | list metadata (`title`, `lastModified`, `messageCount`, `preview`) |
| `session.workbook.v1.<sessionId>` | `workbook/session-association.ts` | canonical workbook id the session belongs to |
| `workbook.latestSession.v1.<workbookId>` | same | newest **non-empty** session for that workbook |
| `session.pendingWorkbook.v1.<sessionId>` | same | marker for a session saved before identity resolved |
| `workbook.identity.v1.alias.<observedId>` | `workbook/identity.ts` | observed URL/path hash → canonical document id |
| `workbook.identity.v1.doc.<canonicalId>` | same | `{ name, firstSeen, lastSeen, observed[] }` |
| `workbook.tabLayout.v1.<observedId>` | `taskpane/tab-layout.ts` | open tabs + active tab |
| `ui.recentlyClosed.v1` | `taskpane/recently-closed-store.ts` | persisted recently-closed ring buffer |
| `storage.repair.v1`, `storage.persistence.v1`, `__pi_migrated_legacy_db_v1` | storage layer | one-time markers |

Database: `pi-for-office`, **version 3**. The upgrade adds any index declared by
the current store config that an existing object store is missing — the missing
`lastModified` index is what made listing throw on older databases.

## Identity ladder

Session association uses the **canonical** workbook id only:

1. **Document-scoped id** (`doc_uuid:<uuid>`) — a UUID stored *inside* the
   document, via Office `document.settings` or WPS `PluginStorage`. It survives
   rename/move, unsaved documents, guest access in Excel on the web, and
   desktop-vs-cloud path differences.
2. **Alias** — the observed hash of the host's document URL/path resolves to the
   document-scoped id. Aliases are written once per document and followed for at
   most one hop.
3. **Observed hash** — `url_sha256:`/`wps_path_sha256:` of the normalized URL or
   path, for documents no document id has ever been seen for.

Normalization (`host/workbook-context.ts`) unifies `file:///C:/…` with
`C:\…`/`c:\…`, decodes percent-escapes once, normalizes UNC paths, lowercases
only the authority/drive letter, and strips query/hash/trailing slashes.

**Scope:** the canonical id is used by *session association only*. Tab layout,
rules, recovery logs, and the Pi bridge keep the observed hash, because changing
their key space would strand existing data and their misses are visible (empty
tabs, no rules) rather than silent.

**Document lineage:** a "Save As" copy inherits the document-scoped id, so the
copy continues the original's session history. This is intentional; the overlay
still lets the user pick sessions explicitly.

## Write path (durability)

`setupSessionPersistence` guarantees:

- **Draft save** — the first user message of a never-written session is written
  immediately, so a long first turn cannot be lost by closing Excel mid-run.
- **Turn boundaries** — `turn_end` and `agent_end` save immediately; assistant
  and tool-result `message_end` events schedule a coalesced save (2 s).
- **Periodic** — a 15 s interval saves only when the transcript changed.
- **Teardown** — `pagehide` / `visibilitychange → hidden` flush every runtime,
  bounded to 150 ms.
- **Serialization** — every save goes through one promise chain, so a slow
  IndexedDB write can never put older state on top of newer state.
- **Commit, not request** — `IndexedDBStorageBackend` awaits
  `transaction.oncomplete`; a save that resolves is a save that committed.
- **Failures are visible** — `getLastSaveError()` feeds a once-per-pane toast
  (every notice is dismissable; see `docs/ui-layout.md`).
  and the diagnostics readout.

### Latest-session pointer rules

- Only a **non-empty** session may set `workbook.latestSession.v1.<id>`.
- A blank session may be written as a durable tab placeholder, but it never
  overwrites an existing record and never claims the pointer.
- Startup restore is **strict**: workbook-linked latest only, no global
  fallback. A blank start is not announced with a toast on every load — the
  pane's empty state carries a persistent **Open session history** action
  (`pi-sidebar.ts` → `onOpenResumePicker`). A toast is shown only when a saved
  tab layout was restored but produced no content at all, i.e. when something
  the user expected to be there is missing.

## Read path (history)

`SessionsStore.getAllMetadataSafe()` never throws and never trusts the index
alone. It unions:

- the `lastModified` index (fast path),
- a full `sessions` scan (missing index / rows without `lastModified`),
- metadata-only rows (`missingData: true`).

Ordering is by parsed timestamp, falling back to `createdAt` for legacy rows.

The resume overlay groups entries into **This workbook / Other workbooks /
Not linked** and renders all of them; "Other workbooks" auto-expands when the
current workbook has none. A storage error renders an inline error with a retry
instead of rejecting the fire-and-forget open call.

## Repair and migration

- **Legacy copy** (`pi-for-excel` → `pi-for-office`) is *insert-if-absent* and
  commits together with its marker in one transaction. A newer row, pointer, or
  tab layout can never be rewound by a retried copy.
- **Repair** (`storage/local/repair.ts`, marker `storage.repair.v1`) rebuilds
  missing metadata, backfills unusable timestamps, repoints stale links through
  the alias table, prunes the identity table, and deletes blank records **only
  when nothing references them**. Non-empty sessions are never deleted.
- **Export/import** (`storage/local/session-transfer.ts`, schema
  `pi-sessions/v1`) is the escape hatch for origin changes, eviction, and
  device moves. Import never overwrites an existing id.

## Diagnostics

The overlay footer's "Storage details" (and the console on demand) reports:
database name/version, record and metadata counts, link/pointer counts,
canonical vs observed identity, persistence grant and quota, markers, and the
last save error. Compare it with devtools when triaging a "missing session".

## Tests that pin these invariants

- `tests/workbook-identity.test.ts` — normalization matrix, alias resolution,
  document-id precedence, dry runs.
- `tests/sessions-store.test.ts` — index fallback, unindexed rows, metadata-only
  rows, ordering.
- `tests/session-persistence.test.ts` — save triggers, blank/pointer guards,
  no id remapping, failure surfacing, serialization, pending association.
- `tests/session-repair.test.ts` — metadata rebuild, timestamp repair,
  referenced-blank safety, dry run, marker.
- `tests/session-transfer.test.ts` — export/import round trip, id collisions,
  invalid payloads.
- `tests/recently-closed-store.test.ts` — persisted ring buffer.
- `tests/resume-grouping.test.ts` — grouped (never hidden) listing.
