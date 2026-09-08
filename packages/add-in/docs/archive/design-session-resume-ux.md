# Design: Session Resume + Tab Recovery UX

> **Related issues:** [#23](https://github.com/tmustier/pi-for-excel/issues/23), [#31](https://github.com/tmustier/pi-for-excel/issues/31), [#12](https://github.com/tmustier/pi-for-excel/issues/12)
> **Status:** Draft
> **Last updated:** 2026-02-11

## Problem

Users can currently lose momentum when they:

1. close a tab by mistake,
2. want to resume a prior session in a **new** tab,
3. open the add-in and need to quickly continue previous work.

Today, `/resume` primarily targets replacing the active runtime, and tab close has no strong recovery affordance.

---

## Goals

1. **Make tab close safe** (easy undo/reopen).
2. **Make resume discoverable** from UI chrome, not only slash commands.
3. **Default to non-destructive resume** (`open in new tab` > `replace current`).
4. Preserve workbook-scoped safety defaults (cross-workbook resumes stay explicit).
5. Keep UX fast in a narrow (~350px) sidebar.

## Non-goals (V1)

- Full Git-like session branching/history graph.
- Cross-device sync semantics beyond existing storage.
- Automatic merge of two sessions.

---

## UX principles

1. **Tabs are temporary, sessions are durable.**
2. **Recovery should be one click.**
3. **Resume should not destroy active work by default.**
4. **Workbook context should always be visible in resume decisions.**

---

## User stories

### A) Accidental close recovery

> “I closed the wrong tab. Bring it back immediately.”

Desired UX:
- Closing a tab shows toast: `Closed Agent 2 — Undo` (8–10s).
- Keyboard fallback: `Cmd+Shift+T` (or `Ctrl+Shift+T` on Windows) reopens the most recently closed tab.
- Header menu includes **Recently closed** list.

### B) Resume into new tab

> “I want old context in a new tab while keeping my current tab.”

Desired UX:
- `+` button has split behavior:
  - Click: New blank tab.
  - Menu: `New tab`, `Resume recent…`, plus up to 3 recent sessions.
- Resume picker default action: **Open in new tab**.
- Secondary action: `Replace current tab` (explicit).

### C) Resume at startup

> “When I reopen the add-in, help me continue quickly.”

Desired UX:
- Startup behavior setting:
  - `Continue last (default)`
  - `Ask every time`
  - `Start fresh`
- If set to `Ask every time`, show compact chooser on boot:
  - Continue last
  - Resume another session
  - Start fresh

---

## Information architecture (entry points)

Users should be able to resume from multiple predictable places:

1. **Tab strip `+` menu** (primary)
2. **Header utilities menu (⋯)** → `Resume…`, `Recently closed`
3. **Slash command `/resume`**
4. **Keyboard `Cmd/Ctrl+Shift+T`** for immediate reopen
5. **Startup chooser** (optional by user setting)

No single entry point should be mandatory.

---

## Interaction spec

## 1) Tab close + undo/reopen

### Close behavior
- On close request:
  1. force-save the session snapshot,
  2. close runtime,
  3. push a `RecentlyClosedItem` entry,
  4. show undo toast.

### Undo toast
- Text: `Closed <tab title> — Undo`
- TTL: 8–10 seconds
- Action: reopens exact session in a new runtime and restores model/thinking/messages from persisted session.

### Recently closed stack
- Keep latest N=10 entries.
- Exposed in header utilities menu.
- Entries show: title, last modified time, workbook match indicator.
- If persisted session is missing/corrupt: show toast `Couldn’t reopen session` and remove stale entry.

### Streaming close guard
- If tab is actively streaming: confirm dialog
  - `Stop and close` / `Cancel`
- If lock is held (`holding_lock`): close action disabled until write completes (avoid ambiguous state).

---

## 2) Resume picker behavior

### Default target
- Default selection target = `Open in new tab`.
- Secondary target = `Replace current tab`.

### Workbook filtering
- Keep current behavior: prioritize current workbook sessions.
- Keep `Show sessions from all workbooks` toggle.
- Keep cross-workbook confirmation on final action.

### Session row content (V1)
- Title
- Message count + relative modified time
- Optional workbook badge when `show all` is enabled (e.g. `This workbook` / `Other workbook`)

### Session row actions
- Primary click: open according to selected target (`new tab` by default).
- Optional row overflow (later): `Open in new tab`, `Replace current`, `Delete`.

---

## 3) Startup resume policy

### New setting key
`ui.startup.resume_behavior.v1`:
- `continue_last`
- `ask`
- `start_fresh`

### Behavior
- `continue_last` (default): current logic + non-blocking banner
  - banner: `Resumed: <title> · Start fresh`
- `ask`: lightweight chooser before first runtime is finalized
- `start_fresh`: create blank runtime, no auto-restore

### Optional future setting
`ui.resume.default_target.v1`:
- `new_tab` (default)
- `replace_current`

---

## Data model

```ts
export interface RecentlyClosedItem {
  sessionId: string;
  title: string;
  closedAt: string; // ISO
  workbookId: string | null;
}
```

Storage strategy:
- In-memory stack for immediate undo.
- Persist a small ring buffer in `SettingsStore` for session continuity across pane refresh (optional in V1, recommended in V1.1).

---

## Runtime semantics

1. **Close tab must force-persist session** even when no assistant response exists yet.
   - This avoids losing drafts/early turns.
2. Reopen should instantiate a **new runtime** and `applyLoadedSession(sessionData)`.
3. Reopen never mutates an existing runtime unless user explicitly picks `replace current`.

---

## Technical changes (file-by-file)

### `src/taskpane/sessions.ts`

- Extend persistence API to support forced saves for close flow.

Proposed shape:
```ts
saveSession(opts?: { force?: boolean }): Promise<void>
```

Behavior:
- Existing guard (`firstAssistantSeen`) remains for autosave path.
- `force: true` allows saving draft sessions on explicit close/recoverable actions.

### `src/taskpane/session-runtime-manager.ts`

- Add close hooks/output for recovery bookkeeping.
- Return enough metadata for `RecentlyClosedItem` capture.

Possible API addition:
```ts
closeRuntime(runtimeId: string, opts?: { reason?: "user_close" | "replace" }): SessionRuntime | null
```

### `src/taskpane/init.ts`

- Add `RecentlyClosedManager` wiring.
- On close:
  - force-save,
  - register recently closed,
  - show undo toast with action callback.
- Add `reopenLastClosed()` handler for keyboard + menu.
- Startup behavior switch based on new setting.

### `src/commands/builtins/overlays.ts`

- Update `showResumeDialog()` to support explicit target mode:

```ts
onOpenInNewTab(sessionData)
onReplaceCurrent(sessionData)
defaultTarget: "new_tab" | "replace_current"
```

- Keep workbook filtering and cross-workbook confirmation.

### `src/commands/builtins/session.ts`

- Keep `/resume`, but semantics become “open picker (default new tab)”.
- Optional new command: `/resume-here` for explicit replace behavior.
- Optional new command: `/reopen` (reopen last closed).

### `src/ui/pi-sidebar.ts`

- Upgrade `+` control to split button/menu.
- Add callbacks:
  - `onOpenResumePicker?: (target: "new_tab" | "replace_current") => void`
  - `onReopenLastClosed?: () => void`

### `src/taskpane/keyboard-shortcuts.ts`

- Add `Cmd/Ctrl+Shift+T` → reopen last closed session/tab.
- Keep busy checks and avoid interfering with text input conventions.

### `src/ui/theme/components.css`

- Add styles for:
  - tab-strip split-button menu,
  - recently-closed menu section,
  - optional small workbook badge in resume rows.

---

## Acceptance criteria

1. Closing a tab always offers undo; undo restores that session in a new tab.
2. `Cmd/Ctrl+Shift+T` reopens the most recently closed session tab.
3. Resume picker defaults to opening in new tab.
4. User can still replace current tab intentionally.
5. Startup behavior follows setting (`continue_last` / `ask` / `start_fresh`).
6. Workbook-safe behavior remains (cross-workbook resume confirmation).
7. No regression in existing `/resume`, `/new`, workbook-scoped latest restore.

---

## Rollout plan

### Phase 1 (high impact, low risk)
- Force-save on close.
- Undo toast on close.
- Reopen last closed (keyboard + command).
- `/resume` default target = new tab.

### Phase 2
- `+` split menu with `Resume recent…` + top recents.
- Header utilities menu entries for Resume/Recently closed.

### Phase 3
- Startup behavior setting + chooser (`ask` mode).
- Optional persisted recently-closed ring.

---

## Open questions

1. Should closing a streaming tab always require confirmation, or silently abort + close?
2. Should we allow closing lock-holding tabs with deferred close semantics after write commit?
3. Is persisted recently-closed history desirable, or should it be memory-only for privacy simplicity?
4. Do we expose delete-from-history in V1 or defer to a session manager surface?
