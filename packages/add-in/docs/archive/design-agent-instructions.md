# Design: Agent Instructions (AGENTS.md equivalent)

> **Issue:** [#30](https://github.com/tmustier/pi-for-excel/issues/30)
> **Status:** Draft
> **Last updated:** 2026-02-10

## Overview

Give users a way to provide persistent instructions to the agent at multiple scopes. Analogous to `AGENTS.md` in Pi TUI — the agent just _sees_ them in its context, and can update them via a tool.

Three levels, two to implement now:

| Level | Scope | Storage | Travels with file? | Agent can write? |
| ------- | ------- | --------- | ------------------- | ----------------- |
| **User** | All workbooks | IndexedDB | No (local to machine) | Yes — auto-save ("memory") |
| **Workbook** | This .xlsx file | IndexedDB (default; keyed by workbook identity) or `workbook.settings` (opt-in) | Optional (opt-in) | Yes — propose & confirm |
| **Sheet/cell** | Specific ranges | Excel comments (`@pi:` prefix) | Yes | Future |

---

## Level 1: User Instructions

### What it is

Personal preferences that apply across all workbooks. Examples:

- "Always use EUR for currencies"
- "Dates as dd-mmm-yyyy"
- "I work in FP&A — use finance terminology"
- "Check circular references after formula writes"

### Storage

IndexedDB via the existing `SettingsStore`, keyed as `"user.instructions"`. Private to the machine, not tied to any workbook.

### Agent visibility

Injected into the system prompt every turn. The agent always sees the current text — no tool call needed to read.

### Agent writes (proactive "memory" pattern)

The agent auto-saves when the user expresses a preference. No confirmation step — user instructions are private.

**Explicit triggers:** "Always do X", "Remember that I prefer...", "From now on..."

**Implicit triggers:** Agent notices repeated corrections (e.g., user keeps fixing date formats) and saves the pattern.

**UX flow:**

```
User:  "Always check for circular references after writing formulas"

Agent: [calls instructions tool: append, level=user]
       "Noted — I've saved that to your instructions."

Toast: 📝 Memory updated: "Check circular refs after formula writes"  [Undo]
```

The toast has an **Undo** link (expires after ~10s). User can always review/edit via `/instructions`.

For inferred preferences (not explicitly requested), the agent should be more explicit about what it captured:

> "I've noticed you keep correcting me on date formats — I've added 'Dates as dd-mmm-yyyy' to your instructions so I remember."

### Char limit

Soft limit: ~2,000 characters (~500 tokens). Show character count in the editor. No hard enforcement — just a warning if the user goes over.

---

## Level 2: Workbook Instructions

### What it is

Workbook-specific guidance that describes the model's structure, conventions, and constraints. Examples:

- "DCF model for Acme Corp, FY2025 projections"
- "Revenue assumptions in Inputs!B5:B15"
- "Column B is GAAP revenue only, not adjusted"
- "Don't modify the Summary sheet without asking"

### Storage (draft): local by default; optional workbook-attached

**Default (recommended): local-only.** Store workbook instructions in IndexedDB, keyed by workbook identity. This avoids accidentally sharing sensitive information when workbooks are emailed or forwarded.

**Optional (opt-in): stored inside the `.xlsx`.** Persist via Excel’s `SettingCollection` API: `workbook.settings`, key `"pi.instructions"`, value: string. Only enable after explicit user action and keep a visible warning in the UI.

**Why `workbook.settings` (when opting in):**

- **Persists inside the `.xlsx` file** — travels with the file, works across machines
- **Per-add-in** — invisible to other add-ins, doesn't clutter the workbook
- **Not user-visible** in Excel UI — no phantom sheets, no Name Manager entries
- **Save As = natural fork** — copies automatically, then diverges independently
- **ExcelApi 1.4** — supported since Excel 2016, very broad compatibility

**Save As behavior (workbook-attached mode):**

1. User has "DCF_CompanyA.xlsx" with workbook instructions
2. File → Save As → "DCF_CompanyB.xlsx"
3. Instructions copy to the new file automatically
4. User edits instructions in the new file → original unaffected
5. Opens original → still has original instructions

Zero edge-case handling needed. This is native fork behavior.

### Agent visibility

Injected into the system prompt every turn, after user instructions.

### Agent writes (propose & confirm)

The file may be shared — the agent must be transparent and cautious.

**The agent always:**

1. Shows the exact text it wants to save
2. Reminds the user where workbook instructions are stored (local vs inside the `.xlsx`). If inside the workbook, warn they travel with the file.
3. Waits for explicit confirmation before saving

**UX flow:**

```
Agent: "I'd like to note this in the workbook instructions:

        > Column B revenue is GAAP (not adjusted).

        These will be saved for this workbook. If stored inside the .xlsx, they'll be visible to anyone you share the file with. OK to add?"

User:  "yes"

Agent: [calls instructions tool: append, level=workbook]
       "Added to workbook instructions."

Toast: 📋 Workbook instruction added
```

**Content guidelines (in system prompt):** Write neutral, professional notes about the workbook's structure and conventions. No personal information, no internal jargon, nothing the user might not want shared.

### Char limit

Soft limit: ~4,000 characters (~1,000 tokens). Same approach — show count, warn, don't enforce.

---

## Level 3: Sheet-level / In-cell Notes (future, zero-cost convention)

Use existing Excel comments with a `@pi:` prefix:

```
Cell B5 comment: "@pi: This is projected revenue, not actual. Source: management guidance."
```

**Why this works:**

- Zero new UI, storage, or API — uses native Excel comments
- Spatial context — notes are attached to the cells they describe
- Travels with the file
- Agent already reads comments in detailed mode (`read_range mode=detailed`)
- Users know how to add comments

**What we'd add:**

1. System prompt line: "Comments prefixed with `@pi:` are instructions from the user. Follow them."
2. Blueprint enhancement: mention sheets that have `@pi:` comments so the agent knows to read detailed when relevant.

**Not implementing now** — just establishing the convention so it's consistent with the rest of the design.

---

## System Prompt Integration

Instructions appear in the system prompt as a natural section. The agent sees them every turn without a tool call:

```markdown
## Your Instructions

You have persistent instructions at two levels. They appear below.
You can update them with the `instructions` tool:
- **User instructions** are private (stored locally). Update freely
  when the user expresses a preference or you notice a repeated
  pattern.
- **Workbook instructions** apply to the current workbook. By default
  they are stored locally; optionally they can be stored inside the .xlsx
  (travels with the file). Always show the user the exact text and get
  confirmation before saving. Write neutral, professional notes — no personal info.

If instructions at the two levels appear to conflict, ask the user
to clarify rather than assuming one overrides the other.

### User
<content or "(No user instructions set.)">

### Workbook
<content or "(No workbook instructions set.)">
```

Preamble cost: ~100 tokens. Pays for itself by preventing confused behavior.

### Conflict resolution

No rigid precedence rules. The agent asks the user and updates the appropriate level:

- **Scope-specific** ("just for this project") → update workbook instructions
- **User instruction was too broad** ("EUR except for USD-denominated models") → refine user instructions
- **One-off exception** ("just use GBP for this cell") → do it, update nothing

---

## The `instructions` Tool

### Schema

```typescript
{
  name: "instructions",
  description: "Update the agent's persistent instructions.",
  parameters: {
    action: "append" | "replace",
    level: "user" | "workbook",
    content: string
  }
}
```

- **`append`** — adds a line/bullet (most common — "memory" style additions)
- **`replace`** — full rewrite (rare — user asks to reorganize or rewrite)

No `read` action — the agent already sees instructions in the system prompt.

### Return value

Returns the updated full text for that level, so the agent sees the new state immediately without waiting for the next turn's system prompt refresh.

### Trust enforcement

The tool itself doesn't enforce the propose-and-confirm pattern for workbook-level — the system prompt does. The agent is instructed to always ask before writing to workbook. This keeps the tool simple. If we find models skip confirmation in practice, we can add a UI confirmation gate later.

### Storage mapping

| Level | Read from | Write to |
|-------|-----------|----------|
| `user` | IndexedDB `SettingsStore` → `"user.instructions"` | Same |
| `workbook` | IndexedDB (keyed by workbook identity) _(default)_; optional `workbook.settings.getItem("pi.instructions")` | IndexedDB _(default)_; optional `workbook.settings.add("pi.instructions", ...)` |

---

## UI: `/instructions` Editor

### Entry point

`/instructions` slash command → opens a full overlay (not crammed into `/settings`).

### Layout

Two tabs: **My Instructions** and **Workbook**.

```
┌─────────────────────────────────────────┐
│  Instructions                        ✕  │
│                                         │
│  [My Instructions]  [Workbook]          │
│  ─────────────────────────────────────  │
│                                         │
│  ┌─────────────────────────────────────┐│
│  │ • Always use EUR for currencies     ││
│  │ • Dates as dd-mmm-yyyy             ││
│  │ • Check circular refs after writes  ││
│  │ • I work in FP&A at Acme Corp      ││
│  │                                     ││
│  │                                     ││
│  └─────────────────────────────────────┘│
│                                         │
│  1,247 / 2,000 chars                   │
│                                         │
│  These instructions are private to      │
│  your machine and apply to all          │
│  workbooks.                             │
│                                         │
│              [Save]  [Cancel]           │
└─────────────────────────────────────────┘
```

### Tab-specific footer text

| Tab | Footer |
|-----|--------|
| **My Instructions** | "Private to your machine. The assistant may update these automatically when you express preferences." |
| **Workbook** | "Saved for this workbook. By default it's private to your machine; you can opt in to store inside the .xlsx (travels with the file). The assistant will always ask before adding." |

### Placeholder text (greyed, disappears on focus)

**My Instructions tab:**

```
Your preferences and habits, e.g.:
• Always use EUR for currencies
• Format dates as dd-mmm-yyyy
• I work in FP&A — use finance terminology
```

**Workbook tab:**

```
Notes about this workbook's structure, e.g.:
• DCF model for Acme Corp, FY2025
• Revenue assumptions in Inputs!B5:B15
• Don't modify the Summary sheet
```

### Discoverability

- **Status bar indicator:** When either level has instructions, show a small indicator (📋 or dot). Click → opens `/instructions`.
- **Working indicator hints:** Include "Tip: Use /instructions to teach the agent your preferences" in the rotating hints.
- **Agent-driven:** The agent can suggest `/instructions` when relevant.

---

## Token Budget

| Component | Estimated tokens | When |
| ----------- | ----------------- | ------ |
| Preamble (instructions section header + guidance) | ~100 | Always |
| User instructions (up to soft limit) | ~500 max | Always |
| Workbook instructions (up to soft limit) | ~1,000 max | Always |
| @pi: comments | 0 (read on demand) | When agent reads relevant ranges |
| **Total overhead** | **~1,600 worst case** | — |

Acceptable for all supported models.

---

## Implementation Sequence

1. **Storage layer** — read/write helpers for user + workbook levels (IndexedDB by default; optional `workbook.settings` when user opts in)
2. **System prompt injection** — extend `buildSystemPrompt()` to accept and render instructions
3. **`instructions` tool** — append/replace with toast notifications
4. **`/instructions` slash command + editor overlay** — two-tab UI
5. **Status bar indicator** — show when instructions are active
6. **System prompt guidance** — memory behavior, workbook caution, conflict resolution
7. **@pi: convention** — system prompt mention + blueprint scan (lightweight)

---

## Related Issues

- [#1](https://github.com/tmustier/pi-for-excel/issues/1) — Conventions storage/exposure (workbook instructions may partially supersede hardcoded conventions)
- [#14](https://github.com/tmustier/pi-for-excel/issues/14) — Agent interface design (instructions are part of context strategy)
- [#20](https://github.com/tmustier/pi-for-excel/issues/20) — Auto-compaction (instructions add to per-turn token cost)
