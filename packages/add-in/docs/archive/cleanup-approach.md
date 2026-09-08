# Code Cleanup Approach (Pi for Excel)

Reusable approach for future cleanup passes. Goal: keep the codebase small, legible, and aligned with Pi’s existing patterns.

## Scope

Three review lenses:
1. **Redundant code** — duplicate, dead, or overly verbose paths
2. **Reinvented logic** — places where Pi already provides better/cleaner infrastructure
3. **Taste / maintainability** — separation of concerns, file size, magic numbers, over‑engineering

## Process

### 1) Baseline scan
- **Find unused symbols**: `npx tsc --noEmit --noUnusedLocals --noUnusedParameters`
- **List large files**: `find src -type f -maxdepth 3 -print0 | xargs -0 wc -l | sort -nr | head`
- **Search for duplicates**: `rg -n "TODO|HACK|FIXME|duplicate|copy" src`

### 2) Redundancy pass
- Remove unused imports/vars
- Consolidate duplicated helpers (e.g., toast helpers, model picker open)
- Replace repeated inline logic with small, named helpers (no speculative abstractions)
- Flag dead files or stubs for removal

### 3) Pi alignment pass
- Prefer Pi’s APIs (sessions, auth, model selection, prompt compaction)
- Avoid parallel implementations when Pi already solves it
- Document every intentional deviation with rationale

### 4) Taste pass
- Enforce **small, coherent modules** (≤ ~400–500 LOC)
- Remove magic numbers by naming intent (e.g., `SELECTION_PADDING_ROWS = 5`)
- Keep UI wiring close to UI code; tool logic close to tool code

### 5) Verification
- `npm run typecheck`
- `npm run lint`
- `npm run test:models` (requires Node 22+ for `--experimental-strip-types`)
- `npx vite build`
- Manual spot test in Excel add‑in

## Current hotspots / refactor candidates (keep a running list)

- `src/taskpane.ts` — **DONE (v0.2.0-pre).** Split into `taskpane/init.ts`, `bootstrap.ts`, `sessions.ts`, `keyboard-shortcuts.ts`, `context-injection.ts`, `default-model.ts`, `queue-display.ts`, `welcome-login.ts`. Builtins split by domain. Zero `any`, zero lint warnings.
- `src/ui/tool-renderers.ts` (444 LOC) — largest file. Could split render functions by tool category if it grows further.
- `src/tools/` — consolidated from 14 → 10 tools. Each file is self-contained and ≤371 LOC.

## Output expectations

- **Small, surgical commits** per cleanup area
- **Notes on intentional deviations** from Pi patterns
- **Short summary** of removed redundancy and refactors

## What to clarify before starting a cleanup pass

- Are deletions allowed this pass? (e.g., stubs, dead files)
- Is behavior change acceptable, or only refactors?
- Priority areas (UI vs tools vs context vs auth)
