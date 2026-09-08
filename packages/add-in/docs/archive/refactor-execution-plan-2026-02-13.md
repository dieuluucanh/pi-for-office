# Refactor Execution Plan (Phase 1, Parallelized)

**Date:** 2026-02-13  
**Scope:** Execute the two highest-ROI refactors from `docs/archive/deep-refactor-review-2026-02-13.md`:
1. Recovery subsystem modularization
2. Shared mutation-tool pipeline extraction

**Status:** Completed (INT1 + CLOSE1 merged). This document is kept as a historical execution record.

**Post-close note:** early PR validation steps in this plan reference `tests/workbook-recovery-log.test.ts`; CLOSE1 replaced that file with:
- `tests/recovery-log-persistence.test.ts`
- `tests/recovery-log-restore.test.ts`
- `tests/recovery-log-format.test.ts`
- `tests/recovery-log-structure.test.ts`

Use `npm run test:recovery` for a reproducible recovery-suite command on current Node versions (including Node v25).

---

## Goals

- Reduce complexity hotspots without changing behavior.
- Keep all public tool contracts and persisted storage keys stable.
- Make recovery and mutation paths easier to extend/test safely.

### Success criteria

- `src/workbook/recovery-states.ts` and `src/workbook/recovery-log.ts` become thin facades.
- Mutation tools share common audit/checkpoint/result-note helpers.
- Existing tests remain green; targeted new tests added where extraction introduces risk.

---

## Non-goals (for this phase)

- No UX redesign.
- No new tool behavior.
- No migration of persisted snapshot schema (`workbook.recovery-snapshots.v1`).
- No performance tuning/chunking work (that is Phase 4 from the review).

---

## Guardrails

- **Strict compatibility:** preserve these APIs/signatures:
  - `getWorkbookRecoveryLog()`
  - `captureFormatCellsState`, `applyFormatCellsState`
  - `captureModifyStructureState`, `applyModifyStructureState`
  - `captureConditionalFormatState`, `applyConditionalFormatState`
  - `captureCommentThreadState`, `applyCommentThreadState`
- **No tool copy changes unless unavoidable** (to avoid snapshot/golden test churn).
- **Small PRs only**: each PR should be reviewable independently and pass gates.
- **Refactor-first commits**: extraction/moves should be isolated from behavior edits.

---

## Parallelization model

## Serial gates (short)

These are intentionally sequential because they unblock all other workstreams:

- **Gate A (PR0):** baseline/guard tests
- **Gate B (PR1):** recovery package scaffolding + compatibility facades

After **Gate B**, run three parallel tracks.

---

## Dependency DAG

```text
PR0 -> PR1 -> {R1, L1, M1}

R1 -> R2 -> R3
L1 -> L2
M1 -> M2 -> M3

{R3, L2, M3} -> INT1 -> CLOSE1
```

Where:
- `R*` = Recovery-state decomposition track
- `L*` = Recovery-log decomposition track
- `M*` = Mutation-pipeline extraction track
- `INT1` = integration/convergence PR
- `CLOSE1` = tests split + docs closeout

---

## Track ownership and file boundaries

To reduce merge conflicts, each track should primarily touch separate file sets.

### Track R (Recovery state decomposition)
Primary files:
- `src/workbook/recovery-states.ts`
- `src/workbook/recovery/format-state.ts`
- `src/workbook/recovery/format-selection.ts`
- `src/workbook/recovery/structure-state.ts`
- `src/workbook/recovery/conditional-format-state.ts`
- `src/workbook/recovery/comment-state.ts`

### Track L (Recovery log decomposition)
Primary files:
- `src/workbook/recovery-log.ts`
- `src/workbook/recovery/log-codec.ts`
- `src/workbook/recovery/log-store.ts`
- `src/workbook/recovery/log-restore.ts`

### Track M (Mutation pipeline extraction)
Primary files:
- `src/tools/mutation/finalize.ts`
- `src/tools/mutation/result-note.ts`
- `src/tools/mutation/types.ts`
- `src/tools/write-cells.ts`
- `src/tools/fill-formula.ts`
- `src/tools/python-transform-range.ts`
- `src/tools/format-cells.ts`
- `src/tools/modify-structure.ts`
- `src/tools/comments.ts`
- `src/tools/conditional-format.ts`
- `src/tools/view-settings.ts`
- `src/tools/workbook-history.ts`

### Track T (Tests/docs closeout; starts near end)
Primary files:
- `tests/recovery-log-persistence.test.ts`
- `tests/recovery-log-restore.test.ts`
- `tests/recovery-log-format.test.ts`
- `tests/recovery-log-structure.test.ts`
- `docs/archive/codebase-simplification-plan.md`
- `src/tools/DECISIONS.md` (only if behavior detail changes)

---

## PR plan (parallel-ready)

## Gate A

### PR0 — Baseline + guard tests (prep)
**Purpose:** freeze expected behavior before extraction.

**Changes**
- Add baseline assertions where coverage is thin around recovery restore paths:
  - structure states (`sheet_absent/present`, rows, columns)
  - conditional format rule round-trips
  - comment thread restore round-trips
- Add short note in `docs/archive/codebase-simplification-plan.md` linking this plan.

**Validation**
- `npm run check`
- `npm run test:context`
- `npm run build`

---

## Gate B

### PR1 — Recovery package scaffolding + facades
**Purpose:** establish modular boundaries with near-zero behavior change.

**Changes**
- Create `src/workbook/recovery/` with foundational modules:
  - `types.ts`, `guards.ts`, `clone.ts`, `address.ts`
- Keep existing entrypoints as compatibility facades.

**Validation**
- `npm run check`
- `npm run test:recovery`
- `npm run build`

---

## Parallel track R (start after PR1)

### R1 — Extract format-state capture/apply
**Changes**
- Add/move to:
  - `src/workbook/recovery/format-state.ts`
  - `src/workbook/recovery/format-selection.ts`
- Keep exports stable via `recovery-states.ts` facade.

**Validation**
- `npm run check`
- `npm run test:recovery`

### R2 — Extract structure-state capture/apply
**Changes**
- Add/move to `src/workbook/recovery/structure-state.ts`
- Keep facade exports stable.

**Validation**
- `npm run check`
- `npm run test:recovery`

### R3 — Extract conditional-format + comment state
**Changes**
- Add/move to:
  - `src/workbook/recovery/conditional-format-state.ts`
  - `src/workbook/recovery/comment-state.ts`
- Reduce `recovery-states.ts` to orchestration/re-exports.

**Validation**
- `npm run check`
- `npm run test:recovery`
- `npm run build`

---

## Parallel track L (start after PR1)

### L1 — Recovery log codec/store split
**Changes**
- Add:
  - `src/workbook/recovery/log-codec.ts`
  - `src/workbook/recovery/log-store.ts`
- Keep schema/key unchanged (`workbook.recovery-snapshots.v1`).

**Validation**
- `npm run check`
- `npm run test:recovery`

### L2 — Recovery restore strategy module
**Changes**
- Add `src/workbook/recovery/log-restore.ts`
- Keep `WorkbookRecoveryLog` and `getWorkbookRecoveryLog()` public API unchanged.
- Make `recovery-log.ts` thin composition root.

**Validation**
- `npm run check`
- `npm run test:recovery`
- `npm run build`

---

## Parallel track M (start after PR1)

### M1 — Shared mutation helper primitives
**Changes**
- Add:
  - `src/tools/mutation/finalize.ts`
  - `src/tools/mutation/result-note.ts`
  - `src/tools/mutation/types.ts`
- Introduce helper APIs without tool migrations yet.

**Validation**
- `npm run check`
- `node --test --experimental-strip-types tests/tool-result-shaping.test.ts tests/workbook-change-audit.test.ts`

### M2 — Migrate first mutation set
**Changes**
- Migrate:
  - `src/tools/write-cells.ts`
  - `src/tools/fill-formula.ts`
  - `src/tools/python-transform-range.ts`

**Validation**
- `npm run check`
- `node --test --experimental-strip-types tests/tool-result-shaping.test.ts tests/python-transform-range-tool.test.ts tests/workbook-change-audit.test.ts`

### M3 — Migrate remaining mutation set
**Changes**
- Migrate where applicable:
  - `src/tools/format-cells.ts`
  - `src/tools/modify-structure.ts`
  - `src/tools/comments.ts`
  - `src/tools/conditional-format.ts`
  - `src/tools/view-settings.ts`
  - `src/tools/workbook-history.ts`

**Validation**
- `npm run check`
- `npm run test:context`
- `npm run build`

---

## Integration and closeout (serial)

### INT1 — Convergence PR
**Purpose:** merge all parallel tracks cleanly and resolve cross-track conflicts.

**Changes**
- Resolve any overlap between `recovery-log` and mutation helper usage.
- Ensure import paths and facades are final.
- Keep behavior snapshots stable.

**Validation**
- `npm run check`
- `npm run test:context`
- `npm run build`

### CLOSE1 — Test decomposition + docs closeout
**Changes**
- Split `tests/workbook-recovery-log.test.ts` into focused suites:
  - `tests/recovery-log-persistence.test.ts`
  - `tests/recovery-log-restore.test.ts`
  - `tests/recovery-log-format.test.ts`
  - `tests/recovery-log-structure.test.ts`
- Update docs and remove stale references.

**Validation**
- `npm run check`
- `npm run test:context`
- `npm run build`

---

## Merge/branch strategy for parallel work

- Branch all tracks from **PR1 merge commit**.
- Use naming convention:
  - `refactor/recovery-r1-format-state`
  - `refactor/recovery-l1-log-codec-store`
  - `refactor/mutation-m2-core-tools`
- Rebase at least daily onto main to keep conflicts small.
- Merge order recommendation:
  1. `R1`, `L1`, `M1`
  2. `R2`, `L2`, `M2`
  3. `R3`, `M3`
  4. `INT1`
  5. `CLOSE1`

---

## CI/test matrix by track

To keep feedback fast, run focused suites per PR, then full gate in convergence.

- **R track PRs:** `npm run test:recovery`
- **L track PRs:** `npm run test:recovery`
- **M track PRs:**
  - `tests/tool-result-shaping.test.ts`
  - `tests/workbook-change-audit.test.ts`
  - plus tool-specific tests touched
- **INT1/CLOSE1:** full `npm run test:context`

Always run:
- `npm run check`
- `npm run build`

---

## Team sizing suggestion

- **3 engineers minimum** for true parallelism:
  - Engineer A: Track R
  - Engineer B: Track L
  - Engineer C: Track M
- Optional 4th engineer for Track T + convergence assistance.

Estimated elapsed wall-clock with parallel tracks: **~3.5 to 4.5 working days** (vs ~6.5 sequential).

---

## Risk register (parallel-specific)

1. **Risk:** cross-track conflict on shared types/exports.  
   **Mitigation:** PR1 establishes stable facade contracts before branching tracks.

2. **Risk:** behavior drift while tools migrate to shared helpers.  
   **Mitigation:** M1 introduces helpers first; M2/M3 are mechanical migrations with unchanged text outputs.

3. **Risk:** integration debt at the end.  
   **Mitigation:** reserve explicit `INT1` convergence PR; do not “sneak merge” large unresolved overlaps.

4. **Risk:** test runtime too slow in parallel CI.  
   **Mitigation:** focused suites per track + full gate only at convergence/closeout.

---

## Ready-to-start sequence

1. Merge **PR0** (baseline tests)
2. Merge **PR1** (scaffolding + facades)
3. Immediately branch and run **R1 / L1 / M1 in parallel**
4. Continue with DAG above until **INT1** and **CLOSE1**
