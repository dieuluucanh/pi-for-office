/**
 * Pure stacking model for toasts.
 *
 * The DOM layer (`ui/toast.ts`) owns elements, timers and hover/focus pausing;
 * everything that decides *which* toasts exist lives here so it can be tested
 * without a DOM. Rules (docs/ui-layout.md):
 *
 * - At most `MAX_VISIBLE_TOASTS` toasts are shown.
 * - A repeated message reuses its toast (timers reset) instead of stacking.
 * - Only the newest action toast stays actionable: adding one demotes the
 *   previous sticky toast, keeping the single-undo invariant that the
 *   Cmd/Ctrl+Z arbitration depends on (`taskpane/keyboard-shortcuts.ts`).
 * - Eviction prefers the oldest plain toast, then the oldest error toast, then
 *   the oldest sticky toast — so an error or an undo never pushes readable
 *   information out of view while a plain notice is on screen.
 */

export type ToastVariant = "info" | "error";

export const MAX_VISIBLE_TOASTS = 3;

export interface ToastRecord {
  id: string;
  message: string;
  variant: ToastVariant;
  /** Auto-dismiss budget in ms; `0` means "until dismissed". */
  duration: number;
  /** Action (undo/retry) toast — keeps its button while it is the newest. */
  sticky: boolean;
  /** Monotonic insertion order; drives eviction and undo targeting. */
  seq: number;
}

export interface ToastDraft {
  message: string;
  variant: ToastVariant;
  duration: number;
  sticky: boolean;
}

export interface PushToastResult {
  list: ToastRecord[];
  record: ToastRecord;
  /** Set when an existing toast was reused (its timer must be reset). */
  reused: boolean;
  /** Toast that must be removed from the DOM. */
  evictedId: string | null;
  /** Toasts whose action button must be dropped after sticky demotion. */
  demotedIds: string[];
}

export function createToastRecord(draft: ToastDraft, seq: number): ToastRecord {
  return {
    id: `toast-${seq}`,
    message: draft.message,
    variant: draft.variant,
    duration: draft.duration,
    sticky: draft.sticky,
    seq,
  };
}

export function pushToast(
  list: readonly ToastRecord[],
  draft: ToastDraft,
  seq: number,
): PushToastResult {
  const existing = list.find((record) => record.message === draft.message);

  if (existing) {
    const updated: ToastRecord = {
      ...existing,
      variant: draft.variant,
      duration: draft.duration,
      // A fresh action toast must win the undo slot even when its message
      // repeats; a plain repeat must not strip an existing action button.
      sticky: draft.sticky ? true : existing.sticky,
    };

    const nextList = list.map((record) => (record.id === existing.id ? updated : record));
    const demotedIds = draft.sticky
      ? demoteOtherSticky(nextList, updated.id)
      : [];

    return {
      list: applyDemotions(nextList, demotedIds),
      record: updated,
      reused: true,
      evictedId: null,
      demotedIds,
    };
  }

  const appended = createToastRecord(draft, seq);
  let nextList: ToastRecord[] = [...list, appended];
  const demotedIds = draft.sticky
    ? demoteOtherSticky(nextList, appended.id)
    : [];
  nextList = applyDemotions(nextList, demotedIds);

  let evictedId: string | null = null;
  if (nextList.length > MAX_VISIBLE_TOASTS) {
    const evicted = pickEviction(nextList);
    if (evicted) {
      evictedId = evicted.id;
      nextList = nextList.filter((record) => record.id !== evicted.id);
    }
  }

  const added = nextList.find((record) => record.id === appended.id);
  return {
    list: nextList,
    record: added ?? appended,
    reused: false,
    evictedId,
    demotedIds,
  };
}

export function removeToast(list: readonly ToastRecord[], id: string): ToastRecord[] {
  return list.filter((record) => record.id !== id);
}

/** The newest sticky toast owns the undo/retry affordance. */
export function pickUndoTarget(list: readonly ToastRecord[]): ToastRecord | null {
  let newest: ToastRecord | null = null;

  for (const record of list) {
    if (!record.sticky) continue;
    if (!newest || record.seq > newest.seq) newest = record;
  }

  return newest;
}

function demoteOtherSticky(list: readonly ToastRecord[], keepId: string): string[] {
  return list
    .filter((record) => record.sticky && record.id !== keepId)
    .map((record) => record.id);
}

function applyDemotions(list: ToastRecord[], demotedIds: readonly string[]): ToastRecord[] {
  if (demotedIds.length === 0) return list;

  return list.map((record) => (
    demotedIds.includes(record.id) ? { ...record, sticky: false } : record
  ));
}

function pickEviction(list: readonly ToastRecord[]): ToastRecord | null {
  const oldest = (candidates: readonly ToastRecord[]): ToastRecord | null => {
    let winner: ToastRecord | null = null;
    for (const record of candidates) {
      if (!winner || record.seq < winner.seq) winner = record;
    }
    return winner;
  };

  return (
    oldest(list.filter((record) => !record.sticky && record.variant !== "error"))
    ?? oldest(list.filter((record) => !record.sticky))
    ?? oldest(list)
  );
}
