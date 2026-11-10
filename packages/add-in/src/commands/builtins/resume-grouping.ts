/**
 * Pure grouping/filtering helpers for the resume (session history) overlay.
 *
 * Extracted so the non-destructive behaviour is unit-testable without a DOM:
 * a session linked to a different workbook (or to no workbook) is **grouped**,
 * never dropped. Dropping those rows is what made recent sessions look missing.
 */

import type { SessionListEntry } from "../../storage/local/sessions-store.js";

export type ResumeGroupId = "this_workbook" | "other_workbooks" | "unlinked";

export interface ResumeGroup {
  id: ResumeGroupId;
  entries: SessionListEntry[];
}

/**
 * Split entries into this-workbook / other-workbook / unlinked buckets.
 *
 * When `workbookId` is null everything linked goes to `other_workbooks` and the
 * unlinked bucket stays separate (the caller renders both).
 */
export function groupSessionsForWorkbook(
  entries: readonly SessionListEntry[],
  links: ReadonlyMap<string, string>,
  workbookId: string | null,
): ResumeGroup[] {
  const thisWorkbook: SessionListEntry[] = [];
  const otherWorkbooks: SessionListEntry[] = [];
  const unlinked: SessionListEntry[] = [];

  for (const entry of entries) {
    const linked = links.get(entry.metadata.id) ?? null;
    if (!linked) {
      unlinked.push(entry);
      continue;
    }
    if (workbookId && linked === workbookId) {
      thisWorkbook.push(entry);
      continue;
    }
    otherWorkbooks.push(entry);
  }

  return [
    { id: "this_workbook", entries: thisWorkbook },
    { id: "other_workbooks", entries: otherWorkbooks },
    { id: "unlinked", entries: unlinked },
  ];
}

/** Case-insensitive title + preview search. Empty query returns a copy. */
export function filterSessionsByQuery(
  entries: readonly SessionListEntry[],
  query: string,
): SessionListEntry[] {
  const needle = query.trim().toLowerCase();
  if (needle.length === 0) return [...entries];

  return entries.filter((entry) => {
    const title = entry.metadata.title ?? "";
    const preview = entry.metadata.preview ?? "";
    return (
      title.toLowerCase().includes(needle) ||
      preview.toLowerCase().includes(needle)
    );
  });
}

/**
 * Auto-expand "Other workbooks" when the current workbook has no sessions.
 *
 * This is the second half of the fix for "recent sessions are missing": the
 * newest session is always rendered expanded if it is the only one there is.
 */
export function shouldExpandOtherWorkbooks(
  groups: readonly ResumeGroup[],
  workbookId: string | null,
): boolean {
  if (!workbookId) return true;
  const thisWorkbook = groups.find((group) => group.id === "this_workbook");
  return (thisWorkbook?.entries.length ?? 0) === 0;
}
