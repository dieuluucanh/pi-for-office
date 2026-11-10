/**
 * Session ↔ workbook association.
 *
 * `SessionsStore` metadata schema is fixed (pi-web-ui), so we store the mapping
 * in `SettingsStore` using a dedicated key prefix.
 *
 * Invariants:
 * - a session is linked to **at most one** workbook, but the link is
 *   *re-linkable* — the canonical identity (`workbook/identity.ts`) may change
 *   when a document-scoped id appears or an observed hash is aliased;
 * - `workbook.latestSession.v1.<id>` is only ever written for a **non-empty**
 *   session (see `setupSessionPersistence.saveSession`);
 * - all reads are batched (`listWithPrefix`) so rendering history is O(1)
 *   storage round-trips instead of one `get` per session.
 */

import type { SettingsStore } from "../storage/local/settings-store.js";

/** Prefix for `session.workbook.v1.<sessionId>` → workbook id. */
export const SESSION_WORKBOOK_PREFIX = "session.workbook.v1.";
/** Prefix for `workbook.latestSession.v1.<workbookId>` → session id. */
export const WORKBOOK_LATEST_SESSION_PREFIX = "workbook.latestSession.v1.";
/** Prefix for sessions saved while no workbook identity was resolvable. */
export const SESSION_PENDING_WORKBOOK_PREFIX = "session.pendingWorkbook.v1.";

export function sessionWorkbookKey(sessionId: string): string {
  return `${SESSION_WORKBOOK_PREFIX}${sessionId}`;
}

export function workbookLatestSessionKey(workbookId: string): string {
  return `${WORKBOOK_LATEST_SESSION_PREFIX}${workbookId}`;
}

export function sessionPendingWorkbookKey(sessionId: string): string {
  return `${SESSION_PENDING_WORKBOOK_PREFIX}${sessionId}`;
}

function normalizeId(value: DynamicValue): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export async function getSessionWorkbookId(
  settings: SettingsStore,
  sessionId: string,
): Promise<string | null> {
  const v = await settings.get(sessionWorkbookKey(sessionId));
  return normalizeId(v);
}

/**
 * Link a session to a workbook.
 *
 * Does not overwrite an existing link (so resuming an old session in a different
 * workbook won't accidentally "move" it). Use `relinkSessionToWorkbook` when the
 * link must be corrected (alias/canonical migration or an explicit user action).
 */
export async function linkSessionToWorkbook(
  settings: SettingsStore,
  sessionId: string,
  workbookId: string,
): Promise<void> {
  const key = sessionWorkbookKey(sessionId);
  const existing = await settings.get(key);
  if (normalizeId(existing)) return;
  await settings.set(key, workbookId);
}

/**
 * Force the link to a workbook id.
 *
 * Used by the alias/canonical migration and by the overlay's explicit
 * "Move to this workbook" action. Only corrects a link the caller has already
 * established is the same document.
 */
export async function relinkSessionToWorkbook(
  settings: SettingsStore,
  sessionId: string,
  workbookId: string,
): Promise<void> {
  await settings.set(sessionWorkbookKey(sessionId), workbookId);
}

/** Remove a link (never deletes the session itself). */
export async function forgetSessionWorkbook(
  settings: SettingsStore,
  sessionId: string,
): Promise<void> {
  await settings.delete(sessionWorkbookKey(sessionId));
}

export async function setLatestSessionForWorkbook(
  settings: SettingsStore,
  workbookId: string,
  sessionId: string,
): Promise<void> {
  await settings.set(workbookLatestSessionKey(workbookId), sessionId);
}

export async function getLatestSessionForWorkbook(
  settings: SettingsStore,
  workbookId: string,
): Promise<string | null> {
  const v = await settings.get(workbookLatestSessionKey(workbookId));
  return normalizeId(v);
}

export interface SessionWorkbookPartition {
  matchingSessionIds: string[];
  unlinkedSessionIds: string[];
  foreignSessionIds: string[];
}

/**
 * Partition session ids by workbook association.
 *
 * - `matchingSessionIds`: linked to this workbook
 * - `unlinkedSessionIds`: legacy/no mapping
 * - `foreignSessionIds`: linked to another workbook
 */
export async function partitionSessionIdsByWorkbook(
  settings: SettingsStore,
  sessionIds: string[],
  workbookId: string,
): Promise<SessionWorkbookPartition> {
  const links = await readSessionWorkbookLinks(settings);

  const matchingSessionIds: string[] = [];
  const unlinkedSessionIds: string[] = [];
  const foreignSessionIds: string[] = [];

  for (const sessionId of sessionIds) {
    const linkedWorkbookId = links.get(sessionId) ?? null;
    if (!linkedWorkbookId) {
      unlinkedSessionIds.push(sessionId);
      continue;
    }

    if (linkedWorkbookId === workbookId) {
      matchingSessionIds.push(sessionId);
      continue;
    }

    foreignSessionIds.push(sessionId);
  }

  return {
    matchingSessionIds,
    unlinkedSessionIds,
    foreignSessionIds,
  };
}

/** Read every `session.workbook.v1.*` link in one settings scan. */
export async function readSessionWorkbookLinks(
  settings: SettingsStore,
): Promise<Map<string, string>> {
  const links = new Map<string, string>();

  let keys: string[];
  try {
    keys = await settings.listWithPrefix(SESSION_WORKBOOK_PREFIX);
  } catch {
    return links;
  }

  for (const key of keys) {
    const sessionId = key.slice(SESSION_WORKBOOK_PREFIX.length);
    if (!sessionId) continue;

    try {
      const value = normalizeId(await settings.get(key));
      if (value) links.set(sessionId, value);
    } catch {
      // A single unreadable link must not blank the whole partition.
    }
  }

  return links;
}

export interface MigrateAliasesResult {
  relinked: number;
  latestMoved: boolean;
}

/**
 * Repoint every session linked to one of `observedIds` at `canonicalId`, and
 * make sure the canonical latest-session pointer is at least as new as any
 * pointer that lived under the observed ids.
 *
 * Called once per document when identity resolution reports an alias
 * (`WorkbookIdentity.aliasedFrom`). Idempotent.
 */
export async function migrateAliasesToCanonical(
  settings: SettingsStore,
  args: {
    canonicalId: string;
    observedIds: string[];
    /** Newest-wins comparison for the latest-session pointer. */
    resolveSessionLastModified?: (sessionId: string) => Promise<string | null>;
  },
): Promise<MigrateAliasesResult> {
  const canonicalId = normalizeId(args.canonicalId);
  const observedIds = args.observedIds
    .map((id) => normalizeId(id))
    .filter((id): id is string => id !== null && id !== canonicalId);

  if (!canonicalId || observedIds.length === 0) {
    return { relinked: 0, latestMoved: false };
  }

  const observedSet = new Set(observedIds);
  let relinked = 0;

  const links = await readSessionWorkbookLinks(settings);
  for (const [sessionId, linkedId] of links) {
    if (!observedSet.has(linkedId)) continue;
    await relinkSessionToWorkbook(settings, sessionId, canonicalId);
    relinked += 1;
  }

  const currentLatest = await getLatestSessionForWorkbook(settings, canonicalId);
  const candidates = new Set<string>();
  if (currentLatest) candidates.add(currentLatest);

  for (const observedId of observedIds) {
    const observedLatest = await getLatestSessionForWorkbook(settings, observedId);
    if (observedLatest) candidates.add(observedLatest);
  }

  if (candidates.size === 0) return { relinked, latestMoved: false };

  let chosen = currentLatest ?? null;
  if (candidates.size > 1 && args.resolveSessionLastModified) {
    let chosenAt = chosen ? await args.resolveSessionLastModified(chosen) : null;

    for (const candidate of candidates) {
      if (candidate === chosen) continue;
      const candidateAt = await args.resolveSessionLastModified(candidate);
      if (!candidateAt) continue;
      if (!chosenAt || candidateAt > chosenAt) {
        chosen = candidate;
        chosenAt = candidateAt;
      }
    }
  } else {
    chosen = chosen ?? [...candidates][0] ?? null;
  }

  if (!chosen || chosen === currentLatest) {
    return { relinked, latestMoved: false };
  }

  await setLatestSessionForWorkbook(settings, canonicalId, chosen);
  return { relinked, latestMoved: true };
}

/** Record that a session was saved while no workbook identity was available. */
export async function markSessionWorkbookPending(
  settings: SettingsStore,
  sessionId: string,
  args: { workbookName: string | null; at: string },
): Promise<void> {
  try {
    await settings.set(sessionPendingWorkbookKey(sessionId), {
      name: args.workbookName,
      at: args.at,
    });
  } catch (error) {
    console.warn("[pi] Failed to record pending workbook association:", error);
  }
}

/** Drop the pending marker once a session has a real link. */
export async function clearSessionWorkbookPending(
  settings: SettingsStore,
  sessionId: string,
): Promise<void> {
  try {
    await settings.delete(sessionPendingWorkbookKey(sessionId));
  } catch {
    // best effort
  }
}
