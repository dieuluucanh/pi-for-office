/**
 * Canonical workbook identity for **session association**.
 *
 * Why this exists
 * ---------------
 * Session history used to be scoped by a hash of `Office.context.document.url`,
 * which is not a stable identity:
 *
 * - it is empty for unsaved documents and for guest users in Excel on the web;
 * - it differs between `file:///C:/…` and `C:\…`, between a local OneDrive path
 *   and `d.docs.live.net/…`, and between a canonical SharePoint path and a
 *   `/:x:/g/…` sharing URL.
 *
 * Any one of those changes silently orphaned every session created under the
 * previous id, and the resume overlay hid the orphaned sessions.
 *
 * The identity ladder implemented here is:
 *
 *   1. **document-scoped id** (`doc_uuid:<uuid>`), stored *inside* the document
 *      by the host (Office `document.settings`, WPS `PluginStorage`) — survives
 *      rename/move/platform changes;
 *   2. **alias**, so the observed URL/path hash of a document we have already
 *      seen resolves to its document-scoped id (and keeps resolving on a boot
 *      where the host cannot read its settings, e.g. protected view);
 *   3. **observed hash**, for documents we have never seen a document id for.
 *
 * Scope note: this canonical id is intentionally used only by session
 * association (`session.workbook.v1.*`, `workbook.latestSession.v1.*`, the
 * resume overlay). Tab layout, rules, recovery logs and the Pi bridge keep the
 * observed hash, because changing their keys would strand existing data; they
 * are also recoverable surfaces where a miss is visible rather than silent.
 */

import type { SettingsStore } from "../storage/local/settings-store.js";
import type { SpreadsheetHost } from "../host/types.js";
import type { WorkbookContext } from "../host/workbook-context.js";

/** `workbook.identity.v1.alias.<observedId>` → canonical id. */
export const WORKBOOK_IDENTITY_ALIAS_PREFIX = "workbook.identity.v1.alias.";
/** `workbook.identity.v1.doc.<canonicalId>` → `{ name, firstSeen, lastSeen, observed }`. */
export const WORKBOOK_IDENTITY_DOC_PREFIX = "workbook.identity.v1.doc.";

/** Aliases are followed for at most one hop (a rewrite is never chained). */
const MAX_ALIAS_HOPS = 1;
/** Per-document cap on remembered observed ids. */
const MAX_OBSERVED_IDS_PER_DOC = 10;

export interface WorkbookIdentity {
  /**
   * Identity used for session association. `null` when the host can provide
   * neither a document id nor a document URL/path (unsaved, unknown host).
   */
  canonicalWorkbookId: string | null;
  /** Raw hash of the host-reported URL/path (diagnostics + alias writes). */
  observedWorkbookId: string | null;
  workbookName: string | null;
  source: WorkbookContext["source"];
  /**
   * Set when the observed hash was mapped to a *different* canonical id, so
   * callers can run the session re-link pass exactly once for that document.
   */
  aliasedFrom: string | null;
}

export interface ResolveWorkbookIdentityDeps {
  host: Pick<
    SpreadsheetHost,
    "getWorkbookContext" | "getDocumentScopedId" | "setDocumentScopedId"
  >;
  settings: SettingsStore;
  /** Resolve without writing aliases (diagnostics, dry runs, tests). */
  dryRun?: boolean;
  now?: () => number;
}

export interface WorkbookIdentityDocRecord {
  name: string | null;
  firstSeen: string;
  lastSeen: string;
  observed: string[];
}

/** Alias pairs already written by this pane, so resolution stays read-only. */
const recordedAliasPairs = new Set<string>();

/** Test seam: forget in-memory alias writes (does not touch storage). */
export function resetWorkbookIdentityWriteCache(): void {
  recordedAliasPairs.clear();
}

export function workbookIdentityAliasKey(observedId: string): string {
  return `${WORKBOOK_IDENTITY_ALIAS_PREFIX}${observedId}`;
}

export function workbookIdentityDocKey(canonicalId: string): string {
  return `${WORKBOOK_IDENTITY_DOC_PREFIX}${canonicalId}`;
}

function normalizeId(value: string | null | undefined): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

/**
 * Resolve `observedId` to its canonical id, following at most one hop.
 * Returns the input id when no alias exists. Self-aliases are ignored.
 */
export async function readWorkbookIdentityAlias(
  settings: SettingsStore,
  observedId: string,
): Promise<string | null> {
  let current = normalizeId(observedId);
  if (!current) return null;

  for (let hop = 0; hop < MAX_ALIAS_HOPS; hop += 1) {
    let next: string | null = null;
    try {
      const stored = await settings.get(workbookIdentityAliasKey(current));
      next = normalizeId(typeof stored === "string" ? stored : null);
    } catch {
      return current;
    }

    if (!next || next === current) return current;
    current = next;
  }

  return current;
}

async function readDocRecord(
  settings: SettingsStore,
  canonicalId: string,
): Promise<WorkbookIdentityDocRecord | null> {
  try {
    const stored = await settings.get(workbookIdentityDocKey(canonicalId));
    if (typeof stored !== "object" || stored === null) return null;

    const record = stored as Partial<WorkbookIdentityDocRecord>;
    const observed = Array.isArray(record.observed)
      ? record.observed.filter((id): id is string => typeof id === "string")
      : [];

    return {
      name: typeof record.name === "string" ? record.name : null,
      firstSeen: typeof record.firstSeen === "string" ? record.firstSeen : "",
      lastSeen: typeof record.lastSeen === "string" ? record.lastSeen : "",
      observed,
    };
  } catch {
    return null;
  }
}

/**
 * Remember that `observedId` belongs to `canonicalId`.
 *
 * Idempotent: an existing alias is left alone (it may have been written by an
 * earlier boot), and the in-memory set short-circuits repeat writes. Never
 * throws — identity must degrade, not block, session persistence.
 */
export async function recordWorkbookIdentityAlias(
  deps: {
    settings: SettingsStore;
    now?: () => number;
  },
  args: {
    observedId: string;
    canonicalId: string;
    workbookName: string | null;
  },
): Promise<void> {
  const observedId = normalizeId(args.observedId);
  const canonicalId = normalizeId(args.canonicalId);
  if (!observedId || !canonicalId || observedId === canonicalId) return;

  const pairKey = `${observedId}>${canonicalId}`;
  if (recordedAliasPairs.has(pairKey)) return;

  const nowIso = new Date(deps.now?.() ?? Date.now()).toISOString();

  try {
    const existing = await deps.settings.get(workbookIdentityAliasKey(observedId));
    if (normalizeId(typeof existing === "string" ? existing : null) !== canonicalId) {
      await deps.settings.set(workbookIdentityAliasKey(observedId), canonicalId);
    }

    const record = await readDocRecord(deps.settings, canonicalId);
    const observed = record?.observed ?? [];
    const nextObserved = observed.includes(observedId)
      ? observed
      : [...observed, observedId].slice(-MAX_OBSERVED_IDS_PER_DOC);

    await deps.settings.set(workbookIdentityDocKey(canonicalId), {
      name: args.workbookName ?? record?.name ?? null,
      firstSeen: record?.firstSeen || nowIso,
      lastSeen: nowIso,
      observed: nextObserved,
    } satisfies WorkbookIdentityDocRecord);

    recordedAliasPairs.add(pairKey);
  } catch (error) {
    console.warn("[pi] Workbook identity alias write failed:", error);
  }
}

/** Observed ids remembered for a canonical document (used by the re-link pass). */
export async function listWorkbookIdentityObservedIds(
  settings: SettingsStore,
  canonicalId: string,
): Promise<string[]> {
  const record = await readDocRecord(settings, canonicalId);
  return record?.observed ?? [];
}

/**
 * Resolve the canonical identity for the current document.
 *
 * Never throws: a host failure degrades to the observed hash which degrades to
 * `null` (sessions then save unlinked and are associated later — see
 * `pendingAssociations` in `taskpane/sessions.ts`).
 */
export async function resolveWorkbookIdentity(
  deps: ResolveWorkbookIdentityDeps,
): Promise<WorkbookIdentity> {
  let context: WorkbookContext;
  try {
    context = await deps.host.getWorkbookContext();
  } catch {
    context = { workbookId: null, workbookName: null, source: "unknown" };
  }

  const observedWorkbookId = normalizeId(context.workbookId);

  let documentScopedId: string | null = null;
  try {
    documentScopedId = normalizeId(await deps.host.getDocumentScopedId());
  } catch {
    documentScopedId = null;
  }

  const aliased = observedWorkbookId
    ? await readWorkbookIdentityAlias(deps.settings, observedWorkbookId)
    : null;

  const canonicalWorkbookId = documentScopedId ?? aliased ?? observedWorkbookId;

  if (!deps.dryRun && canonicalWorkbookId && observedWorkbookId) {
    await recordWorkbookIdentityAlias(deps, {
      observedId: observedWorkbookId,
      canonicalId: canonicalWorkbookId,
      workbookName: context.workbookName,
    });
  }

  return {
    canonicalWorkbookId,
    observedWorkbookId,
    workbookName: context.workbookName,
    source: context.source,
    aliasedFrom:
      observedWorkbookId && canonicalWorkbookId !== observedWorkbookId
        ? observedWorkbookId
        : null,
  };
}
