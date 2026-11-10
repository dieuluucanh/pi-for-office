/**
 * One-time session-storage repair.
 *
 * Fixes the damage left by the pre-fix persistence paths, without ever
 * destroying user work:
 *
 * 1. **Missing metadata** — a `sessions` row whose `sessions-metadata` row was
 *    never written (or whose transaction aborted) is invisible to every
 *    index-based listing. Rebuild it from the record.
 * 2. **Unindexed timestamps** — a metadata row without a valid `lastModified`
 *    is not in the index at all. Backfill it so the row is listable.
 * 3. **Stale workbook links** — repoint links written under an observed
 *    URL/path hash to the canonical document identity (alias migration).
 * 4. **Blank duplicates** — empty records created by the old restore path are
 *    removed *only* when no tab layout, latest-session pointer, or
 *    recently-closed entry references them. Non-empty sessions are never
 *    deleted, ever.
 * 5. **Identity table growth** — prune the per-document identity records.
 *
 * Runs at most once per app version (marker `storage.repair.v1`), and supports
 * `dryRun` so the diagnostics readout can preview it.
 */

import type { SessionsStore, SessionListEntry } from "./sessions-store.js";
import { deriveSessionMetadata } from "./sessions-store.js";
import type { SettingsStore } from "./settings-store.js";
import type { SessionData, SessionMetadata, StorageBackend } from "./types.js";
import { STORAGE_REPAIR_MARKER_KEY } from "./diagnostics.js";
import {
  migrateAliasesToCanonical,
  WORKBOOK_LATEST_SESSION_PREFIX,
} from "../../workbook/session-association.js";
import {
  WORKBOOK_IDENTITY_DOC_PREFIX,
  type WorkbookIdentityDocRecord,
} from "../../workbook/identity.js";

export const SESSION_REPAIR_VERSION = 1;

/**
 * Settings keys that reference session ids. Literals (mirrored from
 * `taskpane/tab-layout.ts` and `taskpane/recently-closed-store.ts`) so the
 * storage layer does not import taskpane modules; `session-repair.test.ts`
 * asserts they still match.
 */
export const TAB_LAYOUT_KEY_PREFIX = "workbook.tabLayout.v1.";
export const RECENTLY_CLOSED_KEY = "ui.recentlyClosed.v1";

const MAX_IDENTITY_DOCS = 50;

export interface SessionRepairReport {
  dryRun: boolean;
  scannedRecords: number;
  metadataRebuilt: number;
  timestampsRepaired: number;
  linksRepointed: number;
  latestMoved: boolean;
  blankDuplicatesRemoved: number;
  identityDocsPruned: number;
  errors: string[];
}

export interface SessionRepairDeps {
  sessions: SessionsStore;
  settings: SettingsStore;
  backend: StorageBackend;
  canonicalWorkbookId?: string | null;
  observedIds?: string[];
  dryRun?: boolean;
  force?: boolean;
  now?: () => number;
}

function parseTimestamp(value: string | null | undefined): number {
  if (typeof value !== "string") return 0;
  const parsed = Date.parse(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

async function readReferencedSessionIds(settings: SettingsStore): Promise<Set<string>> {
  const ids = new Set<string>();

  const addLayoutIds = (value: DynamicValue): void => {
    if (typeof value !== "object" || value === null) return;
    const sessionIds = (value as { sessionIds?: DynamicValue }).sessionIds;
    if (!Array.isArray(sessionIds)) return;
    for (const id of sessionIds) {
      if (typeof id === "string" && id.length > 0) ids.add(id);
    }
    const active = (value as { activeSessionId?: DynamicValue }).activeSessionId;
    if (typeof active === "string" && active.length > 0) ids.add(active);
  };

  try {
    for (const key of await settings.listWithPrefix(TAB_LAYOUT_KEY_PREFIX)) {
      addLayoutIds(await settings.get(key));
    }
  } catch {
    // A missing layout table just means nothing is referenced from it.
  }

  try {
    for (const key of await settings.listWithPrefix(WORKBOOK_LATEST_SESSION_PREFIX)) {
      const value = await settings.get(key);
      if (typeof value === "string" && value.length > 0) ids.add(value);
    }
  } catch {
    // ignore
  }

  try {
    const recent = await settings.get(RECENTLY_CLOSED_KEY);
    if (typeof recent === "object" && recent !== null) {
      const items = (recent as { items?: DynamicValue }).items;
      if (Array.isArray(items)) {
        for (const item of items) {
          if (typeof item !== "object" || item === null) continue;
          const sessionId = (item as { sessionId?: DynamicValue }).sessionId;
          if (typeof sessionId === "string" && sessionId.length > 0) {
            ids.add(sessionId);
          }
        }
      }
    }
  } catch {
    // ignore
  }

  return ids;
}

/**
 * Repair the session stores.
 *
 * Returns `null` when this version already ran and `force`/`dryRun` were not
 * requested.
 */
export async function repairSessionStorage(
  deps: SessionRepairDeps,
): Promise<SessionRepairReport | null> {
  const dryRun = deps.dryRun === true;
  const now = deps.now?.() ?? Date.now();

  if (!dryRun && !deps.force) {
    const marker = await deps.settings.get(STORAGE_REPAIR_MARKER_KEY).catch(() => null);
    if (
      typeof marker === "object" &&
      marker !== null &&
      (marker as { version?: DynamicValue }).version === SESSION_REPAIR_VERSION
    ) {
      return null;
    }
  }

  const report: SessionRepairReport = {
    dryRun,
    scannedRecords: 0,
    metadataRebuilt: 0,
    timestampsRepaired: 0,
    linksRepointed: 0,
    latestMoved: false,
    blankDuplicatesRemoved: 0,
    identityDocsPruned: 0,
    errors: [],
  };

  let records: SessionData[] = [];
  try {
    records = await deps.backend.getAll<SessionData>("sessions");
  } catch (error) {
    report.errors.push(`record scan failed: ${String(error)}`);
    return report;
  }
  report.scannedRecords = records.length;

  let metadataRows: SessionMetadata[] = [];
  try {
    metadataRows = await deps.backend.getAll<SessionMetadata>("sessions-metadata");
  } catch (error) {
    report.errors.push(`metadata scan failed: ${String(error)}`);
  }
  const metadataById = new Map(metadataRows.map((row) => [row.id, row]));

  // 1 + 2: rebuild missing metadata, repair unusable timestamps.
  for (const record of records) {
    if (!record || typeof record.id !== "string") continue;

    const existing = metadataById.get(record.id);
    if (!existing) {
      report.metadataRebuilt += 1;
      if (!dryRun) {
        try {
          const derived = deriveSessionMetadata(record, now);
          await deps.sessions.save(record, derived);
        } catch (error) {
          report.errors.push(`metadata rebuild failed for ${record.id}: ${String(error)}`);
        }
      }
      continue;
    }

    if (parseTimestamp(existing.lastModified) === 0) {
      report.timestampsRepaired += 1;
      if (!dryRun) {
        try {
          const derived = deriveSessionMetadata(record, now);
          await deps.sessions.save(record, derived);
        } catch (error) {
          report.errors.push(`timestamp repair failed for ${record.id}: ${String(error)}`);
        }
      }
    }
  }

  // 3: alias → canonical link migration.
  const observedIds = (deps.observedIds ?? []).filter(
    (id): id is string => typeof id === "string" && id.length > 0,
  );
  if (deps.canonicalWorkbookId && observedIds.length > 0) {
    if (dryRun) {
      const links = await deps.settings
        .listWithPrefix("session.workbook.v1.")
        .catch(() => [] as string[]);
      for (const key of links) {
        const value = await deps.settings.get(key).catch(() => null);
        if (typeof value === "string" && observedIds.includes(value)) {
          report.linksRepointed += 1;
        }
      }
    } else {
      try {
        const migrated = await migrateAliasesToCanonical(deps.settings, {
          canonicalId: deps.canonicalWorkbookId,
          observedIds,
          resolveSessionLastModified: async (sessionId: string) =>
            (await deps.sessions.getMetadata(sessionId))?.lastModified ?? null,
        });
        report.linksRepointed = migrated.relinked;
        report.latestMoved = migrated.latestMoved;
      } catch (error) {
        report.errors.push(`link migration failed: ${String(error)}`);
      }
    }
  }

  // 4: blank duplicates that nothing references.
  const referenced = await readReferencedSessionIds(deps.settings);
  const blankDuplicates = records.filter(
    (record) =>
      record &&
      typeof record.id === "string" &&
      Array.isArray(record.messages) &&
      record.messages.length === 0 &&
      !referenced.has(record.id),
  );

  report.blankDuplicatesRemoved = blankDuplicates.length;
  if (!dryRun) {
    for (const record of blankDuplicates) {
      try {
        await deps.sessions.delete(record.id);
      } catch (error) {
        report.errors.push(`blank delete failed for ${record.id}: ${String(error)}`);
      }
    }
  }

  // 5: prune the identity table.
  if (!dryRun) {
    try {
      const keys = await deps.settings.listWithPrefix(WORKBOOK_IDENTITY_DOC_PREFIX);
      if (keys.length > MAX_IDENTITY_DOCS) {
        const ranked: Array<{ key: string; lastSeen: number }> = [];
        for (const key of keys) {
          const value = await deps.settings.get(key).catch(() => null);
          const record = value as Partial<WorkbookIdentityDocRecord> | null;
          ranked.push({ key, lastSeen: parseTimestamp(record?.lastSeen) });
        }
        ranked.sort((left, right) => right.lastSeen - left.lastSeen);

        for (const stale of ranked.slice(MAX_IDENTITY_DOCS)) {
          await deps.settings.delete(stale.key);
          report.identityDocsPruned += 1;
        }
      }
    } catch (error) {
      report.errors.push(`identity prune failed: ${String(error)}`);
    }
  }

  if (!dryRun) {
    try {
      await deps.settings.set(STORAGE_REPAIR_MARKER_KEY, {
        version: SESSION_REPAIR_VERSION,
        at: new Date(now).toISOString(),
        report,
      });
    } catch (error) {
      report.errors.push(`repair marker write failed: ${String(error)}`);
    }
  }

  return report;
}

/** Format a repair report for the diagnostics readout. */
export function formatRepairReport(report: SessionRepairReport): string {
  const lines = [
    `repair${report.dryRun ? " (dry run)" : ""}: records=${report.scannedRecords}`,
    `metadata rebuilt=${report.metadataRebuilt} · timestamps=${report.timestampsRepaired}`,
    `links repointed=${report.linksRepointed} · latest moved=${report.latestMoved ? "yes" : "no"}`,
    `blank duplicates removed=${report.blankDuplicatesRemoved} · identity docs pruned=${report.identityDocsPruned}`,
  ];
  if (report.errors.length > 0) {
    lines.push(`errors: ${report.errors.slice(0, 5).join(" | ")}`);
  }
  return lines.join("\n");
}

/** Unused export kept for the list type contract used by callers/tests. */
export type RepairSessionListEntry = SessionListEntry;
