/**
 * Session-storage diagnostics.
 *
 * `describeSessionStorage` is intentionally pure (plain data in, plain string
 * out) so it can be unit-tested without IndexedDB; `collectSessionStorageDiagnostics`
 * gathers the live numbers from the storage layer + settings and feeds it.
 *
 * Why this exists: "my session is missing" reports were previously impossible to
 * triage without devtools. The readout answers, in one block, which of the three
 * failure modes is in play — nothing written, written but unindexed, or written
 * under a different workbook identity.
 */

import type { SettingsStore } from "./settings-store.js";

/** Marker written after a successful legacy-database copy. */
export const LEGACY_MIGRATION_MARKER_KEY = "__pi_migrated_legacy_db_v1";
/** Marker written after the one-time storage repair pass. */
export const STORAGE_REPAIR_MARKER_KEY = "storage.repair.v1";
/** Marker recording whether persistent storage was granted. */
export const STORAGE_PERSISTENCE_MARKER_KEY = "storage.persistence.v1";

/** Prefixes owned by the session↔workbook association layer. */
export const SESSION_WORKBOOK_KEY_PREFIX = "session.workbook.v1.";
const WORKBOOK_LATEST_SESSION_KEY_PREFIX = "workbook.latestSession.v1.";

/**
 * Minimal full-store scan surface.
 *
 * Declared locally (instead of depending on the whole `StorageBackend`) so this
 * diagnostics module only needs the one operation it uses — and so an
 * in-memory test double can implement just `getAll`.
 */
export interface StoreScanner {
  getAll<T>(storeName: string): Promise<T[]>;
}

export interface SessionStorageCounts {
  /** Rows in the full `sessions` store (full-store scan). */
  sessions: number;
  /** Rows in the `sessions-metadata` store (full-store scan). */
  metadata: number;
  /** `session.workbook.v1.*` settings keys. */
  links: number;
  /** `workbook.latestSession.v1.*` settings keys. */
  latestPointers: number;
}

/** Values that can be stored as diagnostics markers. */
export type SessionStorageMarkerValue =
  | string
  | number
  | boolean
  | object
  | null
  | undefined;

export interface SessionStorageDiagnosticInput {
  dbName: string;
  dbVersion: number;
  counts: SessionStorageCounts;
  /** Canonical id used for association now. */
  canonicalWorkbookId: string | null;
  /** Raw hash of the host-reported URL/path (may differ from canonical). */
  observedWorkbookId: string | null;
  /** `navigator.storage.persisted()`; null when the API is unavailable. */
  persisted: boolean | null;
  usageBytes: number;
  quotaBytes: number;
  migrationMarker: SessionStorageMarkerValue;
  repairMarker: SessionStorageMarkerValue;
  persistenceMarker: SessionStorageMarkerValue;
  lastSaveError: string | null;
}

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  return `${value.toFixed(value >= 10 || unit === 0 ? 0 : 1)} ${units[unit] ?? "B"}`;
}

function describeMarker(
  value: SessionStorageMarkerValue,
  present: string,
): string {
  if (value === null || value === undefined) return "absent";
  return present;
}

function describeMetadataGap(sessions: number, metadata: number): string {
  const gap = sessions - metadata;
  if (gap === 0) return "";
  if (gap > 0) return ` (${gap} missing metadata)`;
  return ` (${-gap} orphan metadata)`;
}

function describeId(value: string | null): string {
  return value && value.trim().length > 0 ? value : "(none)";
}

function describeIdentityMatch(
  canonicalWorkbookId: string | null,
  observedWorkbookId: string | null,
): string {
  if (!canonicalWorkbookId) return "unknown (sessions saved unlinked)";
  if (observedWorkbookId && canonicalWorkbookId !== observedWorkbookId) {
    return "aliased";
  }
  return "stable";
}

function describePersistence(persisted: boolean | null): string {
  if (persisted === null) return "unavailable";
  if (persisted) return "granted";
  return "NOT granted (eviction possible)";
}

/**
 * Render a deterministic, human-readable diagnostics block.
 *
 * No timestamps: the output is asserted verbatim in tests.
 */
export function describeSessionStorage(
  input: SessionStorageDiagnosticInput,
): string {
  const { counts } = input;
  const quotaPercent =
    input.quotaBytes > 0 ? (input.usageBytes / input.quotaBytes) * 100 : 0;

  const lines = [
    `storage: ${input.dbName}@v${input.dbVersion}`,
    `sessions: ${counts.sessions} records / ${counts.metadata} metadata` +
      describeMetadataGap(counts.sessions, counts.metadata),
    `links: ${counts.links} · latest pointers: ${counts.latestPointers}`,
    `workbook: canonical=${describeId(input.canonicalWorkbookId)} observed=${describeId(input.observedWorkbookId)}`,
    `identity match: ${describeIdentityMatch(input.canonicalWorkbookId, input.observedWorkbookId)}`,
    `persistence: ${describePersistence(input.persisted)} · ${formatBytes(input.usageBytes)} / ${formatBytes(input.quotaBytes)} (${quotaPercent.toFixed(1)}%)`,
    `legacy migration: ${describeMarker(input.migrationMarker, "done")}`,
    `repair: ${describeMarker(input.repairMarker, "done")}`,
    `persistence request: ${describeMarker(input.persistenceMarker, "done")}`,
    `last save error: ${input.lastSaveError ?? "(none)"}`,
  ];

  return lines.join("\n");
}

export interface CollectSessionStorageDiagnosticsDeps {
  backend: StoreScanner;
  settings: SettingsStore;
  dbName: string;
  dbVersion: number;
  canonicalWorkbookId: string | null;
  observedWorkbookId: string | null;
  lastSaveError?: string | null;
}

/**
 * Gather live diagnostics. Never throws: a failing probe degrades to zeroed
 * counts so the readout still renders (and the failure itself is visible as a
 * size mismatch).
 */
export async function collectSessionStorageDiagnostics(
  deps: CollectSessionStorageDiagnosticsDeps,
): Promise<string> {
  const safe = async <T>(fn: () => Promise<T>, fallback: T): Promise<T> => {
    try {
      return await fn();
    } catch {
      return fallback;
    }
  };

  const [sessions, metadata, settingsKeys, persisted, quota] = await Promise.all([
    safe(() => deps.backend.getAll("sessions"), []),
    safe(() => deps.backend.getAll("sessions-metadata"), []),
    safe(() => deps.settings.list(), []),
    safe(async () => {
      const storage = globalThis.navigator?.storage;
      if (!storage?.persisted) return null;
      const granted: boolean = await storage.persisted();
      return granted;
    }, null),
    safe(async () => {
      const storage = globalThis.navigator?.storage;
      if (!storage?.estimate) {
        return { usage: 0, quota: 0 };
      }
      const estimate = await storage.estimate();
      return { usage: estimate.usage ?? 0, quota: estimate.quota ?? 0 };
    }, { usage: 0, quota: 0 }),
  ]);

  const [migrationMarker, repairMarker, persistenceMarker] = await Promise.all([
    safe<SessionStorageMarkerValue>(
      () => deps.settings.get<SessionStorageMarkerValue>(LEGACY_MIGRATION_MARKER_KEY),
      null,
    ),
    safe<SessionStorageMarkerValue>(
      () => deps.settings.get<SessionStorageMarkerValue>(STORAGE_REPAIR_MARKER_KEY),
      null,
    ),
    safe<SessionStorageMarkerValue>(
      () => deps.settings.get<SessionStorageMarkerValue>(STORAGE_PERSISTENCE_MARKER_KEY),
      null,
    ),
  ]);

  return describeSessionStorage({
    dbName: deps.dbName,
    dbVersion: deps.dbVersion,
    counts: {
      sessions: sessions.length,
      metadata: metadata.length,
      links: settingsKeys.filter((key) => key.startsWith(SESSION_WORKBOOK_KEY_PREFIX)).length,
      latestPointers: settingsKeys.filter((key) =>
        key.startsWith(WORKBOOK_LATEST_SESSION_KEY_PREFIX),
      ).length,
    },
    canonicalWorkbookId: deps.canonicalWorkbookId,
    observedWorkbookId: deps.observedWorkbookId,
    persisted,
    usageBytes: quota.usage,
    quotaBytes: quota.quota,
    migrationMarker,
    repairMarker,
    persistenceMarker,
    lastSaveError: deps.lastSaveError ?? null,
  });
}
