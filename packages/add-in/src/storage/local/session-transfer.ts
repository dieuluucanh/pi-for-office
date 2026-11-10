/**
 * Session export / import (JSON).
 *
 * This is the escape hatch for everything the app cannot control:
 * - IndexedDB is per-origin, so a dev build (`localhost:3141`) and the
 *   production add-in (`dieuluucanh.github.io`) have separate histories;
 * - WebView storage can be evicted when persistent storage is not granted;
 * - a profile/device move loses the store entirely.
 *
 * The format is plain: full `SessionData` plus its metadata, so an import
 * restores a session byte-for-byte (messages, model, thinking level,
 * timestamps) without depending on internal keys.
 *
 * Boundary rule (enforced by `check:boundary-casts`): `JSON.parse` lands as
 * `DynamicValue`; `parseSessionTransfer` then refines it field by field. An
 * entry without a usable model/messages is skipped and reported, never trusted.
 */

import type { SessionsStore } from "./sessions-store.js";
import type { SessionData, SessionMetadata } from "./types.js";

export const SESSION_TRANSFER_SCHEMA = "pi-sessions/v1";

export interface SessionTransferEntry {
  data: SessionData;
  /** Optional: recomputed on import from `data` when missing/stale. */
  metadata?: SessionMetadata;
}

export interface SessionTransferFile {
  schema: typeof SESSION_TRANSFER_SCHEMA;
  exportedAt: string;
  sessions: SessionTransferEntry[];
}

export interface ImportSessionsResult {
  imported: number;
  skipped: number;
  /** Human-readable problems, bounded to the first 10. */
  errors: string[];
}

const EMPTY_USAGE: SessionMetadata["usage"] = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
};

interface SessionTransferPayloadShape {
  schema?: DynamicValue;
  exportedAt?: DynamicValue;
  sessions?: DynamicValue;
}

function isSessionTransferPayloadShape(
  value: DynamicValue,
): value is SessionTransferPayloadShape {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Concrete boundary parser for a persisted model descriptor. */
function isSessionModelShape(value: DynamicValue): value is SessionData["model"] {
  if (typeof value !== "object" || value === null) return false;
  const model = value as Partial<SessionData["model"]>;

  return (
    typeof model.id === "string" &&
    typeof model.name === "string" &&
    typeof model.api === "string" &&
    typeof model.provider === "string" &&
    typeof model.baseUrl === "string" &&
    typeof model.contextWindow === "number" &&
    typeof model.maxTokens === "number" &&
    Array.isArray(model.input) &&
    typeof model.cost === "object" &&
    model.cost !== null
  );
}

function isSessionUsageShape(value: DynamicValue): value is SessionMetadata["usage"] {
  if (typeof value !== "object" || value === null) return false;
  const usage = value as Partial<SessionMetadata["usage"]>;

  return (
    typeof usage.cost === "object" &&
    usage.cost !== null &&
    typeof usage.input === "number" &&
    typeof usage.output === "number" &&
    typeof usage.totalTokens === "number"
  );
}

function normalizeThinkingLevel(value: DynamicValue): SessionData["thinkingLevel"] {
  return typeof value === "string" ? (value as SessionData["thinkingLevel"]) : "off";
}

function normalizeId(value: DynamicValue): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function normalizeIso(value: string | null | undefined, fallback: string): string {
  if (typeof value !== "string") return fallback;
  const parsed = Date.parse(value);
  if (!Number.isFinite(parsed)) return fallback;
  return new Date(parsed).toISOString();
}

/** Build the export payload from everything currently stored. */
export async function buildSessionExport(
  sessions: SessionsStore,
  opts?: { now?: () => number },
): Promise<SessionTransferFile> {
  const entries: SessionTransferEntry[] = [];

  for (const { metadata } of await sessions.getAllMetadataSafe()) {
    const data = await sessions.get(metadata.id);
    if (!data) continue;
    entries.push({ metadata, data });
  }

  return {
    schema: SESSION_TRANSFER_SCHEMA,
    exportedAt: new Date(opts?.now?.() ?? Date.now()).toISOString(),
    sessions: entries,
  };
}

export function serializeSessionExport(file: SessionTransferFile): string {
  return JSON.stringify(file, null, 2);
}

/**
 * Parse + validate an export file.
 * Returns a human-readable error instead of throwing so the UI can show it.
 */
export function parseSessionTransfer(
  raw: string,
): { file: SessionTransferFile } | { error: string } {
  let payload: DynamicValue;
  try {
    payload = JSON.parse(raw);
  } catch {
    return { error: "Not valid JSON." };
  }

  if (!isSessionTransferPayloadShape(payload)) {
    return { error: "Session export is not a JSON object." };
  }

  if (payload.schema !== SESSION_TRANSFER_SCHEMA) {
    return {
      error: `Unsupported session export (expected ${SESSION_TRANSFER_SCHEMA}).`,
    };
  }

  if (!Array.isArray(payload.sessions)) {
    return { error: "Session export has no sessions array." };
  }

  const sessions: SessionTransferEntry[] = [];
  for (const candidate of payload.sessions) {
    const entry = refineEntry(candidate);
    if (entry) sessions.push(entry);
  }

  if (sessions.length === 0) {
    return { error: "Session export contains no usable sessions." };
  }

  return {
    file: {
      schema: SESSION_TRANSFER_SCHEMA,
      exportedAt: typeof payload.exportedAt === "string" ? payload.exportedAt : "",
      sessions,
    },
  };
}

/** Refine one exported entry; `null` when it cannot be trusted. */
function refineEntry(value: DynamicValue): SessionTransferEntry | null {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return null;
  }

  const entry = value as { data?: DynamicValue; metadata?: DynamicValue };
  const rawData = entry.data;
  if (typeof rawData !== "object" || rawData === null || Array.isArray(rawData)) {
    return null;
  }

  const data = rawData as Partial<SessionData>;
  const id = normalizeId(data.id);
  if (!id || !Array.isArray(data.messages) || !isSessionModelShape(data.model)) {
    return null;
  }

  const createdAt = typeof data.createdAt === "string" ? data.createdAt : "";
  const lastModified =
    typeof data.lastModified === "string" ? data.lastModified : createdAt;
  const metadata = refineMetadata(entry.metadata);

  return {
    data: {
      id,
      title: typeof data.title === "string" ? data.title : "",
      model: data.model,
      thinkingLevel: normalizeThinkingLevel(data.thinkingLevel),
      messages: data.messages,
      createdAt,
      lastModified,
    },
    ...(metadata ? { metadata } : {}),
  };
}

function refineMetadata(value: DynamicValue): SessionMetadata | null {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return null;
  }

  const metadata = value as Partial<SessionMetadata>;
  const id = normalizeId(metadata.id);
  if (!id) return null;

  return {
    id,
    title: typeof metadata.title === "string" ? metadata.title : "",
    createdAt: typeof metadata.createdAt === "string" ? metadata.createdAt : "",
    lastModified:
      typeof metadata.lastModified === "string" ? metadata.lastModified : "",
    messageCount:
      typeof metadata.messageCount === "number" ? metadata.messageCount : 0,
    usage: isSessionUsageShape(metadata.usage) ? metadata.usage : EMPTY_USAGE,
    thinkingLevel: normalizeThinkingLevel(metadata.thinkingLevel),
    preview: typeof metadata.preview === "string" ? metadata.preview : "",
  };
}

/**
 * Import sessions, keeping local data safe:
 * - an id that already exists is imported under a fresh id (never overwritten);
 * - timestamps are clamped to valid ISO strings;
 * - malformed entries are skipped and reported.
 */
export async function importSessions(
  sessions: SessionsStore,
  file: SessionTransferFile,
  opts?: { now?: () => number },
): Promise<ImportSessionsResult> {
  const nowIso = new Date(opts?.now?.() ?? Date.now()).toISOString();
  const result: ImportSessionsResult = { imported: 0, skipped: 0, errors: [] };

  for (const entry of file.sessions) {
    try {
      const originalId = normalizeId(entry.data.id);
      if (!originalId || !Array.isArray(entry.data.messages)) {
        result.skipped += 1;
        continue;
      }

      const exists = (await sessions.getMetadata(originalId)) !== null;
      const id = exists ? crypto.randomUUID() : originalId;

      const createdAt = normalizeIso(entry.data.createdAt, nowIso);
      const lastModified = normalizeIso(entry.data.lastModified, nowIso);
      const importedMetadata = entry.metadata;

      const data: SessionData = {
        ...entry.data,
        id,
        createdAt,
        lastModified,
      };

      const metadata: SessionMetadata = {
        id,
        title:
          typeof importedMetadata?.title === "string"
            ? importedMetadata.title
            : data.title || "",
        createdAt,
        lastModified,
        messageCount: data.messages.length,
        usage: importedMetadata?.usage ?? EMPTY_USAGE,
        thinkingLevel: data.thinkingLevel,
        preview:
          typeof importedMetadata?.preview === "string"
            ? importedMetadata.preview
            : "",
      };

      await sessions.save(data, metadata);
      result.imported += 1;
    } catch (error) {
      result.skipped += 1;
      if (result.errors.length < 10) {
        result.errors.push(error instanceof Error ? error.message : String(error));
      }
    }
  }

  return result;
}
