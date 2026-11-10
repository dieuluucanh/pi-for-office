/**
 * Store for chat sessions (data and metadata).
 * Uses two object stores: sessions (full data) and sessions-metadata (lightweight).
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 *
 * Listing invariant: the history UI must never be able to "lose" a stored
 * session. `getAllMetadataSafe()` therefore never throws and never trusts the
 * `lastModified` index alone — it unions the index, a full record scan, and
 * metadata-only rows, so a missing index, a row written before the index
 * existed, or a half-completed transaction can only degrade information, never
 * hide a session.
 */

import type { AgentState } from "@earendil-works/pi-agent-core";
import type { Api, Model } from "@earendil-works/pi-ai";

import { Store } from "./store.js";
import type { SessionData, SessionMetadata, StoreConfig } from "./types.js";
import { extractTextFromContent } from "../../utils/content.js";

/** One row of the session history list. */
export interface SessionListEntry {
  metadata: SessionMetadata;
  /** Metadata exists but the full record is gone (or was never written). */
  missingData?: boolean;
  /** True when the ordered index was unavailable and records were scanned. */
  degraded?: boolean;
}

const EMPTY_USAGE: SessionMetadata["usage"] = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
};

const PREVIEW_LIMIT = 2048;

function parseTimestamp(value: string | null | undefined): number {
  if (typeof value !== "string") return 0;
  const parsed = Date.parse(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeIso(value: string | null | undefined, fallback: number): string {
  const parsed = parseTimestamp(value);
  return new Date(parsed > 0 ? parsed : fallback).toISOString();
}

function buildPreview(messages: SessionData["messages"]): string {
  let preview = "";
  for (const message of messages) {
    let text = "";
    if (message.role === "compactionSummary") {
      text = message.summary;
    } else if (message.role === "user" || message.role === "assistant") {
      text = extractTextFromContent(message.content);
    } else {
      continue;
    }
    preview += text + "\n";
    if (preview.length > PREVIEW_LIMIT) return preview.slice(0, PREVIEW_LIMIT);
  }
  return preview;
}

function aggregateUsage(messages: SessionData["messages"]): SessionMetadata["usage"] {
  const usage: SessionMetadata["usage"] = {
    ...EMPTY_USAGE,
    cost: { ...EMPTY_USAGE.cost },
  };

  for (const message of messages) {
    if (message.role !== "assistant") continue;
    const messageUsage = message.usage;

    usage.input += messageUsage.input;
    usage.output += messageUsage.output;
    usage.cacheRead += messageUsage.cacheRead;
    usage.cacheWrite += messageUsage.cacheWrite;
    usage.totalTokens += messageUsage.totalTokens;

    usage.cost.input += messageUsage.cost.input;
    usage.cost.output += messageUsage.cost.output;
    usage.cost.cacheRead += messageUsage.cost.cacheRead;
    usage.cost.cacheWrite += messageUsage.cost.cacheWrite;
    usage.cost.total += messageUsage.cost.total;
  }

  return usage;
}

/** Derive list metadata from a full record (fallback / repair path). */
export function deriveSessionMetadata(
  record: SessionData,
  now = Date.now(),
): SessionMetadata {
  const firstUser = record.messages.find(
    (message) => message.role === "user" || message.role === "user-with-attachments",
  );
  const firstUserText =
    firstUser && "content" in firstUser
      ? extractTextFromContent(firstUser.content).slice(0, 80)
      : null;

  // Ordering falls back to `createdAt` (meaningful) before `now` (arbitrary),
  // so an unindexed legacy row keeps its real position in the list.
  const createdFallback = parseTimestamp(record.createdAt) || now;

  return {
    id: record.id,
    title: record.title || firstUserText || "",
    createdAt: normalizeIso(record.createdAt, now),
    lastModified: normalizeIso(record.lastModified, createdFallback),
    messageCount: record.messages.length,
    usage: aggregateUsage(record.messages),
    thinkingLevel: record.thinkingLevel,
    preview: buildPreview(record.messages),
  };
}

/** Repair a metadata row whose timestamps predate validation. */
function normalizeMetadata(
  metadata: SessionMetadata,
  now = Date.now(),
): SessionMetadata {
  return {
    ...metadata,
    createdAt: normalizeIso(metadata.createdAt, now),
    lastModified: normalizeIso(metadata.lastModified, now),
    messageCount:
      typeof metadata.messageCount === "number" && metadata.messageCount >= 0
        ? metadata.messageCount
        : 0,
    usage:
      typeof metadata.usage === "object" && metadata.usage !== null
        ? metadata.usage
        : EMPTY_USAGE,
    preview: typeof metadata.preview === "string" ? metadata.preview : "",
  };
}

export class SessionsStore extends Store {
  getConfig(): StoreConfig {
    return {
      name: "sessions",
      keyPath: "id",
      indices: [{ name: "lastModified", keyPath: "lastModified" }],
    };
  }

  /**
   * Additional config for sessions-metadata store.
   * Must be included when creating the backend.
   */
  static getMetadataConfig(): StoreConfig {
    return {
      name: "sessions-metadata",
      keyPath: "id",
      indices: [{ name: "lastModified", keyPath: "lastModified" }],
    };
  }

  async save(data: SessionData, metadata: SessionMetadata): Promise<void> {
    await this.getBackend().transaction(
      ["sessions", "sessions-metadata"],
      "readwrite",
      async (tx) => {
        await tx.set("sessions", data.id, data);
        await tx.set("sessions-metadata", metadata.id, metadata);
      },
    );
  }

  async get(id: string): Promise<SessionData | null> {
    return this.getBackend().get<SessionData>("sessions", id);
  }

  async getMetadata(id: string): Promise<SessionMetadata | null> {
    return this.getBackend().get<SessionMetadata>("sessions-metadata", id);
  }

  /**
   * All sessions, newest first, without ever throwing.
   *
   * The union covers the three ways a session can go missing from the naive
   * index read: no index (older database), a row without `lastModified`
   * (not indexed at all), and metadata without a full record.
   */
  async getAllMetadataSafe(): Promise<SessionListEntry[]> {
    const backend = this.getBackend();
    const now = Date.now();

    let indexed: SessionMetadata[] | null = null;
    try {
      indexed = await backend.getAllFromIndex<SessionMetadata>(
        "sessions-metadata",
        "lastModified",
        "desc",
      );
    } catch (error) {
      console.warn(
        "[pi] Session metadata index unavailable — scanning full records:",
        error,
      );
    }

    let records: SessionData[] = [];
    try {
      records = await backend.getAll<SessionData>("sessions");
    } catch (error) {
      console.warn("[pi] Session record scan failed:", error);
    }

    const entries = new Map<string, SessionListEntry>();

    for (const metadata of indexed ?? []) {
      entries.set(metadata.id, {
        metadata: normalizeMetadata(metadata, now),
      });
    }

    const recordIds = new Set<string>();
    for (const record of records) {
      if (!record || typeof record.id !== "string") continue;
      recordIds.add(record.id);

      const existing = entries.get(record.id);
      if (existing) {
        // A metadata row without a usable timestamp can be older than the
        // record it describes; keep the record's timestamps as the source of
        // truth for ordering.
        if (parseTimestamp(existing.metadata.lastModified) === 0) {
          existing.metadata = deriveSessionMetadata(record, now);
        }
        continue;
      }

      entries.set(record.id, {
        metadata: deriveSessionMetadata(record, now),
        ...(indexed === null ? { degraded: true } : {}),
      });
    }

    for (const [id, entry] of entries) {
      if (!recordIds.has(id)) entry.missingData = true;
    }

    return [...entries.values()].sort(
      (left, right) =>
        parseTimestamp(right.metadata.lastModified) -
        parseTimestamp(left.metadata.lastModified),
    );
  }

  async getAllMetadata(): Promise<SessionMetadata[]> {
    const entries = await this.getAllMetadataSafe();
    return entries.map((entry) => entry.metadata);
  }

  async delete(id: string): Promise<void> {
    await this.getBackend().transaction(
      ["sessions", "sessions-metadata"],
      "readwrite",
      async (tx) => {
        await tx.delete("sessions", id);
        await tx.delete("sessions-metadata", id);
      },
    );
  }

  // Alias for backward compatibility
  async deleteSession(id: string): Promise<void> {
    return this.delete(id);
  }

  async updateTitle(id: string, title: string): Promise<void> {
    const metadata = await this.getMetadata(id);
    if (metadata) {
      metadata.title = title;
      await this.getBackend().set("sessions-metadata", id, metadata);
    }
    // Also update in full session data
    const data = await this.get(id);
    if (data) {
      data.title = title;
      await this.getBackend().set("sessions", id, data);
    }
  }

  async getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> {
    return this.getBackend().getQuotaInfo();
  }

  async requestPersistence(): Promise<boolean> {
    return this.getBackend().requestPersistence();
  }

  // Alias methods for backward compatibility
  async saveSession(
    id: string,
    state: AgentState,
    metadata: SessionMetadata | undefined,
    title?: string,
  ): Promise<void> {
    // If metadata is provided, use it; otherwise create it from state
    const meta: SessionMetadata = metadata ?? {
      id,
      title: title ?? "",
      createdAt: new Date().toISOString(),
      lastModified: new Date().toISOString(),
      messageCount: state.messages.length,
      usage: {
        input: 0,
        output: 0,
        cacheRead: 0,
        cacheWrite: 0,
        totalTokens: 0,
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
      },
      thinkingLevel: state.thinkingLevel,
      preview: "",
    };
    // AgentState.model is typed Model<any> upstream; every model that enters
    // our runtime comes from pi-ai's typed registry or a parsed custom-provider
    // boundary, so this narrows an upstream `any` back to the concrete Api union.
    const model = state.model as Model<Api>;
    const data: SessionData = {
      id,
      title: title ?? meta.title,
      model,
      thinkingLevel: state.thinkingLevel,
      messages: state.messages,
      createdAt: meta.createdAt,
      lastModified: new Date().toISOString(),
    };
    await this.save(data, meta);
  }

  async loadSession(id: string): Promise<SessionData | null> {
    return this.get(id);
  }

  async getLatestSessionId(): Promise<string | null> {
    const allMetadata = await this.getAllMetadata();
    if (allMetadata.length === 0) return null;

    const sorted = [...allMetadata].sort(
      (left, right) =>
        parseTimestamp(right.lastModified) - parseTimestamp(left.lastModified),
    );
    return sorted[0]?.id ?? null;
  }
}
