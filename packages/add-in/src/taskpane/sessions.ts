/**
 * Session persistence wiring for one runtime.
 *
 * Owns:
 * - auto-saving agent state to IndexedDB (draft, turn-boundary, periodic, flush)
 * - optional latest-session restore on startup
 * - session identity lifecycle (new / rename / resume)
 * - session ↔ workbook association via the canonical document identity
 *
 * Durability invariants (see docs/session-persistence.md):
 * - a session with content is written at the first user message (draft), at
 *   every turn boundary, and on a bounded interval, not only after the first
 *   assistant reply;
 * - writes are serialized through one chain so a slow IndexedDB write can never
 *   reorder newer state behind older state;
 * - a blank session never overwrites an existing record and never becomes the
 *   workbook's "latest" pointer (that was how restore used to lose sessions);
 * - every failure is surfaced through `getLastSaveError()` instead of only a
 *   `console.warn`.
 */

import type { Models } from "@earendil-works/pi-ai";
import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";
import type { SessionData } from "../storage/local/types.js";
import type { SessionsStore } from "../storage/local/sessions-store.js";
import type { SettingsStore } from "../storage/local/settings-store.js";

import { getCurrentSpreadsheetHost } from "../host/index.js";
import { isDebugEnabled } from "../debug/debug.js";
import { extractTextFromContent } from "../utils/content.js";
import {
  listWorkbookIdentityObservedIds,
  readWorkbookIdentityAlias,
  resolveWorkbookIdentity,
  type WorkbookIdentity,
} from "../workbook/identity.js";

export interface SessionPersistenceController {
  getSessionId: () => string;
  getSessionTitle: () => string;
  getSessionCreatedAt: () => string;
  /** True when the user explicitly named the session via `/name`. */
  hasExplicitTitle: () => boolean;
  startNewSession: () => void;
  renameSession: (title: string) => Promise<void>;
  applyLoadedSession: (sessionData: SessionData) => Promise<void>;
  restoreLatestSession: () => Promise<boolean>;
  saveSession: (opts?: { force?: boolean }) => Promise<void>;
  /**
   * Persist a model selection immediately instead of waiting for the next
   * assistant message (which is otherwise the first save trigger).
   */
  persistModelChange: () => Promise<void>;
  /** Force a durable write of the current state (teardown / visibility change). */
  flush: () => Promise<void>;
  /** Last save failure message, or null when the last save succeeded. */
  getLastSaveError: () => string | null;
  subscribe: (listener: () => void) => () => void;
  dispose: () => void;
}

type SessionId = string;
type PersistedSessionModel = SessionData["model"];

type SaveReason =
  | "agent"
  | "draft"
  | "interval"
  | "rename"
  | "model"
  | "restore"
  | "flush";

/** Identity is cached briefly so a tool-heavy turn doesn't hit host settings. */
const IDENTITY_CACHE_TTL_MS = 5_000;
/** Agent-driven saves are coalesced (a tool-heavy turn emits many events). */
const AGENT_SAVE_DEBOUNCE_MS = 2_000;
/** Periodic dirty-save while a long run is streaming. */
const DIRTY_SAVE_INTERVAL_MS = 15_000;

type UserLikeMessage = AgentMessage & {
  role: "user" | "user-with-attachments";
  content: DynamicValue;
};

/**
 * Re-resolve a persisted model against the current registry so that
 * metadata like `contextWindow` picks up upstream changes (e.g. a dep
 * bump that raised Opus 4.6 from 200k → 1M). Falls back to the
 * persisted model if the runtime no longer has it. Custom and dynamically
 * discovered providers are already part of the same runtime catalogue.
 */
function refreshPersistedModel(
  modelsRuntime: Models,
  persisted: PersistedSessionModel,
): PersistedSessionModel {
  return modelsRuntime.getModel(persisted.provider, persisted.id) ?? persisted;
}

function hasAssistantMessage(messages: AgentMessage[]): boolean {
  return messages.some((m) => m.role === "assistant");
}

function isChatUserMessage(message: AgentMessage): message is UserLikeMessage {
  if (!(message.role === "user" || message.role === "user-with-attachments")) {
    return false;
  }

  return "content" in message;
}

function normalizeSessionId(value: string | null | undefined): string | null {
  if (typeof value !== "string") return null;

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

/** Debug-only trace helper (a single console site, quiet by default). */
function debugLog(message: string, ...args: DynamicValue[]): void {
  if (!isDebugEnabled()) return;
  console.debug(message, ...args);
}

export function shouldPersistSession(opts: {
  firstAssistantSeen: boolean;
  force?: boolean;
}): boolean {
  if (opts.force) return true;
  return opts.firstAssistantSeen;
}

/**
 * What a given agent event should do to persistence.
 *
 * Extracted so the trigger policy is unit-testable without driving a real
 * agent run:
 * - `draft`   — first user message of a never-written session (mid-turn safety);
 * - `agent`   — save soon (debounced) because content changed;
 * - `boundary`— save now (turn boundary);
 * - `boundary_if_dirty` — save at `agent_end` only when something changed.
 */
export type SessionSaveTrigger =
  | "draft"
  | "agent"
  | "boundary"
  | "boundary_if_dirty"
  | "none";

export function resolveSessionSaveTrigger(args: {
  eventType: string;
  messageRole?: string;
  firstAssistantSeen: boolean;
  recordExists: boolean;
}): SessionSaveTrigger {
  if (args.eventType === "message_end") {
    switch (args.messageRole) {
      case "assistant":
        return "agent";
      case "user":
      case "user-with-attachments":
        return args.recordExists ? "none" : "draft";
      case "toolResult":
      case "compactionSummary":
        return args.firstAssistantSeen ? "agent" : "none";
      default:
        return "none";
    }
  }

  if (args.eventType === "turn_end") return "boundary";
  if (args.eventType === "agent_end") return "boundary_if_dirty";
  return "none";
}

/**
 * A model change is worth saving as soon as the session has any stored
 * identity (a record already exists) or content (an assistant reply landed).
 * Never-saved blank tabs keep no record from this path: they are covered by
 * the app-wide remembered model and by the tab-layout force-save on restore.
 */
export function shouldPersistModelChange(opts: {
  firstAssistantSeen: boolean;
  hasExistingRecord: boolean;
}): boolean {
  return opts.firstAssistantSeen || opts.hasExistingRecord;
}

/**
 * Restore candidates for the current runtime.
 *
 * Strict workbook policy: when the workbook identity is known, only the
 * workbook-linked latest session is eligible — there is deliberately no global
 * fallback (a cross-workbook resume must stay an explicit user action in the
 * resume overlay). The caller turns an empty candidate list into a visible
 * "nothing to restore" affordance instead of a silent no-op.
 */
export function getRestoreCandidateSessionIds(args: {
  workbookId: string | null;
  workbookLatestSessionId: string | null;
  globalLatestSessionId: string | null;
}): string[] {
  const candidates: string[] = [];

  const pushUnique = (sessionId: string | null) => {
    if (!sessionId || candidates.includes(sessionId)) return;
    candidates.push(sessionId);
  };

  const workbookLatest = normalizeSessionId(args.workbookLatestSessionId);
  const globalLatest = normalizeSessionId(args.globalLatestSessionId);

  if (args.workbookId) {
    // Workbook is known: restore only the workbook-linked latest session.
    pushUnique(workbookLatest);
    return candidates;
  }

  // Workbook identity unavailable: fall back to global latest behavior.
  pushUnique(globalLatest);
  return candidates;
}

export async function setupSessionPersistence(opts: {
  agent: Agent;
  sessions: SessionsStore;
  settings: SettingsStore;
  models: Models;
  initialSessionId?: string;
  autoRestoreLatest?: boolean;
}): Promise<SessionPersistenceController> {
  const { agent, sessions, settings } = opts;
  const spreadsheetHost = getCurrentSpreadsheetHost();

  const listeners = new Set<() => void>();
  const initialSessionId = normalizeSessionId(opts.initialSessionId ?? null);
  let sessionId: SessionId = initialSessionId ?? crypto.randomUUID();
  let sessionTitle = "";
  let sessionCreatedAt = new Date().toISOString();
  let firstAssistantSeen = false;
  let explicitTitle = false;

  /** Set by agent events; gates the periodic save and is cleared on success. */
  let dirty = false;
  /** True once this session id is known to have a stored record. */
  let recordExists = false;
  let lastSaveError: string | null = null;

  let saveChain: Promise<void> = Promise.resolve();
  let agentSaveTimer: ReturnType<typeof setTimeout> | null = null;
  let dirtySaveInterval: ReturnType<typeof setInterval> | null = null;

  /** Cached canonical identity (see `resolveWorkbookIdentity`). */
  let identityCache: WorkbookIdentity | null = null;
  let identityCachedAt = 0;
  /** Observed ids whose session links were already migrated in this runtime. */
  const migratedAliasIds = new Set<string>();
  /** Sessions saved while no workbook identity was resolvable. */
  const pendingAssociationSessionIds = new Set<string>();
  let pendingLatestSessionId: string | null = null;

  agent.sessionId = sessionId;

  function emitChange(): void {
    for (const listener of listeners) {
      listener();
    }
  }

  async function resolveIdentity(force = false): Promise<WorkbookIdentity> {
    const now = Date.now();
    if (
      !force &&
      identityCache &&
      now - identityCachedAt < IDENTITY_CACHE_TTL_MS
    ) {
      return identityCache;
    }

    identityCache = await resolveWorkbookIdentity({
      host: spreadsheetHost,
      settings,
    });
    identityCachedAt = now;
    return identityCache;
  }

  /**
   * Canonical identity for association, with a one-time migration of session
   * links that were written under an observed (URL/path hash) id.
   */
  async function resolveAssociationWorkbookId(
    force = false,
  ): Promise<string | null> {
    const identity = await resolveIdentity(force);
    const { canonicalWorkbookId, aliasedFrom } = identity;

    if (canonicalWorkbookId && aliasedFrom && !migratedAliasIds.has(aliasedFrom)) {
      migratedAliasIds.add(aliasedFrom);
      try {
        const observedIds = new Set(
          await listWorkbookIdentityObservedIds(settings, canonicalWorkbookId),
        );
        observedIds.add(aliasedFrom);

        await spreadsheetHost.sessionStorage.migrateAliasesToCanonical(settings, {
          canonicalId: canonicalWorkbookId,
          observedIds: [...observedIds],
          resolveSessionLastModified: async (sessionIdToCheck: string) =>
            (await sessions.getMetadata(sessionIdToCheck))?.lastModified ?? null,
        });
      } catch (error) {
        console.warn("[pi] Session link migration failed:", error);
      }
    }

    return canonicalWorkbookId;
  }

  /** Link sessions that were saved before the workbook identity existed. */
  async function flushPendingAssociations(workbookId: string): Promise<void> {
    if (pendingAssociationSessionIds.size === 0) {
      pendingLatestSessionId = null;
      return;
    }

    for (const pendingId of [...pendingAssociationSessionIds]) {
      try {
        await spreadsheetHost.sessionStorage.relinkSessionToWorkbook(
          settings,
          pendingId,
          workbookId,
        );
        await spreadsheetHost.sessionStorage.clearSessionWorkbookPending(
          settings,
          pendingId,
        );
        pendingAssociationSessionIds.delete(pendingId);
      } catch (error) {
        console.warn("[pi] Pending session association failed:", error);
      }
    }

    const latest = pendingLatestSessionId;
    if (latest && !pendingAssociationSessionIds.has(latest)) {
      try {
        await spreadsheetHost.sessionStorage.setLatestSessionForWorkbook(
          settings,
          workbookId,
          latest,
        );
        pendingLatestSessionId = null;
      } catch (error) {
        console.warn("[pi] Pending latest-session pointer failed:", error);
      }
    }
  }

  async function updateWorkbookAssociation(
    savedSessionId: string,
    args: { isEmpty: boolean },
  ): Promise<string | null> {
    // If an earlier save had no identity, force a refresh: the document may
    // have been saved (or Office may have become ready) since then, and the
    // pending sessions should be linked on the very next write.
    const refreshIdentity = pendingAssociationSessionIds.size > 0;
    const identity = await resolveIdentity(refreshIdentity);
    const workbookId = await resolveAssociationWorkbookId(refreshIdentity);

    if (!workbookId) {
      // Identity unavailable (unsaved document, web guest, read-only copy).
      // Remember it so the next successful resolution links it — no guessing
      // across restarts from a workbook name.
      pendingAssociationSessionIds.add(savedSessionId);
      if (!args.isEmpty) pendingLatestSessionId = savedSessionId;
      await spreadsheetHost.sessionStorage.markSessionWorkbookPending(
        settings,
        savedSessionId,
        { workbookName: identity.workbookName, at: new Date().toISOString() },
      );
      return null;
    }

    try {
      await flushPendingAssociations(workbookId);

      const existing = await spreadsheetHost.sessionStorage.getSessionWorkbookId(
        settings,
        savedSessionId,
      );

      if (!existing) {
        await spreadsheetHost.sessionStorage.linkSessionToWorkbook(
          settings,
          savedSessionId,
          workbookId,
        );
      } else if (existing !== workbookId) {
        // Move the link only when the old id is an alias of this document.
        const resolvedExisting = await readWorkbookIdentityAlias(
          settings,
          existing,
        );
        if (resolvedExisting === workbookId) {
          await spreadsheetHost.sessionStorage.relinkSessionToWorkbook(
            settings,
            savedSessionId,
            workbookId,
          );
        }
      }

      if (!args.isEmpty) {
        await spreadsheetHost.sessionStorage.setLatestSessionForWorkbook(
          settings,
          workbookId,
          savedSessionId,
        );
      }

      await spreadsheetHost.sessionStorage.clearSessionWorkbookPending(
        settings,
        savedSessionId,
      );
      return workbookId;
    } catch (err) {
      console.warn("[pi] Workbook/session association update failed:", err);
      return null;
    }
  }

  function buildPreview(messages: AgentMessage[]): string {
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
      if (preview.length > 2048) {
        return preview.slice(0, 2048);
      }
    }
    return preview;
  }

  function aggregateUsage(messages: AgentMessage[]) {
    let inputTokens = 0;
    let outputTokens = 0;
    let cacheReadTokens = 0;
    let cacheWriteTokens = 0;
    let totalTokens = 0;

    let costInput = 0;
    let costOutput = 0;
    let costCacheRead = 0;
    let costCacheWrite = 0;
    let costTotal = 0;

    for (const message of messages) {
      if (message.role !== "assistant") continue;
      const usage = message.usage;

      inputTokens += usage.input;
      outputTokens += usage.output;
      cacheReadTokens += usage.cacheRead;
      cacheWriteTokens += usage.cacheWrite;
      totalTokens += usage.totalTokens;

      costInput += usage.cost.input;
      costOutput += usage.cost.output;
      costCacheRead += usage.cost.cacheRead;
      costCacheWrite += usage.cost.cacheWrite;
      costTotal += usage.cost.total;
    }

    return {
      input: inputTokens,
      output: outputTokens,
      cacheRead: cacheReadTokens,
      cacheWrite: cacheWriteTokens,
      totalTokens,
      cost: {
        input: costInput,
        output: costOutput,
        cacheRead: costCacheRead,
        cacheWrite: costCacheWrite,
        total: costTotal,
      },
    };
  }

  async function hasStoredRecord(): Promise<boolean> {
    if (recordExists) return true;
    try {
      const metadata = await sessions.getMetadata(sessionId);
      recordExists = metadata !== null;
      return recordExists;
    } catch {
      return false;
    }
  }

  async function runSave(optsForSave?: {
    force?: boolean;
    reason?: SaveReason;
  }): Promise<void> {
    const reason = optsForSave?.reason ?? "agent";
    const force = optsForSave?.force === true;

    if (reason === "interval" && !dirty) return;

    const messages = agent.state.messages;
    const isEmpty = messages.length === 0;

    // A blank session is never auto-saved...
    if (isEmpty && !force) return;
    // ...and a forced blank save may create a durable tab placeholder but must
    // never overwrite an existing record.
    if (isEmpty && (await hasStoredRecord())) return;

    if (
      !isEmpty &&
      !shouldPersistSession({
        firstAssistantSeen,
        ...(force ? { force: true } : {}),
      })
    ) {
      return;
    }

    const startedAt = Date.now();

    try {
      const now = new Date().toISOString();

      if (!sessionTitle && messages.length > 0) {
        const firstUser = messages.find((m) => isChatUserMessage(m));
        if (firstUser) {
          const text = extractTextFromContent(firstUser.content);
          sessionTitle = text.slice(0, 80) || "Untitled";
        }
      }

      const savedSessionId = sessionId;

      await sessions.saveSession(
        savedSessionId,
        agent.state,
        {
          id: savedSessionId,
          title: sessionTitle,
          createdAt: sessionCreatedAt,
          lastModified: now,
          messageCount: messages.length,
          usage: aggregateUsage(messages),
          thinkingLevel: agent.state.thinkingLevel || "off",
          preview: buildPreview(messages),
        },
        sessionTitle,
      );

      recordExists = true;
      dirty = false;
      lastSaveError = null;

      const linkedWorkbookId = await updateWorkbookAssociation(savedSessionId, {
        isEmpty,
      });
      emitChange();

      if (isDebugEnabled()) {
        debugLog(
          `[pi] session save id=${savedSessionId} msgs=${messages.length} ` +
            `reason=${reason} empty=${isEmpty} ` +
            `linked=${linkedWorkbookId ?? "none"} ` +
            `title=${JSON.stringify(sessionTitle)} took=${Date.now() - startedAt}ms`,
        );
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      lastSaveError = message;
      console.warn("[pi] Session save failed:", err);
      emitChange();
      throw err;
    }
  }

  function enqueueSave(run: () => Promise<void>): Promise<void> {
    const next = saveChain.then(run, run).catch((error: DynamicValue) => {
      // `runSave` already recorded the failure; keep the chain alive.
      debugLog("[pi] session save chain error:", error);
    });
    saveChain = next;
    return next;
  }

  function saveSession(optsForSave?: {
    force?: boolean;
    reason?: SaveReason;
  }): Promise<void> {
    return enqueueSave(() => runSave(optsForSave));
  }

  function scheduleAgentSave(): void {
    if (agentSaveTimer !== null) return;
    agentSaveTimer = setTimeout(() => {
      agentSaveTimer = null;
      void saveSession({ reason: "agent" });
    }, AGENT_SAVE_DEBOUNCE_MS);
  }

  /**
   * Save a model change without waiting for the next assistant message (the
   * usual first save trigger). Only sessions that already have an identity or
   * content are written; never-saved blank tabs are covered by the app-wide
   * remembered model instead of an empty session record here.
   *
   * Storage-lookup errors propagate so callers own the failure policy.
   */
  async function persistModelChange(): Promise<void> {
    if (!firstAssistantSeen) {
      const hasExistingRecord = await hasStoredRecord();
      if (
        !shouldPersistModelChange({ firstAssistantSeen, hasExistingRecord })
      ) {
        return;
      }
    }

    await saveSession({ force: true, reason: "model" });
  }

  function startNewSession(): void {
    sessionId = crypto.randomUUID();
    sessionTitle = "";
    sessionCreatedAt = new Date().toISOString();
    firstAssistantSeen = false;
    explicitTitle = false;
    dirty = false;
    recordExists = false;
    lastSaveError = null;
    agent.sessionId = sessionId;
    emitChange();
  }

  async function renameSession(title: string): Promise<void> {
    sessionTitle = title.trim();
    explicitTitle = sessionTitle.length > 0;
    emitChange();
    await saveSession({ reason: "rename" });
  }

  async function applyLoadedSession(sessionData: SessionData): Promise<void> {
    // Never remap a stored id: remapping forked the session into a new record
    // on the next save, which is what produced duplicate/ghost history entries.
    sessionId = normalizeSessionId(sessionData.id) ?? crypto.randomUUID();

    sessionTitle = sessionData.title || "";
    explicitTitle = false;
    sessionCreatedAt = sessionData.createdAt;
    firstAssistantSeen = hasAssistantMessage(sessionData.messages);
    recordExists = true;
    dirty = false;
    lastSaveError = null;

    agent.sessionId = sessionId;
    agent.state.messages = sessionData.messages;

    if (sessionData.model) {
      agent.state.model = refreshPersistedModel(opts.models, sessionData.model);
    }
    if (sessionData.thinkingLevel) {
      agent.state.thinkingLevel = sessionData.thinkingLevel;
    }

    await updateWorkbookAssociation(sessionId, {
      isEmpty: sessionData.messages.length === 0,
    });
    emitChange();
  }

  async function restoreLatestSession(): Promise<boolean> {
    try {
      const workbookId = await resolveAssociationWorkbookId(true);
      const workbookLatest = workbookId
        ? await spreadsheetHost.sessionStorage.getLatestSessionForWorkbook(
            settings,
            workbookId,
          )
        : null;
      const globalLatest = workbookId
        ? null
        : await sessions.getLatestSessionId();

      const candidates = getRestoreCandidateSessionIds({
        workbookId,
        workbookLatestSessionId: workbookLatest,
        globalLatestSessionId: globalLatest,
      });

      for (const candidateId of candidates) {
        const sessionData = await sessions.loadSession(candidateId);
        if (!sessionData || sessionData.messages.length === 0) continue;

        await applyLoadedSession(sessionData);
        debugLog(
          `[pi] session restore candidates=${candidates.join(",") || "none"} ` +
            `chosen=${candidateId} msgs=${sessionData.messages.length}`,
        );
        debugLog(
          `[pi] Restored session: ${sessionData.title || candidateId}`,
        );
        return true;
      }

      debugLog(
        `[pi] session restore candidates=${candidates.join(",") || "none"} ` +
          `chosen=none (no restorable candidate)`,
      );
    } catch (err) {
      console.warn("[pi] Session restore failed:", err);
    }

    return false;
  }

  const unsubscribeAgent = agent.subscribe((event) => {
    const messageRole =
      event.type === "message_end" ? event.message.role : undefined;

    const trigger = resolveSessionSaveTrigger({
      eventType: event.type,
      ...(messageRole !== undefined ? { messageRole } : {}),
      firstAssistantSeen,
      recordExists,
    });

    if (trigger === "none") return;

    dirty = true;

    if (event.type === "message_end" && event.message.role === "assistant") {
      firstAssistantSeen = true;
    }

    if (trigger === "draft") {
      // A first turn can run for minutes before the first assistant message
      // lands; without this, closing Excel mid-run lost the whole conversation.
      void saveSession({ force: true, reason: "draft" });
      return;
    }

    if (trigger === "agent") {
      scheduleAgentSave();
      return;
    }

    void saveSession({ reason: "agent" });
  });

  dirtySaveInterval = setInterval(() => {
    if (!dirty) return;
    if (!recordExists && !firstAssistantSeen) return;
    void saveSession({ reason: "interval" });
  }, DIRTY_SAVE_INTERVAL_MS);

  if (opts.autoRestoreLatest) {
    await restoreLatestSession();
  }

  return {
    getSessionId: () => sessionId,
    getSessionTitle: () => sessionTitle,
    getSessionCreatedAt: () => sessionCreatedAt,
    hasExplicitTitle: () => explicitTitle,
    startNewSession,
    renameSession,
    applyLoadedSession,
    restoreLatestSession,
    saveSession,
    persistModelChange,
    flush: () => enqueueSave(() => runSave({ force: true, reason: "flush" })),
    getLastSaveError: () => lastSaveError,
    subscribe(listener: () => void): () => void {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    dispose(): void {
      unsubscribeAgent();

      if (agentSaveTimer !== null) {
        clearTimeout(agentSaveTimer);
        agentSaveTimer = null;
      }
      if (dirtySaveInterval !== null) {
        clearInterval(dirtySaveInterval);
        dirtySaveInterval = null;
      }

      listeners.clear();
    },
  };
}
