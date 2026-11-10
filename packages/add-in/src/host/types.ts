/** Host abstraction types for Office, WPS, and browser/UI-test runtimes. */

import type { SettingsStore } from "../storage/local/settings-store.js";

import type { SessionWorkbookPartition } from "../workbook/session-association.js";
import type { WorkbookContext } from "./workbook-context.js";

export type SpreadsheetHostKind = "office" | "wps" | "browser";

export interface SpreadsheetHostReadyInfo {
  kind: SpreadsheetHostKind;
  nativeHost: string | null;
  nativePlatform: string | null;
  reason: "office-ready" | "office-unavailable" | "office-timeout" | "wps-jsapi" | "browser";
}

export type SpreadsheetHostReadyCallback = (info: SpreadsheetHostReadyInfo) => void;

export interface SpreadsheetHostSessionStorage {
  getSessionWorkbookId(settings: SettingsStore, sessionId: string): Promise<string | null>;
  linkSessionToWorkbook(settings: SettingsStore, sessionId: string, workbookId: string): Promise<void>;
  /** Correct an existing link (alias migration / explicit user action). */
  relinkSessionToWorkbook(settings: SettingsStore, sessionId: string, workbookId: string): Promise<void>;
  /** Remove a link without deleting the session. */
  forgetSessionWorkbook(settings: SettingsStore, sessionId: string): Promise<void>;
  setLatestSessionForWorkbook(settings: SettingsStore, workbookId: string, sessionId: string): Promise<void>;
  getLatestSessionForWorkbook(settings: SettingsStore, workbookId: string): Promise<string | null>;
  /** Batched link read for history rendering (one settings scan). */
  readSessionWorkbookLinks(settings: SettingsStore): Promise<Map<string, string>>;
  migrateAliasesToCanonical(
    settings: SettingsStore,
    args: {
      canonicalId: string;
      observedIds: string[];
      resolveSessionLastModified?: (sessionId: string) => Promise<string | null>;
    },
  ): Promise<{ relinked: number; latestMoved: boolean }>;
  markSessionWorkbookPending(
    settings: SettingsStore,
    sessionId: string,
    args: { workbookName: string | null; at: string },
  ): Promise<void>;
  clearSessionWorkbookPending(settings: SettingsStore, sessionId: string): Promise<void>;
  partitionSessionIdsByWorkbook(
    settings: SettingsStore,
    sessionIds: string[],
    workbookId: string,
  ): Promise<SessionWorkbookPartition>;
}

export interface SpreadsheetHost {
  readonly kind: SpreadsheetHostKind;
  readonly displayName: string;
  readonly sessionStorage: SpreadsheetHostSessionStorage;

  /** Resolve when the host is ready enough for the taskpane to initialize. */
  whenReady(): Promise<SpreadsheetHostReadyInfo>;

  /** Subscribe to a host-ready signal where the host supports one. */
  onReady(callback: SpreadsheetHostReadyCallback): () => void;

  /** Privacy-preserving workbook identity. Never expose or persist raw document URLs. */
  getWorkbookContext(): Promise<WorkbookContext>;

  /**
   * A stable, document-scoped identifier (a UUID the add-in stores *inside*
   * the document), or null when the host cannot provide one.
   *
   * This is the canonical workbook identity for session association: unlike a
   * hash of the document URL/path it survives rename/move, unsaved documents,
   * guest access in Excel on the web, and platform switches (desktop path vs
   * cloud URL).
   */
  getDocumentScopedId(): Promise<string | null>;

  /**
   * Persist a document-scoped identifier inside the document.
   * Resolves false when the host cannot store it durably (read-only,
   * protected view, unsaved document); the value is still usable in-session.
   */
  setDocumentScopedId(id: string): Promise<boolean>;

  /** Host theme signal, if available. Returns null when unknown/unsupported. */
  resolveThemeDark(): boolean | null;
}
