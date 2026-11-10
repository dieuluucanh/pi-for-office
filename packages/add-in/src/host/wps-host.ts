/** WPS Spreadsheets host implementation. */

import { getWpsEtApplication, type WpsPluginStorage } from "./wps/jsapi.js";
import { settingsBackedSessionStorage } from "./session-storage.js";
import {
  createUnknownWorkbookContext,
  getWorkbookContextFromWpsFullName,
  type WorkbookContext,
} from "./workbook-context.js";
import type {
  SpreadsheetHost,
  SpreadsheetHostReadyCallback,
  SpreadsheetHostReadyInfo,
  SpreadsheetHostSessionStorage,
} from "./types.js";

/** Key used for the document-scoped identity in WPS `PluginStorage`. */
const DOCUMENT_SCOPED_ID_KEY = "pi.workbook.identity.v1";
const DOCUMENT_SCOPED_ID_PREFIX = "doc_uuid:";

export const WPS_UNSUPPORTED_PHASE_1_MESSAGE =
  "WPS Spreadsheets Phase 2 (NEXSELL-370) supports only workbook overview, read_range, write_cells, and execute_wps_js so far; this operation is not yet supported on WPS.";

/**
 * Resolve WPS's per-document plugin storage.
 *
 * WPS exposes it on the ET Application and (depending on version) as a global
 * on `wps`; prefer the application-scoped one.
 */
function getWpsPluginStorage(): WpsPluginStorage | null {
  const app = getWpsEtApplication();
  if (app?.PluginStorage) return app.PluginStorage;

  const globalWps = (
    globalThis as { wps?: { PluginStorage?: WpsPluginStorage } }
  ).wps;
  return globalWps?.PluginStorage ?? null;
}

function readDocumentScopedId(storage: WpsPluginStorage): string | null {
  const getItem = storage.getItem;
  if (typeof getItem !== "function") return null;

  try {
    const value = getItem.call(storage, DOCUMENT_SCOPED_ID_KEY);
    return typeof value === "string" && value.trim().length > 0
      ? value.trim()
      : null;
  } catch {
    return null;
  }
}

function writeDocumentScopedId(
  storage: WpsPluginStorage,
  rawId: string,
): boolean {
  const setItem = storage.setItem;
  if (typeof setItem !== "function") return false;

  try {
    setItem.call(storage, DOCUMENT_SCOPED_ID_KEY, rawId);
    return true;
  } catch {
    return false;
  }
}

export class WpsHost implements SpreadsheetHost {
  readonly kind = "wps";
  readonly displayName = "WPS Spreadsheets";
  readonly sessionStorage: SpreadsheetHostSessionStorage = settingsBackedSessionStorage;

  private documentScopedIdCache: { pathKey: string | null; id: string | null } | null =
    null;

  /** Stable document identity from WPS PluginStorage, else null (path hash fallback). */
  getDocumentScopedId(): Promise<string | null> {
    const pathKey = this.currentPathKey();
    const cached = this.documentScopedIdCache;
    if (cached && cached.pathKey === pathKey) return Promise.resolve(cached.id);

    const storage = getWpsPluginStorage();
    if (!storage) {
      this.documentScopedIdCache = { pathKey, id: null };
      return Promise.resolve(null);
    }

    const existing = readDocumentScopedId(storage);
    if (existing) {
      const id = `${DOCUMENT_SCOPED_ID_PREFIX}${existing}`;
      this.documentScopedIdCache = { pathKey, id };
      return Promise.resolve(id);
    }

    const rawId = crypto.randomUUID();
    const id = `${DOCUMENT_SCOPED_ID_PREFIX}${rawId}`;
    this.documentScopedIdCache = { pathKey, id };
    writeDocumentScopedId(storage, rawId);
    return Promise.resolve(id);
  }

  setDocumentScopedId(id: string): Promise<boolean> {
    const rawId = id.startsWith(DOCUMENT_SCOPED_ID_PREFIX)
      ? id.slice(DOCUMENT_SCOPED_ID_PREFIX.length)
      : id;

    this.documentScopedIdCache = {
      pathKey: this.currentPathKey(),
      id: `${DOCUMENT_SCOPED_ID_PREFIX}${rawId}`,
    };

    const storage = getWpsPluginStorage();
    if (!storage) return Promise.resolve(false);
    return Promise.resolve(writeDocumentScopedId(storage, rawId));
  }

  private currentPathKey(): string | null {
    const workbook = getWpsEtApplication()?.ActiveWorkbook;
    if (!workbook) return null;
    if (typeof workbook.FullName === "string") return workbook.FullName;
    if (typeof workbook.fullName === "string") return workbook.fullName;
    return null;
  }

  whenReady(): Promise<SpreadsheetHostReadyInfo> {
    return Promise.resolve({
      kind: "wps",
      nativeHost: "wps",
      nativePlatform: null,
      reason: "wps-jsapi",
    });
  }

  onReady(callback: SpreadsheetHostReadyCallback): () => void {
    let disposed = false;
    queueMicrotask(() => {
      if (disposed) return;
      callback({
        kind: "wps",
        nativeHost: "wps",
        nativePlatform: null,
        reason: "wps-jsapi",
      });
    });

    return () => {
      disposed = true;
    };
  }

  getWorkbookContext(): Promise<WorkbookContext> {
    const workbook = getWpsEtApplication()?.ActiveWorkbook;
    if (!workbook) {
      return Promise.resolve(createUnknownWorkbookContext());
    }

    const fullName = typeof workbook.FullName === "string"
      ? workbook.FullName
      : typeof workbook.fullName === "string"
      ? workbook.fullName
      : null;
    const workbookName = typeof workbook.Name === "string"
      ? workbook.Name
      : typeof workbook.name === "string"
      ? workbook.name
      : null;

    return getWorkbookContextFromWpsFullName(fullName, workbookName);
  }

  resolveThemeDark(): boolean | null {
    // TODO(NEXSELL-370): map WPS theme APIs if available. Unknown keeps the
    // existing media-query fallback in the taskpane UI.
    return null;
  }
}
