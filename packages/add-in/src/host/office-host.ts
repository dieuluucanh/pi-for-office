function isHostOfficeHostPayloadShape(
  value: DynamicValue,
): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Office.js-backed host implementation. */

import { settingsBackedSessionStorage } from "./session-storage.js";
import { resolveOfficeThemeDark } from "./office-theme.js";
import {
  detectOfficeAppFromGlobals,
  officeAppLabel,
  type OfficeApp,
} from "./app.js";
import {
  getWorkbookContextFromDocumentUrl,
  type WorkbookContext,
} from "./workbook-context.js";
import type {
  SpreadsheetHost,
  SpreadsheetHostReadyCallback,
  SpreadsheetHostReadyInfo,
  SpreadsheetHostSessionStorage,
} from "./types.js";

/** Setting name used to store the add-in's stable document identity. */
const DOCUMENT_SCOPED_ID_SETTING = "pi.workbook.identity.v1";
/** Prefixes a document-scoped id so it can never be confused with a URL hash. */
export const DOCUMENT_SCOPED_ID_PREFIX = "doc_uuid:";
/** Upper bound on `Office.Settings.saveAsync`; identity must never block boot. */
const SETTINGS_SAVE_TIMEOUT_MS = 2000;

interface OfficeSettingsLike {
  get(name: string): DynamicValue;
  set(name: string, value: DynamicValue): void;
  saveAsync: ((callback?: (result: DynamicValue) => void) => void) | null;
}

function asOfficeSettings(value: DynamicObject): OfficeSettingsLike | null {
  const rawGet = value.get;
  const rawSet = value.set;
  if (typeof rawGet !== "function" || typeof rawSet !== "function") return null;

  const rawSave = value.saveAsync;

  return {
    get: (name: string): DynamicValue => rawGet.call(value, name) as DynamicValue,
    set: (name: string, next: DynamicValue): void => {
      rawSet.call(value, name, next);
    },
    saveAsync:
      typeof rawSave === "function"
        ? (callback?: (result: DynamicValue) => void): void => {
            rawSave.call(value, callback);
          }
        : null,
  };
}

function getOfficeDocumentSettings(): OfficeSettingsLike | null {
  try {
    const office = (globalThis as { Office?: DynamicValue }).Office;
    if (!isHostOfficeHostPayloadShape(office)) return null;

    const ctx = office.context;
    if (!isHostOfficeHostPayloadShape(ctx)) return null;

    const doc = ctx.document;
    if (!isHostOfficeHostPayloadShape(doc)) return null;

    const settings = doc.settings;
    if (!isHostOfficeHostPayloadShape(settings)) return null;

    return asOfficeSettings(settings);
  } catch {
    return null;
  }
}

function readDocumentScopedId(settings: OfficeSettingsLike): string | null {
  try {
    const value = settings.get(DOCUMENT_SCOPED_ID_SETTING);
    return typeof value === "string" && value.trim().length > 0
      ? value.trim()
      : null;
  } catch {
    return null;
  }
}

/**
 * Write the identity into the document itself (best effort).
 *
 * Resolves false when the host cannot persist it durably (read-only document,
 * protected view, missing `saveAsync`); callers still keep the value for the
 * current pane lifetime.
 */
async function persistDocumentScopedId(
  settings: OfficeSettingsLike,
  rawId: string,
): Promise<boolean> {
  try {
    settings.set(DOCUMENT_SCOPED_ID_SETTING, rawId);
  } catch {
    return false;
  }

  const saveAsync = settings.saveAsync;
  if (!saveAsync) return false;

  return await new Promise<boolean>((resolve) => {
    let settled = false;
    const finish = (ok: boolean) => {
      if (settled) return;
      settled = true;
      resolve(ok);
    };

    try {
      saveAsync((result: DynamicValue) => {
        const status = isHostOfficeHostPayloadShape(result)
          ? result.status
          : undefined;
        finish(status === "succeeded" || status === undefined);
      });
    } catch {
      finish(false);
      return;
    }

    // Never let a silently-dropped callback hold identity resolution open.
    globalThis.setTimeout(() => finish(false), SETTINGS_SAVE_TIMEOUT_MS);
  });
}

function toDocumentScopedId(rawId: string): string {
  return `${DOCUMENT_SCOPED_ID_PREFIX}${rawId}`;
}

function fromDocumentScopedId(id: string): string {
  return id.startsWith(DOCUMENT_SCOPED_ID_PREFIX)
    ? id.slice(DOCUMENT_SCOPED_ID_PREFIX.length)
    : id;
}

function getOfficeDocumentUrl(): string | null {
  try {
    const office = (globalThis as { Office?: DynamicValue }).Office;
    if (!isHostOfficeHostPayloadShape(office)) return null;

    const ctx = office.context;
    if (!isHostOfficeHostPayloadShape(ctx)) return null;

    const doc = ctx.document;
    if (!isHostOfficeHostPayloadShape(doc)) return null;

    const url = doc.url;
    return typeof url === "string" && url.trim().length > 0 ? url : null;
  } catch {
    return null;
  }
}

function nativeValueToString(value: DynamicValue): string | null {
  if (typeof value === "string" && value.trim().length > 0) {
    return value;
  }

  return null;
}

function toError(error: DynamicValue): Error {
  return error instanceof Error ? error : new Error("Office.onReady failed.");
}

interface OfficeReadyInfoLike {
  host?: DynamicValue;
  platform?: DynamicValue;
}

function fromOfficeReadyInfo(
  info: OfficeReadyInfoLike,
): SpreadsheetHostReadyInfo {
  return {
    kind: "office",
    nativeHost: nativeValueToString(info.host),
    nativePlatform: nativeValueToString(info.platform),
    reason: "office-ready",
  };
}

export class OfficeHost implements SpreadsheetHost {
  readonly kind = "office";
  /** "Microsoft Excel" / "Microsoft Word" / "Microsoft PowerPoint", per detected host. */
  readonly displayName: string;
  readonly sessionStorage: SpreadsheetHostSessionStorage =
    settingsBackedSessionStorage;

  private documentScopedIdCache: { urlKey: string | null; id: string | null } | null =
    null;
  private documentScopedIdInFlight: Promise<string | null> | null = null;

  constructor(app: OfficeApp | null = detectOfficeAppFromGlobals()) {
    this.displayName =
      app === "word" || app === "powerpoint"
        ? `Microsoft ${officeAppLabel(app)}`
        : "Microsoft Excel";
  }

  whenReady(): Promise<SpreadsheetHostReadyInfo> {
    if (typeof Office === "undefined") {
      return Promise.reject(new Error("Office.js is unavailable."));
    }

    return new Promise((resolve, reject) => {
      try {
        const readyPromise = Office.onReady((info) => {
          resolve(fromOfficeReadyInfo(info));
        });

        void readyPromise.catch((error: DynamicValue) => {
          reject(
            error instanceof Error
              ? error
              : new Error("Office.onReady failed."),
          );
        });
      } catch (error) {
        reject(toError(error));
      }
    });
  }

  onReady(callback: SpreadsheetHostReadyCallback): () => void {
    if (typeof Office === "undefined") {
      return () => {};
    }

    let disposed = false;
    try {
      const readyPromise = Office.onReady((info) => {
        if (disposed) return;
        callback(fromOfficeReadyInfo(info));
      });

      void readyPromise.catch((error: DynamicValue) => {
        if (!disposed) {
          console.warn("[pi] Office.onReady hook failed:", error);
        }
      });
    } catch (error) {
      console.warn("[pi] Office.onReady hook failed:", error);
    }

    return () => {
      disposed = true;
    };
  }

  getWorkbookContext(): Promise<WorkbookContext> {
    return getWorkbookContextFromDocumentUrl(getOfficeDocumentUrl());
  }

  /**
   * Resolve (and, on first sight, create) the document-scoped identity stored
   * inside the document. Cached per document URL so a save-as/open transitions
   * re-read the persisted value instead of minting a second identity.
   */
  async getDocumentScopedId(): Promise<string | null> {
    const urlKey = getOfficeDocumentUrl();
    const cached = this.documentScopedIdCache;
    if (cached && cached.urlKey === urlKey) return cached.id;

    if (this.documentScopedIdInFlight) return this.documentScopedIdInFlight;

    this.documentScopedIdInFlight = this.resolveDocumentScopedId(urlKey).finally(
      () => {
        this.documentScopedIdInFlight = null;
      },
    );
    return this.documentScopedIdInFlight;
  }

  async setDocumentScopedId(id: string): Promise<boolean> {
    const rawId = fromDocumentScopedId(id);
    this.documentScopedIdCache = {
      urlKey: getOfficeDocumentUrl(),
      id: toDocumentScopedId(rawId),
    };

    const settings = getOfficeDocumentSettings();
    if (!settings) return false;
    const persisted = await persistDocumentScopedId(settings, rawId);
    return persisted;
  }

  private async resolveDocumentScopedId(
    urlKey: string | null,
  ): Promise<string | null> {
    const settings = getOfficeDocumentSettings();
    if (!settings) {
      this.documentScopedIdCache = { urlKey, id: null };
      return null;
    }

    const existing = readDocumentScopedId(settings);
    if (existing) {
      const id = toDocumentScopedId(existing);
      this.documentScopedIdCache = { urlKey, id };
      return id;
    }

    const rawId = crypto.randomUUID();
    const id = toDocumentScopedId(rawId);
    // Cache before persisting: an in-session identity must be consistent even
    // when `saveAsync` fails (read-only document / protected view).
    this.documentScopedIdCache = { urlKey, id };
    await persistDocumentScopedId(settings, rawId);
    return id;
  }

  resolveThemeDark(): boolean | null {
    return resolveOfficeThemeDark();
  }
}
