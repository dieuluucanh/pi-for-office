/**
 * Host-neutral workbook identity primitives.
 *
 * IMPORTANT: callers should persist only `workbookId` (the hash), never a raw
 * host document URL/path. Host adapters pass any raw identity source through
 * this module so the privacy-preserving semantics stay centralized.
 */

import { t } from "../language/index.js";

export interface WorkbookContext {
  /**
   * A stable, local-only identifier when available.
   *
   * Current host strategies:
   * - Office: normalize the host document URL (drop query/hash) and return a SHA-256 hash
   * - WPS: normalize ActiveWorkbook.FullName and return a SHA-256 hash with a WPS-specific prefix
   * - otherwise return null (ephemeral/unknown)
   */
  workbookId: string | null;

  /** Best-effort workbook file name derived from the host document URL/path. */
  workbookName: string | null;

  /** Where the identity came from (useful for debugging / future migration). */
  source: "document.url" | "wps.activeWorkbook.fullName" | "unknown";
}

export function createUnknownWorkbookContext(): WorkbookContext {
  return {
    workbookId: null,
    workbookName: null,
    source: "unknown",
  };
}

function bufferToHex(buf: ArrayBuffer): string {
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function stripQueryAndHash(raw: string): string {
  return raw.split(/[?#]/, 1)[0]?.trim() ?? "";
}

/** `C:\...`, `c:/...` — a Windows drive-qualified path (not a URL). */
const WINDOWS_DRIVE_PATH = /^[a-zA-Z]:[\\/]/;
/** `\\server\share\...` — a UNC path (not a URL). */
const UNC_PATH = /^\\\\[^\\]/;

/** Decode percent-escapes once; unparseable/malformed input falls back to raw. */
function decodeOnce(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

/** Collapse duplicate separators and drop a trailing slash (keeps `/`). */
function normalizePathSeparators(pathname: string): string {
  const collapsed = pathname.replace(/\\/g, "/").replace(/\/{2,}/g, "/");
  if (collapsed.length > 1 && collapsed.endsWith("/")) {
    return collapsed.slice(0, -1);
  }
  return collapsed;
}

/** Lowercase only the drive letter — the rest of a path may be case-sensitive. */
function lowercaseDriveLetter(pathname: string): string {
  return pathname.replace(
    /^\/([a-zA-Z]):/u,
    (_match, drive: string) => `/${drive.toLowerCase()}:`,
  );
}

function getWorkbookNameFromUrl(url: string): string | null {
  const readName = (raw: string): string | null => {
    const cleaned = stripQueryAndHash(raw);
    if (!cleaned) return null;

    try {
      const decoded = decodeURIComponent(cleaned).trim();
      return decoded.length > 0 ? decoded : null;
    } catch {
      return cleaned.length > 0 ? cleaned : null;
    }
  };

  try {
    const parsed = new URL(url);
    const normalizedPathname = parsed.pathname.replace(/\\/g, "/");
    const fromPathname = normalizedPathname.split("/").at(-1);
    if (fromPathname) {
      const name = readName(fromPathname);
      if (name) return name;
    }
  } catch {
    // Fall back to plain path parsing below.
  }

  const normalized = url.replace(/\\/g, "/");
  const fromPath = normalized.split("/").at(-1);
  return fromPath ? readName(fromPath) : null;
}

function normalizeWorkbookIdentitySource(url: string): string {
  const trimmed = url.trim();
  if (trimmed.length === 0) return trimmed;

  const stripped = stripQueryAndHash(trimmed);
  if (stripped.length === 0) return stripped;

  // Windows drive path (Excel desktop often reports a plain path, sometimes a
  // `file:///` URL for the exact same file — both must hash identically).
  if (WINDOWS_DRIVE_PATH.test(stripped)) {
    const decoded = decodeOnce(stripped).replace(/\\/g, "/");
    const drive = decoded.slice(0, 2).toLowerCase();
    const rest = normalizePathSeparators(decoded.slice(2));
    return `file:///${drive}${rest}`;
  }

  // UNC path: `\\server\share\file.xlsx` or `//server/share/file.xlsx`.
  if (UNC_PATH.test(stripped) || (stripped.startsWith("//") && !stripped.startsWith("///"))) {
    const decoded = decodeOnce(stripped).replace(/\\/g, "/");
    const parts = decoded.replace(/^\/+/, "").split("/");
    const host = (parts.shift() ?? "").toLowerCase();
    const rest = normalizePathSeparators(parts.join("/"));
    return `file://${host}/${rest}`;
  }

  // Bare POSIX absolute path (WPS/macOS edge cases).
  if (stripped.startsWith("/")) {
    return `file://${normalizePathSeparators(decodeOnce(stripped))}`;
  }

  try {
    const parsed = new URL(stripped);
    const protocol = parsed.protocol.toLowerCase();

    if (protocol === "file:") {
      const host = parsed.host.toLowerCase();
      const decoded = decodeOnce(parsed.pathname);
      const path = lowercaseDriveLetter(normalizePathSeparators(decoded));
      return host ? `file://${host}${path}` : `file://${path}`;
    }

    // WHATWG already drops default ports from `host`; keeping hostname keeps it
    // explicit and lowercases only the authority (the path stays as-is).
    const authority = parsed.host === "" ? "" : parsed.host.toLowerCase();
    const path = normalizePathSeparators(decodeOnce(parsed.pathname));

    if (protocol === "http:" || protocol === "https:") {
      return `${protocol}//${authority}${path}`;
    }

    return authority ? `${protocol}//${authority}${path}` : `${protocol}${path}`;
  } catch {
    return stripped.replace(/\\/g, "/");
  }
}

export function formatWorkbookLabel(context: WorkbookContext): string {
  if (context.workbookName) return context.workbookName;

  if (context.workbookId) {
    const shortId = context.workbookId.slice(0, 18);
    return `Workbook (${shortId}…)`;
  }

  return t("workbook.currentWorkbook");
}

function fnv1a32Hex(bytes: Uint8Array): string {
  let hash = 0x811c9dc5;
  for (const b of bytes) {
    hash ^= b;
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

async function sha256Hex(input: string): Promise<string> {
  const bytes = new TextEncoder().encode(input);

  // WebCrypto SHA-256 (preferred)
  const subtle = globalThis.crypto?.subtle;
  if (subtle?.digest) {
    try {
      const buf = await subtle.digest("SHA-256", bytes);
      return bufferToHex(buf);
    } catch {
      // fall through to FNV-1a
    }
  }

  // Fallback: non-cryptographic but stable hash. Only used when WebCrypto is unavailable.
  return fnv1a32Hex(bytes);
}

/**
 * Best-effort workbook context from a host-provided document URL/path.
 *
 * IMPORTANT: callers should persist only `workbookId` (the hash), never the raw URL/path.
 */
export async function getWorkbookContextFromDocumentUrl(url: string | null): Promise<WorkbookContext> {
  if (!url) {
    return createUnknownWorkbookContext();
  }

  const identitySource = normalizeWorkbookIdentitySource(url);
  const hash = await sha256Hex(identitySource);
  return {
    workbookId: `url_sha256:${hash}`,
    workbookName: getWorkbookNameFromUrl(url),
    source: "document.url",
  };
}

/**
 * Best-effort workbook context from a WPS ActiveWorkbook.FullName path.
 *
 * IMPORTANT: callers should persist only `workbookId` (the hash), never the raw
 * local path. WPS paths intentionally use their own prefix so they cannot be
 * confused with Office document URL identities.
 */
export async function getWorkbookContextFromWpsFullName(
  fullName: string | null,
  workbookName: string | null,
): Promise<WorkbookContext> {
  const trimmedFullName = fullName?.trim() ?? "";
  if (!trimmedFullName) {
    return {
      workbookId: null,
      workbookName,
      source: "unknown",
    };
  }

  const identitySource = normalizeWorkbookIdentitySource(trimmedFullName);
  const hash = await sha256Hex(identitySource);
  return {
    workbookId: `wps_path_sha256:${hash}`,
    workbookName,
    source: "wps.activeWorkbook.fullName",
  };
}
