/**
 * Pure view model for the "Local Pi agent (bridge)" card — the single source
 * of truth for the header badge AND the status paragraph.
 *
 * Deliberately DOM-free so it can be unit-tested without a fake DOM (and so
 * two writers can never disagree: the card renderer just prints this model).
 *
 * The badge always reflects the **manager's** connection state (the live
 * WebSocket — authoritative). A failed `/health` probe never flips a
 * connected badge; instead the probe verdict is reported *alongside* the
 * connection fact ("Connected over WebSocket · /health unavailable …").
 */

import type { ItemCardBadge } from "./extensions-hub-components.js";
import type { PiBridgeState } from "../bridge/pi-bridge-manager.js";
import type { BridgeConnectFailureReason } from "../bridge/pane-client.js";

/** A translate function: key → localized string with `{var}` substitution. */
export type PiBridgeTranslator = (
  key: string,
  vars?: Record<string, string | number>,
) => string;

/**
 * Classified outcome of the HTTP `/health` probe (see `probeBridgeHealth` in
 * `bridge/pi-bridge-probe.ts`). `idle` means "no probe run yet".
 */
export type PiBridgeProbeResult =
  | { kind: "idle" }
  | { kind: "ok"; serverVersion?: string; paneCount?: number }
  | { kind: "old" }
  | { kind: "timeout" }
  | { kind: "blocked" }
  | { kind: "unreachable"; url: string };

export interface PiBridgeServerInfo {
  serverName?: string;
  serverVersion?: string;
  /** Whether `welcome` advertised the `http-health` capability. */
  hasHttpHealth?: boolean;
  paneCount?: number;
}

export interface PiBridgeRetryInfo {
  attempt: number;
  nextRetryMs?: number;
}

export interface PiBridgeModelOptions {
  /** i18n translator; defaults to an embedded English fallback. */
  t?: PiBridgeTranslator;
  /** Server metadata captured from the `welcome` frame (manager). */
  server?: PiBridgeServerInfo;
  /** Reconnect progress (manager, auto-reconnect). */
  retry?: PiBridgeRetryInfo;
}

export interface PiBridgeCardModel {
  badge: ItemCardBadge;
  statusText: string;
  /** Secondary line (e.g. probe verdict) rendered under `statusText`. */
  detail?: string;
}

/* ── Locale keys used by the model ──────────────────────────────────── */

export const PI_BRIDGE_MODEL_KEYS = {
  badgeOff: "ext-hub-connections.piBridgeBadgeOff",
  badgeDisconnected: "ext-hub-connections.piBridgeBadgeDisconnected",
  badgeConnecting: "ext-hub-connections.piBridgeBadgeConnecting",
  badgeConnected: "ext-hub-connections.piBridgeBadgeConnected",
  badgeError: "ext-hub-connections.piBridgeBadgeError",
  statusOff: "ext-hub-connections.piBridgeStatusOff",
  statusDisconnected: "ext-hub-connections.piBridgeStatusDisconnected",
  statusConnecting: "ext-hub-connections.piBridgeStatusConnecting",
  statusRetrying: "ext-hub-connections.piBridgeStatusRetrying",
  statusConnected: "ext-hub-connections.piBridgeStatusConnected",
  statusWsOkProbe: "ext-hub-connections.piBridgeStatusWsOkProbe",
  serverInfo: "ext-hub-connections.piBridgeServerInfo",
  paneSuffix: "ext-hub-connections.piBridgePaneSuffix",
  probeOk: "ext-hub-connections.piBridgeProbeOk",
  probeOldVersion: "ext-hub-connections.piBridgeProbeOldVersion",
  probeTimeout: "ext-hub-connections.piBridgeProbeTimeout",
  probeBlocked: "ext-hub-connections.piBridgeProbeBlocked",
  probeUnreachable: "ext-hub-connections.piBridgeProbeUnreachable",
  connectRefused: "ext-hub-connections.piBridgeConnectRefused",
  connectTimeout: "ext-hub-connections.piBridgeConnectTimeout",
  connectMismatch: "ext-hub-connections.piBridgeConnectMismatch",
  connectClosed: "ext-hub-connections.piBridgeConnectClosed",
  error: "ext-hub-connections.piBridgeError",
} as const;

/** Embedded English fallback so the model works before locales load. */
const DEFAULT_EN: Record<string, string> = {
  [PI_BRIDGE_MODEL_KEYS.badgeOff]: "Off",
  [PI_BRIDGE_MODEL_KEYS.badgeDisconnected]: "Disconnected",
  [PI_BRIDGE_MODEL_KEYS.badgeConnecting]: "Connecting…",
  [PI_BRIDGE_MODEL_KEYS.badgeConnected]: "Connected",
  [PI_BRIDGE_MODEL_KEYS.badgeError]: "Error",
  [PI_BRIDGE_MODEL_KEYS.statusOff]: "Local Pi bridge is off.",
  [PI_BRIDGE_MODEL_KEYS.statusDisconnected]:
    "Not connected to the Local Pi agent — is Pi running with the bridge extension?",
  [PI_BRIDGE_MODEL_KEYS.statusConnecting]: "Connecting to the Local Pi agent…",
  [PI_BRIDGE_MODEL_KEYS.statusRetrying]: "Reconnecting (attempt {attempt})…",
  [PI_BRIDGE_MODEL_KEYS.statusConnected]: "Connected to the Local Pi agent.",
  [PI_BRIDGE_MODEL_KEYS.statusWsOkProbe]:
    "Connected over WebSocket · {probeNote}",
  [PI_BRIDGE_MODEL_KEYS.serverInfo]: " · bridge v{version}",
  [PI_BRIDGE_MODEL_KEYS.paneSuffix]: " · {count} pane(s) attached",
  [PI_BRIDGE_MODEL_KEYS.probeOk]:
    "HTTP /health OK — v{version}, {count} pane(s) attached",
  [PI_BRIDGE_MODEL_KEYS.probeOldVersion]:
    "/health unavailable — the installed bridge is an older version. Update it with `pi install npm:@dieulc/pi-office-bridge@latest`.",
  [PI_BRIDGE_MODEL_KEYS.probeTimeout]:
    "No answer from /health within 3s — the bridge may be an older version, or Pi is restarting.",
  [PI_BRIDGE_MODEL_KEYS.probeBlocked]:
    "The browser blocked local network access — use desktop Office, or allow localhost access in the browser.",
  [PI_BRIDGE_MODEL_KEYS.probeUnreachable]:
    "Could not reach {url} — check that Pi is running with the bridge extension.",
  [PI_BRIDGE_MODEL_KEYS.connectRefused]:
    "Pi is not running with the bridge extension (connection refused).",
  [PI_BRIDGE_MODEL_KEYS.connectTimeout]:
    "The bridge did not answer in time — is Pi running?",
  [PI_BRIDGE_MODEL_KEYS.connectMismatch]:
    "Protocol mismatch — update the bridge: `pi install npm:@dieulc/pi-office-bridge@latest`.",
  [PI_BRIDGE_MODEL_KEYS.connectClosed]:
    "The bridge closed the connection before the handshake finished.",
  [PI_BRIDGE_MODEL_KEYS.error]: "Connection error: {message}",
};

function defaultTranslator(
  key: string,
  vars?: Record<string, string | number>,
): string {
  const template = DEFAULT_EN[key] ?? key;
  if (!vars) return template;
  return template.replace(/\{(\w+)\}/g, (match, name: string) =>
    name in vars ? String(vars[name]) : match,
  );
}

function hasKnownVersion(version: string | undefined): version is string {
  return (
    typeof version === "string" && version.length > 0 && version !== "unknown"
  );
}

/** Localized action line for a typed connect failure (manager). */
function connectFailureKey(reason: BridgeConnectFailureReason): string {
  switch (reason) {
    case "refused":
      return PI_BRIDGE_MODEL_KEYS.connectRefused;
    case "timeout":
      return PI_BRIDGE_MODEL_KEYS.connectTimeout;
    case "protocol-mismatch":
      return PI_BRIDGE_MODEL_KEYS.connectMismatch;
    case "closed":
      return PI_BRIDGE_MODEL_KEYS.connectClosed;
  }
}

/** Adapt the manager's `welcome` metadata to the model's server shape. */
function serverInfoFromState(
  state: PiBridgeState,
): PiBridgeServerInfo | undefined {
  const meta = state.server;
  if (meta === undefined) return undefined;
  const info: PiBridgeServerInfo = {};
  if (meta.serverVersion !== undefined) {
    info.serverVersion = meta.serverVersion;
  }
  if (meta.capabilities !== undefined) {
    info.hasHttpHealth = meta.capabilities.includes("http-health");
  }
  return info;
}

/**
 * Resolve the card's badge + status text from the manager state and an
 * optional probed `/health` verdict. Pure: no DOM, no side effects.
 */
export function resolvePiBridgeCardModel(
  state: PiBridgeState,
  probe: PiBridgeProbeResult = { kind: "idle" },
  opts: PiBridgeModelOptions = {},
): PiBridgeCardModel {
  const t = opts.t ?? defaultTranslator;
  // Server metadata: explicit opts win; otherwise use what the manager
  // captured from the `welcome` frame.
  const server = opts.server ?? serverInfoFromState(state);

  const badge = resolveBadge(state, t);
  const probeNote = probeNoteText(probe, t);

  // Connected: the WS is authoritative; a bad probe is reported *alongside*
  // the connection fact, never replacing it.
  if (state.status === "connected") {
    if (probe.kind !== "idle" && probe.kind !== "ok") {
      return {
        badge,
        statusText: t(PI_BRIDGE_MODEL_KEYS.statusWsOkProbe, {
          probeNote,
        }),
      };
    }
    let statusText = t(PI_BRIDGE_MODEL_KEYS.statusConnected);
    if (server && hasKnownVersion(server.serverVersion)) {
      statusText += t(PI_BRIDGE_MODEL_KEYS.serverInfo, {
        version: server.serverVersion,
      });
    }
    // Pane count: the welcome frame has no pane list, so fall back to the
    // probe verdict when it succeeded.
    const paneCount =
      server?.paneCount ?? (probe.kind === "ok" ? probe.paneCount : undefined);
    if (typeof paneCount === "number" && paneCount > 0) {
      statusText += t(PI_BRIDGE_MODEL_KEYS.paneSuffix, {
        count: paneCount,
      });
    }
    const model: PiBridgeCardModel = { badge, statusText };
    if (probe.kind === "ok") {
      model.detail = t(PI_BRIDGE_MODEL_KEYS.probeOk, {
        version:
          probe.serverVersion && hasKnownVersion(probe.serverVersion)
            ? probe.serverVersion
            : "unknown",
        count: probe.paneCount ?? 0,
      });
    }
    return model;
  }

  // Not connected: a just-run probe gets the paragraph (today's behaviour),
  // otherwise the state's own message.
  if (probe.kind !== "idle") {
    return { badge, statusText: probeNote };
  }

  if (state.status === "connecting") {
    const attempt = opts.retry?.attempt ?? state.attempt ?? 1;
    const statusText =
      attempt > 1
        ? t(PI_BRIDGE_MODEL_KEYS.statusRetrying, { attempt })
        : t(PI_BRIDGE_MODEL_KEYS.statusConnecting);
    const model: PiBridgeCardModel = { badge, statusText };
    if (state.connectFailure !== undefined) {
      model.detail = t(connectFailureKey(state.connectFailure));
    }
    return model;
  }

  if (state.status === "error") {
    return {
      badge,
      statusText: t(PI_BRIDGE_MODEL_KEYS.error, {
        message: state.error ?? "unknown",
      }),
    };
  }

  // status === "off"
  return {
    badge,
    statusText: state.enabled
      ? t(PI_BRIDGE_MODEL_KEYS.statusDisconnected)
      : t(PI_BRIDGE_MODEL_KEYS.statusOff),
  };
}

function resolveBadge(
  state: PiBridgeState,
  t: PiBridgeTranslator,
): ItemCardBadge {
  if (!state.enabled) {
    return { text: t(PI_BRIDGE_MODEL_KEYS.badgeOff), tone: "muted" };
  }
  switch (state.status) {
    case "connected":
      return { text: t(PI_BRIDGE_MODEL_KEYS.badgeConnected), tone: "ok" };
    case "connecting":
      return { text: t(PI_BRIDGE_MODEL_KEYS.badgeConnecting), tone: "info" };
    case "error":
      return { text: t(PI_BRIDGE_MODEL_KEYS.badgeError), tone: "warn" };
    default:
      // Enabled but not connected (e.g. Pi closed the socket, or startup).
      return { text: t(PI_BRIDGE_MODEL_KEYS.badgeDisconnected), tone: "warn" };
  }
}

function probeNoteText(
  probe: PiBridgeProbeResult,
  t: PiBridgeTranslator,
): string {
  switch (probe.kind) {
    case "ok":
      return t(PI_BRIDGE_MODEL_KEYS.probeOk, {
        version:
          probe.serverVersion && hasKnownVersion(probe.serverVersion)
            ? probe.serverVersion
            : "unknown",
        count: probe.paneCount ?? 0,
      });
    case "old":
      return t(PI_BRIDGE_MODEL_KEYS.probeOldVersion);
    case "timeout":
      return t(PI_BRIDGE_MODEL_KEYS.probeTimeout);
    case "blocked":
      return t(PI_BRIDGE_MODEL_KEYS.probeBlocked);
    case "unreachable":
      return t(PI_BRIDGE_MODEL_KEYS.probeUnreachable, { url: probe.url });
    default:
      return "";
  }
}
