/**
 * Pi bridge manager — single source of truth for the pane↔Pi bridge lifecycle.
 *
 * Owns the singleton `PaneBridgeClient`, its persisted `pi-bridge.enabled`
 * setting, and the connection state the Settings → Connections card renders.
 * Replaces the inline bridge bootstrap that used to live in `taskpane/init.ts`.
 *
 * Host identity is detected from the Office global at connect time, so the
 * pane always registers as the app it is actually running inside (Excel /
 * Word / PowerPoint), and tool calls route to the right pane.
 */

import {
  BridgeConnectError,
  type BridgeConnectFailureReason,
  PaneBridgeClient,
  type PaneBridgeClientCallbacks,
} from "./pane-client.js";
import { ALL_BRIDGE_OPS } from "./registry.js";
import { detectOfficeAppFromGlobals } from "../host/index.js";
import {
  BRIDGE_DEFAULT_PORT,
  type OfficeHostApp,
} from "@dieulc/pi-office-protocol";

const PI_BRIDGE_SETTING_KEY = "pi-bridge.enabled";

/** Persisted bridge WebSocket URL; only loopback ws:// or wss:// is valid. */
export const PI_BRIDGE_URL_SETTING_KEY = "pi-bridge.url";
export const DEFAULT_PI_BRIDGE_URL = `ws://127.0.0.1:${BRIDGE_DEFAULT_PORT}`;

export type PiBridgeStatus = "off" | "connecting" | "connected" | "error";

/** Server metadata captured from the `welcome` frame. */
export interface PiBridgeServerMeta {
  serverName?: string;
  serverVersion?: string;
  /** Advertised capabilities; check for `http-health` to enable the probe. */
  capabilities?: string[];
  piVersion?: string | null;
}

export interface PiBridgeState {
  status: PiBridgeStatus;
  host: OfficeHostApp;
  enabled: boolean;
  error: string | undefined;
  /** Reconnect attempt counter (0 = idle); drives "Reconnecting (attempt N)". */
  attempt: number;
  /** Delay before the next scheduled reconnect, while retrying. */
  nextRetryMs: number | undefined;
  /** Classified reason of the last failed connect, while retrying. */
  connectFailure: BridgeConnectFailureReason | undefined;
  /** Server metadata from the `welcome` frame. */
  server: PiBridgeServerMeta | undefined;
}

export type PiBridgeStateListener = (state: PiBridgeState) => void;

/** Minimal client surface the manager drives (fakeable in unit tests). */
export interface PiBridgeClientLike {
  connect(): Promise<void>;
  disconnect(): void;
}

/**
 * Test-only seams: inject a fake WebSocket client and a virtual clock so the
 * reconnect/backoff state machine can be tested deterministically.
 * Production code never calls `setPiBridgeManagerHooks`.
 */
export interface PiBridgeManagerHooks {
  createClient?: (
    options: { host: OfficeHostApp; url: string },
    callbacks: PaneBridgeClientCallbacks,
  ) => PiBridgeClientLike;
  setTimer?: (fn: () => void, ms: number) => ReturnType<typeof setTimeout>;
  clearTimer?: (handle: ReturnType<typeof setTimeout>) => void;
}

let managerHooks: PiBridgeManagerHooks = {};

/** Test-only: install the fake client/timer hooks. */
export function setPiBridgeManagerHooks(next: PiBridgeManagerHooks): void {
  managerHooks = next;
}

/** Test-only: stop everything and restore pristine module state. */
export function resetPiBridgeManagerForTests(): void {
  generation += 1;
  if (reconnectTimer !== null) {
    clearScheduledTimer(reconnectTimer);
    reconnectTimer = null;
  }
  if (stableTimer !== null) {
    clearScheduledTimer(stableTimer);
    stableTimer = null;
  }
  reconnectAttempt = 0;
  windowTriggersBound = false;
  const current = client;
  client = null;
  if (current) current.disconnect();
  listeners.clear();
  managerHooks = {};
  state = {
    status: "off",
    host: "excel",
    enabled: false,
    error: undefined,
    attempt: 0,
    nextRetryMs: undefined,
    connectFailure: undefined,
    server: undefined,
  };
}

let client: PiBridgeClientLike | null = null;
let state: PiBridgeState = {
  status: "off",
  host: "excel",
  enabled: false,
  error: undefined,
  attempt: 0,
  nextRetryMs: undefined,
  connectFailure: undefined,
  server: undefined,
};
const listeners = new Set<PiBridgeStateListener>();
let generation = 0;

/** Backoff schedule: 1 s → 2 s → 4 s → 8 s → 16 s → 30 s (cap). */
const RETRY_BASE_MS = 1_000;
const RETRY_MAX_MS = 30_000;
/** A connection kept alive this long resets the backoff (no tight loop). */
const STABLE_RESET_MS = 30_000;

let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
let stableTimer: ReturnType<typeof setTimeout> | null = null;
let reconnectAttempt = 0;
let windowTriggersBound = false;

/** Clear console noise vs exceptions: mirror the old init.ts behaviour. */
function warnBridge(message: string): void {
  console.warn(`[pi-bridge] ${message}`);
}

function hostForCurrentApp(): OfficeHostApp {
  const detected = detectOfficeAppFromGlobals();
  return detected === "word" || detected === "powerpoint" ? detected : "excel";
}

function setState(patch: Partial<PiBridgeState>): void {
  state = { ...state, ...patch };
  const snapshot = { ...state };
  for (const listener of [...listeners]) {
    try {
      listener(snapshot);
    } catch (error) {
      console.warn("[pi-bridge] state listener threw:", error);
    }
  }
  // Legacy event kept for any existing consumers.
  document.dispatchEvent(
    new CustomEvent("pi:pi-bridge-state-changed", {
      detail: { connected: state.status === "connected" },
    }),
  );
}

export function getPiBridgeState(): PiBridgeState {
  return { ...state };
}

export function subscribePiBridgeState(
  listener: PiBridgeStateListener,
): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

async function getSettingsStore(): Promise<
  import("../storage/local/settings-store.js").SettingsStore
> {
  const storageModule = await import("../storage/local/app-storage.js");
  return storageModule.getAppStorage().settings;
}

export async function getPiBridgeEnabled(): Promise<boolean> {
  try {
    const settings = await getSettingsStore();
    const value = await settings.get<boolean>(PI_BRIDGE_SETTING_KEY);
    return value === true;
  } catch {
    return false;
  }
}

/** Normalize a bridge URL; only a loopback ws:// or wss:// URL is accepted. */
export function validatePiBridgeUrl(raw: string): string {
  const candidate = raw.trim();
  if (!/^wss?:\/\//i.test(candidate)) {
    throw new Error(
      `Invalid bridge URL: expected ws:// or wss:// (e.g. ${DEFAULT_PI_BRIDGE_URL})`,
    );
  }
  let parsed: URL;
  try {
    parsed = new URL(candidate);
  } catch {
    throw new Error(`Invalid bridge URL: "${candidate}" is not a valid URL`);
  }
  const host = parsed.hostname.replace(/^\[|\]$/gu, "").toLowerCase();
  if (host !== "127.0.0.1" && host !== "localhost" && host !== "::1") {
    throw new Error(
      `Bridge URL must target the local machine (127.0.0.1 or localhost) for security, got "${host}"`,
    );
  }
  // Normalize: drop any path/trailing slashes and default the port.
  const port = parsed.port !== "" ? parsed.port : String(BRIDGE_DEFAULT_PORT);
  return `${parsed.protocol}//${host}:${port}`;
}

/** Resolve the effective bridge URL (setting if valid, else the default). */
export async function getPiBridgeUrl(): Promise<string> {
  try {
    const settings = await getSettingsStore();
    const value = await settings.get<string>(PI_BRIDGE_URL_SETTING_KEY);
    if (typeof value === "string" && value.trim().length > 0) {
      return validatePiBridgeUrl(value);
    }
  } catch (error) {
    warnBridge(
      `invalid pi-bridge.url setting, using ${DEFAULT_PI_BRIDGE_URL}: ${String(error)}`,
    );
  }
  return DEFAULT_PI_BRIDGE_URL;
}

/** Timer seams — tests substitute a virtual clock (see PiBridgeManagerHooks). */
function setScheduledTimer(
  fn: () => void,
  ms: number,
): ReturnType<typeof setTimeout> {
  return (managerHooks.setTimer ?? setTimeout)(fn, ms);
}

function clearScheduledTimer(handle: ReturnType<typeof setTimeout>): void {
  (managerHooks.clearTimer ?? clearTimeout)(handle);
}

/** Production client factory (tests override via `createClient`). */
function defaultCreateClient(
  options: { host: OfficeHostApp; url: string },
  callbacks: PaneBridgeClientCallbacks,
): PiBridgeClientLike {
  return new PaneBridgeClient(
    { host: options.host, registry: ALL_BRIDGE_OPS, url: options.url },
    callbacks,
  );
}

/** Exponential backoff: 1 s, 2 s, 4 s, … capped at 30 s. */
function retryDelayMs(attempt: number): number {
  return Math.min(RETRY_BASE_MS * 2 ** (attempt - 1), RETRY_MAX_MS);
}

/**
 * Schedule the next reconnect attempt while the bridge is enabled.
 * Guards against stale generations (disable/re-enable races).
 */
function scheduleReconnect(
  gen: number,
  connectFailure?: BridgeConnectFailureReason,
): void {
  if (gen !== generation || !state.enabled) return;
  const attempt = reconnectAttempt + 1;
  reconnectAttempt = attempt;
  const delay = retryDelayMs(attempt);
  setState({
    status: "connecting",
    attempt,
    nextRetryMs: delay,
    connectFailure,
  });
  if (reconnectTimer !== null) clearScheduledTimer(reconnectTimer);
  reconnectTimer = setScheduledTimer(() => {
    reconnectTimer = null;
    if (gen !== generation || !state.enabled) return;
    void startBridge();
  }, delay);
}

/** After 30 s of stable connection, reset the backoff to 1 s. */
function armStableReset(): void {
  if (stableTimer !== null) clearScheduledTimer(stableTimer);
  stableTimer = setScheduledTimer(() => {
    stableTimer = null;
    if (state.status === "connected") reconnectAttempt = 0;
  }, STABLE_RESET_MS);
}

/** Immediate retry on network regain / tab becoming visible. */
function retryNow(): void {
  if (!state.enabled || state.status === "connected") return;
  if (reconnectTimer !== null) {
    clearScheduledTimer(reconnectTimer);
    reconnectTimer = null;
  }
  void startBridge();
}

/** Bind once; the handlers no-op while the bridge is disabled. */
function bindWindowTriggers(): void {
  if (windowTriggersBound) return;
  // Non-browser runtimes (tests/SSR) simply skip the triggers.
  if (typeof window === "undefined" || typeof document === "undefined") return;
  windowTriggersBound = true;
  window.addEventListener("online", () => retryNow());
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") retryNow();
  });
}

function classifyConnectFailure(
  error: DynamicValue,
): BridgeConnectFailureReason | undefined {
  return error instanceof BridgeConnectError ? error.reason : undefined;
}

/** Called once at boot by init.ts. Starts only if pi-bridge.enabled is set. */
export async function initPiBridgeFromSettings(): Promise<void> {
  const enabled = await getPiBridgeEnabled();
  if (!enabled) return;
  await startBridge();
}

/** Enable + persist + connect (used by the Settings → Connections toggle). */
export async function enablePiBridge(): Promise<void> {
  try {
    const settings = await getSettingsStore();
    await settings.set(PI_BRIDGE_SETTING_KEY, true);
  } catch (error) {
    setState({
      status: "error",
      error: `could not save setting: ${String(error)}`,
    });
    return;
  }
  bindWindowTriggers();
  await startBridge();
}

/** Disable + persist + disconnect. */
export async function disablePiBridge(): Promise<void> {
  try {
    const settings = await getSettingsStore();
    await settings.set(PI_BRIDGE_SETTING_KEY, false);
  } catch (error) {
    setState({
      status: "error",
      error: `could not save setting: ${String(error)}`,
    });
    return;
  }
  stopBridge();
}

async function startBridge(): Promise<void> {
  const gen = ++generation;
  const host = hostForCurrentApp();
  setState({
    status: "connecting",
    host,
    enabled: true,
    error: undefined,
    nextRetryMs: undefined,
    server: undefined,
  });

  if (client) {
    client.disconnect();
    client = null;
  }

  const createClient = managerHooks.createClient ?? defaultCreateClient;
  const next = createClient(
    { host, url: await getPiBridgeUrl() },
    {
      onWelcome: (welcome) => {
        if (gen !== generation) return;
        // Build the meta object; optional fields are assigned conditionally
        // (exactOptionalPropertyTypes rejects explicit `undefined`).
        const server: PiBridgeServerMeta = { serverName: welcome.serverName };
        if (welcome.capabilities !== undefined) {
          server.capabilities = [...welcome.capabilities];
        }
        if (welcome.serverVersion !== undefined) {
          server.serverVersion = welcome.serverVersion;
        }
        if (welcome.piVersion !== undefined) {
          server.piVersion = welcome.piVersion;
        }
        setState({ server });
      },
      onStatusChange: (connected) => {
        if (gen !== generation) return; // stale connect/disconnect race
        if (connected) {
          armStableReset();
          setState({
            status: "connected",
            host,
            error: undefined,
            nextRetryMs: undefined,
            connectFailure: undefined,
          });
        } else if (state.status === "connected") {
          // Unexpected drop (server close or watchdog) — reconnect w/ backoff.
          warnBridge("connection lost — scheduling reconnect");
          scheduleReconnect(gen);
        }
      },
      onServerError: (error) => {
        if (gen !== generation) return;
        // Non-fatal frames (e.g. bad_message) must NOT flip a healthy
        // connection to "error"; only connect-level failures surface.
        warnBridge(`server error: ${error.code}: ${error.message}`);
      },
    },
  );
  client = next;

  try {
    await next.connect();
  } catch (error) {
    if (gen !== generation) return;
    const failure = classifyConnectFailure(error);
    const detail = error instanceof Error ? error.message : String(error);
    warnBridge(`could not connect (${failure ?? "unknown"}): ${detail}`);
    scheduleReconnect(gen, failure);
  }
}

function stopBridge(): void {
  generation += 1; // invalidate any in-flight connect/reconnect
  if (reconnectTimer !== null) {
    clearScheduledTimer(reconnectTimer);
    reconnectTimer = null;
  }
  if (stableTimer !== null) {
    clearScheduledTimer(stableTimer);
    stableTimer = null;
  }
  reconnectAttempt = 0;
  const current = client;
  client = null;
  if (current) current.disconnect();
  setState({
    status: "off",
    enabled: false,
    error: undefined,
    attempt: 0,
    nextRetryMs: undefined,
    connectFailure: undefined,
    server: undefined,
  });
}
