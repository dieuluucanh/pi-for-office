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

import { PaneBridgeClient } from "./pane-client.js";
import { ALL_BRIDGE_OPS } from "./registry.js";
import { detectOfficeAppFromGlobals } from "../host/index.js";
import type { OfficeHostApp } from "@dieulc/pi-office-protocol";

const PI_BRIDGE_SETTING_KEY = "pi-bridge.enabled";

export type PiBridgeStatus = "off" | "connecting" | "connected" | "error";

export interface PiBridgeState {
  status: PiBridgeStatus;
  host: OfficeHostApp;
  enabled: boolean;
  error: string | undefined;
}

export type PiBridgeStateListener = (state: PiBridgeState) => void;

let client: PaneBridgeClient | null = null;
let state: PiBridgeState = {
  status: "off",
  host: "excel",
  enabled: false,
  error: undefined,
};
const listeners = new Set<PiBridgeStateListener>();
let generation = 0;

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
  setState({ status: "connecting", host, enabled: true, error: undefined });

  if (client) {
    client.disconnect();
    client = null;
  }

  const next = new PaneBridgeClient(
    { host, registry: ALL_BRIDGE_OPS },
    {
      onStatusChange: (connected) => {
        if (gen !== generation) return; // stale connect/disconnect race
        if (connected) {
          setState({ status: "connected", host, error: undefined });
        } else if (state.status === "connected") {
          setState({ status: "off", host, error: undefined });
        }
      },
      onServerError: (error) => {
        if (gen !== generation) return;
        warnBridge(`server error: ${error.code}: ${error.message}`);
        setState({ status: "error", host, error: error.message });
      },
    },
  );
  client = next;

  try {
    await next.connect();
  } catch (error) {
    if (gen !== generation) return;
    const message = error instanceof Error ? error.message : String(error);
    warnBridge(`could not connect (is the Pi process running?): ${message}`);
    setState({ status: "error", host, error: message });
  }
}

function stopBridge(): void {
  generation += 1; // invalidate any in-flight connect
  const current = client;
  client = null;
  if (current) current.disconnect();
  setState({ status: "off", enabled: false, error: undefined });
}
