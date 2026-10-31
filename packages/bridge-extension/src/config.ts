/**
 * Bridge configuration: port precedence and the opt-in autostart switch.
 *
 * Kept free of Pi imports so the precedence rules are testable and identical in
 * the TUI and Pi Desktop hosts.
 */

import { BRIDGE_DEFAULT_PORT } from "./protocol.js";

export const FLAG_PORT = "office-bridge-port";
export const FLAG_AUTOSTART = "office-bridge-autostart";

export type FlagReader = (name: string) => boolean | string | undefined;

export type BridgeEnv = Record<string, string | undefined>;

function parsePort(raw: unknown): number | null {
  const text = typeof raw === "string" ? raw.trim() : "";
  if (text === "") return null;
  if (!/^\d+$/u.test(text)) return null;
  const value = Number.parseInt(text, 10);
  return Number.isInteger(value) && value > 0 && value < 65536 ? value : null;
}

/**
 * Port precedence: `--office-bridge-port` → `PI_OFFICE_BRIDGE_PORT` → 38617.
 * An invalid explicit value falls through to the next source rather than
 * crashing the session.
 */
export function resolveBridgePort(readFlag: FlagReader, env: BridgeEnv): number {
  return (
    parsePort(readFlag(FLAG_PORT)) ??
    parsePort(env.PI_OFFICE_BRIDGE_PORT) ??
    BRIDGE_DEFAULT_PORT
  );
}

function truthy(raw: string | undefined): boolean {
  if (typeof raw !== "string") return false;
  const value = raw.trim().toLowerCase();
  return value === "1" || value === "true" || value === "yes" || value === "on";
}

/**
 * Autostart is **off by default**. It returns only when the user passed the
 * extension flag or set the environment variable; there is no persisted GUI
 * setting in this release.
 */
export function resolveBridgeAutostart(readFlag: FlagReader, env: BridgeEnv): boolean {
  if (readFlag(FLAG_AUTOSTART) === true) return true;
  return truthy(env.PI_OFFICE_BRIDGE_AUTOSTART);
}
