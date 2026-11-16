/**
 * Legacy local-service URL repair.
 *
 * Older builds used a dev port profile (bridge 38618, proxy 3004, python 3350,
 * tmux 3351). That profile was revoked: dev and production builds now share one
 * loopback port profile. A settings value an older dev build could have stored
 * would point at a dead port, so on boot we rewrite exact matches to the
 * current URL.
 *
 * Rules:
 * - Runs in every build (the hosted add-in origin may also hold a dev URL).
 * - Exact match only; custom URLs and the org proxy override are never touched.
 * - `from` values are frozen historical literals; `to` values follow the
 *   current port constants.
 * - Idempotent: a repaired value equals no `from`, so a second boot writes
 *   nothing.
 *
 * The Python-bridge approval grant is repaired too, but only for the exact
 * loopback URL of the same local service on its old dev port; any custom
 * approved URL keeps requiring a fresh approval.
 *
 * Pure planner + tiny async applier, so the rewrite rule is unit-testable
 * without a storage backend.
 */

import {
  LOCAL_PROXY_URL,
  PI_BRIDGE_URL,
  PROXY_PORT,
  PYTHON_BRIDGE_URL,
  TMUX_BRIDGE_URL,
} from "./local-service-defaults.js";

export interface LegacyServiceUrlRewrite {
  /** Setting key holding a local-service URL. */
  settingKey: string;
  /** Frozen value an older dev build could have stored. */
  from: string;
  /** Current shared-port value. */
  to: string;
}

/**
 * Every URL an older dev-port build could have persisted.
 * `pi-bridge.url` / `python.bridge.url` / `tmux.bridge.url` are user-editable;
 * `proxy.url` is seeded on boot; `python.bridge.approved.url` is the stored
 * approval grant (same local service, old dev port).
 */
export const LEGACY_SERVICE_URL_REWRITES: readonly LegacyServiceUrlRewrite[] = [
  {
    settingKey: "pi-bridge.url",
    from: "ws://127.0.0.1:38618",
    to: PI_BRIDGE_URL,
  },
  {
    settingKey: "proxy.url",
    from: "https://localhost:3004",
    to: LOCAL_PROXY_URL,
  },
  {
    settingKey: "proxy.url",
    from: "http://10.0.2.2:3004",
    to: `http://10.0.2.2:${PROXY_PORT}`,
  },
  {
    settingKey: "python.bridge.url",
    from: "https://localhost:3350",
    to: PYTHON_BRIDGE_URL,
  },
  {
    settingKey: "python.bridge.approved.url",
    from: "https://localhost:3350",
    to: PYTHON_BRIDGE_URL,
  },
  {
    settingKey: "tmux.bridge.url",
    from: "https://localhost:3351",
    to: TMUX_BRIDGE_URL,
  },
];

/**
 * Compute the writes needed to repair a stored settings snapshot. Pure:
 * `stored` maps setting key → stored value (`null`/`undefined` = unset).
 * Only exact legacy matches produce a write.
 */
export function planLegacyServiceUrlRepairs(
  stored: Readonly<Record<string, string | null | undefined>>,
): ReadonlyArray<{ settingKey: string; value: string }> {
  const writes: { settingKey: string; value: string }[] = [];
  for (const { settingKey, from, to } of LEGACY_SERVICE_URL_REWRITES) {
    const raw = stored[settingKey];
    if (typeof raw !== "string") continue;
    if (raw.trim() === from) {
      writes.push({ settingKey, value: to });
    }
  }
  return writes;
}

/** Minimal settings accessor shapes (matches the app storage settings store). */
export interface ServiceSettingsAccess {
  get(key: string): Promise<string | null>;
  set(key: string, value: string): Promise<void>;
}

/**
 * Apply the legacy URL repair at boot. Best-effort: individual read/write
 * errors are ignored so an unhealthy storage page never blocks startup.
 */
export async function applyLegacyServiceUrlRepairs(
  settings: ServiceSettingsAccess,
): Promise<void> {
  const keys = new Set(
    LEGACY_SERVICE_URL_REWRITES.map(({ settingKey }) => settingKey),
  );
  const stored: Record<string, string | null | undefined> = {};
  for (const settingKey of keys) {
    try {
      stored[settingKey] = await settings.get(settingKey);
    } catch {
      stored[settingKey] = undefined;
    }
  }
  for (const write of planLegacyServiceUrlRepairs(stored)) {
    try {
      await settings.set(write.settingKey, write.value);
    } catch {
      // ignore — repair is best-effort
    }
  }
}
