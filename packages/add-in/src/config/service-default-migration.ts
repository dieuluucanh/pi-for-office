/**
 * One-time dev-build migration of stored prod-default service URLs.
 *
 * When the add-in is built in dev mode (Vite dev server) the local-service
 * defaults point at the dev ports (bridge 38618, proxy 3004, python 3350,
 * tmux 3351). A user who previously saved the *prod* defaults into settings
 * would keep hitting the prod services, so on boot — dev builds only — any
 * stored value that is exactly a prod default is rewritten to its dev
 * default. Custom URLs are never touched; in prod builds this is a no-op.
 *
 * Pure planner + tiny async applier, so the rewrite rule is unit-testable
 * without a storage backend.
 */

import {
  DEV_LOCAL_PROXY_URL,
  DEV_PI_BRIDGE_URL,
  DEV_PYTHON_BRIDGE_URL,
  DEV_TMUX_BRIDGE_URL,
  PROD_LOCAL_PROXY_URL,
  PROD_PI_BRIDGE_URL,
  PROD_PYTHON_BRIDGE_URL,
  PROD_TMUX_BRIDGE_URL,
} from "./local-service-defaults.js";

export interface ServiceDefaultMigration {
  /** Setting key holding a local-service URL. */
  settingKey: string;
  /** Production default this migration replaces. */
  prodDefault: string;
  /** Dev default written instead. */
  devDefault: string;
}

/**
 * The four local-service URL settings, prod default → dev default.
 * Keys mirror the canonical constants (`PI_BRIDGE_URL_SETTING_KEY`,
 * `"proxy.url"`, `PYTHON_BRIDGE_URL_SETTING_KEY`, `TMUX_BRIDGE_URL_SETTING_KEY`).
 */
export const DEV_SERVICE_DEFAULT_MIGRATIONS: readonly ServiceDefaultMigration[] = [
  {
    settingKey: "pi-bridge.url",
    prodDefault: PROD_PI_BRIDGE_URL,
    devDefault: DEV_PI_BRIDGE_URL,
  },
  {
    settingKey: "proxy.url",
    prodDefault: PROD_LOCAL_PROXY_URL,
    devDefault: DEV_LOCAL_PROXY_URL,
  },
  {
    settingKey: "python.bridge.url",
    prodDefault: PROD_PYTHON_BRIDGE_URL,
    devDefault: DEV_PYTHON_BRIDGE_URL,
  },
  {
    settingKey: "tmux.bridge.url",
    prodDefault: PROD_TMUX_BRIDGE_URL,
    devDefault: DEV_TMUX_BRIDGE_URL,
  },
];

/**
 * Compute the writes needed to migrate a stored settings snapshot to the dev
 * defaults. Pure: `stored` maps setting key → stored value (`null`/`undefined`
 * = unset). Only exact prod-default matches produce a write.
 */
export function planDevServiceDefaultMigrations(
  stored: Readonly<Record<string, string | null | undefined>>,
  isDev: boolean,
): ReadonlyArray<{ settingKey: string; value: string }> {
  if (!isDev) return [];
  const writes: { settingKey: string; value: string }[] = [];
  for (const { settingKey, prodDefault, devDefault } of DEV_SERVICE_DEFAULT_MIGRATIONS) {
    const raw = stored[settingKey];
    if (typeof raw !== "string") continue;
    if (raw.trim() === prodDefault) {
      writes.push({ settingKey, value: devDefault });
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
 * Apply the dev migration at boot. Best-effort: individual read/write errors
 * are ignored so an unhealthy storage page never blocks startup.
 */
export async function applyDevServiceDefaultMigrations(
  settings: ServiceSettingsAccess,
  isDev: boolean,
): Promise<void> {
  if (!isDev) return;
  const stored: Record<string, string | null | undefined> = {};
  for (const { settingKey } of DEV_SERVICE_DEFAULT_MIGRATIONS) {
    try {
      stored[settingKey] = await settings.get(settingKey);
    } catch {
      stored[settingKey] = undefined;
    }
  }
  for (const write of planDevServiceDefaultMigrations(stored, true)) {
    try {
      await settings.set(write.settingKey, write.value);
    } catch {
      // ignore — migration is best-effort
    }
  }
}