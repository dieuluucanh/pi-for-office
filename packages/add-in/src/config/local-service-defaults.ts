/**
 * Local-service ports — one loopback profile shared by dev and production
 * builds.
 *
 * Port profile:
 *
 * | Service        | Port  |
 * | -------------- | ----- |
 * | Pi bridge WS   | 38617 |
 * | CORS proxy     | 3003  |
 * | Python bridge  | 3340  |
 * | Tmux bridge    | 3341  |
 *
 * The build mode never changes a port: a Vite dev add-in and the hosted add-in
 * both target the same helper services, so one process per port serves either.
 * Only the setup-command copy differs (repo scripts for a dev build, the
 * published one-liners otherwise).
 *
 * Older builds used a dev port profile (bridge 38618, proxy 3004, python 3350,
 * tmux 3351). URLs they stored are repaired once on boot by
 * `service-url-repair.ts`.
 */

/** Whether this build is a dev build (Vite dev server). */
export const IS_DEV_BUILD = readIsDevBuild();

function readIsDevBuild(): boolean {
 try {
  return import.meta.env.DEV === true;
 } catch {
  return false;
 }
}

/* ── Pi bridge (WebSocket) ──────────────────────────────────────────── */

/** Pi bridge port. Must match `BRIDGE_DEFAULT_PORT` in the protocol package. */
export const BRIDGE_PORT = 38617;

export const PI_BRIDGE_URL = `ws://127.0.0.1:${BRIDGE_PORT}`;

/* ── CORS proxy (HTTP/S) ────────────────────────────────────────────── */

export const PROXY_PORT = 3003;

export const LOCAL_PROXY_URL = `https://localhost:${PROXY_PORT}`;

/* ── Python bridge ──────────────────────────────────────────────────── */

export const PYTHON_PORT = 3340;

export const PYTHON_BRIDGE_URL = `https://localhost:${PYTHON_PORT}`;

/* ── Tmux bridge ────────────────────────────────────────────────────── */

export const TMUX_PORT = 3341;

export const TMUX_BRIDGE_URL = `https://localhost:${TMUX_PORT}`;

/* ── Setup commands (one-liners the user runs to start a service) ───── */

/**
 * Python bridge setup command. Dev builds point at the repo script
 * (`npm run python:bridge:https` from packages/add-in); otherwise the
 * published one-liner. Both bind the shared port.
 */
export function resolvePythonBridgeSetupCommand(isDev = IS_DEV_BUILD): string {
 return isDev
  ? "npm run python:bridge:https"
  : "npx pi-for-office-python-bridge";
}

/** Tmux bridge setup command (dev script vs published one-liner). */
export function resolveTmuxBridgeSetupCommand(isDev = IS_DEV_BUILD): string {
 return isDev
  ? "npm run tmux:bridge:https"
  : "npx pi-for-office-tmux-bridge";
}

/** CORS proxy setup command (dev script vs published one-liner). */
export function resolveProxySetupCommand(isDev = IS_DEV_BUILD): string {
 return isDev ? "npm run proxy:https" : "npx pi-for-office-proxy";
}
