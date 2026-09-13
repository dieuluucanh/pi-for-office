/**
 * Local-service defaults — the single source of truth for which ports the
 * add-in's local helper services use, and how those defaults differ between
 * dev and production builds.
 *
 * Mode signal: Vite's `import.meta.env.DEV` (true under `vite dev`, false for
 * `vite build`, including the GitHub Pages deploy and `serve:dist`). Under
 * Node/tests `import.meta.env` is undefined, so the guarded read fails closed
 * to the production profile — existing tests that assert prod defaults keep
 * passing.
 *
 * Port profiles:
 *
 * | Service        | Prod | Dev |
 * | --------------- | ------ | ------ |
 * | Pi bridge WS   | 38617 | 38618 |
 * | CORS proxy     | 3003  | 3004  |
 * | Python bridge  | 3340  | 3350  |
 * | Tmux bridge    | 3341  | 3351  |
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

export const BRIDGE_PROD_PORT = 38617;
export const BRIDGE_DEV_PORT = 38618;

export const PROD_PI_BRIDGE_URL = `ws://127.0.0.1:${BRIDGE_PROD_PORT}`;
export const DEV_PI_BRIDGE_URL = `ws://127.0.0.1:${BRIDGE_DEV_PORT}`;

export function defaultBridgePort(isDev = IS_DEV_BUILD): number {
  return isDev ? BRIDGE_DEV_PORT : BRIDGE_PROD_PORT;
}

export function resolvePiBridgeUrl(isDev = IS_DEV_BUILD): string {
  return isDev ? DEV_PI_BRIDGE_URL : PROD_PI_BRIDGE_URL;
}

/* ── CORS proxy (HTTP/S) ────────────────────────────────────────────── */

export const PROXY_PROD_PORT = 3003;
export const PROXY_DEV_PORT = 3004;

export const PROD_LOCAL_PROXY_URL = `https://localhost:${PROXY_PROD_PORT}`;
export const DEV_LOCAL_PROXY_URL = `https://localhost:${PROXY_DEV_PORT}`;

export function resolveLocalProxyUrl(isDev = IS_DEV_BUILD): string {
  return isDev ? DEV_LOCAL_PROXY_URL : PROD_LOCAL_PROXY_URL;
}

/* ── Python bridge ──────────────────────────────────────────────────── */

export const PYTHON_PROD_PORT = 3340;
export const PYTHON_DEV_PORT = 3350;

export const PROD_PYTHON_BRIDGE_URL = `https://localhost:${PYTHON_PROD_PORT}`;
export const DEV_PYTHON_BRIDGE_URL = `https://localhost:${PYTHON_DEV_PORT}`;

export function resolvePythonBridgeUrl(isDev = IS_DEV_BUILD): string {
  return isDev ? DEV_PYTHON_BRIDGE_URL : PROD_PYTHON_BRIDGE_URL;
}

/* ── Tmux bridge ────────────────────────────────────────────────────── */

export const TMUX_PROD_PORT = 3341;
export const TMUX_DEV_PORT = 3351;

export const PROD_TMUX_BRIDGE_URL = `https://localhost:${TMUX_PROD_PORT}`;
export const DEV_TMUX_BRIDGE_URL = `https://localhost:${TMUX_DEV_PORT}`;

export function resolveTmuxBridgeUrl(isDev = IS_DEV_BUILD): string {
  return isDev ? DEV_TMUX_BRIDGE_URL : PROD_TMUX_BRIDGE_URL;
}

/* ── Setup commands (one-liners the user runs to start a service) ───── */

/**
 * Python bridge setup command. Dev builds point at the repo dev script
 * (`npm run python:bridge:dev:https` from packages/add-in); otherwise the
 * published one-liner.
 */
export function resolvePythonBridgeSetupCommand(isDev = IS_DEV_BUILD): string {
  return isDev
    ? "npm run python:bridge:dev:https"
    : "npx pi-for-office-python-bridge";
}

/** Tmux bridge setup command (dev script vs published one-liner). */
export function resolveTmuxBridgeSetupCommand(isDev = IS_DEV_BUILD): string {
  return isDev
    ? "npm run tmux:bridge:dev:https"
    : "npx pi-for-office-tmux-bridge";
}

/** CORS proxy setup command (dev script vs published one-liner). */
export function resolveProxySetupCommand(isDev = IS_DEV_BUILD): string {
  return isDev ? "npm run proxy:dev" : "npx pi-for-office-proxy";
}