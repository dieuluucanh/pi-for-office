/**
 * HTTP `/health` probe for the Local Pi bridge.
 *
 * The bridge server (v0.2.0+) serves `GET /health` with JSON metadata on its
 * loopback HTTP surface. Older bridges (0.1.0) have no HTTP listener at all —
 * the TCP port accepts WebSocket handshakes only — so a probe against them
 * hangs (→ timeout) or is refused. The probe classifies every outcome so the
 * card never shows a bare "Could not reach the bridge":
 *
 * - `ok`          — 200 + JSON with `ok: true` → bridge is current.
 * - `old`         — server answered but not with our `/health` (404 / non-JSON
 *                   / 5xx) → running an older bridge version.
 * - `timeout`     — no answer within the budget (older bridge or Pi restarting).
 * - `unreachable` — TCP-level refusal (no bridge on the port at all).
 * - `blocked`     — fetch rejected without ECONN* (browser blocked local
 *                   network access — Office on the web / LNA, or CSP).
 */

import type { PiBridgeProbeResult } from "../ui/pi-bridge-card-model.js";

export interface PiBridgeProbeOptions {
  /** Total budget for the request (default 3 s). */
  timeoutMs?: number;
  /** Injectable for tests; defaults to the global fetch. */
  fetchImpl?: typeof fetch;
}

const DEFAULT_TIMEOUT_MS = 3000;

/** Translate a bridge WS URL to its /health HTTP URL. */
export function healthUrlFromWsUrl(wsUrl: string): string {
  const normalized = wsUrl.replace(/^ws:/, "http:").replace(/^wss:/, "https:");
  const base = normalized.replace(/\/+$/u, "");
  return `${base}/health`;
}

/**
 * `targetAddressSpace: "loopback"` lets a secure public origin (hosted GitHub
 * Pages) reach the loopback bridge under Chrome's Local Network Access rules.
 * Older WebView2/Chromium runtimes reject the option — probe once.
 */
function canUseTargetAddressSpace(): boolean {
  // SAFETY: `targetAddressSpace` is Chromium's Local Network Access opt-in; the
  // TS lib does not declare it yet, so describe the option explicitly.
  const probeOptions: RequestInit & { targetAddressSpace?: string } = {
    targetAddressSpace: "loopback",
  };
  try {
    // Constructing a Request validates RequestInit; runtimes that do not know
    // the option throw TypeError, which is exactly the signal we want.
    void new Request("http://127.0.0.1/", probeOptions);
    return true;
  } catch {
    return false;
  }
}

function isConnectionRefused(error: DynamicValue): boolean {
  if (typeof error !== "object" || error === null) return false;
  const cause = (error as { cause?: DynamicValue }).cause;
  if (typeof cause !== "object" || cause === null) return false;
  const code = (cause as { code?: DynamicValue }).code;
  return typeof code === "string" && code.startsWith("ECONN");
}

export async function probeBridgeHealth(
  url: string,
  opts: PiBridgeProbeOptions = {},
): Promise<PiBridgeProbeResult> {
  const timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const fetchImpl = opts.fetchImpl ?? fetch;

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  // AbortController-based timeout (not AbortSignal.timeout-only) so older
  // WebView2 runtimes behave the same.
  const requestInit: RequestInit & { targetAddressSpace?: string } = {
    mode: "cors",
    cache: "no-store",
    signal: controller.signal,
  };
  if (canUseTargetAddressSpace()) {
    requestInit.targetAddressSpace = "loopback";
  }

  try {
    const res = await fetchImpl(url, requestInit);
    if (!res.ok) return { kind: "old" };

    let payload: DynamicValue;
    try {
      payload = await res.json();
    } catch {
      return { kind: "old" };
    }

    if (typeof payload !== "object" || payload === null) return { kind: "old" };
    const health = payload as {
      ok?: DynamicValue;
      serverVersion?: DynamicValue;
      panes?: DynamicValue;
    };
    if (health.ok !== true) return { kind: "old" };

    const result: PiBridgeProbeResult = {
      kind: "ok",
      paneCount: Array.isArray(health.panes) ? health.panes.length : 0,
    };
    if (typeof health.serverVersion === "string") {
      result.serverVersion = health.serverVersion;
    }
    return result;
  } catch (error) {
    if (controller.signal.aborted) return { kind: "timeout" };
    if (isConnectionRefused(error)) return { kind: "unreachable", url };
    return { kind: "blocked" };
  } finally {
    clearTimeout(timer);
  }
}
