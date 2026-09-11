/**
 * Pure model tests for the "Local Pi agent" card.
 *
 * The model is the single source of truth for the header badge + status
 * paragraph, so these tests cover every manager state × probe-outcome
 * combination — especially the regression that started this work:
 * `connected + failed probe` must keep the Connected badge AND show both
 * facts (never a bare "Could not reach the bridge.").
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  PI_BRIDGE_MODEL_KEYS,
  resolvePiBridgeCardModel,
  type PiBridgeProbeResult,
} from "../src/ui/pi-bridge-card-model.ts";
import type { PiBridgeState } from "../src/bridge/pi-bridge-manager.ts";

function state(overrides: Partial<PiBridgeState> = {}): PiBridgeState {
  return {
    status: "off",
    host: "excel",
    enabled: false,
    error: undefined,
    attempt: 0,
    nextRetryMs: undefined,
    connectFailure: undefined,
    server: undefined,
    ...overrides,
  };
}

const IDLE_PROBE: PiBridgeProbeResult = { kind: "idle" };
const OK_PROBE: PiBridgeProbeResult = {
  kind: "ok",
  serverVersion: "0.2.0",
  paneCount: 2,
};
const OLD_PROBE: PiBridgeProbeResult = { kind: "old" };
const TIMEOUT_PROBE: PiBridgeProbeResult = { kind: "timeout" };
const BLOCKED_PROBE: PiBridgeProbeResult = { kind: "blocked" };
const UNREACHABLE_PROBE: PiBridgeProbeResult = {
  kind: "unreachable",
  url: "http://127.0.0.1:38617/health",
};

/* ── Connection states ─────────────────────────────────────────────── */

void test("disabled bridge renders a muted Off badge", () => {
  const model = resolvePiBridgeCardModel(state({ enabled: false }));
  assert.equal(model.badge.text, "Off");
  assert.equal(model.badge.tone, "muted");
  assert.match(model.statusText, /off/i);
  assert.equal(model.detail, undefined);
});

void test("enabled but not connected renders Disconnected/warn", () => {
  const model = resolvePiBridgeCardModel(
    state({ enabled: true, status: "off" }),
  );
  assert.equal(model.badge.text, "Disconnected");
  assert.equal(model.badge.tone, "warn");
  assert.match(model.statusText, /not connected/i);
});

void test("connecting renders an info badge and connecting text", () => {
  const model = resolvePiBridgeCardModel(
    state({ enabled: true, status: "connecting", attempt: 1 }),
  );
  assert.equal(model.badge.text, "Connecting…");
  assert.equal(model.badge.tone, "info");
  assert.match(model.statusText, /Connecting to the Local Pi agent/);
  assert.equal(model.detail, undefined);
});

void test("retrying includes the attempt number", () => {
  const model = resolvePiBridgeCardModel(
    state({
      enabled: true,
      status: "connecting",
      attempt: 4,
      nextRetryMs: 8_000,
    }),
  );
  assert.equal(model.badge.tone, "info");
  assert.match(model.statusText, /attempt 4/);
});

void test("typed connect failures surface a localized detail line", () => {
  const cases: Array<[PiBridgeState["connectFailure"], RegExp]> = [
    ["refused", /connection refused/i],
    ["timeout", /did not answer/i],
    ["protocol-mismatch", /Protocol mismatch/i],
    ["closed", /before the handshake/i],
  ];
  for (const [reason, pattern] of cases) {
    const model = resolvePiBridgeCardModel(
      state({
        enabled: true,
        status: "connecting",
        attempt: 2,
        connectFailure: reason,
      }),
    );
    assert.match(model.detail ?? "", pattern, `detail for ${String(reason)}`);
    assert.equal(model.badge.tone, "info");
  }
});

void test("error state renders Error/warn with the message", () => {
  const model = resolvePiBridgeCardModel(
    state({ enabled: true, status: "error", error: "boom" }),
  );
  assert.equal(model.badge.text, "Error");
  assert.equal(model.badge.tone, "warn");
  assert.match(model.statusText, /boom/);
});

/* ── Connected + probe interplay (the regression) ──────────────────── */

void test("connected + idle probe shows the connection and server metadata", () => {
  const model = resolvePiBridgeCardModel(
    state({
      enabled: true,
      status: "connected",
      server: {
        serverVersion: "0.2.0",
        capabilities: ["http-health"],
      },
    }),
    IDLE_PROBE,
  );
  assert.equal(model.badge.text, "Connected");
  assert.equal(model.badge.tone, "ok");
  assert.match(model.statusText, /Connected to the Local Pi agent/);
  assert.match(model.statusText, /bridge v0\.2\.0/);
  assert.equal(model.detail, undefined);
});

void test("connected + ok probe keeps the badge and adds the probe detail", () => {
  const model = resolvePiBridgeCardModel(
    state({ enabled: true, status: "connected" }),
    OK_PROBE,
  );
  assert.equal(model.badge.text, "Connected");
  assert.equal(model.badge.tone, "ok");
  // The probe is the only source of the pane count (welcome carries no panes).
  assert.match(model.statusText, /2 pane/);
  assert.match(model.detail ?? "", /HTTP \/health OK/);
  assert.match(model.detail ?? "", /v0\.2\.0/);
  assert.match(model.detail ?? "", /2 pane/);
});

for (const [name, probe] of [
  ["old-version", OLD_PROBE],
  ["timeout", TIMEOUT_PROBE],
  ["blocked", BLOCKED_PROBE],
  ["unreachable", UNREACHABLE_PROBE],
] as const) {
  void test(`connected + ${name} probe shows BOTH facts (badge stays Connected)`, () => {
    const model = resolvePiBridgeCardModel(
      state({ enabled: true, status: "connected" }),
      probe,
    );
    assert.equal(
      model.badge.text,
      "Connected",
      "the live WebSocket is authoritative — a bad probe must not flip the badge",
    );
    assert.equal(model.badge.tone, "ok");
    assert.match(model.statusText, /Connected over WebSocket/);
    assert.notEqual(model.statusText.trim(), "Could not reach the bridge.");
  });
}

void test("unreachable probe note includes the probed URL", () => {
  const model = resolvePiBridgeCardModel(
    state({ enabled: true, status: "connected" }),
    UNREACHABLE_PROBE,
  );
  assert.match(model.statusText, /127\.0\.0\.1:38617\/health/);
});

void test("old-version probe note tells the user how to update", () => {
  const model = resolvePiBridgeCardModel(
    state({ enabled: true, status: "off" }),
    OLD_PROBE,
  );
  assert.match(
    model.statusText,
    /pi install npm:@dieulc\/pi-office-bridge@latest/,
  );
  // The badge still reflects the manager state (enabled but disconnected).
  assert.equal(model.badge.text, "Disconnected");
});

void test("a probe run while offline replaces the paragraph but not the badge", () => {
  const model = resolvePiBridgeCardModel(
    state({ enabled: true, status: "off" }),
    TIMEOUT_PROBE,
  );
  assert.match(model.statusText, /No answer from \/health/);
  assert.equal(model.badge.text, "Disconnected");
});

/* ── Global invariants ─────────────────────────────────────────────── */

void test("no state/probe combination ever emits a bare 'Could not reach the bridge.'", () => {
  const probes: PiBridgeProbeResult[] = [
    IDLE_PROBE,
    OK_PROBE,
    OLD_PROBE,
    TIMEOUT_PROBE,
    BLOCKED_PROBE,
    UNREACHABLE_PROBE,
  ];
  const states = [
    state(),
    state({ enabled: true }),
    state({ enabled: true, status: "connecting", attempt: 2 }),
    state({ enabled: true, status: "connected" }),
    state({ enabled: true, status: "error", error: "x" }),
  ];
  for (const s of states) {
    for (const p of probes) {
      const model = resolvePiBridgeCardModel(s, p);
      assert.notEqual(model.statusText.trim(), "Could not reach the bridge.");
      assert.ok(model.badge.text.length > 0, "badge text is never empty");
      assert.ok(model.statusText.length > 0, "status text is never empty");
    }
  }
});

void test("an injected translator receives the model's locale keys", () => {
  const seen: string[] = [];
  const t = (key: string): string => {
    seen.push(key);
    return key;
  };
  const model = resolvePiBridgeCardModel(
    state({ enabled: true, status: "connected" }),
    OK_PROBE,
    { t },
  );
  assert.ok(seen.includes(PI_BRIDGE_MODEL_KEYS.badgeConnected));
  assert.ok(seen.includes(PI_BRIDGE_MODEL_KEYS.statusConnected));
  assert.ok(seen.includes(PI_BRIDGE_MODEL_KEYS.probeOk));
  assert.equal(model.badge.text, PI_BRIDGE_MODEL_KEYS.badgeConnected);
  assert.equal(model.detail, PI_BRIDGE_MODEL_KEYS.probeOk);
});
