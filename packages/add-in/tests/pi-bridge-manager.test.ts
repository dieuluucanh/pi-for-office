/**
 * Manager state-machine tests with an injected fake client + virtual clock.
 *
 * Covers the self-healing contract added for the Local Pi bridge:
 *  - enable starts a connect; welcome metadata is captured
 *  - a failed connect schedules an exponential backoff (1s → 30s cap)
 *  - the backoff resets after 30 s of stability
 *  - disable cancels pending reconnects and clears state
 *  - an unexpected disconnect reconnects automatically
 *  - callbacks from a superseded client are ignored (generation guard)
 *  - non-fatal server error frames never flip a connected bridge to "error"
 *  - an `online` window event retries immediately
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  disablePiBridge,
  enablePiBridge,
  getPiBridgeState,
  resetPiBridgeManagerForTests,
  setPiBridgeManagerHooks,
  type PiBridgeClientLike,
} from "../src/bridge/pi-bridge-manager.ts";
import {
  BridgeConnectError,
  type PaneBridgeClientCallbacks,
} from "../src/bridge/pane-client.ts";
import {
  setAppStorage,
  type AppStorage,
} from "../src/storage/local/app-storage.ts";
import type { WelcomeMessage } from "@dieulc/pi-office-protocol";
import { installFakeDom } from "./fake-dom.test.ts";

/* ── Harness ───────────────────────────────────────────────────────── */

interface FakeClient extends PiBridgeClientLike {
  connectCount: number;
  disconnectCount: number;
  options: { host: string; url: string };
  callbacks: PaneBridgeClientCallbacks;
}

interface FakeClock {
  setTimer: (fn: () => void, ms: number) => ReturnType<typeof setTimeout>;
  clearTimer: (handle: ReturnType<typeof setTimeout>) => void;
  pending(): number;
  advance(ms: number): void;
}

interface Harness {
  clock: FakeClock;
  clients: FakeClient[];
  setConnectScript(fn: () => Promise<void>): void;
  firstClient(): FakeClient;
  lastClient(): FakeClient;
  restore(): void;
}

const tick = (): Promise<void> =>
  new Promise((resolve) => {
    setImmediate(resolve);
  });

function createFakeClock(): FakeClock {
  let now = 0;
  let nextId = 1;
  const timers = new Map<number, { fn: () => void; at: number }>();
  const idOf = (handle: ReturnType<typeof setTimeout>): number =>
    handle as DynamicValue as number;
  return {
    setTimer: (fn, ms) => {
      const id = nextId++;
      timers.set(id, { fn, at: now + ms });
      return id as DynamicValue as ReturnType<typeof setTimeout>;
    },
    clearTimer: (handle) => {
      timers.delete(idOf(handle));
    },
    pending: () => timers.size,
    advance: (ms) => {
      now += ms;
      for (const [id, timer] of [...timers]) {
        if (timer.at <= now) {
          timers.delete(id);
          timer.fn();
        }
      }
    },
  };
}

function installMemoryStorage(): void {
  const data = new Map<string, DynamicValue>();
  const settings = {
    get: (key: string) => Promise.resolve(data.get(key)),
    set: (key: string, value: DynamicValue) => {
      data.set(key, value);
      return Promise.resolve();
    },
    delete: (key: string) => {
      data.delete(key);
      return Promise.resolve();
    },
  };
  const appStorage: DynamicValue = { settings };
  setAppStorage(appStorage as AppStorage);
}

function setup(): Harness {
  const dom = installFakeDom();
  const previousWindow = Reflect.get(globalThis, "window");
  Reflect.set(globalThis, "window", new EventTarget());
  installMemoryStorage();

  const clock = createFakeClock();
  const clients: FakeClient[] = [];
  let connectScript: () => Promise<void> = () => Promise.resolve();

  resetPiBridgeManagerForTests();
  setPiBridgeManagerHooks({
    createClient: (options, callbacks) => {
      const client: FakeClient = {
        connectCount: 0,
        disconnectCount: 0,
        options,
        callbacks,
        connect: () => {
          client.connectCount += 1;
          return connectScript();
        },
        disconnect: () => {
          client.disconnectCount += 1;
        },
      };
      clients.push(client);
      return client;
    },
    setTimer: clock.setTimer,
    clearTimer: clock.clearTimer,
  });

  const firstClient = (): FakeClient => {
    const client = clients[0];
    assert.ok(client, "expected at least one client");
    return client;
  };
  const lastClient = (): FakeClient => {
    const client = clients[clients.length - 1];
    assert.ok(client, "expected at least one client");
    return client;
  };

  return {
    clock,
    clients,
    setConnectScript: (fn) => {
      connectScript = fn;
    },
    firstClient,
    lastClient,
    restore: () => {
      resetPiBridgeManagerForTests();
      Reflect.set(globalThis, "window", previousWindow);
      dom.restore();
    },
  };
}

const WELCOME: WelcomeMessage = {
  type: "welcome",
  protocolVersion: 1,
  piVersion: "0.85.1-test",
  serverName: "pi-office-bridge",
  serverVersion: "0.2.0",
  capabilities: ["http-health"],
};

/* ── Tests ─────────────────────────────────────────────────────────── */

void test("enable starts a connect and welcome metadata is captured", async () => {
  const h = setup();
  try {
    await enablePiBridge();
    assert.equal(h.clients.length, 1);
    assert.equal(h.firstClient().connectCount, 1);
    assert.equal(h.firstClient().options.url, "ws://127.0.0.1:38617");
    assert.equal(getPiBridgeState().status, "connecting");

    h.firstClient().callbacks.onWelcome?.(WELCOME);
    h.firstClient().callbacks.onStatusChange?.(true);

    const state = getPiBridgeState();
    assert.equal(state.status, "connected");
    assert.equal(state.enabled, true);
    assert.equal(state.server?.serverVersion, "0.2.0");
    assert.deepEqual(state.server?.capabilities, ["http-health"]);
    assert.equal(state.attempt, 0);
  } finally {
    h.restore();
  }
});

void test("a failed connect schedules a backoff with a typed reason", async () => {
  const h = setup();
  try {
    h.setConnectScript(() =>
      Promise.reject(new BridgeConnectError("refused", "nothing listening")),
    );
    await enablePiBridge();
    await tick();

    const state = getPiBridgeState();
    assert.equal(state.status, "connecting");
    assert.equal(state.attempt, 1);
    assert.equal(state.nextRetryMs, 1_000);
    assert.equal(state.connectFailure, "refused");
    assert.equal(h.clock.pending(), 1);

    // Nothing happens before the delay elapses…
    h.clock.advance(999);
    await tick();
    assert.equal(h.clients.length, 1);

    // …and the retry fires exactly at the delay.
    h.clock.advance(1);
    await tick();
    assert.equal(h.clients.length, 2);
  } finally {
    h.restore();
  }
});

void test("backoff grows 1s → 2s → 4s … and caps at 30s", async () => {
  const h = setup();
  try {
    h.setConnectScript(() =>
      Promise.reject(new BridgeConnectError("refused", "nothing listening")),
    );
    await enablePiBridge();
    await tick();

    const delays: number[] = [];
    for (let i = 0; i < 7; i += 1) {
      const state = getPiBridgeState();
      const delay = state.nextRetryMs;
      assert.ok(typeof delay === "number", "expected a scheduled retry delay");
      delays.push(delay);
      h.clock.advance(delay);
      await tick();
    }

    assert.deepEqual(
      delays,
      [1_000, 2_000, 4_000, 8_000, 16_000, 30_000, 30_000],
    );
    // The loop advanced through attempt 7, firing the next (8th) retry.
    assert.equal(getPiBridgeState().attempt, 8);
  } finally {
    h.restore();
  }
});

void test("backoff resets after 30s of stable connection", async () => {
  const h = setup();
  try {
    h.setConnectScript(() =>
      Promise.reject(new BridgeConnectError("refused", "nothing listening")),
    );
    await enablePiBridge();
    await tick();
    h.clock.advance(1_000);
    await tick();
    h.clock.advance(2_000);
    await tick();
    assert.equal(getPiBridgeState().attempt, 3);

    // Next retry succeeds; a stable connection follows.
    h.setConnectScript(() => Promise.resolve());
    h.clock.advance(4_000);
    await tick();
    const client = h.lastClient();
    client.callbacks.onStatusChange?.(true);
    assert.equal(getPiBridgeState().status, "connected");

    // 30 s of stability resets the backoff…
    h.clock.advance(30_000);
    // …so the next drop starts again at attempt 1 (1 s).
    client.callbacks.onStatusChange?.(false);
    const state = getPiBridgeState();
    assert.equal(state.status, "connecting");
    assert.equal(state.attempt, 1, "backoff should have been reset");
    assert.equal(state.nextRetryMs, 1_000);
  } finally {
    h.restore();
  }
});

void test("disable cancels pending reconnects and clears state", async () => {
  const h = setup();
  try {
    h.setConnectScript(() =>
      Promise.reject(new BridgeConnectError("refused", "nothing listening")),
    );
    await enablePiBridge();
    await tick();
    assert.equal(h.clock.pending(), 1);

    await disablePiBridge();

    const state = getPiBridgeState();
    assert.equal(state.status, "off");
    assert.equal(state.enabled, false);
    assert.equal(state.attempt, 0);
    assert.equal(state.nextRetryMs, undefined);
    assert.equal(state.connectFailure, undefined);
    assert.equal(h.clock.pending(), 0);

    const clientCount = h.clients.length;
    h.clock.advance(120_000);
    await tick();
    assert.equal(h.clients.length, clientCount, "no retry after disable");
  } finally {
    h.restore();
  }
});

void test("an unexpected disconnect reconnects automatically", async () => {
  const h = setup();
  try {
    await enablePiBridge();
    h.firstClient().callbacks.onStatusChange?.(true);
    assert.equal(getPiBridgeState().status, "connected");

    h.firstClient().callbacks.onStatusChange?.(false);
    assert.equal(getPiBridgeState().status, "connecting");
    assert.equal(getPiBridgeState().attempt, 1);

    h.clock.advance(1_000);
    await tick();
    assert.equal(h.clients.length, 2);
    assert.equal(h.firstClient().disconnectCount, 1, "old client is dropped");
  } finally {
    h.restore();
  }
});

void test("callbacks from a superseded client are ignored", async () => {
  const h = setup();
  try {
    await enablePiBridge();
    const stale = h.firstClient();
    await disablePiBridge();
    await enablePiBridge();
    assert.equal(h.clients.length, 2);

    // The old socket's late "connected" must not resurrect the bridge.
    stale.callbacks.onStatusChange?.(true);
    assert.equal(getPiBridgeState().status, "connecting");

    h.lastClient().callbacks.onStatusChange?.(true);
    assert.equal(getPiBridgeState().status, "connected");
  } finally {
    h.restore();
  }
});

void test("a non-fatal server error frame keeps the bridge connected", async () => {
  const h = setup();
  try {
    await enablePiBridge();
    h.firstClient().callbacks.onStatusChange?.(true);
    h.firstClient().callbacks.onServerError?.({
      code: "bad_message",
      message: "bridge: unknown client message type 'nope'",
    });

    const state = getPiBridgeState();
    assert.equal(state.status, "connected");
    assert.equal(state.error, undefined);
  } finally {
    h.restore();
  }
});

void test("an 'online' event retries immediately (cancelling the pending delay)", async () => {
  const h = setup();
  try {
    h.setConnectScript(() =>
      Promise.reject(new BridgeConnectError("refused", "nothing listening")),
    );
    await enablePiBridge();
    await tick();
    assert.equal(getPiBridgeState().attempt, 1);
    assert.equal(h.clients.length, 1);

    Reflect.get(globalThis, "window").dispatchEvent(new Event("online"));
    await tick();

    assert.equal(h.clients.length, 2, "online should retry without waiting");
    assert.equal(
      h.clock.pending(),
      1,
      "the old delay was replaced by the new one",
    );
  } finally {
    h.restore();
  }
});
