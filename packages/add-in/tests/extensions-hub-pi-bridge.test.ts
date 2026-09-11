/**
 * DOM test for the "Local Pi agent" card inside the Connections tab.
 *
 * Uses the existing fake-dom + memory settings pattern, plus the manager's
 * test hooks (a fake WebSocket client + virtual clock) so the card's badge
 * and status paragraph are driven by a real but controllable state machine.
 *
 * The three behaviours under test (the original bug reports):
 *  - toggle on → the header badge flips (it used to be a static "Off");
 *  - a failed /health probe never clobbers a Connected badge (both facts are
 *    rendered: badge = WS state, paragraph = probe verdict);
 *  - running a probe while disconnected re-renders the paragraph but keeps
 *    the badge reflecting the manager state.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { renderConnectionsTab } from "../src/commands/builtins/extensions-hub-connections.ts";
import { ConnectionManager } from "../src/connections/manager.ts";
import type { ExtensionsHubDependencies } from "../src/commands/builtins/settings-pages/dependencies.ts";
import {
  resetPiBridgeManagerForTests,
  setPiBridgeManagerHooks,
  type PiBridgeClientLike,
} from "../src/bridge/pi-bridge-manager.ts";
import type { PaneBridgeClientCallbacks } from "../src/bridge/pane-client.ts";
import {
  setAppStorage,
  type AppStorage,
} from "../src/storage/local/app-storage.ts";
import type { SettingsStore } from "../src/storage/local/settings-store.ts";
import type { WelcomeMessage } from "@dieulc/pi-office-protocol";
import { installFakeDom } from "./fake-dom.test.ts";

/* ── Minimal structural view of the fake DOM (document.createElement parts) ── */

interface FakeNodeLike {
  children: FakeNodeLike[];
  className: string;
  textContent: string;
  parentElement: FakeNodeLike | null;
  querySelectorAll(selector: string): Element[];
  dispatchEvent(event: Event): boolean;
}

const asNode = (el: Element): FakeNodeLike =>
  el as DynamicValue as FakeNodeLike;

const tick = (): Promise<void> =>
  new Promise((resolve) => {
    setImmediate(resolve);
  });

/* ── Storage, deps, manager hooks ──────────────────────────────────── */

function createMemorySettings(): SettingsStore {
  const data = new Map<string, DynamicValue>();
  const store: DynamicValue = {
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
  return store as SettingsStore;
}

function installAppStorage(settings: SettingsStore): void {
  const appStorage: DynamicValue = { settings };
  setAppStorage(appStorage as AppStorage);
}

function installManagerHooks(clients: PiBridgeClientLike[]): void {
  resetPiBridgeManagerForTests();
  const clock: { timers: Set<ReturnType<typeof setTimeout>> } = {
    timers: new Set(),
  };
  setPiBridgeManagerHooks({
    createClient: (_options, callbacks: PaneBridgeClientCallbacks) => {
      const client: PiBridgeClientLike = {
        connect: () => Promise.resolve(),
        disconnect: () => {},
      };
      (
        client as DynamicValue as { callbacks: PaneBridgeClientCallbacks }
      ).callbacks = callbacks;
      clients.push(client);
      return client;
    },
    setTimer: (fn, ms) => {
      const handle = setTimeout(fn, ms);
      clock.timers.add(handle);
      return handle;
    },
    clearTimer: (handle) => {
      clock.timers.delete(handle);
      clearTimeout(handle);
    },
  });
}

function createDeps(settings: SettingsStore): ExtensionsHubDependencies {
  const connectionManager = new ConnectionManager({ settings });
  const extensionManager: DynamicValue = { list: () => [] };
  return {
    getActiveSessionId: () => null,
    resolveWorkbookContext: () =>
      Promise.resolve({ workbookId: null, workbookLabel: "" }),
    connectionManager,
    extensionManager:
      extensionManager as ExtensionsHubDependencies["extensionManager"],
  };
}

/* ── Finding the pi-bridge card in the rendered tree ───────────────── */

function findPiBridgeCard(): { card: FakeNodeLike } {
  const doc = globalThis.document as DynamicValue as {
    querySelectorAll(selector: string): Element[];
  };
  const names = doc.querySelectorAll(".pi-item-card__name");
  assert.ok(names.length > 0, "expected at least one item card");
  const nameEl = names.find(
    (el) => asNode(el).textContent === "Local Pi agent (advanced)",
  );
  assert.ok(nameEl, "pi-bridge card not rendered");

  let card: FakeNodeLike | null = asNode(nameEl).parentElement;
  while (card && !card.className.split(/\s+/u).includes("pi-item-card")) {
    card = card.parentElement;
  }
  assert.ok(card, "pi-bridge card root not found in the tree");
  return { card };
}

function badgeOf(card: FakeNodeLike): FakeNodeLike {
  const visit = (node: FakeNodeLike): FakeNodeLike | null => {
    if (node.className.split(/\s+/u).includes("pi-overlay-badge")) {
      return node;
    }
    for (const child of node.children) {
      const found = visit(child);
      if (found) return found;
    }
    return null;
  };
  const badge = visit(card);
  assert.ok(badge, "badge not found in the pi-bridge card");
  return badge;
}

function toggleInputOf(card: FakeNodeLike): HTMLInputElement {
  const visit = (node: FakeNodeLike): HTMLInputElement | null => {
    if (node.className.split(/\s+/u).includes("pi-toggle__input")) {
      return node as DynamicValue as HTMLInputElement;
    }
    for (const child of node.children) {
      const found = visit(child);
      if (found) return found;
    }
    return null;
  };
  const input = visit(card);
  assert.ok(input, "pi-bridge toggle input not found");
  return input;
}

function statusTextsOf(card: FakeNodeLike): string[] {
  const texts: string[] = [];
  const visit = (node: FakeNodeLike): void => {
    if (node.className.split(/\s+/u).includes("pi-hub-bridge-setup__hint")) {
      texts.push(node.textContent);
    }
    for (const child of node.children) {
      visit(child);
    }
  };
  visit(card);
  return texts;
}

function probeButtonOf(card: FakeNodeLike): HTMLButtonElement {
  const buttons = card.querySelectorAll("button");
  const btn = Array.from(buttons).find(
    (b) => asNode(b).textContent === "Test connection",
  );
  assert.ok(btn, "Test connection button not found");
  return btn as DynamicValue as HTMLButtonElement;
}

function toneOf(badge: FakeNodeLike): string {
  const match = badge.className.match(/pi-overlay-badge--(\w+)/u);
  assert.ok(match, `badge has no tone class: ${badge.className}`);
  return match[1] as string;
}

function connectLastClient(clients: PiBridgeClientLike[]): void {
  const client = clients[clients.length - 1];
  assert.ok(client, "expected a client");
  const callbacks = (
    client as DynamicValue as {
      callbacks: PaneBridgeClientCallbacks;
    }
  ).callbacks;
  const welcome: WelcomeMessage = {
    type: "welcome",
    protocolVersion: 1,
    piVersion: "0.85.1-test",
    serverName: "pi-office-bridge",
    serverVersion: "0.2.0",
    capabilities: ["http-health"],
  };
  callbacks.onWelcome?.(welcome);
  callbacks.onStatusChange?.(true);
}

async function renderCard(settings: SettingsStore): Promise<void> {
  const container = document.createElement("div");
  document.body.appendChild(container);
  await renderConnectionsTab({
    container,
    settings,
    deps: createDeps(settings),
    isBusy: () => false,
    runMutation: async (action) => {
      await action();
    },
  });
}

/* ── Tests ─────────────────────────────────────────────────────────── */

void test("toggle on flips the header badge to Connected (ok tone)", async () => {
  const dom = installFakeDom();
  const clients: PiBridgeClientLike[] = [];
  try {
    const settings = createMemorySettings();
    installAppStorage(settings);
    installManagerHooks(clients);
    await renderCard(settings);

    const { card } = findPiBridgeCard();
    const badge = badgeOf(card);
    assert.equal(badge.textContent, "Off");
    assert.equal(toneOf(badge), "muted");

    // Toggle on → manager connects (fake client resolves) → Connecting.
    const toggle = toggleInputOf(card);
    toggle.checked = true;
    toggle.dispatchEvent(new Event("change"));
    await tick();
    assert.equal(badgeOf(card).textContent, "Connecting…");
    assert.equal(toneOf(badgeOf(card)), "info");

    // Server welcome + connected → badge goes Connected/ok.
    connectLastClient(clients);
    await tick();
    const connectedBadge = badgeOf(card);
    assert.equal(connectedBadge.textContent, "Connected");
    assert.equal(toneOf(connectedBadge), "ok");
    assert.ok(
      statusTextsOf(card).some((t) =>
        t.includes("Connected to the Local Pi agent"),
      ),
      "status paragraph shows the connection state",
    );
  } finally {
    resetPiBridgeManagerForTests();
    dom.restore();
  }
});

void test("a failed probe never clobbers a Connected badge (both facts shown)", async () => {
  const dom = installFakeDom();
  const clients: PiBridgeClientLike[] = [];
  const previousFetch = Reflect.get(globalThis, "fetch");
  try {
    const settings = createMemorySettings();
    installAppStorage(settings);
    installManagerHooks(clients);
    await renderCard(settings);

    const { card } = findPiBridgeCard();
    const toggle = toggleInputOf(card);
    toggle.checked = true;
    toggle.dispatchEvent(new Event("change"));
    await tick();
    connectLastClient(clients);
    await tick();
    assert.equal(badgeOf(card).textContent, "Connected");

    // Probe fails at the fetch level (browser-style TypeError) →
    // classified "blocked" → the badge must stay Connected.
    Reflect.set(globalThis, "fetch", () =>
      Promise.reject(new TypeError("Failed to fetch")),
    );
    probeButtonOf(card).dispatchEvent(new Event("click"));
    await tick();
    await tick();

    const badge = badgeOf(card);
    assert.equal(badge.textContent, "Connected", "WS state is authoritative");
    assert.equal(toneOf(badge), "ok");
    assert.ok(
      statusTextsOf(card).some(
        (t) => t.includes("Connected over WebSocket") && t.includes("blocked"),
      ),
      "both facts rendered in the same paragraph",
    );
  } finally {
    Reflect.set(globalThis, "fetch", previousFetch);
    resetPiBridgeManagerForTests();
    dom.restore();
  }
});

void test("a probe while disconnected re-renders the paragraph but not the badge", async () => {
  const dom = installFakeDom();
  const clients: PiBridgeClientLike[] = [];
  const previousFetch = Reflect.get(globalThis, "fetch");
  try {
    const settings = createMemorySettings();
    installAppStorage(settings);
    installManagerHooks(clients);
    await renderCard(settings);

    const { card } = findPiBridgeCard();
    assert.equal(badgeOf(card).textContent, "Off");

    // A healthy /health response while the bridge is off → the paragraph
    // shows the probe verdict, the badge still says the manager state.
    Reflect.set(globalThis, "fetch", () =>
      Promise.resolve({
        ok: true,
        json: () =>
          Promise.resolve({ ok: true, serverVersion: "0.2.0", panes: [{}] }),
      }),
    );
    probeButtonOf(card).dispatchEvent(new Event("click"));
    await tick();
    await tick();

    const badge = badgeOf(card);
    assert.equal(
      badge.textContent,
      "Off",
      "badge still reflects the manager state",
    );
    assert.equal(toneOf(badge), "muted");
    assert.ok(
      statusTextsOf(card).some((t) => t.includes("HTTP /health OK")),
      "probe verdict rendered in the paragraph",
    );
  } finally {
    Reflect.set(globalThis, "fetch", previousFetch);
    resetPiBridgeManagerForTests();
    dom.restore();
  }
});
