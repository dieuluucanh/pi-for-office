/**
 * Extension-level tests with a fake Pi API.
 *
 * These pin the user-visible contract: a fresh session does not start the
 * server, `/office connect` and `/office close` drive the lifecycle, and the
 * active tool family follows it.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { createOfficeBridgeExtension } from "../dist/index.js";

const PANE = {
  ws: null,
  host: "excel",
  paneId: "pane-a",
  clientName: "test-pane",
  protocolVersion: 3,
  connectedAt: 1,
  lastSeen: 2,
  ops: ["excel.read_range"],
  catalogVersion: 3,
  opsIgnoredCount: 0,
  documentId: "url_sha256:abc",
  documentName: "Book1.xlsx",
  documentSource: "document.url",
  targetId: "url_sha256:abc",
  focused: true,
  focusUpdatedAt: 2,
  visible: true,
};

function fakeServer(port, handlers, servers) {
  const server = {
    isRunning: false,
    actualPort: null,
    panes: [],
    starts: 0,
    stops: 0,
    handlers,
    attachedPanes() {
      return server.panes;
    },
    async start() {
      server.starts += 1;
      server.isRunning = true;
      server.actualPort = port;
    },
    async stop() {
      server.stops += 1;
      server.isRunning = false;
      server.actualPort = null;
      server.panes = [];
      handlers.onPanesChanged?.(server.panes);
    },
    async callOfficeTool() {
      return { text: "ok" };
    },
    sendToPane() {
      return true;
    },
  };
  servers.push(server);
  return server;
}

function makeHarness(options = {}) {
  const events = new Map();
  const tools = new Map();
  const commands = new Map();
  const flagDefaults = new Map();
  const flags = new Map();
  const servers = [];
  const notifications = [];
  const statuses = new Map();
  const sent = [];
  let activeTools = ["bash", "read"];

  const api = {
    on(event, handler) {
      const list = events.get(event) ?? [];
      list.push(handler);
      events.set(event, list);
      return () => undefined;
    },
    registerTool(tool) {
      tools.set(tool.name, tool);
    },
    registerCommand(name, command) {
      commands.set(name, command);
    },
    registerFlag(name, flag) {
      if (flag.default !== undefined) flagDefaults.set(name, flag.default);
    },
    getFlag(name) {
      return flags.has(name) ? flags.get(name) : flagDefaults.get(name);
    },
    getActiveTools() {
      return [...activeTools];
    },
    setActiveTools(next) {
      activeTools = [...next];
    },
    getAllTools() {
      return [...tools.values()];
    },
    sendUserMessage(text) {
      sent.push(text);
    },
  };

  const ctx = {
    ui: {
      notify(message, type) {
        notifications.push({ message, type });
      },
      setStatus(key, text) {
        statuses.set(key, text);
      },
      addAutocompleteProvider() {
        // The provider is irrelevant to lifecycle behaviour.
      },
    },
    abort() {
      // Pane abort is covered by pane-replies tests.
    },
  };

  createOfficeBridgeExtension(api, {
    env: options.env ?? {},
    now: () => 1000,
    createServer: (port, handlers) => fakeServer(port, handlers, servers),
  });

  return {
    api,
    ctx,
    events,
    tools,
    commands,
    flags,
    servers,
    notifications,
    statuses,
    activeTools: () => activeTools,
    async startSession() {
      const handler = events.get("session_start")?.[0];
      assert.ok(handler, "session_start handler must be registered");
      await handler({ type: "session_start" }, ctx);
    },
    async shutdown() {
      const handler = events.get("session_shutdown")?.[0];
      assert.ok(handler, "session_shutdown handler must be registered");
      await handler({ type: "session_shutdown", reason: "quit" }, ctx);
    },
    async runCommand(args) {
      const command = commands.get("office");
      assert.ok(command, "/office command must be registered");
      return command.handler(args, ctx);
    },
  };
}

test("a fresh session does not start the server or advertise office tools", async () => {
  const harness = makeHarness();
  await harness.startSession();
  assert.equal(harness.servers.length, 0);
  assert.deepEqual(harness.activeTools(), ["bash", "read"]);
  assert.equal(harness.statuses.get("office-bridge"), undefined);
});

test("autostart flag restores the old start-on-session behaviour", async () => {
  const harness = makeHarness();
  harness.flags.set("office-bridge-autostart", true);
  await harness.startSession();
  assert.equal(harness.servers.length, 1);
  assert.equal(harness.servers[0].isRunning, true);
  assert.ok(harness.activeTools().includes("office_list_documents"));
  assert.ok(harness.activeTools().includes("office_excel_read_range") === false);
});

test("autostart env value works and an explicit flag is not required", async () => {
  const harness = makeHarness({ env: { PI_OFFICE_BRIDGE_AUTOSTART: "1" } });
  await harness.startSession();
  assert.equal(harness.servers.length, 1);
  assert.equal(harness.servers[0].isRunning, true);
});

test("/office connect starts the bridge and activates only discovery until a pane attaches", async () => {
  const harness = makeHarness();
  await harness.startSession();
  await harness.runCommand("connect 38617");
  assert.equal(harness.servers.length, 1);
  assert.equal(harness.servers[0].isRunning, true);
  assert.ok(harness.activeTools().includes("office_list_documents"));
  assert.ok(!harness.activeTools().includes("office_excel_read_range"));
  assert.match(
    harness.notifications.at(-1)?.message ?? "",
    /listening on ws:\/\/127\.0\.0\.1:38617/,
  );
});

test("a pane attaches: its advertised ops become active", async () => {
  const harness = makeHarness();
  await harness.startSession();
  await harness.runCommand("connect 38617");
  const server = harness.servers[0];
  server.panes.push(PANE);
  server.handlers.onPanesChanged(server.panes);
  assert.ok(harness.activeTools().includes("office_list_documents"));
  assert.ok(harness.activeTools().includes("office_excel_read_range"));
});

test("/office close stops the server and removes every bridge tool", async () => {
  const harness = makeHarness();
  await harness.startSession();
  await harness.runCommand("connect 38617");
  const server = harness.servers[0];
  server.panes.push(PANE);
  server.handlers.onPanesChanged(server.panes);
  await harness.runCommand("close");
  assert.equal(server.isRunning, false);
  assert.equal(server.stops, 1);
  assert.deepEqual(
    harness.activeTools().filter((name) => name.startsWith("office_")),
    [],
  );
  assert.equal(harness.statuses.get("office-bridge"), undefined);
});

test("/office status says how to start when off", async () => {
  const harness = makeHarness();
  await harness.startSession();
  await harness.runCommand("status");
  assert.match(harness.notifications.at(-1)?.message ?? "", /is off/);
  assert.match(harness.notifications.at(-1)?.message ?? "", /\/office connect/);
});

test("session_shutdown stops a connected bridge", async () => {
  const harness = makeHarness();
  await harness.startSession();
  await harness.runCommand("connect 38617");
  const server = harness.servers[0];
  await harness.shutdown();
  assert.equal(server.isRunning, false);
  assert.equal(server.stops, 1);
});

test("office tools refuse to run while the bridge is off", async () => {
  const harness = makeHarness();
  await harness.startSession();
  const tool = harness.tools.get("office_excel_read_range");
  assert.ok(tool);
  await assert.rejects(
    () => tool.execute("call-1", {}, undefined, undefined, harness.ctx),
    /server is not running/,
  );
  const list = harness.tools.get("office_list_documents");
  assert.ok(list);
  await assert.rejects(
    () => list.execute("call-2", {}, undefined, undefined, harness.ctx),
    /server is not running/,
  );
});

test("invalid subcommands are thrown so Pi Desktop receives the reason", async () => {
  const harness = makeHarness();
  await harness.startSession();
  await assert.rejects(() => harness.runCommand("pause"), /Unknown subcommand/);
});

test("/office-tools marks the family off until connect", async () => {
  const harness = makeHarness();
  await harness.startSession();
  const command = harness.commands.get("office-tools");
  assert.ok(command);
  await command.handler("", harness.ctx);
  const text = harness.notifications.at(-1)?.message ?? "";
  assert.match(text, /\[off\] office_excel_read_range/);
  assert.match(text, /\[off\] office_list_documents/);
});
