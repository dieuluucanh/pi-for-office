/**
 * Unit tests for the on-demand bridge lifecycle controller (bridge 0.6+).
 * Run after `npm run build` (imports the compiled dist/lifecycle.js).
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  BridgeLifecycleError,
  createBridgeController,
  describeBridgeStartError,
  normalizeBridgePort,
} from "../dist/lifecycle.js";

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function fakeServer(options = {}) {
  const server = {
    isRunning: false,
    actualPort: null,
    panes: options.panes ?? [],
    starts: 0,
    stops: 0,
    attachedPanes() {
      return server.panes;
    },
    async start() {
      server.starts += 1;
      if (options.startDelay) await options.startDelay.promise;
      if (options.fail) throw options.fail;
      server.isRunning = true;
      server.actualPort = options.actualPort ?? options.port ?? 38617;
    },
    async stop() {
      server.stops += 1;
      if (options.stopDelay) await options.stopDelay.promise;
      if (options.stopFail) throw options.stopFail;
      server.isRunning = false;
      server.actualPort = null;
    },
  };
  return server;
}

function controllerFor(servers, options = {}) {
  const changes = [];
  let index = 0;
  const controller = createBridgeController({
    defaultPort: 38617,
    now: () => 1000,
    createServer(port) {
      const server = fakeServer({ ...options, port });
      servers.push(server);
      return server;
    },
    onChange: (snapshot) => changes.push(snapshot),
    ...options.controller,
  });
  return { controller, changes, next: () => servers[index++] };
}

test("starts off and reports the default port", () => {
  const servers = [];
  const { controller, changes } = controllerFor(servers);
  assert.deepEqual(controller.snapshot(), {
    state: "off",
    port: 38617,
    actualPort: null,
    panes: 0,
    lastError: null,
    startedAt: null,
  });
  assert.equal(changes.length, 0);
});

test("connect starts once, reports on, and is idempotent", async () => {
  const servers = [];
  const { controller, changes } = controllerFor(servers);
  const first = await controller.connect();
  assert.equal(first.state, "on");
  assert.equal(first.actualPort, 38617);
  assert.equal(first.startedAt, 1000);
  assert.equal(servers.length, 1);
  assert.equal(servers[0].starts, 1);

  const second = await controller.connect();
  assert.equal(second.state, "on");
  assert.equal(servers.length, 1);
  assert.equal(servers[0].starts, 1);
  assert.deepEqual(
    changes.map((entry) => entry.state),
    ["starting", "on"],
  );
});

test("concurrent connect calls share one start", async () => {
  const startDelay = deferred();
  const servers = [];
  const { controller } = controllerFor(servers, { startDelay });
  const a = controller.connect();
  const b = controller.connect();
  startDelay.resolve();
  const [first, second] = await Promise.all([a, b]);
  assert.equal(first.state, "on");
  assert.equal(second.state, "on");
  assert.equal(servers.length, 1);
  assert.equal(servers[0].starts, 1);
});

test("connect on an explicit port while on a different port asks for restart", async () => {
  const servers = [];
  const { controller } = controllerFor(servers);
  await controller.connect();
  await assert.rejects(
    () => controller.connect(38620),
    (error) =>
      error instanceof BridgeLifecycleError &&
      error.code === "already_running" &&
      /restart 38620/.test(error.message),
  );
  assert.equal(servers.length, 1);
});

test("a failed start returns to off, records the reason and can be retried", async () => {
  const servers = [];
  const { controller, changes } = controllerFor(servers, {
    fail: Object.assign(new Error("address already in use"), { code: "EADDRINUSE" }),
  });
  await assert.rejects(() => controller.connect(), /port 38617 is already in use/);
  const failed = controller.snapshot();
  assert.equal(failed.state, "off");
  assert.equal(failed.actualPort, null);
  assert.match(failed.lastError, /already in use/);
  assert.equal(servers.length, 1);
  assert.deepEqual(
    changes.map((entry) => entry.state),
    ["starting", "off"],
  );

  // A retry creates a fresh server instead of reusing the failed one.
  const retry = controllerFor(servers, {});
  const result = await retry.controller.connect();
  assert.equal(result.state, "on");
});

test("close stops the server, clears panes and is idempotent", async () => {
  const servers = [];
  const { controller, changes } = controllerFor(servers, {
    panes: [{ host: "excel" }],
  });
  await controller.connect();
  assert.equal(controller.snapshot().panes, 1);
  const closed = await controller.close();
  assert.equal(closed.state, "off");
  assert.equal(closed.panes, 0);
  assert.equal(servers[0].stops, 1);
  await controller.close();
  assert.equal(servers[0].stops, 1);
  assert.deepEqual(
    changes.map((entry) => entry.state),
    ["starting", "on", "closing", "off"],
  );
});

test("restart closes the old server and binds the requested port", async () => {
  const servers = [];
  const { controller } = controllerFor(servers);
  await controller.connect();
  const restarted = await controller.restart(38620);
  assert.equal(restarted.state, "on");
  assert.equal(restarted.port, 38620);
  assert.equal(restarted.actualPort, 38620);
  assert.equal(servers.length, 2);
  assert.equal(servers[0].stops, 1);
  assert.equal(servers[1].starts, 1);
});

test("close while starting waits for the start and then stops it", async () => {
  const startDelay = deferred();
  const servers = [];
  const { controller } = controllerFor(servers, { startDelay });
  const connecting = controller.connect();
  const closing = controller.close();
  startDelay.resolve();
  await connecting;
  const result = await closing;
  assert.equal(result.state, "off");
  assert.equal(servers[0].starts, 1);
  assert.equal(servers[0].stops, 1);
});

test("a stop failure still ends off and records the reason", async () => {
  const servers = [];
  const { controller } = controllerFor(servers, { stopFail: new Error("socket stuck") });
  await controller.connect();
  await assert.rejects(() => controller.close(), /did not close cleanly/);
  const state = controller.snapshot();
  assert.equal(state.state, "off");
  assert.match(state.lastError, /socket stuck/);
});

test("normalizeBridgePort validates the command argument", () => {
  assert.equal(normalizeBridgePort(undefined), null);
  assert.equal(normalizeBridgePort(null), null);
  assert.equal(normalizeBridgePort(38617), 38617);
  for (const bad of [0, -1, 65536, 1.5, Number.NaN]) {
    assert.throws(
      () => normalizeBridgePort(bad),
      (error) => error instanceof BridgeLifecycleError && error.code === "bad_port",
    );
  }
});

test("describeBridgeStartError names the port owner for EADDRINUSE", () => {
  const message = describeBridgeStartError(
    Object.assign(new Error("bind failed"), { code: "EADDRINUSE" }),
    38617,
  );
  assert.match(message, /port 38617 is already in use/);
  assert.match(message, /\/office status/);
  assert.match(message, /\/office connect <port>/);
});
