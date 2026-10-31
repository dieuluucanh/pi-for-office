/**
 * Server-level lifecycle regressions:
 * - a failed bind does not leave `isRunning` true or a half-open server behind
 * - `startedAt`/`uptimeMs` restart with the server
 */

import assert from "node:assert/strict";
import { test } from "node:test";
import { request as httpRequest } from "node:http";

import { OfficeBridgeServer } from "../dist/bridge-server.js";

function httpGet(port, path) {
  return new Promise((resolve, reject) => {
    const req = httpRequest({ host: "127.0.0.1", port, path, method: "GET" }, (res) => {
      let body = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => (body += chunk));
      res.on("end", () => resolve({ status: res.statusCode, body }));
    });
    req.on("error", reject);
    req.end();
  });
}

function makeServer(port) {
  return new OfficeBridgeServer({
    port,
    serverName: "pi-office-bridge-test",
    serverVersion: "0.6.0-test",
    handlers: {
      onUserMessage: () => undefined,
    },
  });
}

test("a failed bind leaves the server off and stop() resolves", async () => {
  const owner = makeServer(0);
  await owner.start();
  const port = owner.actualPort;
  assert.ok(port);

  const blocked = makeServer(port);
  await assert.rejects(() => blocked.start(), (error) => error.code === "EADDRINUSE");
  assert.equal(blocked.isRunning, false);
  assert.equal(blocked.actualPort, null);
  await assert.doesNotReject(() => blocked.stop());

  await owner.stop();
});

test("startedAt resets on restart and health reports a fresh uptime", async () => {
  const server = makeServer(0);
  await server.start();
  const firstPort = server.actualPort;
  const first = JSON.parse((await httpGet(firstPort, "/health")).body);
  assert.equal(typeof first.startedAt, "number");
  assert.ok(first.uptimeMs >= 0);

  await server.stop();
  await new Promise((resolve) => setTimeout(resolve, 20));
  await server.start();
  const secondPort = server.actualPort;
  const second = JSON.parse((await httpGet(secondPort, "/health")).body);
  assert.ok(second.startedAt >= first.startedAt);
  assert.ok(second.uptimeMs < 5_000);
  assert.equal(second.owner.pid, process.pid);
  assert.ok(Array.isArray(second.capabilities));
  assert.ok(second.capabilities.includes("manual-lifecycle"));

  await server.stop();
});
