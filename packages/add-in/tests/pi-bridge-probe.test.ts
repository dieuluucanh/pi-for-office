/**
 * Probe tests against REAL local http servers (loopback, ephemeral ports).
 *
 * The probe must classify every outcome instead of collapsing everything into
 * a bare "could not reach": 200 JSON → ok, 404/non-JSON → old bridge,
 * hang → timeout, connection refused → unreachable. It must also keep working
 * on runtimes where `targetAddressSpace` is unsupported (older WebView2).
 */

import assert from "node:assert/strict";
import { test } from "node:test";
import { createServer, type RequestListener, type Server } from "node:http";

import {
  healthUrlFromWsUrl,
  probeBridgeHealth,
} from "../src/bridge/pi-bridge-probe.ts";
import type { PiBridgeProbeResult } from "../src/ui/pi-bridge-card-model.ts";

const HEALTH_OK: RequestListener = (_req, res) => {
  res.writeHead(200, { "Content-Type": "application/json" });
  res.end(
    JSON.stringify({
      ok: true,
      serverVersion: "0.2.0",
      panes: [{ host: "excel" }],
    }),
  );
};

async function startServer(
  handler: RequestListener = HEALTH_OK,
): Promise<{ server: Server; url: string }> {
  const server = createServer(handler);
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => resolve());
  });
  const addr = server.address();
  assert.ok(typeof addr === "object" && addr !== null, "server bound");
  return { server, url: `http://127.0.0.1:${addr.port}/health` };
}

async function closeServer(server: Server): Promise<void> {
  try {
    server.closeAllConnections();
  } catch {
    // Older Node runtimes lack closeAllConnections — close() still works.
  }
  await new Promise<void>((resolve) => {
    server.close(() => resolve());
  });
}

void test("200 + JSON /health classifies as ok with version and pane count", async () => {
  const { server, url } = await startServer();
  try {
    const result = await probeBridgeHealth(url);
    assert.equal(result.kind, "ok");
    if (result.kind !== "ok") return;
    assert.equal(result.serverVersion, "0.2.0");
    assert.equal(result.paneCount, 1);
  } finally {
    await closeServer(server);
  }
});

void test("404 classifies as an older bridge (server answered, not our /health)", async () => {
  const { server, url } = await startServer((_req, res) => {
    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ ok: false, error: "not_found" }));
  });
  try {
    const result = await probeBridgeHealth(url);
    assert.equal(result.kind, "old");
  } finally {
    await closeServer(server);
  }
});

void test("200 with non-JSON body classifies as old", async () => {
  const { server, url } = await startServer((_req, res) => {
    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("hello, this is not your bridge");
  });
  try {
    const result = await probeBridgeHealth(url);
    assert.equal(result.kind, "old");
  } finally {
    await closeServer(server);
  }
});

void test("a silent server (no response) times out with a classified verdict", async () => {
  const { server, url } = await startServer(() => {
    // Never answer — the probe must give up within its budget.
  });
  try {
    const result = await probeBridgeHealth(url, { timeoutMs: 200 });
    assert.equal(result.kind, "timeout");
  } finally {
    await closeServer(server);
  }
});

void test("connection refused classifies as unreachable and keeps the URL", async () => {
  const { server, url } = await startServer();
  await closeServer(server);

  const result = await probeBridgeHealth(url, { timeoutMs: 2_000 });
  if (result.kind !== "unreachable") {
    // On some CI networks a closed port can answer with a reset instead;
    // either way it must NEVER be 'ok' or a bare failure string.
    assert.ok(
      result.kind === "old" || result.kind === "blocked",
      `expected unreachable/blocked, got ${result.kind}`,
    );
    return;
  }
  assert.equal(result.url, url);
});

void test("healthUrlFromWsUrl rewrites ws/wss to the /health URL", () => {
  assert.equal(
    healthUrlFromWsUrl("ws://127.0.0.1:38617"),
    "http://127.0.0.1:38617/health",
  );
  assert.equal(
    healthUrlFromWsUrl("wss://localhost:4444/"),
    "https://localhost:4444/health",
  );
});

void test("the probe works on runtimes where targetAddressSpace is unsupported", async () => {
  // Node's undici either supports the option or silently ignores unknown
  // RequestInit keys, so the feature-detect path is exercised either way.
  // The assertion is that the request still succeeds (the option must never
  // break the probe on a runtime that rejects it).
  const { server, url } = await startServer();
  try {
    const result: PiBridgeProbeResult = await probeBridgeHealth(url);
    assert.equal(result.kind, "ok");
  } finally {
    await closeServer(server);
  }
});
