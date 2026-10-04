import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import http from "node:http";
import net from "node:net";
import { fileURLToPath } from "node:url";
import { setTimeout as delay } from "node:timers/promises";

const ORIGIN = "https://localhost:3141";
const PROXY_SCRIPT_PATH = fileURLToPath(
  new URL("../scripts/cors-proxy-server.mjs", import.meta.url),
);

async function getFreePort() {
  return new Promise((resolve, reject) => {
    const srv = net.createServer();
    srv.once("error", reject);
    srv.listen(0, "127.0.0.1", () => {
      const addr = srv.address();
      if (!addr || typeof addr === "string") {
        srv.close();
        reject(new Error("Failed to get free port"));
        return;
      }
      const { port } = addr;
      srv.close((err) => {
        if (err) reject(err);
        else resolve(port);
      });
    });
  });
}

async function stopChild(child) {
  if (!child.killed) {
    child.kill("SIGTERM");
  }

  const exited = Promise.race([
    once(child, "exit"),
    delay(2000).then(() => {
      if (!child.killed) {
        child.kill("SIGKILL");
      }
    }),
  ]);

  await exited.catch(() => {});
}

async function startProxy(extraEnv = {}) {
  const port = await getFreePort();

  const child = spawn(process.execPath, [PROXY_SCRIPT_PATH], {
    env: {
      ...process.env,
      HOST: "127.0.0.1",
      PORT: String(port),
      ALLOWED_ORIGINS: ORIGIN,
      OAUTH_CALLBACK_SERVER: "0",
      ALLOW_LOOPBACK_TARGETS: "1",
      ALLOW_PRIVATE_TARGETS: "1",
      ...extraEnv,
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");

  const ready = new Promise((resolve, reject) => {
    let stdout = "";
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
      if (
        stdout.includes("CORS proxy listening on") &&
        (stdout.includes("Allowed target hosts") ||
          stdout.includes("target host allowlisting disabled"))
      ) {
        resolve(undefined);
      }
    });
    child.once("exit", (code, signal) => {
      reject(
        new Error(
          `proxy exited before ready (code=${String(code)} signal=${String(signal)})`,
        ),
      );
    });
  });

  await Promise.race([
    ready,
    delay(5000).then(() => {
      throw new Error("proxy start timeout");
    }),
  ]);

  return {
    port,
    stop: () => stopChild(child),
  };
}

/**
 * A loopback HTTP target driven by a request handler. Records upstream
 * request-close events with an unfinished response, so tests can assert that
 * the proxy tore down the upstream request.
 */
async function startMockTarget(handler) {
  const port = await getFreePort();
  const lifecycle = { upstreamRequestsClosedPrematurely: 0 };

  const server = http.createServer(async (upstreamReq, res) => {
    upstreamReq.once("close", () => {
      if (!res.writableEnded) {
        lifecycle.upstreamRequestsClosedPrematurely += 1;
      }
    });
    await handler(upstreamReq, res);
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", () => resolve(undefined));
  });

  const stop = async () => {
    await new Promise((resolve) => {
      server.close(() => resolve(undefined));
      server.closeAllConnections();
    });
  };

  return { port, stop, lifecycle };
}

test("proxied non-streaming responses still work end-to-end", async (t) => {
  const mockTarget = await startMockTarget(async (_req, res) => {
    res.setHeader("Content-Type", "application/json");
    res.end('{"ok":true}');
  });
  t.after(async () => {
    await mockTarget.stop();
  });

  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const proxiedUrl = encodeURIComponent(
    `http://127.0.0.1:${mockTarget.port}/v1/models`,
  );
  const response = await fetch(
    `http://127.0.0.1:${proxy.port}/?url=${proxiedUrl}`,
    { headers: { Origin: ORIGIN } },
  );

  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), { ok: true });
});

test("slow-drip SSE responses pass through without being cut", async (t) => {
  const mockTarget = await startMockTarget(async (_req, res) => {
    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache");
    res.flushHeaders();
    for (let index = 0; index < 5; index += 1) {
      res.write(`data: chunk-${index}\n\n`);
      await delay(200);
    }
    res.end();
  });
  t.after(async () => {
    await mockTarget.stop();
  });

  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const proxiedUrl = encodeURIComponent(
    `http://127.0.0.1:${mockTarget.port}/v1/chat/completions`,
  );
  const response = await fetch(
    `http://127.0.0.1:${proxy.port}/?url=${proxiedUrl}`,
    { headers: { Origin: ORIGIN } },
  );

  assert.equal(response.status, 200);
  const text = await response.text();
  for (let index = 0; index < 5; index += 1) {
    assert.ok(
      text.includes(`data: chunk-${index}`),
      `expected chunk-${index} in the proxied body`,
    );
  }
});

test("client abort propagates upstream so abandoned streams are torn down", async (t) => {
  const mockTarget = await startMockTarget(async (upstreamReq, res) => {
    res.setHeader("Content-Type", "text/event-stream");
    res.flushHeaders();
    // Drip forever; the test aborts after the first chunk.
    let index = 0;
    const drip = setInterval(() => {
      res.write(`data: drip-${index}\n\n`);
      index += 1;
    }, 100);
    upstreamReq.once("close", () => clearInterval(drip));
  });
  t.after(async () => {
    await mockTarget.stop();
  });

  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const proxiedUrl = encodeURIComponent(
    `http://127.0.0.1:${mockTarget.port}/v1/chat/completions`,
  );
  const abort = new AbortController();
  const response = await fetch(
    `http://127.0.0.1:${proxy.port}/?url=${proxiedUrl}`,
    { headers: { Origin: ORIGIN }, signal: abort.signal },
  );

  assert.equal(response.status, 200);
  // Read the first chunk, then abandon the stream.
  const reader = response.body.getReader();
  await reader.read();
  abort.abort();

  // The proxy must tear down the upstream request: the mock target records
  // request close events with an unfinished response.
  await delay(1000);
  assert.ok(
    mockTarget.lifecycle.upstreamRequestsClosedPrematurely >= 1,
    "expected the proxied upstream request to be closed prematurely after the client aborted",
  );
});

test("upstream errors mid-response surface as real connection failures", async (t) => {
  const mockTarget = await startMockTarget(async (_req, res) => {
    res.setHeader("Content-Type", "text/event-stream");
    res.flushHeaders();
    res.write("data: first\n\n");
    await delay(200);
    // Cut the upstream connection mid-stream (provider disconnect).
    res.destroy(new Error("upstream cut"));
  });
  t.after(async () => {
    await mockTarget.stop();
  });

  const proxy = await startProxy();
  t.after(async () => {
    await proxy.stop();
  });

  const proxiedUrl = encodeURIComponent(
    `http://127.0.0.1:${mockTarget.port}/v1/chat/completions`,
  );
  const response = await fetch(
    `http://127.0.0.1:${proxy.port}/?url=${proxiedUrl}`,
    { headers: { Origin: ORIGIN } },
  );

  assert.equal(response.status, 200);
  // The proxied body must fail — not end cleanly — when the upstream dies
  // mid-response. A clean close would silently truncate SSE responses.
  const reader = response.body.getReader();
  const first = await reader.read();
  assert.equal(first.done, false);

  let bodyFailed = false;
  try {
    for (;;) {
      const read = await reader.read();
      if (read.done) break;
    }
  } catch {
    bodyFailed = true;
  }
  assert.ok(
    bodyFailed,
    "expected the proxied body read to fail when the upstream connection died mid-response",
  );
});
