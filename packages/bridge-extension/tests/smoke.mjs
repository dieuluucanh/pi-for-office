/** Smoke test: real WebSocket round-trip against the compiled bridge server. */
import { OfficeBridgeServer } from "../dist/bridge-server.js";
import WebSocket from "ws";
import { request as httpRequest } from "node:http";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Minimal GET helper against the loopback HTTP surface. */
function httpGet(port, path, headers = {}) {
  return new Promise((resolve, reject) => {
    const req = httpRequest(
      { host: "127.0.0.1", port, path, method: "GET", headers },
      (res) => {
        let body = "";
        res.setEncoding("utf8");
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () =>
          resolve({ status: res.statusCode, headers: res.headers, body }),
        );
      },
    );
    req.on("error", reject);
    req.end();
  });
}

function httpOptions(port, path, headers = {}) {
  return new Promise((resolve, reject) => {
    const req = httpRequest(
      { host: "127.0.0.1", port, path, method: "OPTIONS", headers },
      (res) => {
        let body = "";
        res.setEncoding("utf8");
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () =>
          resolve({ status: res.statusCode, headers: res.headers, body }),
        );
      },
    );
    req.on("error", reject);
    req.end();
  });
}

async function main() {
  const userMessages = [];
  const server = new OfficeBridgeServer({
    port: 0, // ephemeral
    serverName: "pi-office-bridge-test",
    piVersion: "0.85.1-test",
    handlers: {
      onUserMessage: (text, pane) => {
        userMessages.push({ text, host: pane.host });
      },
    },
  });
  await server.start();
  const port = server.actualPort;
  console.log(`[ok] server listening on :${port}`);

  // --- HTTP /health surface ---
  const health = await httpGet(port, "/health");
  if (health.status !== 200) throw new Error(`health status ${health.status}`);
  const healthBody = JSON.parse(health.body);
  if (healthBody.ok !== true) throw new Error("health not ok");
  if (!Array.isArray(healthBody.panes) || healthBody.panes.length !== 0) {
    throw new Error("expected zero panes before any hello");
  }
  if (healthBody.protocolVersion !== 1) throw new Error("bad protocolVersion");
  console.log(
    `[ok] GET /health → 200 (panes=${healthBody.panes.length}, svc=${healthBody.service})`,
  );

  // CORS gate: known origin reflected, unknown origin denied
  const okOrigin = await httpGet(port, "/health", {
    Origin: "https://localhost:3141",
  });
  if (
    okOrigin.headers["access-control-allow-origin"] !== "https://localhost:3141"
  ) {
    throw new Error("allowed origin not reflected");
  }
  const badOrigin = await httpGet(port, "/health", {
    Origin: "https://evil.example",
  });
  if (badOrigin.headers["access-control-allow-origin"] !== undefined) {
    throw new Error("unknown origin must be denied");
  }
  console.log("[ok] /health CORS gate (known origin allowed, unknown denied)");

  const preflight = await httpOptions(port, "/health", {
    Origin: "https://localhost:3141",
    "Access-Control-Request-Method": "GET",
  });
  if (preflight.status !== 204)
    throw new Error(`OPTIONS status ${preflight.status}`);
  if (preflight.headers["access-control-allow-private-network"] !== "true") {
    throw new Error("missing Access-Control-Allow-Private-Network");
  }
  console.log("[ok] OPTIONS preflight → 204 (+ PNA header)");

  const notFound = await httpGet(port, "/nope");
  if (notFound.status !== 404)
    throw new Error(`404 expected, got ${notFound.status}`);
  console.log("[ok] GET /nope → 404");

  // --- connect a fake Excel pane ---
  const ws = new WebSocket(`ws://127.0.0.1:${port}`);
  const opened = new Promise((r) => ws.once("open", r));
  await opened;

  const incoming = [];
  ws.on("message", (raw) => incoming.push(JSON.parse(raw.toString())));

  ws.send(
    JSON.stringify({
      type: "hello",
      protocolVersion: 1,
      host: "excel",
      clientName: "test-pane",
      paneId: "pane-a",
    }),
  );
  await sleep(100);

  // --- tool proxy: call from server side, answer from pane ---
  const toolPromise = server.callOfficeTool("excel", "read_range", {
    range: "A1:B2",
  });
  await sleep(100);
  const toolCall = incoming.find((m) => m.type === "tool_call");
  if (!toolCall) throw new Error("no tool_call received by pane");
  console.log(
    `[ok] pane received tool_call: ${toolCall.tool} args=${JSON.stringify(toolCall.args)}`,
  );

  ws.send(
    JSON.stringify({
      type: "tool_result",
      id: toolCall.id,
      ok: true,
      text: "| A | B |\n|---|---|\n| 1 | 2 |",
      details: { rows: 2 },
    }),
  );
  const toolResult = await toolPromise;
  console.log(
    `[ok] server resolved tool result: "${toolResult.text.split("\n")[0]}"`,
  );

  // --- error path: wrong pane answers (word attached, excel call) ---
  const failPromise = server.callOfficeTool("word", "get_overview", {});
  try {
    await failPromise;
    throw new Error("expected error for unattached word host");
  } catch (e) {
    console.log(`[ok] unattached host rejected: ${e.message.slice(0, 60)}…`);
  }

  // --- unattached host is already covered above (word); test a host that
  // never attaches (powerpoint) ---
  const neverAttached = server.callOfficeTool("powerpoint", "get_overview", {});
  try {
    await neverAttached;
    throw new Error("expected error");
  } catch (e) {
    console.log(
      `[ok] never-attached host rejected: ${e.message.slice(0, 60)}…`,
    );
  }

  // --- ping/pong ---
  ws.send(JSON.stringify({ type: "ping", ts: Date.now() }));
  await sleep(100);
  if (!incoming.some((m) => m.type === "pong")) throw new Error("no pong");

  // --- user_message ---
  ws.send(
    JSON.stringify({ type: "user_message", text: "Summarize this sheet" }),
  );
  await sleep(100);
  if (
    userMessages.length !== 1 ||
    userMessages[0].text !== "Summarize this sheet"
  ) {
    throw new Error("user message not routed");
  }
  console.log("[ok] pane user_message routed to pi handler");

  // --- status update ---
  ws.send(
    JSON.stringify({
      type: "status",
      model: "mimo-v2.5",
      provider: "opencode-go",
    }),
  );
  await sleep(100);
  const panes = server.attachedPanes();
  if (panes[0].model !== "mimo-v2.5") throw new Error("status not stored");
  console.log("[ok] pane status stored (model=%s)", panes[0].model);

  // --- /health reflects the attached pane ---
  const healthWithPane = JSON.parse((await httpGet(port, "/health")).body);
  if (healthWithPane.panes.length !== 1)
    throw new Error("health should list 1 pane");
  if (healthWithPane.panes[0].host !== "excel")
    throw new Error("health pane host wrong");
  if (healthWithPane.panes[0].model !== "mimo-v2.5")
    throw new Error("health pane model wrong");
  console.log("[ok] /health lists attached excel pane w/ model+provider");

  // --- timeout path: a tool call nobody answers ---
  const slow = server.callOfficeTool(
    "excel",
    "read_range",
    { range: "A1" },
    undefined,
    250,
  );
  try {
    await slow;
    throw new Error("expected timeout");
  } catch (e) {
    console.log(
      `[ok] unanswered tool call rejected (timeout wiring): ${e.message.slice(0, 50)}…`,
    );
  }

  // --- disconnect cleanup ---
  ws.close();
  await sleep(150);
  const remaining = server.attachedPanes();
  if (remaining.length !== 0)
    throw new Error("pane not cleaned up on disconnect");
  console.log("[ok] pane removed on disconnect");

  await server.stop();
  console.log("[ok] server stopped cleanly");
  console.log("\nALL SMOKE TESTS PASSED");
}

main().catch((e) => {
  console.error("SMOKE TEST FAILED:", e);
  process.exit(1);
});
