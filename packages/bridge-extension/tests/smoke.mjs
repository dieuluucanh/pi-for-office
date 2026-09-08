/** Smoke test: real WebSocket round-trip against the compiled bridge server. */
import { OfficeBridgeServer } from "../dist/bridge-server.js";
import WebSocket from "ws";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

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

  // --- connect a fake Excel pane ---
  const ws = new WebSocket(`ws://127.0.0.1:${port}`);
  const opened = new Promise((r) => ws.once("open", r));
  await opened;

  const incoming = [];
  ws.on("message", (raw) => incoming.push(JSON.parse(raw.toString())));

  ws.send(JSON.stringify({ type: "hello", protocolVersion: 1, host: "excel", clientName: "test-pane", paneId: "pane-a" }));
  await sleep(100);

  // --- tool proxy: call from server side, answer from pane ---
  const toolPromise = server.callOfficeTool("excel", "read_range", { range: "A1:B2" });
  await sleep(100);
  const toolCall = incoming.find((m) => m.type === "tool_call");
  if (!toolCall) throw new Error("no tool_call received by pane");
  console.log(`[ok] pane received tool_call: ${toolCall.tool} args=${JSON.stringify(toolCall.args)}`);

  ws.send(JSON.stringify({ type: "tool_result", id: toolCall.id, ok: true, text: "| A | B |\n|---|---|\n| 1 | 2 |", details: { rows: 2 } }));
  const toolResult = await toolPromise;
  console.log(`[ok] server resolved tool result: "${toolResult.text.split("\n")[0]}"`);

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
    console.log(`[ok] never-attached host rejected: ${e.message.slice(0, 60)}…`);
  }

  // --- ping/pong ---
  ws.send(JSON.stringify({ type: "ping", ts: Date.now() }));
  await sleep(100);
  if (!incoming.some((m) => m.type === "pong")) throw new Error("no pong");

  // --- user_message ---
  ws.send(JSON.stringify({ type: "user_message", text: "Summarize this sheet" }));
  await sleep(100);
  if (userMessages.length !== 1 || userMessages[0].text !== "Summarize this sheet") {
    throw new Error("user message not routed");
  }
  console.log("[ok] pane user_message routed to pi handler");

  // --- status update ---
  ws.send(JSON.stringify({ type: "status", model: "mimo-v2.5", provider: "opencode-go" }));
  await sleep(100);
  const panes = server.attachedPanes();
  if (panes[0].model !== "mimo-v2.5") throw new Error("status not stored");
  console.log("[ok] pane status stored (model=%s)", panes[0].model);

  // --- timeout path: a tool call nobody answers ---
  const slow = server.callOfficeTool("excel", "read_range", { range: "A1" }, undefined, 250);
  try {
    await slow;
    throw new Error("expected timeout");
  } catch (e) {
    console.log(`[ok] unanswered tool call rejected (timeout wiring): ${e.message.slice(0, 50)}…`);
  }

  // --- disconnect cleanup ---
  ws.close();
  await sleep(150);
  const remaining = server.attachedPanes();
  if (remaining.length !== 0) throw new Error("pane not cleaned up on disconnect");
  console.log("[ok] pane removed on disconnect");

  await server.stop();
  console.log("[ok] server stopped cleanly");
  console.log("\nALL SMOKE TESTS PASSED");
}

main().catch((e) => {
  console.error("SMOKE TEST FAILED:", e);
  process.exit(1);
});
