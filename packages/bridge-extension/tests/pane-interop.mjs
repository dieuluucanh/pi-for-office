/**
 * End-to-end interop test: the real add-in PaneBridgeClient ↔ the real bridge
 * server. Verifies the shared protocol (`@dieulc/pi-office-protocol`) is
 * symmetric across both packages.
 *
 * Run: npm run test:interop (after building dist).
 */
import { OfficeBridgeServer } from "../dist/bridge-server.js";
import { PaneBridgeClient } from "../../add-in/src/bridge/pane-client.ts";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const userMessages = [];
  const server = new OfficeBridgeServer({
    port: 0,
    serverName: "interop",
    piVersion: "0.85.1-test",
    handlers: {
      onUserMessage: (text, pane) => userMessages.push({ text, host: pane.host }),
    },
  });
  await server.start();
  const port = server.actualPort;

  // A fake Word op executor (no Office in this test).
  const fakeRegistry = new Map([
    [
      "word.read_document",
      async (args) => ({
        text: `Fake doc text (${args.scope ?? "all"})`,
        details: { scope: args.scope ?? "all" },
      }),
    ],
  ]);

  let assistantFinal = null;
  const client = new PaneBridgeClient(
    { host: "word", registry: fakeRegistry, url: `ws://127.0.0.1:${port}`, paneId: "pane-interop" },
    { onAssistantFinal: (text) => { assistantFinal = text; } },
  );

  await client.connect();
  console.log("[ok] client connected (welcome received)");

  // Server sees the pane as host "word".
  const panes = server.attachedPanes();
  if (panes.length !== 1 || panes[0].host !== "word") throw new Error("server did not register word pane");
  console.log("[ok] server registered word pane");

  // Tool proxy: server → client → fake executor → result.
  const result = await server.callOfficeTool("word", "read_document", { scope: "all" });
  if (!result.text.includes("Fake doc text (all)")) throw new Error("tool result mismatch");
  console.log("[ok] tool_call routed through shared protocol → executor → tool_result");

  // Pane → Pi chat.
  client.sendPrompt("Summarize the doc");
  await sleep(100);
  if (userMessages.length !== 1 || userMessages[0].text !== "Summarize the doc") {
    throw new Error("user_message not delivered to server");
  }
  console.log("[ok] pane user_message delivered to server handler");

  // Pi → pane reply.
  server.broadcast({ type: "agent_message", kind: "final", text: "Here is the summary." });
  await sleep(100);
  if (assistantFinal !== "Here is the summary.") throw new Error("final reply not delivered to pane");
  console.log("[ok] agent final reply delivered to pane");

  client.disconnect();
  await sleep(100);
  if (server.attachedPanes().length !== 0) throw new Error("pane not cleaned up on disconnect");
  console.log("[ok] pane cleaned up on disconnect");

  await server.stop();
  console.log("\nALL INTEROP TESTS PASSED");
}

main().catch((e) => {
  console.error("INTEROP TEST FAILED:", e);
  process.exit(1);
});
