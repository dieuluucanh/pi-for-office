/**
 * Multi-document routing test — protocol v2 target resolution.
 *
 * Covers: focus-first default, explicit name/#name/id/substring targets,
 * ambiguity + unknown errors with candidate lists, selector stripping,
 * duplicate windows of one file, unattached hosts, /health diagnostics and the
 * v1/absent-document handshake gates.
 *
 * Run: npm run test:multi (after building dist).
 */
import { OfficeBridgeServer } from "../dist/bridge-server.js";
import WebSocket from "ws";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Connect a fake pane and return its sockets/frames. */
function connectPane(port, hello) {
  const ws = new WebSocket(`ws://127.0.0.1:${port}`);
  const incoming = [];
  const ready = new Promise((r) => ws.once("open", r));
  ws.on("message", (raw) => incoming.push(JSON.parse(raw.toString())));
  return {
    ws,
    incoming,
    ready,
    greet: () =>
      ws.send(JSON.stringify({ type: "hello", protocolVersion: 2, ...hello })),
  };
}

/** Poll a pane's frames for an unanswered tool_call at/after `fromIndex`. */
async function waitForToolCall(pane, fromIndex = 0) {
  for (let i = 0; i < 50; i++) {
    const call = pane.incoming
      .slice(fromIndex)
      .find((m) => m.type === "tool_call");
    if (call) return call;
    await sleep(20);
  }
  return null;
}

/** Assert a promise rejects with a message containing `needle`. */
async function expectReject(promise, needle) {
  try {
    await promise;
    throw new Error(`expected rejection containing "${needle}"`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (!message.includes(needle)) {
      throw new Error(`wrong rejection (wanted "${needle}"): ${message}`);
    }
    return message;
  }
}

/** Answer a received tool_call so the server promise resolves. */
function answer(pane, call, text = "ok") {
  pane.ws.send(
    JSON.stringify({ type: "tool_result", id: call.id, ok: true, text }),
  );
}

async function main() {
  const server = new OfficeBridgeServer({
    port: 0,
    serverName: "multi-document-test",
    handlers: { onUserMessage() {} },
  });
  await server.start();
  const port = server.actualPort;

  const alpha = connectPane(port, {
    host: "excel",
    clientName: "alpha",
    paneId: "pane-alpha",
    document: {
      id: "url_sha256:alpha",
      name: "Alpha.xlsx",
      source: "document.url",
    },
  });
  const beta = connectPane(port, {
    host: "excel",
    clientName: "beta",
    paneId: "pane-beta",
    document: {
      id: "url_sha256:beta",
      name: "Beta.xlsx",
      source: "document.url",
    },
  });
  await Promise.all([alpha.ready, beta.ready]);
  alpha.greet();
  beta.greet();
  await sleep(100);
  if (server.attachedPanes().length !== 2) {
    throw new Error("expected two attached Excel panes");
  }
  console.log("[ok] two Excel documents attached (no target is ambiguous)");

  // 1. No target and no focus → never guess.
  const ambiguous = await expectReject(
    server.callOfficeTool("excel", "read_range", { range: "A1" }),
    "none is focused",
  );
  if (!ambiguous.includes("Alpha.xlsx") || !ambiguous.includes("Beta.xlsx")) {
    throw new Error("ambiguous error must list the open documents");
  }
  console.log("[ok] no target + no focus → error listing both documents");

  // 2. Focus Beta via pane_state → default target follows focus.
  beta.ws.send(
    JSON.stringify({
      type: "pane_state",
      focused: true,
      visible: true,
      activeAt: Date.now(),
    }),
  );
  await sleep(50);
  const betaIndex = beta.incoming.length;
  const focusCall = server.callOfficeTool("excel", "read_range", {
    range: "A1",
  });
  const routed = await waitForToolCall(beta, betaIndex);
  if (!routed) throw new Error("focused pane did not receive the call");
  answer(beta, routed);
  await focusCall;
  if (alpha.incoming.some((m) => m.type === "tool_call")) {
    throw new Error("unfocused pane must not receive the call");
  }
  console.log(
    "[ok] pane_state focus → default target routed to the focused document",
  );

  // 3. Explicit targets: name, #name, exact id, unique substring.
  for (const target of [
    "Alpha.xlsx",
    "#Alpha.xlsx",
    "url_sha256:alpha",
    "Al",
  ]) {
    const index = alpha.incoming.length;
    const promise = server.callOfficeTool("excel", "read_range", {
      range: "A1",
      document: target,
    });
    const call = await waitForToolCall(alpha, index);
    if (!call) throw new Error(`target "${target}" did not reach Alpha`);
    if ("document" in call.args) {
      throw new Error(`document selector leaked to the pane for "${target}"`);
    }
    if (call.documentId !== "url_sha256:alpha") {
      throw new Error(`documentId echo wrong for "${target}"`);
    }
    answer(alpha, call);
    await promise;
  }
  console.log(
    "[ok] name / #name / id / unique-substring targets route correctly and strip the selector",
  );

  // 4. Unknown target → actionable error with candidates.
  const unknown = await expectReject(
    server.callOfficeTool("excel", "read_range", {
      range: "A1",
      document: "Nope.xlsx",
    }),
    "no excel document matches",
  );
  if (!unknown.includes("Alpha.xlsx") || !unknown.includes("Beta.xlsx")) {
    throw new Error("unknown-target error must list the candidates");
  }
  console.log("[ok] unknown target → error with the candidate list");

  // 5. Duplicate name → ambiguous.
  const alphaCopyName = connectPane(port, {
    host: "excel",
    clientName: "alpha-copy-name",
    paneId: "pane-alpha-copy",
    document: {
      id: "url_sha256:alpha-copy",
      name: "Alpha.xlsx",
      source: "document.url",
    },
  });
  await alphaCopyName.ready;
  alphaCopyName.greet();
  await sleep(100);
  await expectReject(
    server.callOfficeTool("excel", "read_range", {
      range: "A1",
      document: "Alpha.xlsx",
    }),
    "matches several",
  );
  console.log("[ok] duplicate document name → ambiguous error");
  alphaCopyName.ws.close();
  await sleep(100);

  // 6. Same file in two windows (same document id), none of them focused.
  const alphaWindow = connectPane(port, {
    host: "excel",
    clientName: "alpha-window-2",
    paneId: "pane-alpha-window",
    document: {
      id: "url_sha256:alpha",
      name: "Alpha.xlsx",
      source: "document.url",
    },
  });
  await alphaWindow.ready;
  alphaWindow.greet();
  await sleep(100);
  await expectReject(
    server.callOfficeTool("excel", "read_range", {
      range: "A1",
      document: "url_sha256:alpha",
    }),
    "windows",
  );
  console.log(
    "[ok] same file in two windows + no focus → error, never a silent guess",
  );

  // 7. Unattached host.
  await expectReject(
    server.callOfficeTool("word", "get_overview", {}),
    "no word workbook/document is attached",
  );
  console.log("[ok] unattached host rejected");

  // 8. /health reports the duplicate-window count and document identity.
  const health = await (await fetch(`http://127.0.0.1:${port}/health`)).json();
  const alphaEntries = health.panes.filter(
    (pane) => pane.document?.id === "url_sha256:alpha",
  );
  if (alphaEntries.length !== 2 || alphaEntries[0].windows !== 2) {
    throw new Error(
      `health windows wrong: ${JSON.stringify(alphaEntries.map((p) => p.windows))}`,
    );
  }
  console.log("[ok] /health documents identity + duplicate-window count");

  // 9. Handshake gates: v1 rejected, v2 without document rejected.
  const wsV1 = new WebSocket(`ws://127.0.0.1:${port}`);
  await new Promise((r) => wsV1.once("open", r));
  const v1Frames = [];
  wsV1.on("message", (raw) => v1Frames.push(JSON.parse(raw.toString())));
  wsV1.send(
    JSON.stringify({
      type: "hello",
      protocolVersion: 1,
      host: "excel",
      clientName: "v1",
      paneId: "pane-v1",
      document: { id: null, name: null, source: "unknown" },
    }),
  );
  await sleep(100);
  if (
    !v1Frames.some((m) => m.type === "error" && m.code === "protocol_mismatch")
  ) {
    throw new Error("v1 hello must be rejected");
  }
  wsV1.close();

  const wsNoDoc = new WebSocket(`ws://127.0.0.1:${port}`);
  await new Promise((r) => wsNoDoc.once("open", r));
  const noDocFrames = [];
  wsNoDoc.on("message", (raw) => noDocFrames.push(JSON.parse(raw.toString())));
  wsNoDoc.send(
    JSON.stringify({
      type: "hello",
      protocolVersion: 2,
      host: "excel",
      clientName: "no-document",
      paneId: "pane-nodoc",
    }),
  );
  await sleep(100);
  if (
    !noDocFrames.some((m) => m.type === "error" && m.code === "bad_message")
  ) {
    throw new Error("v2 hello without document must be rejected");
  }
  wsNoDoc.close();
  console.log("[ok] handshake gates: v1 mismatch + missing document rejected");

  for (const pane of [alpha, beta, alphaWindow]) pane.ws.close();
  await sleep(150);
  await server.stop();
  console.log("\nALL MULTI-DOCUMENT TESTS PASSED");
}

main().catch((error) => {
  console.error("MULTI-DOCUMENT TEST FAILED:", error);
  process.exit(1);
});
