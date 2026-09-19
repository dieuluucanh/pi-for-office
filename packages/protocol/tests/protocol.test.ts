/**
 * Wire-protocol contract tests — version constants, the v2/v3 compatibility
 * helper, and the client-frame parser (including the protocol v3 `abort` frame).
 * Run with `npm test` in packages/protocol.
 */

import assert from "node:assert/strict";
import { test } from "node:test";
import {
  BRIDGE_PROTOCOL_VERSION,
  MIN_SUPPORTED_PROTOCOL_VERSION,
  isProtocolVersionSupported,
  parseClientMessage,
  type AgentDeltaMessage,
  type AgentFinalMessage,
  type ClientAbortMessage,
  type ClientMessage,
  type ServerMessage,
} from "../src/protocol.ts";

test("protocol version is v3 and still accepts v2 clients", () => {
  assert.equal(BRIDGE_PROTOCOL_VERSION, 3);
  assert.equal(MIN_SUPPORTED_PROTOCOL_VERSION, 2);
  assert.ok(
    MIN_SUPPORTED_PROTOCOL_VERSION < BRIDGE_PROTOCOL_VERSION,
    "the supported window must be non-empty",
  );
});

test("isProtocolVersionSupported accepts the supported window only", () => {
  for (const version of [2, 3]) {
    assert.equal(isProtocolVersionSupported(version), true, `v${version} must be accepted`);
  }
  for (const version of [0, 1, 4, 3.5, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
    assert.equal(
      isProtocolVersionSupported(version),
      false,
      `v${String(version)} must be rejected`,
    );
  }
  for (const notAVersion of [undefined, null, "3", {}, [], true]) {
    assert.equal(
      isProtocolVersionSupported(notAVersion),
      false,
      `${JSON.stringify(notAVersion) ?? String(notAVersion)} must be rejected`,
    );
  }
});

test("parseClientMessage accepts an abort frame with and without promptId", () => {
  const bare = parseClientMessage(JSON.stringify({ type: "abort" }));
  assert.equal(bare.type, "abort");
  assert.equal((bare as ClientAbortMessage).promptId, undefined);

  const correlated = parseClientMessage(
    JSON.stringify({ type: "abort", promptId: "prompt-1" }),
  );
  assert.deepEqual(correlated, { type: "abort", promptId: "prompt-1" });
});

test("parseClientMessage rejects a malformed abort frame", () => {
  assert.throws(
    () => parseClientMessage(JSON.stringify({ type: "abort", promptId: 5 })),
    /abort\.promptId must be a string/,
  );
});

test("parser keeps accepting the v2 client frames", () => {
  const hello = parseClientMessage(
    JSON.stringify({
      type: "hello",
      protocolVersion: 2,
      host: "excel",
      document: { id: "abc", name: "Book1.xlsx", source: "document.url" },
    }),
  );
  assert.equal(hello.type, "hello");

  assert.deepEqual(parseClientMessage(JSON.stringify({ type: "ping", ts: 1 })), {
    type: "ping",
    ts: 1,
  });
  assert.deepEqual(
    parseClientMessage(JSON.stringify({ type: "status", model: "m", provider: "p" })),
    { type: "status", model: "m", provider: "p" },
  );
  assert.equal(parseClientMessage(JSON.stringify({ type: "pane_state", focused: true })).type, "pane_state");
  assert.equal(
    parseClientMessage(JSON.stringify({ type: "tool_result", id: "c-1", ok: true, text: "ok" })).type,
    "tool_result",
  );
  assert.equal(
    parseClientMessage(JSON.stringify({ type: "user_message", text: "hi" })).type,
    "user_message",
  );
});

test("parser rejects junk and unknown frames", () => {
  assert.throws(() => parseClientMessage("not json"), /not valid JSON/);
  assert.throws(() => parseClientMessage("null"), /expected a JSON object/);
  assert.throws(() => parseClientMessage("42"), /expected a JSON object/);
  assert.throws(() => parseClientMessage("[]"), /missing 'type'/);
  assert.throws(() => parseClientMessage(JSON.stringify({})), /missing 'type'/);
  assert.throws(
    () => parseClientMessage(JSON.stringify({ type: "nope" })),
    /unknown client message type/,
  );
});

test("v3 delta and abort shapes are wired into the message unions", () => {
  // These assignments are compile-time assertions: they fail `npm run typecheck`
  // if a future change drops a v3 field or removes a frame from a union.
  const delta: AgentDeltaMessage = {
    type: "agent_message",
    kind: "delta",
    text: "hi",
    promptId: "p1",
    messageId: "m1",
  };
  const deltaWithoutIds: AgentDeltaMessage = {
    type: "agent_message",
    kind: "delta",
    text: "hi",
  };
  const abort: ClientAbortMessage = { type: "abort" };
  const clientFrames: ClientMessage[] = [abort, { type: "ping", ts: 0 }];
  const final: AgentFinalMessage = {
    type: "agent_message",
    kind: "final",
    text: "hi",
    promptId: "p1",
  };
  const serverFrames: ServerMessage[] = [delta, final];

  assert.equal(delta.messageId, "m1");
  assert.equal(deltaWithoutIds.promptId, undefined);
  assert.equal(final.promptId, "p1");
  assert.equal(clientFrames.length, 2);
  assert.equal(serverFrames.length, 2);
});
