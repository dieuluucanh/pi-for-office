import assert from "node:assert/strict";
import { test } from "node:test";

import {
  extractBridgeErrorMessage,
  isAbortError,
  joinBridgeUrl,
  tryParseBridgeJson,
} from "../src/tools/bridge-http-utils.ts";

void test("tryParseBridgeJson parses valid json and ignores invalid payloads", () => {
  assert.deepEqual(tryParseBridgeJson('{"ok":true}'), { ok: true });
  assert.equal(tryParseBridgeJson("   "), null);
  assert.equal(tryParseBridgeJson("not-json"), null);
});

void test("extractBridgeErrorMessage prefers structured error field", () => {
  assert.equal(extractBridgeErrorMessage({ error: "bridge down" }), "bridge down");
  assert.equal(extractBridgeErrorMessage(" timeout "), "timeout");
  assert.equal(extractBridgeErrorMessage({ error: 123 }), null);
  assert.equal(extractBridgeErrorMessage({ message: "oops" }), null);
});

void test("joinBridgeUrl normalizes trailing slashes", () => {
  assert.equal(joinBridgeUrl("https://localhost:7777", "/v1/test"), "https://localhost:7777/v1/test");
  assert.equal(joinBridgeUrl("https://localhost:7777/", "/v1/test"), "https://localhost:7777/v1/test");
  assert.equal(joinBridgeUrl("https://localhost:7777///", "/v1/test"), "https://localhost:7777/v1/test");
});

void test("isAbortError matches DOMException and Error abort names", () => {
  assert.equal(isAbortError(new DOMException("aborted", "AbortError")), true);
  assert.equal(isAbortError(new Error("aborted")), false);

  const namedError = new Error("aborted");
  namedError.name = "AbortError";
  assert.equal(isAbortError(namedError), true);

  assert.equal(isAbortError("AbortError"), false);
});
