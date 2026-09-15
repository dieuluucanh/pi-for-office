/**
 * Unit tests for the `#` document-mention autocomplete provider.
 *
 * Run: npm run test:mentions (after building dist).
 */
import test from "node:test";
import assert from "node:assert/strict";

import {
  createDocumentMentionProvider,
  documentMentionItems,
  extractDocumentToken,
  shortDocumentId,
} from "../dist/document-mentions.js";

const target = (overrides = {}) => ({
  host: "excel",
  hostLabel: "Excel",
  name: "Alpha.xlsx",
  documentId: "url_sha256:alpha",
  targetId: "url_sha256:alpha",
  focused: false,
  windows: 1,
  ...overrides,
});

/** Fake built-in provider that records delegation. */
function fakeCurrent() {
  const calls = { get: 0, apply: 0 };
  return {
    calls,
    provider: {
      async getSuggestions() {
        calls.get += 1;
        return null;
      },
      applyCompletion(lines, cursorLine, cursorCol) {
        calls.apply += 1;
        return { lines: [...lines], cursorLine, cursorCol };
      },
    },
  };
}

test("extractDocumentToken only matches a # token after whitespace or line start", () => {
  assert.equal(extractDocumentToken("#Al", 3), "Al");
  assert.equal(extractDocumentToken("see #Al", 7), "Al");
  assert.equal(extractDocumentToken("#", 1), "");
  assert.equal(extractDocumentToken("foo#bar", 7), undefined);
  assert.equal(extractDocumentToken("plain text", 10), undefined);
});

test("shortDocumentId strips the hash prefix and truncates", () => {
  assert.equal(shortDocumentId("url_sha256:abcdef0123456789"), "abcdef01");
  assert.equal(shortDocumentId("ephemeral:pane-1"), "pane-1");
});

test("documentMentionItems inserts #name when the name is unique", () => {
  const items = documentMentionItems(
    [target(), target({ name: "Beta.xlsx", documentId: "url_sha256:beta" })],
    "",
  );
  assert.equal(items.length, 2);
  assert.equal(items[0].value, "#Alpha.xlsx");
  assert.match(items[0].description, /Excel · alpha/);
});

test("documentMentionItems falls back to #id for duplicate names", () => {
  const items = documentMentionItems(
    [
      target(),
      target({
        documentId: "url_sha256:alpha2",
        targetId: "url_sha256:alpha2",
      }),
    ],
    "",
  );
  assert.deepEqual(items.map((item) => item.value).sort(), [
    "#url_sha256:alpha",
    "#url_sha256:alpha2",
  ]);
});

test("documentMentionItems sorts focused first and filters by query", () => {
  const items = documentMentionItems(
    [
      target(),
      target({
        name: "Beta.xlsx",
        documentId: "url_sha256:beta",
        focused: true,
      }),
    ],
    "bet",
  );
  assert.equal(items.length, 1);
  assert.equal(items[0].value, "#Beta.xlsx");
  assert.match(items[0].description, /focused/);
});

test("mention provider delegates when there is no # token", async () => {
  const current = fakeCurrent();
  const provider = createDocumentMentionProvider(current.provider, () => [
    target(),
  ]);
  assert.equal(
    await provider.getSuggestions(["plain"], 0, 5, {
      signal: new AbortController().signal,
    }),
    null,
  );
  assert.equal(current.calls.get, 1);
});

test("mention provider delegates when no documents are attached", async () => {
  const current = fakeCurrent();
  const provider = createDocumentMentionProvider(current.provider, () => []);
  assert.equal(
    await provider.getSuggestions(["#al"], 0, 3, {
      signal: new AbortController().signal,
    }),
    null,
  );
  assert.equal(current.calls.get, 1);
});

test("mention provider returns document items for a # token", async () => {
  const current = fakeCurrent();
  const provider = createDocumentMentionProvider(current.provider, () => [
    target(),
  ]);
  const suggestions = await provider.getSuggestions(["#al"], 0, 3, {
    signal: new AbortController().signal,
  });
  assert.equal(suggestions.prefix, "#al");
  assert.equal(suggestions.items[0].value, "#Alpha.xlsx");
});

test("applyCompletion inserts the mention for a # prefix", () => {
  const current = fakeCurrent();
  const provider = createDocumentMentionProvider(current.provider, () => [
    target(),
  ]);
  const item = { value: "#Alpha.xlsx", label: "#Alpha.xlsx" };
  const result = provider.applyCompletion(["see #Al"], 0, 7, item, "#Al");
  assert.deepEqual(result.lines, ["see #Alpha.xlsx"]);
  assert.equal(result.cursorCol, "see #Alpha.xlsx".length);
  assert.equal(current.calls.apply, 0);
});

test("applyCompletion delegates for non-# prefixes", () => {
  const current = fakeCurrent();
  const provider = createDocumentMentionProvider(current.provider, () => [
    target(),
  ]);
  const item = { value: "src/app.ts", label: "src/app.ts" };
  provider.applyCompletion(["@app"], 0, 4, item, "@app");
  assert.equal(current.calls.apply, 1);
});
