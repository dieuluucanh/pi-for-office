/**
 * Unit tests for the pane-prompt ↔ reply tracker and the document tag.
 *
 * Run: npm run test:replies (after building dist).
 */
import test from "node:test";
import assert from "node:assert/strict";

import {
  DEFAULT_PROMPT_TTL_MS,
  PanePromptTracker,
  formatDocumentTag,
} from "../dist/pane-replies.js";

const entry = (sentText, promptId = "prompt-1", pane = { id: "pane-1" }) => ({
  pane,
  promptId,
  sentText,
});

test("binds an exactly delivered user message to the pending pane prompt", () => {
  const tracker = new PanePromptTracker();
  const queued = entry("hello");
  tracker.enqueue(queued, 1_000);

  assert.equal(tracker.active(), null);
  const bound = tracker.bindDeliveredUserMessage("hello", 1_100);
  assert.equal(bound?.promptId, "prompt-1");
  assert.equal(tracker.active()?.sentText, "hello");
  assert.equal(tracker.pendingCount, 0);
  assert.equal(tracker.complete()?.promptId, "prompt-1");
  assert.equal(tracker.active(), null);
});

test("binds when the delivered text extends the injected prompt (prefix match)", () => {
  const tracker = new PanePromptTracker();
  tracker.enqueue(entry("tagged prompt"), 0);
  const bound = tracker.bindDeliveredUserMessage("tagged prompt + extra", 10);
  assert.equal(bound?.promptId, "prompt-1");
});

test("does not bind or consume for an unrelated (TUI) user message", () => {
  const tracker = new PanePromptTracker();
  tracker.enqueue(entry("pane prompt"), 0);
  assert.equal(
    tracker.bindDeliveredUserMessage("a typed TUI prompt", 10),
    null,
  );
  assert.equal(tracker.pendingCount, 1);
  assert.equal(tracker.active(), null);
});

test("prompts are bound in FIFO order", () => {
  const tracker = new PanePromptTracker();
  tracker.enqueue(entry("first", "prompt-1"), 0);
  tracker.enqueue(entry("second", "prompt-2"), 0);
  assert.equal(
    tracker.bindDeliveredUserMessage("first", 1)?.promptId,
    "prompt-1",
  );
  assert.equal(tracker.active()?.promptId, "prompt-1");
  assert.equal(
    tracker.bindDeliveredUserMessage("second", 2)?.promptId,
    "prompt-2",
  );
  assert.equal(tracker.pendingCount, 0);
});

test("drops prompts that are never delivered within the TTL", () => {
  const tracker = new PanePromptTracker(1_000);
  tracker.enqueue(entry("stale"), 0);
  assert.equal(
    tracker.bindDeliveredUserMessage("stale", DEFAULT_PROMPT_TTL_MS + 1),
    null,
  );
  assert.equal(tracker.pendingCount, 0);
});

test("clear drops queue and active binding", () => {
  const tracker = new PanePromptTracker();
  tracker.enqueue(entry("one"), 0);
  tracker.bindDeliveredUserMessage("one", 1);
  assert.notEqual(tracker.active(), null);
  tracker.clear();
  assert.equal(tracker.active(), null);
  assert.equal(tracker.pendingCount, 0);
});

test("formatDocumentTag includes host, name and id when known", () => {
  assert.equal(
    formatDocumentTag({
      host: "excel",
      documentName: "MBI_2026.xlsx",
      documentId: "url_sha256:ab12cd34",
    }),
    '[office:excel "MBI_2026.xlsx" id=url_sha256:ab12cd34]',
  );
});

test("formatDocumentTag degrades for unsaved/unknown documents", () => {
  assert.equal(
    formatDocumentTag({ host: "word", documentName: null, documentId: null }),
    '[office:word "untitled"]',
  );
});
