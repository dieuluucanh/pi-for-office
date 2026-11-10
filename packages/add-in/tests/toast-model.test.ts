import assert from "node:assert/strict";
import { test } from "node:test";

import {
  MAX_VISIBLE_TOASTS,
  pickUndoTarget,
  pushToast,
  removeToast,
  type ToastDraft,
  type ToastRecord,
} from "../src/ui/toast-model.ts";

const info = (message: string): ToastDraft => ({
  message,
  variant: "info",
  duration: 2000,
  sticky: false,
});

const action = (message: string): ToastDraft => ({
  message,
  variant: "info",
  duration: 9000,
  sticky: true,
});

const error = (message: string): ToastDraft => ({
  message,
  variant: "error",
  duration: 6000,
  sticky: false,
});

function pushAll(drafts: readonly ToastDraft[]): ToastRecord[] {
  let list: ToastRecord[] = [];
  let seq = 0;

  for (const draft of drafts) {
    seq += 1;
    list = pushToast(list, draft, seq).list;
  }

  return list;
}

void test("caps the stack and evicts the oldest plain toast first", () => {
  const list = pushAll([info("one"), info("two"), info("three"), info("four")]);

  assert.equal(list.length, MAX_VISIBLE_TOASTS);
  assert.deepEqual(
    list.map((record) => record.message),
    ["two", "three", "four"],
  );
});

void test("errors survive an info overflow", () => {
  const list = pushAll([error("LLM error"), info("saved"), info("copied"), info("renamed")]);

  assert.deepEqual(
    list.map((record) => record.message),
    ["LLM error", "copied", "renamed"],
  );
  assert.equal(list[0]?.variant, "error");
});

void test("four errors evict the oldest error", () => {
  const list = pushAll([error("a"), error("b"), error("c"), error("d")]);

  assert.deepEqual(
    list.map((record) => record.message),
    ["b", "c", "d"],
  );
});

void test("a repeated message reuses its toast instead of stacking", () => {
  const first = pushToast([], info("Session moved"), 1);
  const second = pushToast(first.list, info("Session moved"), 2);

  assert.equal(second.reused, true);
  assert.equal(second.record.id, first.record.id);
  assert.equal(second.list.length, 1);
  assert.equal(second.evictedId, null);
});

void test("only the newest action toast stays sticky", () => {
  const first = pushToast([], action("Closed Chat 1"), 1);
  const second = pushToast(first.list, action("Closed Chat 2"), 2);

  assert.deepEqual(second.demotedIds, [first.record.id]);
  assert.equal(second.list[0]?.sticky, false);
  assert.equal(pickUndoTarget(second.list)?.message, "Closed Chat 2");
});

void test("an action toast is never evicted while a plain toast exists", () => {
  const list = pushAll([action("Undo close"), info("saved"), info("copied"), info("renamed")]);

  assert.equal(list.length, MAX_VISIBLE_TOASTS);
  assert.equal(list[0]?.message, "Undo close");
  assert.equal(list[0]?.sticky, true);
  assert.deepEqual(
    list.slice(1).map((record) => record.message),
    ["copied", "renamed"],
  );
});

void test("removeToast drops one record and keeps insertion order", () => {
  const list = pushAll([info("one"), info("two"), info("three")]);
  const removed = removeToast(list, list[1]?.id ?? "");

  assert.deepEqual(
    removed.map((record) => record.message),
    ["one", "three"],
  );
  assert.equal(pickUndoTarget(removed), null);
});
