import assert from "node:assert/strict";
import { test } from "node:test";

import { resolveTabTitle } from "../src/taskpane/session-title.ts";

void test("resolveTabTitle uses explicit title when provided", () => {
  assert.equal(
    resolveTabTitle({
      hasExplicitTitle: true,
      sessionTitle: "  Revenue Model  ",
      defaultTabNumber: 3,
    }),
    "Revenue Model",
  );
});

void test("resolveTabTitle falls back to Chat N when no explicit title", () => {
  assert.equal(
    resolveTabTitle({
      hasExplicitTitle: false,
      sessionTitle: "Please can we do ...",
      defaultTabNumber: 1,
    }),
    "Chat 1",
  );
});

void test("resolveTabTitle falls back to Chat N when explicit title is blank", () => {
  assert.equal(
    resolveTabTitle({
      hasExplicitTitle: true,
      sessionTitle: "   ",
      defaultTabNumber: 4,
    }),
    "Chat 4",
  );
});

void test("resolveTabTitle falls back to Chat 1 for invalid default numbers", () => {
  assert.equal(
    resolveTabTitle({
      hasExplicitTitle: false,
      sessionTitle: "",
      defaultTabNumber: 0,
    }),
    "Chat 1",
  );
});
