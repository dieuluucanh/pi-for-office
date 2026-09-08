import assert from "node:assert/strict";
import { test } from "node:test";

import {
  normalizeWorkbookTabLayout,
  parseWorkbookTabLayout,
  workbookTabLayoutKey,
} from "../src/taskpane/tab-layout.ts";

void test("workbookTabLayoutKey uses workbook id when available", () => {
  assert.equal(
    workbookTabLayoutKey("url_sha256:abc"),
    "workbook.tabLayout.v1.url_sha256:abc",
  );
});

void test("workbookTabLayoutKey falls back to global key when workbook id is missing", () => {
  assert.equal(workbookTabLayoutKey(null), "workbook.tabLayout.v1.__global__");
  assert.equal(workbookTabLayoutKey("   "), "workbook.tabLayout.v1.__global__");
});

void test("normalizeWorkbookTabLayout trims ids and falls back active tab", () => {
  const normalized = normalizeWorkbookTabLayout({
    sessionIds: ["  session-a  ", "", " session-b "],
    activeSessionId: "missing",
  });

  assert.deepEqual(normalized, {
    sessionIds: ["session-a", "session-b"],
    activeSessionId: "session-a",
  });
});

void test("parseWorkbookTabLayout accepts valid shape and normalizes active id", () => {
  const parsed = parseWorkbookTabLayout({
    sessionIds: [" session-a ", "session-b"],
    activeSessionId: "session-b",
  });

  assert.deepEqual(parsed, {
    sessionIds: ["session-a", "session-b"],
    activeSessionId: "session-b",
  });
});

void test("parseWorkbookTabLayout rejects invalid shapes", () => {
  assert.equal(parseWorkbookTabLayout(null), null);
  assert.equal(parseWorkbookTabLayout({}), null);
  assert.equal(parseWorkbookTabLayout({ sessionIds: "not-an-array" }), null);
  assert.equal(parseWorkbookTabLayout({ sessionIds: ["", "  "] }), null);
});
