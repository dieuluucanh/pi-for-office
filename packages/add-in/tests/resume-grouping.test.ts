import assert from "node:assert/strict";
import { test } from "node:test";

import {
  filterSessionsByQuery,
  groupSessionsForWorkbook,
  shouldExpandOtherWorkbooks,
} from "../src/commands/builtins/resume-grouping.ts";
import type { SessionListEntry } from "../src/storage/local/sessions-store.ts";

function entry(id: string, title: string, preview = ""): SessionListEntry {
  return {
    metadata: {
      id,
      title,
      createdAt: "2026-01-01T00:00:00.000Z",
      lastModified: "2026-01-02T00:00:00.000Z",
      messageCount: 2,
      usage: {
        input: 0,
        output: 0,
        cacheRead: 0,
        cacheWrite: 0,
        totalTokens: 0,
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
      },
      thinkingLevel: "off",
      preview,
    },
  };
}

void test("sessions are grouped, never dropped", () => {
  const entries = [entry("this", "This"), entry("other", "Other"), entry("none", "None")];
  const links = new Map([
    ["this", "wb-1"],
    ["other", "wb-2"],
  ]);

  const groups = groupSessionsForWorkbook(entries, links, "wb-1");
  assert.deepEqual(
    groups.map((group) => [group.id, group.entries.length]),
    [
      ["this_workbook", 1],
      ["other_workbooks", 1],
      ["unlinked", 1],
    ],
  );
  assert.equal(groups.flatMap((group) => group.entries).length, entries.length);
});

void test("with no workbook identity everything linked is grouped as other", () => {
  const entries = [entry("a", "A"), entry("b", "B")];
  const links = new Map([["a", "wb-1"]]);

  const groups = groupSessionsForWorkbook(entries, links, null);
  assert.equal(groups.find((group) => group.id === "this_workbook")?.entries.length, 0);
  assert.equal(groups.find((group) => group.id === "other_workbooks")?.entries.length, 1);
  assert.equal(groups.find((group) => group.id === "unlinked")?.entries.length, 1);
});

void test("other workbooks auto-expand when this workbook has nothing", () => {
  const entries = [entry("other", "Other")];
  const groups = groupSessionsForWorkbook(entries, new Map([["other", "wb-2"]]), "wb-1");

  assert.equal(shouldExpandOtherWorkbooks(groups, "wb-1"), true);

  const withMatch = groupSessionsForWorkbook(
    [...entries, entry("this", "This")],
    new Map([
      ["other", "wb-2"],
      ["this", "wb-1"],
    ]),
    "wb-1",
  );
  assert.equal(shouldExpandOtherWorkbooks(withMatch, "wb-1"), false);
});

void test("search matches titles and previews, and empty search is a copy", () => {
  const entries = [
    entry("a", "Quarterly budget", "revenue by region"),
    entry("b", "Charts", "add a bar chart"),
  ];

  assert.equal(filterSessionsByQuery(entries, "budget").length, 1);
  assert.equal(filterSessionsByQuery(entries, "BAR CHART").length, 1);
  assert.equal(filterSessionsByQuery(entries, "  ").length, 2);
  assert.notEqual(filterSessionsByQuery(entries, "  "), entries);
});
