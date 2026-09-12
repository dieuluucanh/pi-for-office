/**
 * Unit tests for the pure active-tool reconciler. Run after `npm run build`
 * (imports the compiled dist/active-tools.js).
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  activeOfficeToolNames,
  isOfficeToolName,
  reconcileOfficeToolActivation,
} from "../dist/active-tools.js";
import { OFFICE_TOOL_NAMES } from "../dist/office-tools.js";

const LEGACY_WORD_TOOLS = [
  "office_word_get_overview",
  "office_word_read_document",
  "office_word_insert_text",
  "office_word_replace_text",
];

const MODERN_WORD_TOOLS = [
  ...LEGACY_WORD_TOOLS,
  "office_word_format_range",
  "office_word_insert_blocks",
  "office_word_insert_table",
  "office_word_insert_page_break",
  "office_word_insert_image",
  "office_word_insert_hyperlink",
];

test("isOfficeToolName recognizes only catalog tool names", () => {
  for (const name of OFFICE_TOOL_NAMES) {
    assert.equal(
      isOfficeToolName(name),
      true,
      `${name} should be an office tool`,
    );
  }
  assert.equal(isOfficeToolName("office_word_does_not_exist"), false);
  assert.equal(isOfficeToolName("bash"), false);
  assert.equal(isOfficeToolName("read"), false);
  assert.equal(isOfficeToolName(""), false);
});

test("activeOfficeToolNames: modern pane advertises exactly its ops", () => {
  const names = activeOfficeToolNames([
    { host: "word", ops: ["word.get_overview", "word.insert_blocks"] },
  ]);
  assert.deepEqual(names, [
    "office_word_get_overview",
    "office_word_insert_blocks",
  ]);
});

test("activeOfficeToolNames: legacy pane (null ops) gets the v1 set only", () => {
  const names = activeOfficeToolNames([{ host: "word", ops: null }]);
  assert.deepEqual(names.sort(), [...LEGACY_WORD_TOOLS].sort());
  assert.ok(!names.includes("office_word_format_range"));
  assert.ok(!names.includes("office_word_insert_blocks"));
});

test("activeOfficeToolNames: multi-host union, deterministic order", () => {
  const names = activeOfficeToolNames([
    { host: "excel", ops: ["excel.read_range"] },
    { host: "word", ops: ["word.insert_text"] },
  ]);
  assert.deepEqual(names, [
    "office_excel_read_range",
    "office_word_insert_text",
  ]);
});

test("activeOfficeToolNames: foreign/unknown ops are ignored", () => {
  const names = activeOfficeToolNames([
    {
      host: "word",
      ops: ["word.get_overview", "excel.read_range", "word.bogus"],
    },
  ]);
  assert.deepEqual(names, ["office_word_get_overview"]);
});

test("reconcile preserves non-office tools and keeps them first", () => {
  const next = reconcileOfficeToolActivation(
    ["bash", "read", "office_word_insert_text"],
    [{ host: "word", ops: ["word.get_overview"] }],
  );
  assert.ok(next.includes("bash"));
  assert.ok(next.includes("read"));
  assert.ok(!next.includes("office_word_insert_text"));
  assert.ok(next.includes("office_word_get_overview"));
  // Non-office tools come before office tools (deterministic).
  assert.ok(next.indexOf("bash") < next.indexOf("office_word_get_overview"));
});

test("reconcile deactivates office tools when the pane detaches", () => {
  const active = [
    "bash",
    "office_word_insert_text",
    "office_word_format_range",
  ];
  const next = reconcileOfficeToolActivation(active, []);
  assert.deepEqual(next, ["bash"]);
});

test("reconcile is idempotent on the same pane set", () => {
  const panes = [{ host: "word", ops: MODERN_WORD_TOOLS }];
  const first = reconcileOfficeToolActivation(
    ["bash", "office_word_insert_text"],
    panes,
  );
  const second = reconcileOfficeToolActivation(first, panes);
  assert.deepEqual(second, first);
});

test("reconcile multi-host: attached word + excel union", () => {
  const next = reconcileOfficeToolActivation(
    ["bash"],
    [
      { host: "word", ops: ["word.insert_text"] },
      { host: "excel", ops: ["excel.read_range", "excel.format_cells"] },
    ],
  );
  assert.deepEqual(next, [
    "bash",
    "office_excel_format_cells",
    "office_excel_read_range",
    "office_word_insert_text",
  ]);
});

test("reconcile does not add tools for a legacy pane beyond v1", () => {
  const next = reconcileOfficeToolActivation(
    ["bash"],
    [{ host: "word", ops: null }],
  );
  const office = next.filter((n) => isOfficeToolName(n));
  assert.deepEqual(office.sort(), [...LEGACY_WORD_TOOLS].sort());
});
