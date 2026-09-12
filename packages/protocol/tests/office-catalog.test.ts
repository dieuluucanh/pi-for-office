/**
 * Catalog contract tests — the office op catalog must stay internally
 * consistent: unique names/op ids, non-empty copy, well-formed schemas, exact
 * per-host membership, and a valid legacy subset. Run with
 * `npm test` in packages/protocol.
 */

import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import {
  OFFICE_CATALOG,
  OFFICE_CATALOG_BY_OP,
  OFFICE_TOOL_NAMES,
  LEGACY_V1_OPS,
  CATALOG_VERSION,
  catalogForHost,
  catalogOpIdsForHost,
  hostForToolName,
  officeToolName,
  type OfficeCatalogEntry,
} from "../src/office-catalog.ts";

const HOSTS = ["excel", "word", "powerpoint"] as const;
type Host = (typeof HOSTS)[number];

/** Number of legacy v1 ops per host (bridge 0.2.x catalog). */
const LEGACY_COUNT_BY_HOST: Record<Host, number> = {
  excel: 4,
  word: 4,
  powerpoint: 4,
};

function schemaIsObject(parameters: unknown): boolean {
  return (
    typeof parameters === "object" &&
    parameters !== null &&
    !Array.isArray(parameters)
  );
}

test("catalog hosts are exactly excel/word/powerpoint", () => {
  const hosts = new Set(OFFICE_CATALOG.map((entry) => entry.host));
  assert.deepEqual([...hosts].sort(), [...HOSTS].sort());
});

test("every entry has unique op id and unique pi tool name", () => {
  const opIds = OFFICE_CATALOG.map((entry) => `${entry.host}.${entry.op}`);
  const names = OFFICE_CATALOG.map((entry) => entry.name);
  assert.equal(new Set(opIds).size, opIds.length, "duplicate op id");
  assert.equal(new Set(names).size, names.length, "duplicate tool name");
});

test("every entry has non-empty label/description and an object schema", () => {
  for (const entry of OFFICE_CATALOG) {
    assert.ok(entry.label.length > 0, `${entry.name}: empty label`);
    assert.ok(entry.description.length > 0, `${entry.name}: empty description`);
    assert.ok(
      schemaIsObject(entry.parameters),
      `${entry.name}: schema not an object`,
    );
  }
});

test("every op id maps back to its entry via OFFICE_CATALOG_BY_OP", () => {
  for (const entry of OFFICE_CATALOG) {
    const byOp = OFFICE_CATALOG_BY_OP.get(`${entry.host}.${entry.op}`);
    assert.equal(byOp, entry);
  }
});

test("officeToolName / hostForToolName are inverse on every entry", () => {
  for (const entry of OFFICE_CATALOG) {
    assert.equal(officeToolName(entry.host, entry.op), entry.name);
    assert.equal(hostForToolName(entry.name), entry.host);
  }
  assert.equal(hostForToolName("office_excel_read_range"), "excel");
  assert.equal(hostForToolName("office_word_insert_text"), "word");
  assert.equal(hostForToolName("office_powerpoint_add_slide"), "powerpoint");
  assert.equal(hostForToolName("not_an_office_tool"), null);
});

test("OFFICE_TOOL_NAMES matches the catalog exactly", () => {
  assert.deepEqual(
    [...OFFICE_TOOL_NAMES].sort(),
    OFFICE_CATALOG.map((e) => e.name).sort(),
  );
});

test("catalogForHost / catalogOpIdsForHost are exact per host", () => {
  for (const host of HOSTS) {
    const entries = catalogForHost(host);
    assert.ok(entries.length > 0, `${host}: no entries`);
    assert.ok(
      entries.every((e) => e.host === host),
      `${host}: foreign entry`,
    );
    assert.deepEqual(
      catalogOpIdsForHost(host),
      entries.map((e) => `${e.host}.${e.op}`),
    );
  }
});

test("Word catalog exposes the full formatting/composition surface", () => {
  const wordOps = new Set(catalogForHost("word").map((e) => e.op));
  for (const op of [
    "get_overview",
    "read_document",
    "insert_text",
    "replace_text",
    "format_range",
    "insert_blocks",
    "insert_table",
    "insert_page_break",
    "insert_image",
    "insert_hyperlink",
  ]) {
    assert.ok(wordOps.has(op), `missing word.${op}`);
  }
});

test("Word formatting schemas accept the documented params", () => {
  const formatRange = OFFICE_CATALOG_BY_OP.get("word.format_range");
  assert.ok(formatRange);
  const params = formatRange.parameters as {
    properties?: Record<string, unknown>;
  };
  assert.ok(params.properties, "format_range: schema has no properties");
  for (const key of [
    "bold",
    "italic",
    "underline",
    "size",
    "name",
    "color",
    "alignment",
    "style",
  ]) {
    assert.ok(key in params.properties, `format_range: missing param ${key}`);
  }
});

test("Legacy v1 ops match the 0.2.x bridge catalog exactly", () => {
  assert.equal(LEGACY_V1_OPS.length, 12);
  assert.equal(new Set(LEGACY_V1_OPS).size, LEGACY_V1_OPS.length);
  assert.ok(LEGACY_V1_OPS.includes("word.get_overview"));
  assert.ok(LEGACY_V1_OPS.includes("word.read_document"));
  assert.ok(LEGACY_V1_OPS.includes("word.insert_text"));
  assert.ok(LEGACY_V1_OPS.includes("word.replace_text"));
  assert.ok(!LEGACY_V1_OPS.includes("word.format_range"));
  assert.ok(!LEGACY_V1_OPS.includes("word.insert_blocks"));
  assert.ok(!LEGACY_V1_OPS.includes("excel.format_cells"));
  assert.ok(!LEGACY_V1_OPS.includes("powerpoint.format_slide"));
  // Every legacy op must be a real catalog op for its host.
  for (const legacyOp of LEGACY_V1_OPS) {
    const [host, ...opParts] = legacyOp.split(".");
    const op = opParts.join(".");
    const entry = OFFICE_CATALOG_BY_OP.get(legacyOp);
    assert.ok(entry, `legacy op ${legacyOp} not in catalog`);
    assert.equal(
      entry.legacyV1,
      true,
      `legacy op ${legacyOp} must be marked legacyV1`,
    );
    assert.equal(hostForToolName(officeToolName(host as Host, op)), host);
  }
  // Per-host legacy counts (the union over entries marked legacyV1).
  for (const host of HOSTS) {
    const count = OFFICE_CATALOG.filter(
      (e) => e.host === host && e.legacyV1 === true,
    ).length;
    assert.equal(count, LEGACY_COUNT_BY_HOST[host], `${host} legacy count`);
  }
});

test("every schema round-trips through TypeBox validation as an object", () => {
  for (const entry of OFFICE_CATALOG) {
    assert.ok(schemaIsObject(entry.parameters), `${entry.name}: not an object`);
    assert.ok(
      typeof entry.parameters.type === "string" || "anyOf" in entry.parameters,
      `${entry.name}: not a TypeBox schema`,
    );
  }
});

test("catalog version is a positive integer", () => {
  assert.equal(typeof CATALOG_VERSION, "number");
  assert.ok(Number.isInteger(CATALOG_VERSION) && CATALOG_VERSION >= 1);
});

test("promptGuidelines name their tool when present", () => {
  for (const entry of OFFICE_CATALOG as OfficeCatalogEntry[]) {
    for (const guideline of entry.promptGuidelines ?? []) {
      assert.ok(
        guideline.includes(entry.name),
        `${entry.name}: guideline "${guideline}" does not name the tool`,
      );
    }
  }
});
