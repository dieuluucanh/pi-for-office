/**
 * format_cells tests — exercised against the fake `Excel` global.
 *
 * Covers the completion added after the original capability audit:
 *  - `border_style` (all eight ExcelApi-1.1 `BorderLineStyle` values)
 *  - underline styles (boolean legacy alias + named `RangeUnderlineStyle`)
 *  - all eight `HorizontalAlignment` values
 *  - the ExcelApi 1.9 gate on `strikethrough` / `indent_level` (fail-loud on an
 *    older host, applied on a capable one)
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  EXCEL_BORDER_LINE_STYLES,
  EXCEL_HORIZONTAL_ALIGNMENTS,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { createFormatCellsTool } from "../src/tools/format-cells.ts";
import type { FormatCellsDetails } from "../src/tools/tool-details.ts";
import {
  installFakeAppStorage,
  withFakeExcel,
  type FakeContext,
} from "./helpers/fake-excel.ts";

installFakeAppStorage();

function toolText(result: AgentToolResult<FormatCellsDetails>): string {
  const first = result.content[0];
  if (!first || first.type !== "text") {
    throw new Error("Expected first content block to be text.");
  }
  return first.text;
}

/** The range the tool itself created (the capture path creates one first). */
function mainRange(context: FakeContext) {
  const sheet = context.workbook.worksheets.items[0];
  if (!sheet) throw new Error("No sheet in fake context.");
  const range = sheet.createdRanges.at(-1);
  if (!range) throw new Error("format_cells did not create a range.");
  return range;
}

void test("border_style reaches the border item for a non-Continuous style", async () => {
  const result = await withFakeExcel(async (_context) => {
    const tool = createFormatCellsTool();
    return tool.execute("tc-border-style", {
      range: "Sheet1!A1:B2",
      border_top: "thin",
      border_style: "Dash",
    });
  }, {});

  // Built after the run so the recorded range is available to assert on.
  assert.ok(toolText(result).startsWith("Formatted"));
});

void test("border_style applies the named line style to the requested edges", async () => {
  const seen: Record<string, { style?: string; weight?: string }> = {};

  await withFakeExcel(async (context) => {
    const tool = createFormatCellsTool();
    const result = await tool.execute("tc-border-style-2", {
      range: "Sheet1!A1:B2",
      border_bottom: "medium",
      border_style: "DashDot",
    });
    assert.match(toolText(result), /Formatted/);
    const range = mainRange(context);
    for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]) {
      const border = range.format.borders.items.get(edge);
      if (border) seen[edge] = { style: border.style, weight: border.weight };
    }
    return result;
  }, {});

  assert.deepEqual(seen["EdgeBottom"], { style: "DashDot", weight: "Medium" });
  assert.equal(seen["EdgeTop"], undefined);
});

void test("weight 'none' forces BorderLineStyle None regardless of border_style", async () => {
  let style: string | undefined;

  await withFakeExcel(async (context) => {
    const tool = createFormatCellsTool();
    await tool.execute("tc-border-none", {
      range: "Sheet1!A1",
      border_top: "none",
      border_style: "Double",
    });
    style = mainRange(context).format.borders.items.get("EdgeTop")?.style;
  }, {});

  assert.equal(style, "None");
});

void test("every advertised border style is accepted by the schema", () => {
  // The catalog constant is the single source both the tool schema and the
  // implementation read; a `Hair` member (which does not exist in Office.js)
  // must never come back.
  assert.deepEqual(
    [...EXCEL_BORDER_LINE_STYLES].sort(),
    [
      "Continuous",
      "Dash",
      "DashDot",
      "DashDotDot",
      "Double",
      "Dot",
      "None",
      "SlantDashDot",
    ].sort(),
  );
  assert.ok(!(EXCEL_BORDER_LINE_STYLES as readonly string[]).includes("Hair"));
});

void test("underline boolean alias maps to Single/None", async () => {
  const values: Array<string | undefined> = [];

  await withFakeExcel(async (context) => {
    const tool = createFormatCellsTool();
    await tool.execute("tc-underline-true", {
      range: "Sheet1!A1",
      underline: true,
    });
    values.push(mainRange(context).format.font.underline);
    await tool.execute("tc-underline-false", {
      range: "Sheet1!A1",
      underline: false,
    });
    values.push(mainRange(context).format.font.underline);
  }, {});

  assert.deepEqual(values, ["Single", "None"]);
});

void test("a named underline style passes through unchanged", async () => {
  let value: string | undefined;

  await withFakeExcel(async (context) => {
    const tool = createFormatCellsTool();
    await tool.execute("tc-underline-style", {
      range: "Sheet1!A1",
      underline: "DoubleAccountant",
    });
    value = mainRange(context).format.font.underline;
  }, {});

  assert.equal(value, "DoubleAccountant");
});

void test("an unknown underline string is never cast onto the range font", async () => {
  let value: string | undefined;

  await withFakeExcel(async (context) => {
    const tool = createFormatCellsTool();
    // Bypass the schema union: stored conventions can carry stale strings.
    const badParams: DynamicValue = {
      range: "Sheet1!A1",
      underline: "WavyHeavy",
    };
    await tool.execute(
      "tc-underline-bad",
      badParams as Parameters<typeof tool.execute>[1],
    );
    value = mainRange(context).format.font.underline;
  }, {});

  assert.equal(value, "None");
});

void test("all eight horizontal alignments are applied", async () => {
  const applied: Array<string | undefined> = [];

  await withFakeExcel(async (context) => {
    const tool = createFormatCellsTool();
    for (const alignment of EXCEL_HORIZONTAL_ALIGNMENTS) {
      await tool.execute("tc-align", {
        range: "Sheet1!A1",
        horizontal_alignment: alignment,
      });
      applied.push(mainRange(context).format.horizontalAlignment);
    }
  }, {});

  assert.deepEqual(applied, [...EXCEL_HORIZONTAL_ALIGNMENTS]);
});

void test("strikethrough and indent_level fail loudly below ExcelApi 1.9", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createFormatCellsTool();
      return tool.execute("tc-gate-unsupported", {
        range: "Sheet1!A1",
        strikethrough: true,
        indent_level: 2,
      });
    },
    { capabilities: { supported: { ExcelApi: "1.8" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs ExcelApi 1\.9/);
  assert.match(text, /platform: /);
});

void test("strikethrough and indent_level apply when the host has ExcelApi 1.9", async () => {
  let strikethrough: boolean | undefined;
  let indentLevel: number | undefined;

  await withFakeExcel(
    async (context) => {
      const tool = createFormatCellsTool();
      const result = await tool.execute("tc-gate-supported", {
        range: "Sheet1!A1",
        strikethrough: true,
        indent_level: 2,
      });
      assert.match(toolText(result), /Formatted/);
      const format = mainRange(context).format;
      strikethrough = format.font.strikethrough;
      indentLevel = format.indentLevel;
    },
    { capabilities: { supported: { ExcelApi: "1.9" } } },
  );

  assert.equal(strikethrough, true);
  assert.equal(indentLevel, 2);
});
