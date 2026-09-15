/**
 * conditional_format tests — one per `Excel.ConditionalFormatType` family plus
 * the ExcelApi 1.6 gate.
 *
 * Every family is ExcelApi 1.6, so the gate is exercised once on the add path
 * and once on the clear path rather than six times.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { createConditionalFormatTool } from "../src/tools/conditional-format.ts";
import type { ConditionalFormatDetails } from "../src/tools/tool-details.ts";
import {
  installFakeAppStorage,
  withFakeExcel,
  type FakeConditionalFormat,
  type FakeContext,
} from "./helpers/fake-excel.ts";

installFakeAppStorage();

function toolText(result: AgentToolResult<ConditionalFormatDetails>): string {
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
  if (!range) throw new Error("conditional_format did not create a range.");
  return range;
}

/** The conditional format the tool added to the main range. */
function addedFormat(context: FakeContext): FakeConditionalFormat {
  const added = mainRange(context).conditionalFormats.added;
  const last = added.at(-1);
  if (!last) throw new Error("No conditional format was added.");
  return last;
}

void test("formula and cell_value still work (regression)", async () => {
  const types: string[] = [];
  let cellValueRule: {
    formula1?: string;
    formula2?: string;
    operator?: string;
  } = {};

  await withFakeExcel(async (context) => {
    const tool = createConditionalFormatTool();
    const formula = await tool.execute("tc-cf-formula", {
      action: "add",
      range: "Sheet1!A1:B2",
      type: "formula",
      formula: "=A1>0",
    });
    assert.match(toolText(formula), /formula rule/);
    types.push(addedFormat(context).type);

    const cellValue = await tool.execute("tc-cf-cell", {
      action: "add",
      range: "Sheet1!A1:B2",
      type: "cell_value",
      operator: "Between",
      value: 1,
      value2: 10,
    });
    assert.match(toolText(cellValue), /cell value rule/);
    cellValueRule = addedFormat(context).cellValue.rule;
  }, {});

  assert.deepEqual(types, ["custom"]);
  assert.deepEqual(cellValueRule, {
    formula1: "1",
    formula2: "10",
    operator: "Between",
  });
});

void test("data_bar applies direction, fill colors and axis color", async () => {
  let dataBar: FakeConditionalFormat["dataBar"] | undefined;

  await withFakeExcel(async (context) => {
    const tool = createConditionalFormatTool();
    const result = await tool.execute("tc-cf-databar", {
      action: "add",
      range: "Sheet1!A1:A10",
      type: "data_bar",
      data_bar_direction: "LeftToRight",
      data_bar_fill_color: "#638EC6",
      data_bar_negative_fill_color: "#FF0000",
      data_bar_show_only: true,
      data_bar_axis_color: "#000000",
    });
    assert.match(toolText(result), /data bar/);
    dataBar = addedFormat(context).dataBar;
  }, {});

  assert.equal(dataBar?.barDirection, "LeftToRight");
  assert.equal(dataBar?.showDataBarOnly, true);
  assert.equal(dataBar?.axisColor, "#000000");
  assert.equal(dataBar?.positiveFormat.fillColor, "#638EC6");
  assert.equal(dataBar?.negativeFormat.fillColor, "#FF0000");
});

void test("color_scale applies minimum/midpoint/maximum criteria", async () => {
  let criteria: FakeConditionalFormat["colorScale"]["criteria"] | undefined;

  await withFakeExcel(async (context) => {
    const tool = createConditionalFormatTool();
    await tool.execute("tc-cf-colorscale", {
      action: "add",
      range: "Sheet1!A1:A10",
      type: "color_scale",
      color_scale_min_color: "#FF0000",
      color_scale_mid_color: "#FFFF00",
      color_scale_max_color: "#00FF00",
    });
    criteria = addedFormat(context).colorScale.criteria;
  }, {});

  assert.deepEqual(criteria?.minimum, {
    type: "LowestValue",
    color: "#FF0000",
  });
  assert.deepEqual(criteria?.midpoint, {
    type: "Percentile",
    formula: "50",
    color: "#FFFF00",
  });
  assert.deepEqual(criteria?.maximum, {
    type: "HighestValue",
    color: "#00FF00",
  });
});

void test("color_scale omits the midpoint for a two-color scale", async () => {
  let criteria: FakeConditionalFormat["colorScale"]["criteria"] | undefined;

  await withFakeExcel(async (context) => {
    const tool = createConditionalFormatTool();
    await tool.execute("tc-cf-colorscale-2", {
      action: "add",
      range: "Sheet1!A1:A10",
      type: "color_scale",
    });
    criteria = addedFormat(context).colorScale.criteria;
  }, {});

  assert.equal(criteria?.midpoint, undefined);
});

void test("icon_set applies style, reverse order and icon-only", async () => {
  let iconSet: FakeConditionalFormat["iconSet"] | undefined;

  await withFakeExcel(async (context) => {
    const tool = createConditionalFormatTool();
    await tool.execute("tc-cf-iconset", {
      action: "add",
      range: "Sheet1!A1:A10",
      type: "icon_set",
      icon_set_style: "FourRating",
      icon_set_reverse_order: true,
      icon_set_show_icon_only: true,
    });
    iconSet = addedFormat(context).iconSet;
  }, {});

  assert.equal(iconSet?.style, "FourRating");
  assert.equal(iconSet?.reverseIconOrder, true);
  assert.equal(iconSet?.showIconOnly, true);
});

void test("top_bottom applies the rank rule and font formatting", async () => {
  let topBottom: FakeConditionalFormat["topBottom"] | undefined;

  await withFakeExcel(async (context) => {
    const tool = createConditionalFormatTool();
    const result = await tool.execute("tc-cf-topbottom", {
      action: "add",
      range: "Sheet1!A1:A10",
      type: "top_bottom",
      top_bottom_type: "BottomPercent",
      top_bottom_rank: 25,
      fill_color: "#FFC7CE",
      bold: true,
    });
    assert.match(toolText(result), /top\/bottom rule/);
    topBottom = addedFormat(context).topBottom;
  }, {});

  assert.deepEqual(topBottom?.rule, { type: "BottomPercent", rank: 25 });
  assert.equal(topBottom?.format.fill.color, "#FFC7CE");
  assert.equal(topBottom?.format.font.bold, true);
});

void test("preset_criteria applies the criterion and formatting", async () => {
  let preset: FakeConditionalFormat["preset"] | undefined;

  await withFakeExcel(async (context) => {
    const tool = createConditionalFormatTool();
    await tool.execute("tc-cf-preset", {
      action: "add",
      range: "Sheet1!A1:A10",
      type: "preset_criteria",
      preset_criterion: "DuplicateValues",
      font_color: "#9C0006",
    });
    preset = addedFormat(context).preset;
  }, {});

  assert.deepEqual(preset?.rule, { criterion: "DuplicateValues" });
  assert.equal(preset?.format.font.color, "#9C0006");
});

void test("contains_text applies operator and text, and requires text", async () => {
  let textComparison: FakeConditionalFormat["textComparison"] | undefined;

  await withFakeExcel(async (context) => {
    const tool = createConditionalFormatTool();
    const ok = await tool.execute("tc-cf-text", {
      action: "add",
      range: "Sheet1!A1:A10",
      type: "contains_text",
      text_operator: "BeginsWith",
      text: "ERR",
    });
    assert.match(toolText(ok), /text rule/);
    textComparison = addedFormat(context).textComparison;

    const missing = await tool.execute("tc-cf-text-missing", {
      action: "add",
      range: "Sheet1!A1:A10",
      type: "contains_text",
    });
    assert.match(toolText(missing), /text is required/);
  }, {});

  assert.deepEqual(textComparison?.rule, {
    operator: "BeginsWith",
    text: "ERR",
  });
});

void test("every conditional-format family needs ExcelApi 1.6 (fail loud below it)", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createConditionalFormatTool();
      return tool.execute("tc-cf-gate", {
        action: "add",
        range: "Sheet1!A1",
        type: "data_bar",
      });
    },
    { capabilities: { supported: { ExcelApi: "1.5" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs ExcelApi 1\.6/);
  assert.match(text, /platform: /);
  assert.match(text, /condition-based formats/);
});

void test("clear is also gated on ExcelApi 1.6", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createConditionalFormatTool();
      return tool.execute("tc-cf-clear-gate", {
        action: "clear",
        range: "Sheet1!A1",
      });
    },
    { capabilities: { supported: { ExcelApi: "1.5" } } },
  );

  assert.match(toolText(result), /needs ExcelApi 1\.6/);
});

void test("a host with ExcelApi 1.6 applies the family", async () => {
  let type: string | undefined;

  await withFakeExcel(
    async (context) => {
      const tool = createConditionalFormatTool();
      const result = await tool.execute("tc-cf-supported", {
        action: "add",
        range: "Sheet1!A1:A10",
        type: "icon_set",
        icon_set_style: "ThreeFlags",
      });
      assert.match(toolText(result), /Added conditional format/);
      type = addedFormat(context).type;
    },
    { capabilities: { supported: { ExcelApi: "1.6" } } },
  );

  assert.equal(type, "iconSet");
});
