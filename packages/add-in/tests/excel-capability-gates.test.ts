/**
 * Excel capability gates (step 9).
 *
 * Every gate needs BOTH paths: an unsupported host must fail loudly (capability
 * + set + version + platform), and a capable host must still run the operation.
 * A gate tested only on the unsupported path can silently break a host where
 * the feature worked.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { createViewSettingsTool } from "../src/tools/view-settings.ts";
import { createModifyStructureTool } from "../src/tools/modify-structure.ts";
import { createFillFormulaTool } from "../src/tools/fill-formula.ts";
import { createCommentsTool } from "../src/tools/comments.ts";
import { createChartsTool } from "../src/tools/charts.ts";
import { createFormatCellsTool } from "../src/tools/format-cells.ts";
import { buildOverview } from "../src/tools/get-workbook-overview.ts";
import type { ViewSettingsDetails } from "../src/tools/tool-details.ts";
import {
  installFakeAppStorage,
  withFakeExcel,
  type FakeContext,
} from "./helpers/fake-excel.ts";

installFakeAppStorage();

function toolText(result: {
  content: Array<{ type: string; text?: DynamicValue }>;
}): string {
  const first = result.content[0];
  if (!first || first.type !== "text" || typeof first.text !== "string") {
    throw new Error("Expected first content block to be text.");
  }
  return first.text;
}

/* ── view_settings: gridlines/headings (ExcelApi 1.8) ───────────────────── */

void test("view_settings gridlines fail loudly below ExcelApi 1.8", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createViewSettingsTool();
      return tool.execute("tc-vs-gate", { action: "hide_gridlines" });
    },
    { capabilities: { supported: { ExcelApi: "1.7" } } },
  );

  const text = toolText(result);
  assert.match(text, /ExcelApi 1\.8/);
  assert.match(text, /gridline/);
  assert.match(text, /platform: /);
});

void test("view_settings gridlines apply on ExcelApi 1.8", async () => {
  const sheet = await withFakeExcel(
    async (context) => {
      const tool = createViewSettingsTool();
      const result = (await tool.execute("tc-vs-ok", {
        action: "hide_gridlines",
      })) as AgentToolResult<ViewSettingsDetails>;
      assert.match(toolText(result), /Gridlines hidden/);
      return context.workbook.worksheets.items[0];
    },
    { capabilities: { supported: { ExcelApi: "1.8" } } },
  );

  assert.equal(sheet?.showGridlines, false);
});

/* ── modify_structure: worksheet visibility setter (ExcelApi 1.2) ────────── */

void test("hide_sheet fails loudly below ExcelApi 1.2", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createModifyStructureTool();
      return tool.execute("tc-ms-gate", {
        action: "hide_sheet",
        sheet: "Sheet1",
      });
    },
    { capabilities: { supported: { ExcelApi: "1.1" } } },
  );

  const text = toolText(result);
  assert.match(text, /ExcelApi 1\.2/);
  assert.match(text, /hiding and unhiding worksheets/);
});

void test("hide_sheet applies on ExcelApi 1.2", async () => {
  const sheet = await withFakeExcel(
    async (context) => {
      const tool = createModifyStructureTool();
      const result = await tool.execute("tc-ms-ok", {
        action: "hide_sheet",
        sheet: "Sheet1",
      });
      assert.match(toolText(result), /Hidden sheet/);
      return context.workbook.worksheets.items[0];
    },
    { capabilities: { supported: { ExcelApi: "1.2" } } },
  );

  assert.equal(sheet?.visibility, "Hidden");
});

/* ── fill_formula: autoFill (ExcelApi 1.9) ──────────────────────────────── */

void test("fill_formula fails loudly below ExcelApi 1.9", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createFillFormulaTool();
      return tool.execute("tc-ff-gate", {
        range: "Sheet1!B2:B20",
        formula: "=A2*2",
      });
    },
    { capabilities: { supported: { ExcelApi: "1.8" } } },
  );

  const text = toolText(result);
  assert.match(text, /ExcelApi 1\.9/);
  assert.match(text, /autofill/i);
});

void test("fill_formula reaches the host on ExcelApi 1.9", async () => {
  const outcome = await withFakeExcel(
    async (context) => {
      const tool = createFillFormulaTool();
      const result = await tool.execute("tc-ff-ok", {
        range: "Sheet1!B2:B20",
        formula: "=A2*2",
      });
      return { text: toolText(result), syncs: context.syncCount };
    },
    { capabilities: { supported: { ExcelApi: "1.9" } } },
  );

  // The gate must not pre-empt the call: the operation ran against the host.
  assert.ok(!/ExcelApi 1\.9/.test(outcome.text), outcome.text);
  assert.ok(outcome.syncs > 0, "expected at least one Office.js sync");
});

/* ── charts: getImage (ExcelApi 1.7) ────────────────────────────────────── */

void test("charts get_image fails loudly below ExcelApi 1.7", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createChartsTool();
      return tool.execute("tc-ch-gate", { action: "get_image", name: "Sales" });
    },
    { capabilities: { supported: { ExcelApi: "1.6" } } },
  );

  const text = toolText(result);
  assert.match(text, /ExcelApi 1\.7/);
  assert.match(text, /image/i);
});

void test("charts get_image reaches the host on ExcelApi 1.7", async () => {
  const outcome = await withFakeExcel(
    async (context) => {
      const tool = createChartsTool();
      const result = await tool.execute("tc-ch-ok", {
        action: "get_image",
        name: "Sales",
      });
      return { text: toolText(result), syncs: context.syncCount };
    },
    { capabilities: { supported: { ExcelApi: "1.7" } } },
  );

  assert.ok(!/ExcelApi 1\.7/.test(outcome.text), outcome.text);
  assert.ok(outcome.syncs > 0, "expected at least one Office.js sync");
});

/* ── comments (ExcelApi 1.10) ───────────────────────────────────────────── */

void test("comments fail loudly below ExcelApi 1.10", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createCommentsTool();
      return tool.execute("tc-cm-gate", { action: "read", range: "Sheet1!A1" });
    },
    { capabilities: { supported: { ExcelApi: "1.9" } } },
  );

  const text = toolText(result);
  assert.match(text, /ExcelApi 1\.10/);
  assert.match(text, /comments/i);
});

void test("comments reach the host on ExcelApi 1.10", async () => {
  const outcome = await withFakeExcel(
    async (context) => {
      const tool = createCommentsTool();
      const result = await tool.execute("tc-cm-ok", {
        action: "read",
        range: "Sheet1!A1",
      });
      return { text: toolText(result), syncs: context.syncCount };
    },
    { capabilities: { supported: { ExcelApi: "1.10" } } },
  );

  assert.ok(!/ExcelApi 1\.10/.test(outcome.text), outcome.text);
  assert.ok(outcome.syncs > 0, "expected at least one Office.js sync");
});

/* ── format_cells: multi-area ranges (RangeAreas, ExcelApi 1.9) ─────────── */

void test("multi-range format_cells fails loudly below ExcelApi 1.9", async () => {
  const result = await withFakeExcel(
    async () => {
      const tool = createFormatCellsTool();
      return tool.execute("tc-fc-gate", {
        range: "Sheet1!A1:B2, D1:D2",
        bold: true,
      });
    },
    { capabilities: { supported: { ExcelApi: "1.8" } } },
  );

  const text = toolText(result);
  assert.match(text, /ExcelApi 1\.9/);
  assert.match(text, /multi-area range operations/);
});

void test("multi-range format_cells applies on ExcelApi 1.9", async () => {
  const outcome = await withFakeExcel(
    async (context) => {
      const tool = createFormatCellsTool();
      const result = await tool.execute("tc-fc-ok", {
        range: "Sheet1!A1:B2, D1:D2",
        bold: true,
      });
      const sheet = context.workbook.worksheets.items[0];
      return {
        text: toolText(result),
        areaFormat: sheet?.createdAreas.at(-1)?.format.font.bold,
        areaCount: sheet?.createdAreas.at(-1)?.areas.items.length,
      };
    },
    { capabilities: { supported: { ExcelApi: "1.9" } } },
  );

  assert.match(outcome.text, /Formatted/);
  assert.equal(outcome.areaCount, 2);
  assert.equal(outcome.areaFormat, true);
});

/* ── get_workbook_overview: shapes (ExcelApi 1.9) ───────────────────────── */

void test("overview reports the shape count as unavailable below ExcelApi 1.9", async () => {
  const text = await withFakeExcel(async () => buildOverview(), {
    capabilities: { supported: { ExcelApi: "1.8" } },
  });

  assert.match(text, /shape count unavailable \(needs ExcelApi 1\.9\)/);
});

void test("overview reports a real shape count on ExcelApi 1.9", async () => {
  const text = await withFakeExcel(
    async (context: FakeContext) => {
      const sheet = context.workbook.worksheets.items[0];
      if (!sheet) throw new Error("No sheet in fake context.");
      sheet.shapes.items.push({ name: "Oval 1" });
      return buildOverview();
    },
    { capabilities: { supported: { ExcelApi: "1.9" } } },
  );

  assert.match(text, /1 shape\(s\)/);
  assert.ok(!/unavailable/.test(text), text);
});
