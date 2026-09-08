import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildExplainFormulaNarrative,
  extractFormulaFunctionNames,
  previewCellValue,
} from "../src/tools/explain-formula-logic.ts";

void test("extractFormulaFunctionNames returns ordered unique function names", () => {
  const names = extractFormulaFunctionNames("=IFERROR(SUM(A1:A3) / SUM(B1:B3), 0)");

  assert.deepEqual(names, ["IFERROR", "SUM"]);
});

void test("previewCellValue normalizes blanks and truncates long text", () => {
  assert.equal(previewCellValue(""), "(blank)");
  assert.equal(previewCellValue(null), "(blank)");

  const preview = previewCellValue("x".repeat(200));
  assert.equal(preview.endsWith("…"), true);
});

void test("buildExplainFormulaNarrative mentions reference count and truncation", () => {
  const narrative = buildExplainFormulaNarrative({
    valuePreview: "1520",
    functionNames: ["SUMIFS"],
    referenceCount: 14,
    truncated: true,
  });

  assert.match(narrative, /Current value: 1520\./u);
  assert.match(narrative, /14 direct references/u);
  assert.match(narrative, /truncated/u);
});
