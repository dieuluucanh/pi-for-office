import assert from "node:assert/strict";
import { test } from "node:test";

import {
  extractFormulaReferences,
  formulaReferencesTargetCell,
  normalizeTraceMode,
  summarizeTraceTree,
} from "../src/tools/trace-dependencies-logic.ts";
import type { DepNodeDetail } from "../src/tools/tool-details.ts";

void test("normalizeTraceMode defaults to precedents", () => {
  assert.equal(normalizeTraceMode(undefined), "precedents");
  assert.equal(normalizeTraceMode("dependents"), "dependents");
  assert.equal(normalizeTraceMode("anything-else"), "precedents");
});

void test("extractFormulaReferences parses local and cross-sheet references", () => {
  const refs = extractFormulaReferences(
    "=SUM(A1:B2,'Input Data'!$C$5,Sheet2!D4:D9)",
    "Calc",
  );

  assert.equal(refs.length, 3);

  assert.equal(refs[0]?.startAddress, "Calc!A1");
  assert.equal(refs[0]?.endCol, 1);
  assert.equal(refs[0]?.endRow, 2);

  assert.equal(refs[1]?.startAddress, "'Input Data'!C5");
  assert.equal(refs[1]?.sheet, "Input Data");

  assert.equal(refs[2]?.startAddress, "Sheet2!D4");
  assert.equal(refs[2]?.endRow, 9);
});

void test("formulaReferencesTargetCell matches cells inside ranges", () => {
  const formula = "=SUM(A1:B3, Sheet2!D5:D8)";

  assert.equal(formulaReferencesTargetCell(formula, "Calc", "Calc!B2"), true);
  assert.equal(formulaReferencesTargetCell(formula, "Calc", "Calc!C4"), false);
  assert.equal(formulaReferencesTargetCell(formula, "Calc", "Sheet2!D7"), true);
  assert.equal(formulaReferencesTargetCell(formula, "Calc", "Sheet2!E7"), false);
});

void test("extractFormulaReferences ignores A1-like tokens inside quoted literals", () => {
  const formula = "=IF(A1=\"B2\", A1, 0)";
  const refs = extractFormulaReferences(formula, "Calc");

  assert.deepEqual(refs.map((ref) => ref.startAddress), ["Calc!A1"]);
  assert.equal(formulaReferencesTargetCell(formula, "Calc", "Calc!B2"), false);
});

void test("summarizeTraceTree returns node and edge counts", () => {
  const tree: DepNodeDetail = {
    address: "Sheet1!D10",
    value: 42,
    formula: "=B10+C10",
    precedents: [
      {
        address: "Sheet1!B10",
        value: 20,
        precedents: [],
      },
      {
        address: "Sheet1!C10",
        value: 22,
        formula: "=C5+C6",
        precedents: [
          {
            address: "Sheet1!C5",
            value: 10,
            precedents: [],
          },
        ],
      },
    ],
  };

  assert.deepEqual(summarizeTraceTree(tree), { nodeCount: 4, edgeCount: 3 });
});
