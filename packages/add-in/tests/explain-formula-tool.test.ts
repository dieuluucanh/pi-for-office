import assert from "node:assert/strict";
import { test } from "node:test";

import { isSingleCellReference } from "../src/tools/explain-formula.ts";

void test("isSingleCellReference accepts quoted sheet names containing commas", () => {
  assert.equal(isSingleCellReference("'Sales, Q1'!A1"), true);
});

void test("isSingleCellReference rejects ranges and multi-area addresses", () => {
  assert.equal(isSingleCellReference("Sheet1!A1:B2"), false);
  assert.equal(isSingleCellReference("A1,B2"), false);
});
