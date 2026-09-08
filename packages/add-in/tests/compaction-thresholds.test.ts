import assert from "node:assert/strict";
import { test } from "node:test";

import { getCompactionThresholds } from "../src/compaction/defaults.ts";

void test("applies earlier hard compaction cap for 200k-context models", () => {
  const thresholds = getCompactionThresholds(200_000);

  assert.equal(thresholds.contextWindow, 200_000);
  assert.equal(thresholds.reserveTokens, 16_384);
  assert.equal(thresholds.hardTriggerTokens, 170_000);
  assert.equal(thresholds.softWarningTokens, 160_000);
});

void test("keeps reserve-based trigger for smaller context windows", () => {
  const thresholds = getCompactionThresholds(32_768);

  assert.equal(thresholds.reserveTokens, 16_384);
  assert.equal(thresholds.hardTriggerTokens, 16_384);
  assert.equal(thresholds.softWarningTokens, 14_336);
});

void test("normalizes invalid context window inputs to fallback values", () => {
  const thresholds = getCompactionThresholds(Number.NaN);

  assert.equal(thresholds.contextWindow, 200_000);
  assert.equal(thresholds.reserveTokens, 16_384);
  assert.equal(thresholds.hardTriggerTokens, 170_000);
  assert.equal(thresholds.softWarningTokens, 160_000);
});
