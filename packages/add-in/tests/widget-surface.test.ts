import assert from "node:assert/strict";
import { test } from "node:test";

import {
  normalizeWidgetHeightBounds,
  resolveWidgetCollapseState,
  resolveWidgetHeightBoundsForUpsert,
} from "../src/extensions/internal/widget-surface.ts";

void test("normalizeWidgetHeightBounds coerces max to min when max < min", () => {
  const normalized = normalizeWidgetHeightBounds(320, 120);

  assert.deepEqual(normalized, {
    minHeightPx: 320,
    maxHeightPx: 320,
  });
});

void test("resolveWidgetHeightBoundsForUpsert preserves existing bounds when values are omitted", () => {
  const normalized = resolveWidgetHeightBoundsForUpsert(
    undefined,
    undefined,
    {
      minHeightPx: 180,
      maxHeightPx: 420,
    },
  );

  assert.deepEqual(normalized, {
    minHeightPx: 180,
    maxHeightPx: 420,
  });
});

void test("resolveWidgetHeightBoundsForUpsert supports clearing bounds with null", () => {
  const normalized = resolveWidgetHeightBoundsForUpsert(
    null,
    null,
    {
      minHeightPx: 180,
      maxHeightPx: 420,
    },
  );

  assert.deepEqual(normalized, {
    minHeightPx: null,
    maxHeightPx: null,
  });
});

void test("resolveWidgetHeightBoundsForUpsert clamps values into safe host range", () => {
  const normalized = resolveWidgetHeightBoundsForUpsert(
    12,
    2000,
    null,
  );

  assert.deepEqual(normalized, {
    minHeightPx: 72,
    maxHeightPx: 640,
  });
});

void test("resolveWidgetCollapseState keeps collapsed false when widget is not collapsible", () => {
  const state = resolveWidgetCollapseState(
    false,
    true,
    {
      collapsible: true,
      collapsed: true,
    },
  );

  assert.deepEqual(state, {
    collapsible: false,
    collapsed: false,
  });
});

void test("resolveWidgetCollapseState preserves existing state when fields are omitted", () => {
  const state = resolveWidgetCollapseState(
    undefined,
    undefined,
    {
      collapsible: true,
      collapsed: true,
    },
  );

  assert.deepEqual(state, {
    collapsible: true,
    collapsed: true,
  });
});
