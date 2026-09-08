import assert from "node:assert/strict";
import { test } from "node:test";

import {
  detectOfficeAppFromGlobals,
  officeAppLabel,
  parseOfficeApp,
  type OfficeApp,
} from "../src/host/index.ts";

test("parseOfficeApp maps Office host strings", () => {
  assert.equal(parseOfficeApp("Excel"), "excel");
  assert.equal(parseOfficeApp("excel"), "excel");
  assert.equal(parseOfficeApp("Word"), "word");
  assert.equal(parseOfficeApp("PowerPoint"), "powerpoint");
  assert.equal(parseOfficeApp("Power BI"), "powerbi");
  assert.equal(parseOfficeApp("Whatever"), "other");
  assert.equal(parseOfficeApp(null), null);
  assert.equal(parseOfficeApp(undefined), null);
  assert.equal(parseOfficeApp(""), null);
});

test("officeAppLabel renders human labels", () => {
  const labels: Record<OfficeApp, string> = {
    excel: "Excel",
    word: "Word",
    powerpoint: "PowerPoint",
    powerbi: "Power BI",
    other: "Other",
  };
  for (const [app, label] of Object.entries(labels)) {
    assert.equal(officeAppLabel(app as OfficeApp), label);
  }
  assert.equal(officeAppLabel(null), "Unknown");
});

test("detectOfficeAppFromGlobals reads Office.context.host", () => {
  const scope = {
    Office: { context: { host: "Word" } },
  } as unknown as Record<string, unknown>;
  assert.equal(detectOfficeAppFromGlobals(scope), "word");
});

test("detectOfficeAppFromGlobals is null without an Office global", () => {
  assert.equal(detectOfficeAppFromGlobals({}), null);
});
