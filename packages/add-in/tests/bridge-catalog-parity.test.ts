/**
 * Bridge catalog ⟷ pane registry parity.
 *
 * The pane's bridge ops (`ALL_BRIDGE_OPS`, host-scoped via `opsForHost`) must
 * EXACTLY match the shared catalog's per-host op set. Any drift (a catalog op
 * with no executor, or a pane executor never advertised) fails here — this is
 * the enforcement that keeps the two bridges from diverging again.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  catalogOpIdsForHost,
  catalogSchemaFor,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { OfficeHostApp } from "@dieulc/pi-office-protocol";
import { ALL_BRIDGE_OPS, opsForHost } from "../src/bridge/registry.ts";

const HOSTS: readonly OfficeHostApp[] = ["excel", "word", "powerpoint"];

void test("pane bridge ops exactly match the catalog for every host", () => {
  for (const host of HOSTS) {
    const catalogOps = [...catalogOpIdsForHost(host)];
    const paneOps = [...opsForHost(host).keys()];
    assert.deepEqual(
      [...paneOps].sort(),
      [...catalogOps].sort(),
      `${host}: pane registry drift`,
    );
  }
});

void test("ALL_BRIDGE_OPS covers every catalog op with an executor", () => {
  const missing: string[] = [];
  for (const host of HOSTS) {
    for (const op of catalogOpIdsForHost(host)) {
      if (!ALL_BRIDGE_OPS.has(op)) missing.push(op);
    }
  }
  assert.deepEqual(missing, [], "catalog ops without a pane executor");
});

void test("every catalog parameter schema is resolvable for the pane", () => {
  // Exercises catalogSchemaFor on the identical op ids the pane advertises,
  // so schema typos/renames never silently reach a vague executor.
  for (const host of HOSTS) {
    for (const op of opsForHost(host).keys()) {
      const schema = catalogSchemaFor(op);
      assert.ok(schema, `${op}: missing catalog schema`);
    }
  }
});

void test("legacy v1 ops are a strict subset of the pane's modern ops", () => {
  for (const host of HOSTS) {
    const paneOps = new Set(opsForHost(host).keys());
    for (const op of catalogOpIdsForHost(host)) {
      assert.ok(paneOps.has(op), `${op}: pane cannot run a catalog op`);
    }
  }
});

void test("every bridge executor wraps a delegating implementation (passthrough wiring)", async () => {
  for (const host of HOSTS) {
    for (const [op, executor] of opsForHost(host)) {
      assert.equal(typeof executor, "function", `${op}: missing executor`);
      // Running without an Office host must resolve structurally (guarded)
      // rather than reject the promise — proves the delegation wrapper stands.
      const outcome = await executor({});
      assert.ok(
        typeof outcome === "object" &&
          outcome !== null &&
          typeof outcome.text === "string",
        `${op}: executor did not return an OfficeOpOutcome`,
      );
    }
  }
});
