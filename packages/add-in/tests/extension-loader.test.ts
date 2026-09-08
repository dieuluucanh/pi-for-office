import assert from "node:assert/strict";
import { test } from "node:test";

import {
  collectActivationCleanups,
  createLoadedExtensionHandle,
  getExtensionActivator,
  getExtensionDeactivator,
} from "../src/commands/extension-loader.ts";

void test("collectActivationCleanups accepts undefined, function, and function arrays", () => {
  const cleanupA = () => undefined;
  const cleanupB = () => undefined;

  assert.deepEqual(collectActivationCleanups(undefined), []);
  assert.deepEqual(collectActivationCleanups(cleanupA), [cleanupA]);
  assert.deepEqual(collectActivationCleanups([cleanupA, cleanupB]), [cleanupA, cleanupB]);
});

void test("collectActivationCleanups rejects invalid return values", () => {
  assert.throws(
    () => collectActivationCleanups("not-valid"),
    /activate\(api\) must return void, a cleanup function, or an array of cleanup functions/,
  );

  assert.throws(
    () => collectActivationCleanups([() => undefined, "bad"]),
    /activate\(api\) returned an invalid cleanup entry; expected a function/,
  );
});

void test("getExtensionActivator resolves activate then default", () => {
  const activate = () => undefined;
  const fallback = () => undefined;

  assert.equal(getExtensionActivator({ activate }), activate);
  assert.equal(getExtensionActivator({ default: fallback }), fallback);
  assert.equal(getExtensionActivator({}), null);
});

void test("getExtensionDeactivator resolves optional deactivate", () => {
  const deactivate = () => undefined;

  assert.equal(getExtensionDeactivator({ deactivate }), deactivate);
  assert.equal(getExtensionDeactivator({}), null);
  assert.equal(getExtensionDeactivator(null), null);
});

void test("createLoadedExtensionHandle executes cleanup in reverse order and only once", async () => {
  const calls: string[] = [];

  const handle = createLoadedExtensionHandle([
    () => {
      calls.push("cleanup-1");
    },
    () => {
      calls.push("cleanup-2");
    },
  ], () => {
    calls.push("module-deactivate");
  });

  await handle.deactivate();
  await handle.deactivate();

  assert.deepEqual(calls, ["cleanup-2", "cleanup-1", "module-deactivate"]);
});

void test("createLoadedExtensionHandle aggregates cleanup failures", async () => {
  const handle = createLoadedExtensionHandle([
    () => {
      throw new Error("cleanup-a");
    },
  ], () => {
    throw new Error("cleanup-b");
  });

  await assert.rejects(
    handle.deactivate(),
    /Extension cleanup failed:\n- cleanup-a\n- cleanup-b/,
  );
});
