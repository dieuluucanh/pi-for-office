import assert from "node:assert/strict";
import { test } from "node:test";

import {
  ALWAYS_ALLOW_TOOLS_SETTING_KEY,
  MAX_ALWAYS_ALLOWED_TOOLS,
  createAlwaysAllowController,
  getStoredAlwaysAllowedTools,
  normalizeAlwaysAllowedTools,
  setStoredAlwaysAllowedTools,
  type AlwaysAllowStore,
} from "../src/execution/always-allow.ts";

function createMemoryStore(initial?: DynamicValue): AlwaysAllowStore & {
  data: Map<string, DynamicValue>;
} {
  const data = new Map<string, DynamicValue>();
  if (initial !== undefined) {
    data.set(ALWAYS_ALLOW_TOOLS_SETTING_KEY, initial);
  }

  return {
    data,
    get: (key: string) => Promise.resolve(data.get(key)),
    set: (key: string, value: DynamicValue) => {
      data.set(key, value);
      return Promise.resolve();
    },
  };
}

void test("normalizeAlwaysAllowedTools drops invalid entries and dedupes", () => {
  assert.deepEqual(normalizeAlwaysAllowedTools(null), []);
  assert.deepEqual(normalizeAlwaysAllowedTools("execute_office_js"), []);
  assert.deepEqual(normalizeAlwaysAllowedTools([]), []);
  assert.deepEqual(
    normalizeAlwaysAllowedTools([
      "execute_office_js",
      "",
      "execute_office_js",
      42,
      "write_cells",
      null,
    ]),
    ["execute_office_js", "write_cells"],
  );
});

void test("normalizeAlwaysAllowedTools caps the list length", () => {
  const input = Array.from({ length: MAX_ALWAYS_ALLOWED_TOOLS + 20 }, (_, i) => `tool_${i}`);
  const normalized = normalizeAlwaysAllowedTools(input);

  assert.equal(normalized.length, MAX_ALWAYS_ALLOWED_TOOLS);
});

void test("getStoredAlwaysAllowedTools returns an empty list when unset", async () => {
  const store = createMemoryStore();
  assert.deepEqual(await getStoredAlwaysAllowedTools(store), []);
  assert.deepEqual(await getStoredAlwaysAllowedTools(createMemoryStore("nope")), []);
});

void test("setStoredAlwaysAllowedTools persists a normalized list", async () => {
  const store = createMemoryStore();
  const saved = await setStoredAlwaysAllowedTools(store, [
    "execute_office_js",
    "execute_office_js",
    "",
  ]);

  assert.deepEqual(saved, ["execute_office_js"]);
  assert.deepEqual(
    await getStoredAlwaysAllowedTools(store),
    ["execute_office_js"],
  );
});

void test("controller add/remove/clear persists and reports tools", async () => {
  const store = createMemoryStore();
  const controller = await createAlwaysAllowController({ settings: store });

  assert.deepEqual(controller.getTools(), []);
  assert.equal(controller.hasTool("execute_office_js"), false);

  await controller.addTool("execute_office_js");
  assert.deepEqual(controller.getTools(), ["execute_office_js"]);
  assert.equal(controller.hasTool("execute_office_js"), true);
  assert.deepEqual(
    await getStoredAlwaysAllowedTools(store),
    ["execute_office_js"],
  );

  await controller.addTool("execute_office_js");
  assert.deepEqual(controller.getTools(), ["execute_office_js"]);

  await controller.addTool("write_cells");
  assert.deepEqual(controller.getTools(), ["execute_office_js", "write_cells"]);

  await controller.removeTool("execute_office_js");
  assert.deepEqual(controller.getTools(), ["write_cells"]);
  await controller.removeTool("execute_office_js");
  assert.deepEqual(controller.getTools(), ["write_cells"]);

  await controller.clear();
  assert.deepEqual(controller.getTools(), []);
  assert.deepEqual(await getStoredAlwaysAllowedTools(store), []);
});

void test("controller getTools returns a copy", async () => {
  const controller = await createAlwaysAllowController({
    settings: createMemoryStore(),
  });
  await controller.addTool("execute_office_js");

  const snapshot = controller.getTools();
  snapshot.push("write_cells");

  assert.deepEqual(controller.getTools(), ["execute_office_js"]);
});

void test("controller reverts the cache when persistence fails", async () => {
  const store: AlwaysAllowStore = {
    get: () => Promise.resolve(undefined),
    set: () => Promise.reject(new Error("write failed")),
  };
  const controller = await createAlwaysAllowController({ settings: store });

  await assert.rejects(
    () => controller.addTool("execute_office_js"),
    /write failed/u,
  );
  assert.deepEqual(controller.getTools(), []);
});

void test("controller keeps the previous list when a later write fails", async () => {
  let writeCount = 0;
  const store: AlwaysAllowStore = {
    get: () => Promise.resolve(undefined),
    set: () => {
      writeCount += 1;
      if (writeCount === 1) {
        return Promise.resolve();
      }
      return Promise.reject(new Error("write failed"));
    },
  };
  const controller = await createAlwaysAllowController({ settings: store });
  await controller.addTool("execute_office_js");

  await assert.rejects(
    () => controller.clear(),
    /write failed/u,
  );
  assert.deepEqual(controller.getTools(), ["execute_office_js"]);
});

void test("controller rejects grants beyond the cap", async () => {
  const tools = Array.from(
    { length: MAX_ALWAYS_ALLOWED_TOOLS },
    (_, i) => `tool_${i}`,
  );
  const store = createMemoryStore(tools);
  const controller = await createAlwaysAllowController({ settings: store });

  await assert.rejects(
    () => controller.addTool("execute_office_js"),
    /always-allow list is full/iu,
  );
  assert.equal(controller.hasTool("execute_office_js"), false);
  assert.equal(controller.getTools().length, MAX_ALWAYS_ALLOWED_TOOLS);
});
