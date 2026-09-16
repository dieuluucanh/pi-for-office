import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import { test } from "node:test";

import type { Api, Model } from "@earendil-works/pi-ai";
import { builtinModels } from "@earendil-works/pi-ai/providers/all";

import {
  LAST_USED_MODEL_SETTING_KEY,
  getStoredLastUsedModel,
  isSameStoredLastUsedModel,
  normalizeStoredLastUsedModel,
  pickStartupModel,
  setStoredLastUsedModel,
  type LastUsedModelStore,
} from "../src/models/last-used-model.ts";

class MemoryStore {
  private readonly values = new Map<string, DynamicValue>();

  get(key: string): Promise<DynamicValue> {
    return Promise.resolve(
      this.values.has(key) ? (this.values.get(key) ?? null) : null,
    );
  }

  set(key: string, value: DynamicValue): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }
}

const modelsRuntime = builtinModels();

function firstModelId(provider: string): string {
  const first = modelsRuntime.getModels(provider)[0];
  if (!first) throw new Error(`No builtin models for provider: ${provider}`);
  return first.id;
}

function getModel(provider: string, id: string): Model<Api> {
  const model = modelsRuntime.getModel(provider, id);
  if (!model) throw new Error(`Missing builtin test model: ${provider}/${id}`);
  return model;
}

// Ids are read from the installed registry instead of being hardcoded, so an
// upstream pi-ai bump cannot silently invalidate these expectations.
const openaiId = firstModelId("openai");
const anthropicId = firstModelId("anthropic");
const anthropicModel = getModel("anthropic", anthropicId);

void test("normalizes a valid remembered-model record and trims whitespace", () => {
  assert.deepEqual(
    normalizeStoredLastUsedModel({ provider: " openai ", id: " gpt-5.6-sol " }),
    { provider: "openai", id: "gpt-5.6-sol" },
  );
});

void test("rejects malformed remembered-model records", () => {
  const invalid: DynamicValue[] = [
    undefined,
    null,
    "",
    42,
    true,
    [],
    ["openai", "gpt-5.6-sol"],
    {},
    { provider: "openai" },
    { id: "gpt-5.6-sol" },
    { provider: "", id: "gpt-5.6-sol" },
    { provider: "   ", id: "gpt-5.6-sol" },
    { provider: "openai", id: "" },
    { provider: "openai", id: "   " },
    { provider: 1, id: "gpt-5.6-sol" },
    { provider: "openai", id: null },
  ];

  for (const value of invalid) {
    assert.equal(
      normalizeStoredLastUsedModel(value),
      null,
      `expected null for ${JSON.stringify(value)}`,
    );
  }
});

void test("stores and reads the remembered model identity", async () => {
  const store = new MemoryStore();
  assert.equal(await getStoredLastUsedModel(store), null);

  const stored = await setStoredLastUsedModel(store, {
    provider: "openai",
    id: "gpt-5.6-sol",
  });
  assert.deepEqual(stored, { provider: "openai", id: "gpt-5.6-sol" });

  assert.deepEqual(await getStoredLastUsedModel(store), {
    provider: "openai",
    id: "gpt-5.6-sol",
  });
  assert.deepEqual(await store.get(LAST_USED_MODEL_SETTING_KEY), {
    provider: "openai",
    id: "gpt-5.6-sol",
  });
});

void test("a malformed stored value reads as not-remembered", async () => {
  const store = new MemoryStore();
  await store.set(LAST_USED_MODEL_SETTING_KEY, { provider: "openai" });

  assert.equal(await getStoredLastUsedModel(store), null);
});

void test("storage read failures degrade to not-remembered", async () => {
  const failingStore: LastUsedModelStore = {
    get: () => Promise.reject(new Error("storage unavailable")),
    set: () => Promise.reject(new Error("storage unavailable")),
  };

  assert.equal(await getStoredLastUsedModel(failingStore), null);
});

void test("isSameStoredLastUsedModel compares provider and id", () => {
  const stored = { provider: "openai", id: "gpt-5.6-sol" };

  assert.equal(isSameStoredLastUsedModel(null, stored), false);
  assert.equal(isSameStoredLastUsedModel(stored, stored), true);
  assert.equal(
    isSameStoredLastUsedModel(stored, {
      provider: "openai",
      id: "gpt-5.6-terra",
    }),
    false,
  );
  assert.equal(
    isSameStoredLastUsedModel(stored, {
      provider: "openai-codex",
      id: "gpt-5.6-sol",
    }),
    false,
  );
});

void test("pickStartupModel prefers the remembered model when its provider is usable", () => {
  const picked = pickStartupModel({
    modelsRuntime,
    availableProviders: ["openai", "anthropic"],
    stored: { provider: "openai", id: openaiId },
    defaultModel: anthropicModel,
  });

  assert.equal(picked.provider, "openai");
  assert.equal(picked.id, openaiId);
  // Re-resolved from the registry (not the stored stub) so upstream metadata
  // updates such as contextWindow/maxTokens are picked up automatically.
  assert.equal(picked, modelsRuntime.getModel("openai", openaiId));
});

void test("pickStartupModel falls back to the computed default when nothing is remembered", () => {
  const picked = pickStartupModel({
    modelsRuntime,
    availableProviders: ["openai", "anthropic"],
    stored: null,
    defaultModel: anthropicModel,
  });

  assert.equal(picked, anthropicModel);
});

void test("pickStartupModel falls back when the remembered model left the registry", () => {
  const retired: Array<{ provider: string; id: string }> = [
    { provider: "openai", id: "retired-model-id" },
    { provider: "retired-provider", id: openaiId },
  ];

  for (const stored of retired) {
    const picked = pickStartupModel({
      modelsRuntime,
      availableProviders: ["openai", "anthropic"],
      stored,
      defaultModel: anthropicModel,
    });

    assert.equal(
      picked,
      anthropicModel,
      `expected fallback for ${stored.provider}/${stored.id}`,
    );
  }
});

void test("pickStartupModel falls back when the remembered provider is not usable", () => {
  // Mirrors upstream pi-mono's hasConfiguredAuth() gate in findInitialModel().
  const picked = pickStartupModel({
    modelsRuntime,
    availableProviders: ["anthropic"],
    stored: { provider: "openai", id: openaiId },
    defaultModel: anthropicModel,
  });

  assert.equal(picked, anthropicModel);
});

void test("pickStartupModel keeps the remembered model while providers are unresolved", () => {
  // An empty provider list means "startup credential restore has not finished",
  // not "no provider is configured".
  const picked = pickStartupModel({
    modelsRuntime,
    availableProviders: [],
    stored: { provider: "openai", id: openaiId },
    defaultModel: anthropicModel,
  });

  assert.equal(picked.provider, "openai");
  assert.equal(picked.id, openaiId);
});

function readInitSource(): string {
  return readFileSync(
    path.resolve(process.cwd(), "src/taskpane/init.ts"),
    "utf8",
  );
}

void test("init.ts resolves new runtime models through the remembered model", () => {
  // Wiring pin: runtime creation must consult the app-wide remembered model
  // instead of falling straight back to the computed default.
  const initSource = readInitSource();

  assert.notEqual(
    initSource.indexOf("const initialModel = resolveInitialRuntimeModel();"),
    -1,
    "expected createRuntime to use resolveInitialRuntimeModel()",
  );
  assert.notEqual(
    initSource.indexOf("stored: lastUsedModel,"),
    -1,
    "expected pickStartupModel to receive the remembered model",
  );
  assert.equal(
    initSource.indexOf("getActiveRuntime()?.agent.state.model ?? defaultModel"),
    -1,
    "expected the raw defaultModel fallback to be replaced",
  );
});

void test("init.ts remembers explicit selections and persists them immediately", () => {
  const initSource = readInitSource();

  // Both in-place branches of applyModelSelection (same-identity metadata
  // refresh and the normal switch) persist the selection right away so a
  // session never comes back on a stale model.
  assert.equal(
    (
      initSource.match(
        /await rememberModelSelection\(nextModel, runtime\);/gu,
      ) ?? []
    ).length,
    2,
    "expected both in-place model-selection branches to persist",
  );
  // The fork branch clones the session (already saved) and only records the
  // app-wide choice.
  assert.notEqual(
    initSource.indexOf("await rememberModelSelection(nextModel, null);"),
    -1,
    "expected the fork branch to record the app-wide selection",
  );
});

void test("init.ts writes the remembered model from exactly one helper", () => {
  // Session restore, tab switches and provider reconciliation must never
  // overwrite the user's choice with a fallback default.
  const initSource = readInitSource();
  const writes = initSource.match(/await rememberLastUsedModel\(/gu) ?? [];

  assert.equal(
    writes.length,
    1,
    "expected exactly one remembered-model writer",
  );
});
