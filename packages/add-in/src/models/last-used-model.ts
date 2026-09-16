/**
 * Remembers the model the user last selected, app-wide.
 *
 * This is the browser-storage equivalent of upstream pi-mono's
 * `defaultProvider` / `defaultModel` settings keys
 * (`~/.pi/agent/settings.json`). The intentional divergence is the trigger:
 * upstream persists only when the user presses Ctrl+S ("save as default") in
 * its model selector, while Pi for Office persists on every explicit model
 * selection. See docs/upstream-divergences.md §9.
 *
 * Only model identity (`provider` + `id`) is stored. The model is re-resolved
 * against the live registry whenever it is used, so upstream metadata updates
 * (context window, max tokens) are picked up automatically — the same
 * rationale as `refreshPersistedModel` in `src/taskpane/sessions.ts`.
 */

import type { Api, Model, Models } from "@earendil-works/pi-ai";

export const LAST_USED_MODEL_SETTING_KEY = "model.lastUsed.v1";

export interface StoredLastUsedModel {
 provider: string;
 id: string;
}

/**
 * Minimal settings-store contract. Structurally satisfied by
 * `SettingsStore`; kept narrow so this module stays free of storage wiring.
 */
export interface LastUsedModelStore {
 get: (key: string) => Promise<DynamicValue>;
 set: (key: string, value: DynamicValue) => Promise<void>;
}

function isRememberedModelPayload(value: DynamicValue): value is DynamicObject {
 return typeof value === "object" && value !== null && !Array.isArray(value);
}

function normalizeNonEmptyString(value: DynamicValue): string | null {
 if (typeof value !== "string") return null;

 const trimmed = value.trim();
 return trimmed.length > 0 ? trimmed : null;
}

/**
 * Parse a persisted value at the storage boundary. Malformed, partial or
 * empty records normalize to `null` so callers fall back to the computed
 * default model instead of failing startup.
 */
export function normalizeStoredLastUsedModel(
 value: DynamicValue,
): StoredLastUsedModel | null {
 if (!isRememberedModelPayload(value)) {
  return null;
 }

 const provider = normalizeNonEmptyString(value.provider);
 const id = normalizeNonEmptyString(value.id);
 if (!provider || !id) {
  return null;
 }

 return { provider, id };
}

/** Read the remembered model. Storage failures degrade to "not remembered". */
export async function getStoredLastUsedModel(
 store: LastUsedModelStore,
): Promise<StoredLastUsedModel | null> {
 try {
  const raw = await store.get(LAST_USED_MODEL_SETTING_KEY);
  return normalizeStoredLastUsedModel(raw);
 } catch {
  return null;
 }
}

/** Persist the remembered model. Returns the normalized stored identity. */
export async function setStoredLastUsedModel(
 store: LastUsedModelStore,
 model: { provider: string; id: string },
): Promise<StoredLastUsedModel> {
 const stored: StoredLastUsedModel = { provider: model.provider, id: model.id };
 await store.set(LAST_USED_MODEL_SETTING_KEY, stored);
 return stored;
}

export function isSameStoredLastUsedModel(
 left: StoredLastUsedModel | null,
 right: { provider: string; id: string },
): boolean {
 if (!left) return false;
 return left.provider === right.provider && left.id === right.id;
}

/**
 * Resolve the model a freshly created runtime should start on.
 *
 * Order: remembered model → caller-supplied computed default. The remembered
 * model wins only while it still exists in the registry and its provider is
 * usable, mirroring upstream's `hasConfiguredAuth()` gate in
 * `model-resolver.findInitialModel()`. An empty `availableProviders` list
 * means "not resolved yet" (startup credential restore), so the remembered
 * model is kept rather than discarded.
 */
export function pickStartupModel(opts: {
 modelsRuntime: Models;
 availableProviders: readonly string[];
 stored: StoredLastUsedModel | null;
 defaultModel: Model<Api>;
}): Model<Api> {
 const { modelsRuntime, availableProviders, stored, defaultModel } = opts;

 if (!stored) {
  return defaultModel;
 }

 const model = modelsRuntime.getModel(stored.provider, stored.id);
 if (!model) {
  return defaultModel;
 }

 if (
  availableProviders.length > 0 &&
  !availableProviders.includes(model.provider)
 ) {
  return defaultModel;
 }

 return model;
}
