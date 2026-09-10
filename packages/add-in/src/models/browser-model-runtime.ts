import {
  createModels,
  createProvider,
  lazyApi,
  type Api,
  type ApiKeyAuth,
  type Model,
  type ModelsRefreshOptions,
  type ModelsRefreshResult,
  type ModelsStore,
  type ModelsStoreEntry,
  type MutableModels,
  type Provider,
  type ProviderStreams,
  type RefreshModelsContext,
} from "@earendil-works/pi-ai";
import { builtinProviders } from "@earendil-works/pi-ai/providers/all";

import { originalFetch } from "../auth/cors-proxy.js";
import { normalizeProxyUrl } from "../auth/proxy-validation.js";
import type { CustomProvider } from "../storage/local/custom-providers-store.js";
import {
  ProviderCredentialsStore,
  type ProviderKeysStoreLike,
} from "../storage/local/provider-credentials-store.js";

const DEFAULT_DISCOVERED_CONTEXT_WINDOW = 256_000;
const DEFAULT_DISCOVERED_MAX_TOKENS = 4_096;
const MODEL_DISCOVERY_TIMEOUT_MS = 8_000;
const MAX_MODEL_DISCOVERY_RESPONSE_BYTES = 2 * 1024 * 1024;
const MAX_DISCOVERED_MODELS = 2_000;
const MAX_DISCOVERED_MODEL_ID_LENGTH = 256;

const openAiCompletionsStreams: ProviderStreams = lazyApi(
  () => import("@earendil-works/pi-ai/api/openai-completions"),
);
const openAiResponsesStreams: ProviderStreams = lazyApi(
  () => import("@earendil-works/pi-ai/api/openai-responses"),
);
const anthropicMessagesStreams: ProviderStreams = lazyApi(
  () => import("@earendil-works/pi-ai/api/anthropic-messages"),
);

export type BrowserProviderApi =
  | "openai-completions"
  | "openai-responses"
  | "anthropic-messages";

export interface BrowserProviderModelDefinition {
  id: string;
  name?: string;
  /** Per-model API override. When set, takes precedence over the registration-level api. */
  api?: BrowserProviderApi | undefined;
  reasoning?: boolean;
  input?: readonly ("text" | "image")[];
  contextWindow?: number;
  maxTokens?: number;
}

export interface BrowserProviderRegistration {
  id: string;
  name: string;
  api: BrowserProviderApi;
  baseUrl: string;
  models: readonly BrowserProviderModelDefinition[];
  /** Defaults to `${baseUrl}/models` for OpenAI-compatible providers. */
  modelsUrl?: string;
  resolveApiKey: () => Promise<string | undefined>;
  /** Keyless local providers can opt in while still satisfying Pi AI auth semantics. */
  allowKeyless?: boolean;
  /** Extra headers sent with every request (e.g. x-opencode-session for custom gateways). */
  headers?: Record<string, string>;
}

export interface CreateBrowserModelRuntimeOptions {
  providerKeys: ProviderKeysStoreLike;
  modelCatalogs: ModelsStore;
  getProxyUrl: () => Promise<string | undefined>;
  fetchFn?: typeof globalThis.fetch;
}

interface DynamicProviderCatalogSpec {
  api: BrowserProviderApi;
  baseUrl: string;
}

function normalizeNonEmpty(value: string, label: string): string {
  const normalized = value.trim();
  if (normalized.length === 0) {
    throw new Error(`${label} cannot be empty.`);
  }
  return normalized;
}

function normalizeHttpUrl(value: string, label: string): string {
  let parsed: URL;
  try {
    parsed = new URL(normalizeNonEmpty(value, label));
  } catch {
    throw new Error(`${label} must be a valid http:// or https:// URL.`);
  }

  if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
    throw new Error(`${label} must use http:// or https://.`);
  }

  parsed.hash = "";
  const normalized = parsed.toString();
  return normalized.endsWith("/") ? normalized.slice(0, -1) : normalized;
}

function streamsForApi(api: BrowserProviderApi): ProviderStreams {
  if (api === "openai-completions") return openAiCompletionsStreams;
  if (api === "openai-responses") return openAiResponsesStreams;
  return anthropicMessagesStreams;
}

function createStoredApiKeyAuth(args: {
  name: string;
  resolveApiKey: () => Promise<string | undefined>;
  allowKeyless: boolean;
}): ApiKeyAuth {
  return {
    name: args.name,
    async resolve() {
      const key = (await args.resolveApiKey())?.trim();
      if (!key && !args.allowKeyless) return undefined;

      return {
        auth: key ? { apiKey: key } : {},
        source: key ? "Browser credential store" : "Keyless provider",
      };
    },
  };
}

function createBrowserAdapterProvider(provider: Provider): Provider {
  const authName = provider.auth.oauth?.name ?? provider.name;
  return createProvider({
    id: provider.id,
    name: provider.name,
    ...(provider.baseUrl !== undefined ? { baseUrl: provider.baseUrl } : {}),
    ...(provider.headers !== undefined ? { headers: provider.headers } : {}),
    auth: {
      apiKey: {
        name: authName,
        resolve({ credential }) {
          if (credential?.key) {
            return Promise.resolve({
              auth: { apiKey: credential.key },
              source: "Browser credential store",
            });
          }
          return Promise.resolve(undefined);
        },
      },
    },
    models: provider.getModels(),
    api: {
      stream: (model, context, options) =>
        provider.stream(model, context, options),
      streamSimple: (model, context, options) =>
        provider.streamSimple(model, context, options),
    },
  });
}

function createModel(args: {
  providerId: string;
  api: BrowserProviderApi;
  baseUrl: string;
  definition: BrowserProviderModelDefinition;
}): Model<Api> {
  const id = normalizeNonEmpty(args.definition.id, "Model id");
  const contextWindow =
    args.definition.contextWindow ?? DEFAULT_DISCOVERED_CONTEXT_WINDOW;
  const maxTokens =
    args.definition.maxTokens ??
    Math.min(DEFAULT_DISCOVERED_MAX_TOKENS, contextWindow);

  if (!Number.isInteger(contextWindow) || contextWindow < 1_024) {
    throw new Error(
      `Model ${id} contextWindow must be an integer of at least 1024.`,
    );
  }
  if (!Number.isInteger(maxTokens) || maxTokens < 1) {
    throw new Error(`Model ${id} maxTokens must be a positive integer.`);
  }

  // Per-model api override takes precedence over registration-level default.
  const api = args.definition.api ?? args.api;

  return {
    id,
    name: args.definition.name?.trim() || id,
    api,
    provider: args.providerId,
    baseUrl: args.baseUrl,
    reasoning: args.definition.reasoning ?? false,
    input: args.definition.input ? [...args.definition.input] : ["text"],
    cost: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
    },
    contextWindow,
    maxTokens,
  };
}

function deriveModelsUrl(
  baseUrl: string,
  api: BrowserProviderApi,
): string | undefined {
  if (api !== "openai-completions" && api !== "openai-responses") {
    return undefined;
  }
  return `${baseUrl}/models`;
}

function sanitizeCatalogEntry(
  providerId: string,
  spec: DynamicProviderCatalogSpec,
  entry: ModelsStoreEntry,
): ModelsStoreEntry {
  const models: Model<Api>[] = [];
  for (const model of entry.models) {
    try {
      models.push(
        createModel({
          providerId,
          api: spec.api,
          baseUrl: spec.baseUrl,
          definition: {
            id: model.id,
            name: model.name,
            reasoning: model.reasoning,
            input: model.input,
            contextWindow: model.contextWindow,
            maxTokens: model.maxTokens,
          },
        }),
      );
    } catch {
      // Ignore malformed/stale entries rather than restoring unsafe metadata.
    }
  }

  return {
    models,
    ...(entry.checkedAt !== undefined ? { checkedAt: entry.checkedAt } : {}),
  };
}

/** Rebind cached discovery results to the currently registered transport. */
class BrowserModelCatalogsStore implements ModelsStore {
  private readonly persisted: ModelsStore;
  private readonly specs = new Map<string, DynamicProviderCatalogSpec>();

  constructor(persisted: ModelsStore) {
    this.persisted = persisted;
  }

  configure(providerId: string, spec: DynamicProviderCatalogSpec): void {
    this.specs.set(providerId, spec);
  }

  unconfigure(providerId: string): void {
    this.specs.delete(providerId);
  }

  async read(providerId: string): Promise<ModelsStoreEntry | undefined> {
    const entry = await this.persisted.read(providerId);
    const spec = this.specs.get(providerId);
    return entry && spec
      ? sanitizeCatalogEntry(providerId, spec, entry)
      : entry;
  }

  write(providerId: string, entry: ModelsStoreEntry): Promise<void> {
    const spec = this.specs.get(providerId);
    return this.persisted.write(
      providerId,
      spec ? sanitizeCatalogEntry(providerId, spec, entry) : entry,
    );
  }

  delete(providerId: string): Promise<void> {
    return this.persisted.delete(providerId);
  }
}

function isModelsResponse(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseModelIds(value: DynamicValue): string[] {
  if (!isModelsResponse(value) || !Array.isArray(value.data)) {
    throw new Error("Model discovery response must contain a data array.");
  }
  if (value.data.length > MAX_DISCOVERED_MODELS) {
    throw new Error(
      `Model discovery returned more than ${MAX_DISCOVERED_MODELS} entries.`,
    );
  }

  const ids: string[] = [];
  for (const entry of value.data) {
    if (!isModelsResponse(entry) || typeof entry.id !== "string") continue;
    const id = entry.id.trim();
    if (id.length === 0) continue;
    if (id.length > MAX_DISCOVERED_MODEL_ID_LENGTH) {
      throw new Error(
        `Discovered model id exceeds ${MAX_DISCOVERED_MODEL_ID_LENGTH} characters.`,
      );
    }
    ids.push(id);
  }

  return Array.from(new Set(ids)).sort((left, right) =>
    left.localeCompare(right),
  );
}

async function readLimitedDiscoveryJson(
  response: Response,
): Promise<DynamicValue> {
  const declaredLengthRaw = response.headers.get("content-length");
  const declaredLength =
    declaredLengthRaw === null ? null : Number.parseInt(declaredLengthRaw, 10);
  if (
    declaredLength !== null &&
    Number.isFinite(declaredLength) &&
    declaredLength > MAX_MODEL_DISCOVERY_RESPONSE_BYTES
  ) {
    throw new Error(
      `Model discovery response exceeds ${MAX_MODEL_DISCOVERY_RESPONSE_BYTES} bytes.`,
    );
  }

  if (!response.body) {
    const text = await response.text();
    if (
      new TextEncoder().encode(text).byteLength >
      MAX_MODEL_DISCOVERY_RESPONSE_BYTES
    ) {
      throw new Error(
        `Model discovery response exceeds ${MAX_MODEL_DISCOVERY_RESPONSE_BYTES} bytes.`,
      );
    }
    let parsed: DynamicValue;
    try {
      parsed = JSON.parse(text);
    } catch {
      throw new Error("Model discovery returned invalid JSON.");
    }
    return parsed;
  }

  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let totalBytes = 0;

  while (true) {
    const chunk = await reader.read();
    if (chunk.done) break;
    if (!chunk.value) continue;

    totalBytes += chunk.value.byteLength;
    if (totalBytes > MAX_MODEL_DISCOVERY_RESPONSE_BYTES) {
      await reader.cancel();
      throw new Error(
        `Model discovery response exceeds ${MAX_MODEL_DISCOVERY_RESPONSE_BYTES} bytes.`,
      );
    }
    chunks.push(chunk.value);
  }

  const body = new Uint8Array(totalBytes);
  let offset = 0;
  for (const chunk of chunks) {
    body.set(chunk, offset);
    offset += chunk.byteLength;
  }

  let parsed: DynamicValue;
  try {
    parsed = JSON.parse(new TextDecoder().decode(body));
  } catch {
    throw new Error("Model discovery returned invalid JSON.");
  }
  return parsed;
}

async function resolveDiscoveryRequestUrl(
  targetUrl: string,
  getProxyUrl: () => Promise<string | undefined>,
): Promise<string> {
  const proxyUrl = await getProxyUrl();
  if (!proxyUrl) return targetUrl;
  return `${normalizeProxyUrl(proxyUrl)}/?url=${encodeURIComponent(targetUrl)}`;
}

async function fetchWithDiscoveryTimeout(
  fetchFn: typeof globalThis.fetch,
  requestUrl: string,
  headers: Record<string, string>,
  signal?: AbortSignal,
  lifecycleSignal?: AbortSignal,
): Promise<Response> {
  const controller = new AbortController();
  const handleAbort = (): void => controller.abort();
  const signals: readonly (AbortSignal | undefined)[] = [
    signal,
    lifecycleSignal,
  ];
  for (const sourceSignal of signals) {
    if (sourceSignal?.aborted) controller.abort();
    sourceSignal?.addEventListener("abort", handleAbort, { once: true });
  }
  const timeoutId = setTimeout(
    () => controller.abort(),
    MODEL_DISCOVERY_TIMEOUT_MS,
  );

  try {
    return await fetchFn(requestUrl, {
      headers,
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timeoutId);
    for (const sourceSignal of signals) {
      sourceSignal?.removeEventListener("abort", handleAbort);
    }
  }
}

/** Live-catalog config for the OpenCode Zen/Go gateways. */
interface OpenCodeDiscoveryConfig {
  /** Live model catalog endpoint (OpenAI-style `{ data: [{ id }] }`). */
  modelsUrl: string;
  /** API root without `/v1` — baseUrl for anthropic-messages models. */
  rootWithoutV1: string;
  /** API root with `/v1` — baseUrl for the other streaming APIs. */
  rootWithV1: string;
}

function opencodeDiscoveryConfig(
  providerId: string,
): OpenCodeDiscoveryConfig | undefined {
  if (providerId === "opencode") {
    return {
      modelsUrl: "https://opencode.ai/zen/v1/models",
      rootWithoutV1: "https://opencode.ai/zen",
      rootWithV1: "https://opencode.ai/zen/v1",
    };
  }
  if (providerId === "opencode-go") {
    return {
      modelsUrl: "https://opencode.ai/zen/go/v1/models",
      rootWithoutV1: "https://opencode.ai/zen/go",
      rootWithV1: "https://opencode.ai/zen/go/v1",
    };
  }
  return undefined;
}

/**
 * The `/models` endpoint returns bare IDs without API metadata, so infer the
 * streaming API from the documented per-family endpoint table:
 * "gpt-", "grok-", "muse-" prefixes → Responses, "claude-", "qwen" prefixes →
 * Messages, "gemini-" (Zen) → Generative AI, everything else → Completions.
 */
function inferOpenCodeModelApi(modelId: string, providerId: string): Api {
  const id = modelId.toLowerCase();
  if (
    id.startsWith("gpt-") ||
    id.startsWith("grok-") ||
    id.startsWith("muse-")
  ) {
    return "openai-responses";
  }
  if (id.startsWith("claude-") || id.startsWith("qwen")) {
    return "anthropic-messages";
  }
  if (id.startsWith("gemini-") && providerId === "opencode") {
    return "google-generative-ai";
  }
  return "openai-completions";
}

/**
 * Prefix tables mirroring the static pi-ai OpenCode catalogs, most-specific
 * first. `freeContextWindow` overrides the value for `-free` id variants.
 */
const OPENCODE_CONTEXT_WINDOW_RULES: readonly {
  readonly prefixes: readonly string[];
  readonly contextWindow: number;
  readonly freeContextWindow?: number;
}[] = [
  // ~1M-token families.
  { prefixes: ["gemini-"], contextWindow: 1_048_576 },
  { prefixes: ["gpt-5.5", "gpt-5.6", "gpt-5.4-pro"], contextWindow: 1_050_000 },
  { prefixes: ["gpt-5.4"], contextWindow: 272_000 },
  { prefixes: ["gpt-5"], contextWindow: 400_000 },
  { prefixes: ["grok-"], contextWindow: 500_000 },
  {
    prefixes: [
      "claude-opus-4-6",
      "claude-opus-4-7",
      "claude-opus-4-8",
      "claude-opus-5",
      "claude-sonnet-4-6",
      "claude-sonnet-5",
      "claude-fable-",
    ],
    contextWindow: 1_000_000,
  },
  { prefixes: ["deepseek-v4-pro"], contextWindow: 1_000_000 },
  {
    prefixes: ["deepseek-v4-flash"],
    contextWindow: 1_000_000,
    freeContextWindow: 200_000,
  },
  { prefixes: ["glm-5.2"], contextWindow: 1_000_000 },
  { prefixes: ["kimi-k3"], contextWindow: 1_000_000 },
  { prefixes: ["nemotron-3-ultra"], contextWindow: 1_000_000 },
  { prefixes: ["mimo-"], contextWindow: 1_000_000, freeContextWindow: 200_000 },
  { prefixes: ["qwen3.7"], contextWindow: 1_000_000 },
  { prefixes: ["minimax-m3"], contextWindow: 512_000 },
  // 200k-262k families (free variants sit at the low end).
  {
    prefixes: [
      "claude-",
      "qwen",
      "kimi-",
      "glm-",
      "minimax-",
      "deepseek-",
      "big-pickle",
      "laguna-",
      "ling-",
      "north-mini",
      "hy",
      "muse-",
      "longcat",
      "omen-",
    ],
    contextWindow: 262_144,
    freeContextWindow: 200_000,
  },
];

/**
 * The `/models` endpoint returns bare IDs without context metadata, so infer a
 * conservative context window from the model-family prefix, using the static
 * pi-ai catalogs as the source of truth. Unknown families fall back to
 * DEFAULT_DISCOVERED_CONTEXT_WINDOW.
 */
export function inferOpenCodeContextWindow(modelId: string): number {
  const id = modelId.toLowerCase();
  const isFree = id.endsWith("-free");
  for (const rule of OPENCODE_CONTEXT_WINDOW_RULES) {
    if (rule.prefixes.some((prefix) => id.startsWith(prefix))) {
      return isFree
        ? (rule.freeContextWindow ?? rule.contextWindow)
        : rule.contextWindow;
    }
  }
  return DEFAULT_DISCOVERED_CONTEXT_WINDOW;
}

function buildDiscoveredOpenCodeModel(
  providerId: string,
  modelId: string,
  config: OpenCodeDiscoveryConfig,
): Model<Api> {
  const api = inferOpenCodeModelApi(modelId, providerId);
  const contextWindow = inferOpenCodeContextWindow(modelId);
  return {
    id: modelId,
    name: modelId,
    api,
    provider: providerId,
    baseUrl:
      api === "anthropic-messages" ? config.rootWithoutV1 : config.rootWithV1,
    reasoning: true,
    input: ["text"],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow,
    maxTokens: Math.min(DEFAULT_DISCOVERED_MAX_TOKENS, contextWindow),
  };
}

/**
 * Wrap a builtin OpenCode provider with live catalog discovery.
 *
 * Mirrors pi-ai's `createProvider` fetchModels wiring (restore stored catalog →
 * fetch when network is allowed → persist), while keeping the static catalog
 * authoritative for known IDs so verified per-model metadata (api, cost,
 * contextWindow, compat) is preserved. Only newly listed IDs are added.
 */
function wrapWithOpenCodeDiscovery(
  provider: Provider,
  config: OpenCodeDiscoveryConfig,
  getProxyUrl: () => Promise<string | undefined>,
  fetchFn: typeof globalThis.fetch,
): Provider {
  let dynamicModels: readonly Model<Api>[] = [];
  let inflightRefresh: Promise<void> | undefined;

  const mergeModels = (): readonly Model<Api>[] => {
    const staticModels = provider.getModels();
    const staticIds = new Set(staticModels.map((model) => model.id));
    const extras = dynamicModels.filter((model) => !staticIds.has(model.id));
    return [...staticModels, ...extras];
  };

  return {
    ...provider,
    getModels: mergeModels,
    refreshModels: (context: RefreshModelsContext): Promise<void> => {
      inflightRefresh ??= (async () => {
        try {
          const stored = await context.store.read();
          if (stored) {
            dynamicModels = stored.models.filter(
              (model) => model.provider === provider.id,
            );
          }
          if (!context.allowNetwork || context.signal?.aborted) return;

          // OpenCode gateway responses omit CORS headers, so route discovery
          // through the office proxy when one is configured.
          const requestUrl = await resolveDiscoveryRequestUrl(
            config.modelsUrl,
            getProxyUrl,
          );
          const headers: Record<string, string> = {
            Accept: "application/json",
          };
          if (
            context.credential?.type === "api_key" &&
            context.credential.key
          ) {
            headers.Authorization = `Bearer ${context.credential.key}`;
          }
          const response = await fetchWithDiscoveryTimeout(
            fetchFn,
            requestUrl,
            headers,
            context.signal,
          );
          if (context.signal?.aborted) return;
          if (!response.ok) {
            throw new Error(
              `Model discovery failed with HTTP ${response.status}.`,
            );
          }
          const ids = parseModelIds(await readLimitedDiscoveryJson(response));
          const staticIds = new Set(provider.getModels().map((m) => m.id));
          dynamicModels = ids
            .filter((id) => !staticIds.has(id))
            .map((id) => buildDiscoveredOpenCodeModel(provider.id, id, config));
          await context.store.write({
            models: dynamicModels,
            checkedAt: Date.now(),
          });
        } finally {
          inflightRefresh = undefined;
        }
      })();
      return inflightRefresh;
    },
  };
}

function createRegisteredProvider(
  registration: BrowserProviderRegistration,
  getProxyUrl: () => Promise<string | undefined>,
  fetchFn: typeof globalThis.fetch,
  lifecycleSignal: AbortSignal,
): Provider {
  const id = normalizeNonEmpty(registration.id, "Provider id");
  const name = normalizeNonEmpty(registration.name, "Provider name");
  const baseUrl = normalizeHttpUrl(registration.baseUrl, "Provider baseUrl");
  const baselineModels = registration.models.map((definition) => {
    const model = createModel({
      providerId: id,
      api: registration.api,
      baseUrl,
      definition,
    });
    // Merge provider-level headers onto each model so openai-completions' createClient()
    // picks them up via model.headers (Provider.headers is NOT auto-propagated to API-key auth).
    if (registration.headers) {
      return {
        ...model,
        headers: { ...model.headers, ...registration.headers },
      };
    }
    return model;
  });
  // Determine the provider-level API type(s).
  // When all models share the same api, use a single stream (optimized path).
  // When models have mixed apis, build a map for pi-ai's multi-API dispatch.
  const uniqueApis = new Set<BrowserProviderApi>(
    baselineModels.map((m) => m.api as BrowserProviderApi),
  );
  const hasMixedApis = uniqueApis.size > 1;
  const providerApiValue = hasMixedApis ? undefined : registration.api;

  const modelsUrlRaw =
    registration.modelsUrl ?? deriveModelsUrl(baseUrl, registration.api);
  const modelsUrl = modelsUrlRaw
    ? normalizeHttpUrl(modelsUrlRaw, "Provider modelsUrl")
    : undefined;

  return createProvider({
    id,
    name,
    baseUrl,
    ...(registration.headers !== undefined
      ? { headers: registration.headers }
      : {}),
    auth: {
      apiKey: createStoredApiKeyAuth({
        name: `${name} credential`,
        resolveApiKey: registration.resolveApiKey,
        allowKeyless: registration.allowKeyless === true,
      }),
    },
    models: baselineModels as readonly Model<BrowserProviderApi>[],
    ...(modelsUrl !== undefined
      ? {
          fetchModels: async ({ credential, signal }) => {
            const requestUrl = await resolveDiscoveryRequestUrl(
              modelsUrl,
              getProxyUrl,
            );
            const headers: Record<string, string> = {
              Accept: "application/json",
            };
            if (credential?.type === "api_key" && credential.key) {
              headers.Authorization = `Bearer ${credential.key}`;
            }

            const response = await fetchWithDiscoveryTimeout(
              fetchFn,
              requestUrl,
              headers,
              signal,
              lifecycleSignal,
            );
            if (!response.ok) {
              throw new Error(
                `Model discovery failed with HTTP ${response.status}.`,
              );
            }

            const ids = parseModelIds(await readLimitedDiscoveryJson(response));
            const template = registration.models[0];
            return ids.map((modelId) =>
              createModel({
                providerId: id,
                api: registration.api,
                baseUrl,
                definition: {
                  id: modelId,
                  ...(template?.reasoning !== undefined
                    ? { reasoning: template.reasoning }
                    : {}),
                  ...(template?.input !== undefined
                    ? { input: template.input }
                    : {}),
                  ...(template?.contextWindow !== undefined
                    ? { contextWindow: template.contextWindow }
                    : {}),
                  ...(template?.maxTokens !== undefined
                    ? { maxTokens: template.maxTokens }
                    : {}),
                },
              }),
            ) as readonly Model<BrowserProviderApi>[];
          },
        }
      : {}),
    api: hasMixedApis
      ? (Object.fromEntries(
          Array.from(uniqueApis).map((a) => [a, streamsForApi(a)]),
        ) as Partial<Record<BrowserProviderApi, ProviderStreams>>)
      : streamsForApi(providerApiValue!),
  });
}

// Stable session ID sent to custom gateways (required by OpenCode Go, harmless for others).
const GATEWAY_SESSION_ID = globalThis.crypto.randomUUID();

function customProviderRegistrations(
  provider: CustomProvider,
): BrowserProviderRegistration[] {
  const storedModels = provider.models ?? [];
  const modelsByProvider = new Map<string, BrowserProviderModelDefinition[]>();

  // Per-model API type map from probing (stored on the CustomProvider)
  const modelApiMap = provider.modelApiMap as
    | Record<string, BrowserProviderApi>
    | undefined;

  for (const model of storedModels) {
    const definitions = modelsByProvider.get(model.provider) ?? [];

    // Per-model API override: modelApiMap > stored model.api > undefined (falls back to provider type)
    const perModelApi =
      modelApiMap?.[model.id] ?? (model.api as BrowserProviderApi | undefined);

    definitions.push({
      id: model.id,
      name: model.name,
      api: perModelApi,
      reasoning: model.reasoning,
      input: model.input,
      contextWindow: model.contextWindow,
      maxTokens: model.maxTokens,
    });
    modelsByProvider.set(model.provider, definitions);
  }

  if (modelsByProvider.size === 0) return [];
  if (
    provider.type !== "openai-completions" &&
    provider.type !== "openai-responses" &&
    provider.type !== "anthropic-messages"
  ) {
    return [];
  }

  const providerApi: BrowserProviderApi = provider.type;
  return Array.from(modelsByProvider.entries()).map(([providerId, models]) => ({
    id: providerId,
    name: provider.name,
    api: providerApi,
    baseUrl: provider.baseUrl,
    models,
    resolveApiKey: () => Promise.resolve(provider.apiKey),
    allowKeyless: !provider.apiKey,
    headers: { "x-opencode-session": GATEWAY_SESSION_ID },
  }));
}

/** Browser-native Pi AI provider runtime with IndexedDB credentials/catalogues. */
export class BrowserModelRuntime {
  readonly models: MutableModels;

  private readonly modelCatalogs: BrowserModelCatalogsStore;
  private readonly getProxyUrl: () => Promise<string | undefined>;
  private readonly fetchFn: typeof globalThis.fetch;
  private readonly builtinProviderIds = new Set<string>();
  private readonly customProviderIds = new Set<string>();
  private readonly extensionProviderOwners = new Map<string, string>();
  private readonly dynamicProviderAbortControllers = new Map<
    string,
    AbortController
  >();

  constructor(options: CreateBrowserModelRuntimeOptions) {
    this.modelCatalogs = new BrowserModelCatalogsStore(options.modelCatalogs);
    this.getProxyUrl = options.getProxyUrl;
    this.fetchFn = options.fetchFn ?? originalFetch ?? globalThis.fetch;
    this.models = createModels({
      credentials: new ProviderCredentialsStore(options.providerKeys),
      modelsStore: this.modelCatalogs,
      authContext: {
        env: () => Promise.resolve(undefined),
        fileExists: () => Promise.resolve(false),
      },
    });

    for (const provider of builtinProviders()) {
      let browserProvider = provider.auth.apiKey
        ? provider
        : createBrowserAdapterProvider(provider);

      // OpenCode catalogs move fast (new/free models rotate weekly); augment
      // the static pi-ai catalog with live discovery from the /models endpoint.
      const discovery = opencodeDiscoveryConfig(browserProvider.id);
      if (discovery) {
        browserProvider = wrapWithOpenCodeDiscovery(
          browserProvider,
          discovery,
          this.getProxyUrl,
          this.fetchFn,
        );
      }

      this.models.setProvider(browserProvider);
      this.builtinProviderIds.add(browserProvider.id);
    }
  }

  private createDynamicProvider(
    registration: BrowserProviderRegistration,
  ): Provider {
    const controller = new AbortController();
    const provider = createRegisteredProvider(
      registration,
      this.getProxyUrl,
      this.fetchFn,
      controller.signal,
    );

    this.dynamicProviderAbortControllers.get(provider.id)?.abort();
    this.dynamicProviderAbortControllers.set(provider.id, controller);
    return provider;
  }

  private stopDynamicProvider(providerId: string): void {
    this.dynamicProviderAbortControllers.get(providerId)?.abort();
    this.dynamicProviderAbortControllers.delete(providerId);
  }

  async syncCustomProviders(
    customProviders: readonly CustomProvider[],
  ): Promise<void> {
    const nextIds = new Set<string>();

    for (const customProvider of customProviders) {
      for (const registration of customProviderRegistrations(customProvider)) {
        if (this.builtinProviderIds.has(registration.id)) {
          throw new Error(
            `Custom provider id conflicts with built-in provider: ${registration.id}`,
          );
        }
        if (this.extensionProviderOwners.has(registration.id)) {
          throw new Error(
            `Custom provider id conflicts with extension provider: ${registration.id}`,
          );
        }

        const provider = this.createDynamicProvider(registration);
        this.modelCatalogs.configure(provider.id, {
          api: registration.api,
          baseUrl:
            provider.baseUrl ??
            normalizeHttpUrl(registration.baseUrl, "Provider baseUrl"),
        });
        this.models.setProvider(provider);
        nextIds.add(registration.id);
      }
    }

    for (const previousId of this.customProviderIds) {
      if (nextIds.has(previousId)) continue;
      this.stopDynamicProvider(previousId);
      this.models.deleteProvider(previousId);
      try {
        await this.modelCatalogs.delete(previousId);
      } finally {
        this.modelCatalogs.unconfigure(previousId);
      }
    }

    this.customProviderIds.clear();
    for (const providerId of nextIds) this.customProviderIds.add(providerId);
  }

  registerExtensionProvider(
    ownerId: string,
    registration: BrowserProviderRegistration,
  ): void {
    const providerId = normalizeNonEmpty(registration.id, "Provider id");
    if (
      this.builtinProviderIds.has(providerId) ||
      this.customProviderIds.has(providerId)
    ) {
      throw new Error(`Provider id is reserved: ${providerId}`);
    }

    const existingOwner = this.extensionProviderOwners.get(providerId);
    if (existingOwner && existingOwner !== ownerId) {
      throw new Error(
        `Provider id is already registered by another extension: ${providerId}`,
      );
    }

    const provider = this.createDynamicProvider(registration);
    this.modelCatalogs.configure(provider.id, {
      api: registration.api,
      baseUrl:
        provider.baseUrl ??
        normalizeHttpUrl(registration.baseUrl, "Provider baseUrl"),
    });
    this.models.setProvider(provider);
    this.extensionProviderOwners.set(providerId, ownerId);
  }

  isExtensionProvider(providerId: string): boolean {
    return this.extensionProviderOwners.has(providerId);
  }

  shouldProxyProvider(providerId: string): boolean {
    return (
      this.customProviderIds.has(providerId) ||
      this.extensionProviderOwners.has(providerId)
    );
  }

  async unregisterExtensionProvider(
    ownerId: string,
    providerId: string,
  ): Promise<void> {
    if (this.extensionProviderOwners.get(providerId) !== ownerId) {
      throw new Error(`Provider is not owned by this extension: ${providerId}`);
    }

    this.extensionProviderOwners.delete(providerId);
    this.stopDynamicProvider(providerId);
    this.models.deleteProvider(providerId);
    try {
      await this.modelCatalogs.delete(providerId);
    } finally {
      this.modelCatalogs.unconfigure(providerId);
    }
  }

  async unregisterExtensionProviders(ownerId: string): Promise<void> {
    const ownedIds = Array.from(this.extensionProviderOwners.entries())
      .filter(([, owner]) => owner === ownerId)
      .map(([providerId]) => providerId);

    for (const providerId of ownedIds) {
      await this.unregisterExtensionProvider(ownerId, providerId);
    }
  }

  refresh(options?: ModelsRefreshOptions): Promise<ModelsRefreshResult> {
    return this.models.refresh(options);
  }
}
