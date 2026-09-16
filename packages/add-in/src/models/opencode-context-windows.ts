/**
 * Context-window inference for OpenCode Zen/Go models discovered at runtime.
 *
 * The gateway `/models` endpoints return bare ids without metadata, so the
 * window for an id missing from the bundled pi-ai catalogs is inferred from its
 * family prefix. Two tables feed the lookup:
 *
 * - `OPENCODE_BASE_CONTEXT_WINDOW_RULES` mirrors the conservative family
 *   defaults of the static pi-ai OpenCode catalogs (Zen semantics). Static
 *   catalog entries always win in the runtime's `mergeModels()`, so this table
 *   only decides newly listed ids.
 * - `OPENCODE_GO_CONTEXT_WINDOW_RULES` mirrors models.dev for OpenCode Go
 *   (`providers/opencode-go` page and the `api.json` provider key, verified
 *   2026-09-16) where Go differs from the base table.
 *
 * Lookup rules:
 * - the longest matching prefix wins across both tables, so a coarse provider
 *   rule (e.g. Go `glm-5`) never shadows a finer base rule (`glm-5.2`);
 * - on an equal-length match the provider table wins;
 * - a provider rule replaces the base rule entirely, so it must repeat
 *   `freeContextWindow` itself when the family has a cheaper `-free` tier.
 */

export interface OpenCodeContextWindowRule {
  /** Family prefixes; the longest matching prefix decides the window. */
  readonly prefixes: readonly string[];
  /** Context window for non-`-free` ids of this family. */
  readonly contextWindow: number;
  /** Window for `-free` ids of this family; defaults to `contextWindow`. */
  readonly freeContextWindow?: number;
}

/** Window assumed for ids whose family matches no rule. */
export const OPENCODE_FALLBACK_CONTEXT_WINDOW = 256_000;

/** Family defaults shared by both gateways. */
export const OPENCODE_BASE_CONTEXT_WINDOW_RULES: readonly OpenCodeContextWindowRule[] =
  [
    // ~1M-token families.
    { prefixes: ["gemini-"], contextWindow: 1_048_576 },
    {
      prefixes: ["gpt-5.5", "gpt-5.6", "gpt-5.4-pro"],
      contextWindow: 1_050_000,
    },
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
    { prefixes: ["glm-5.2", "glm-5.3"], contextWindow: 1_000_000 },
    { prefixes: ["kimi-k3"], contextWindow: 1_048_576 },
    { prefixes: ["nemotron-3-ultra"], contextWindow: 1_000_000 },
    { prefixes: ["mimo-v2-pro", "mimo-v2.5-pro"], contextWindow: 1_048_576 },
    {
      prefixes: ["mimo-"],
      contextWindow: 1_000_000,
      freeContextWindow: 200_000,
    },
    { prefixes: ["muse-spark"], contextWindow: 1_048_576 },
    { prefixes: ["longcat-2"], contextWindow: 1_000_000 },
    { prefixes: ["qwen3.7"], contextWindow: 1_000_000 },
    // Mid tiers.
    { prefixes: ["gpt-5.4"], contextWindow: 272_000 },
    { prefixes: ["gpt-5"], contextWindow: 400_000 },
    { prefixes: ["grok-build", "grok-code"], contextWindow: 256_000 },
    { prefixes: ["grok-"], contextWindow: 500_000 },
    { prefixes: ["hy"], contextWindow: 256_000, freeContextWindow: 190_000 },
    {
      prefixes: ["minimax-m3"],
      contextWindow: 512_000,
      freeContextWindow: 200_000,
    },
    // 200k-262k families (free variants sit at the low end).
    { prefixes: ["claude-"], contextWindow: 200_000 },
    {
      prefixes: ["deepseek-"],
      contextWindow: 262_144,
      freeContextWindow: 200_000,
    },
    { prefixes: ["glm-5"], contextWindow: 204_800 },
    { prefixes: ["glm-"], contextWindow: 204_800 },
    { prefixes: ["kimi-"], contextWindow: 262_144, freeContextWindow: 200_000 },
    {
      prefixes: ["minimax-m2"],
      contextWindow: 204_800,
    },
    {
      prefixes: ["minimax-"],
      contextWindow: 204_800,
      freeContextWindow: 200_000,
    },
    {
      prefixes: ["mimo-v2-omni", "mimo-v2-flash"],
      contextWindow: 262_144,
    },
    { prefixes: ["qwen"], contextWindow: 262_144, freeContextWindow: 200_000 },
    {
      prefixes: ["big-pickle"],
      contextWindow: 262_144,
      freeContextWindow: 200_000,
    },
    {
      prefixes: ["laguna-"],
      contextWindow: 262_144,
      freeContextWindow: 200_000,
    },
    { prefixes: ["ling-"], contextWindow: 262_144, freeContextWindow: 200_000 },
    {
      prefixes: ["longcat"],
      contextWindow: 262_144,
      freeContextWindow: 200_000,
    },
    { prefixes: ["muse-"], contextWindow: 262_144, freeContextWindow: 200_000 },
    {
      prefixes: ["north-mini"],
      contextWindow: 262_144,
      freeContextWindow: 200_000,
    },
    { prefixes: ["omen-"], contextWindow: 262_144, freeContextWindow: 200_000 },
  ];

/**
 * OpenCode Go overrides for families where models.dev Go differs from the
 * shared base table. Only differences are listed here.
 */
export const OPENCODE_GO_CONTEXT_WINDOW_RULES: readonly OpenCodeContextWindowRule[] =
  [
    { prefixes: ["deepseek-v4.1-flash"], contextWindow: 1_000_000 },
    { prefixes: ["glm-5"], contextWindow: 202_752 },
    { prefixes: ["hy4"], contextWindow: 1_024_000 },
    {
      prefixes: ["minimax-m3"],
      contextWindow: 1_000_000,
      freeContextWindow: 200_000,
    },
    { prefixes: ["omen-"], contextWindow: 500_000, freeContextWindow: 200_000 },
    { prefixes: ["ox-alpha"], contextWindow: 1_000_000 },
    { prefixes: ["qwen3.6-plus"], contextWindow: 1_000_000 },
    {
      prefixes: ["qwen3.8"],
      contextWindow: 1_000_000,
      freeContextWindow: 200_000,
    },
  ];

interface OpenCodeContextWindowMatch {
  readonly rule: OpenCodeContextWindowRule;
  readonly prefixLength: number;
}

function findLongestPrefixMatch(
  rules: readonly OpenCodeContextWindowRule[],
  id: string,
): OpenCodeContextWindowMatch | undefined {
  let best: OpenCodeContextWindowMatch | undefined;
  for (const rule of rules) {
    for (const prefix of rule.prefixes) {
      if (!id.startsWith(prefix)) continue;
      if (best === undefined || prefix.length > best.prefixLength) {
        best = { rule, prefixLength: prefix.length };
      }
    }
  }
  return best;
}

/**
 * Infer a conservative context window for a discovered OpenCode model id.
 *
 * `providerId` selects the provider-specific override table (`"opencode-go"`);
 * every other id uses the shared base table. Unknown families fall back to
 * `OPENCODE_FALLBACK_CONTEXT_WINDOW`.
 */
export function inferOpenCodeContextWindow(
  modelId: string,
  providerId: string,
): number {
  const id = modelId.toLowerCase();
  const rules =
    providerId === "opencode-go"
      ? [
          ...OPENCODE_GO_CONTEXT_WINDOW_RULES,
          ...OPENCODE_BASE_CONTEXT_WINDOW_RULES,
        ]
      : OPENCODE_BASE_CONTEXT_WINDOW_RULES;
  const match = findLongestPrefixMatch(rules, id);
  if (match === undefined) return OPENCODE_FALLBACK_CONTEXT_WINDOW;
  if (!id.endsWith("-free")) return match.rule.contextWindow;
  return match.rule.freeContextWindow ?? match.rule.contextWindow;
}
