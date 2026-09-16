import assert from "node:assert/strict";
import { test } from "node:test";

import { inferOpenCodeContextWindow } from "../src/models/opencode-context-windows.ts";

/**
 * Every context window listed by models.dev for the `opencode-go` provider
 * (page + `api.json` provider key, verified 2026-09-16). Static pi-ai catalogue
 * ids are included because inference must agree with them for new variants.
 */
const OPENCODE_GO_EXPECTED_CONTEXT_WINDOWS: readonly (readonly [
 string,
 number,
])[] = [
 ["qwen3.7-max", 1_000_000],
 ["longcat-2.0", 1_000_000],
 ["deepseek-v4-flash-vision-exp", 1_000_000],
 ["qwen3.6-plus", 1_000_000],
 ["muse-spark-1.2-contributor", 1_048_576],
 ["minimax-m2.7", 204_800],
 ["kimi-k2.6", 262_144],
 ["glm-5.2", 1_000_000],
 ["minimax-m2.5", 204_800],
 ["minimax-m3", 1_000_000],
 ["deepseek-v4-flash", 1_000_000],
 ["kimi-k2.7-code", 262_144],
 ["grok-4.5", 500_000],
 ["ox-alpha-free", 1_000_000],
 ["deepseek-v4.1-flash", 1_000_000],
 ["hy3", 256_000],
 ["hy4-preview", 1_024_000],
 ["omen-alpha", 500_000],
 ["gpt-5.6-luna", 1_050_000],
 ["kimi-k3", 1_048_576],
 ["glm-5.3-flash", 1_000_000],
 ["muse-spark-1.3-contributor", 1_048_576],
 ["mimo-v2-pro", 1_048_576],
 ["qwen3.8-flash", 1_000_000],
 ["grok-4.6", 500_000],
 ["glm-5", 202_752],
 ["mimo-v2-omni", 262_144],
 ["qwen3.8-max", 1_000_000],
 ["kimi-k2.5", 262_144],
 ["glm-5.1", 202_752],
 ["qwen3.7-plus", 1_000_000],
 ["deepseek-v4-pro", 1_000_000],
 ["glm-5.3", 1_000_000],
 ["mimo-v2.5", 1_000_000],
 ["mimo-v2.5-pro", 1_048_576],
 ["qwen3.5-plus", 262_144],
];

void test("inferOpenCodeContextWindow matches the models.dev OpenCode Go catalogue", () => {
 assert.equal(OPENCODE_GO_EXPECTED_CONTEXT_WINDOWS.length, 36);
 for (const [id, contextWindow] of OPENCODE_GO_EXPECTED_CONTEXT_WINDOWS) {
  assert.equal(
   inferOpenCodeContextWindow(id, "opencode-go"),
   contextWindow,
   `${id} should resolve to ${contextWindow}`,
  );
 }
});

void test("inferOpenCodeContextWindow keeps Zen and Go family values apart", () => {
 // Families where the two gateways publish different windows.
 assert.equal(
  inferOpenCodeContextWindow("minimax-m3", "opencode-go"),
  1_000_000,
 );
 assert.equal(inferOpenCodeContextWindow("minimax-m3", "opencode"), 512_000);
 assert.equal(
  inferOpenCodeContextWindow("qwen3.6-plus", "opencode-go"),
  1_000_000,
 );
 assert.equal(inferOpenCodeContextWindow("qwen3.6-plus", "opencode"), 262_144);
 assert.equal(inferOpenCodeContextWindow("glm-5", "opencode-go"), 202_752);
 assert.equal(inferOpenCodeContextWindow("glm-5", "opencode"), 204_800);
 assert.equal(inferOpenCodeContextWindow("glm-5.1", "opencode-go"), 202_752);
 assert.equal(inferOpenCodeContextWindow("glm-5.1", "opencode"), 204_800);
 assert.equal(inferOpenCodeContextWindow("omen-alpha", "opencode-go"), 500_000);
 assert.equal(inferOpenCodeContextWindow("omen-alpha", "opencode"), 262_144);
 // Go-only family: Zen falls back to the shared `hy` tier (256,000).
 assert.equal(
  inferOpenCodeContextWindow("hy4-preview", "opencode-go"),
  1_024_000,
 );
 assert.equal(inferOpenCodeContextWindow("hy4-preview", "opencode"), 256_000);
 // Go-only free model: models.dev lists it at 1M, so no 200k free cap applies.
 assert.equal(
  inferOpenCodeContextWindow("ox-alpha-free", "opencode-go"),
  1_000_000,
 );
 // Shared families resolve through the base table for both gateways.
 assert.equal(
  inferOpenCodeContextWindow("longcat-2.0", "opencode-go"),
  1_000_000,
 );
 assert.equal(inferOpenCodeContextWindow("longcat-2.0", "opencode"), 1_000_000);
 assert.equal(inferOpenCodeContextWindow("kimi-k3", "opencode-go"), 1_048_576);
 assert.equal(inferOpenCodeContextWindow("kimi-k3", "opencode"), 1_048_576);
});

void test("inferOpenCodeContextWindow maps base families to conservative tiers", () => {
 // ~1M families.
 assert.equal(
  inferOpenCodeContextWindow("claude-opus-4-6", "opencode"),
  1_000_000,
 );
 assert.equal(
  inferOpenCodeContextWindow("claude-opus-5", "opencode"),
  1_000_000,
 );
 assert.equal(
  inferOpenCodeContextWindow("claude-fable-5", "opencode"),
  1_000_000,
 );
 assert.equal(
  inferOpenCodeContextWindow("claude-sonnet-5", "opencode"),
  1_000_000,
 );
 assert.equal(
  inferOpenCodeContextWindow("gemini-3.7-flash", "opencode"),
  1_048_576,
 );
 assert.equal(inferOpenCodeContextWindow("gpt-5.5", "opencode"), 1_050_000);
 assert.equal(
  inferOpenCodeContextWindow("gpt-5.6-terra", "opencode"),
  1_050_000,
 );
 assert.equal(inferOpenCodeContextWindow("gpt-5.4-pro", "opencode"), 1_050_000);
 assert.equal(
  inferOpenCodeContextWindow("deepseek-v4-flash", "opencode"),
  1_000_000,
 );
 assert.equal(
  inferOpenCodeContextWindow("deepseek-v4-pro", "opencode"),
  1_000_000,
 );
 assert.equal(inferOpenCodeContextWindow("glm-5.2", "opencode"), 1_000_000);
 assert.equal(inferOpenCodeContextWindow("glm-5.3", "opencode"), 1_000_000);
 assert.equal(
  inferOpenCodeContextWindow("nemotron-3-ultra-free", "opencode"),
  1_000_000,
 );
 assert.equal(inferOpenCodeContextWindow("mimo-v2.5", "opencode"), 1_000_000);
 assert.equal(inferOpenCodeContextWindow("mimo-v2-pro", "opencode"), 1_048_576);
 assert.equal(
  inferOpenCodeContextWindow("muse-spark-1.3", "opencode"),
  1_048_576,
 );
 assert.equal(
  inferOpenCodeContextWindow("qwen3.7-plus", "opencode"),
  1_000_000,
 );

 // Mid tiers.
 assert.equal(inferOpenCodeContextWindow("grok-4.5", "opencode"), 500_000);
 assert.equal(inferOpenCodeContextWindow("grok-4.6", "opencode"), 500_000);
 assert.equal(inferOpenCodeContextWindow("grok-code", "opencode"), 256_000);
 assert.equal(inferOpenCodeContextWindow("gpt-5.4", "opencode"), 272_000);
 assert.equal(inferOpenCodeContextWindow("gpt-5.9-codex", "opencode"), 400_000);
 assert.equal(inferOpenCodeContextWindow("gpt-5", "opencode"), 400_000);
 assert.equal(inferOpenCodeContextWindow("hy3", "opencode"), 256_000);
 assert.equal(inferOpenCodeContextWindow("minimax-m3", "opencode"), 512_000);

 // 200k-262k families (free variants sit at the low end).
 assert.equal(
  inferOpenCodeContextWindow("claude-haiku-6", "opencode"),
  200_000,
 );
 assert.equal(inferOpenCodeContextWindow("qwen3.8-plus", "opencode"), 262_144);
 assert.equal(inferOpenCodeContextWindow("kimi-k4", "opencode"), 262_144);
 assert.equal(inferOpenCodeContextWindow("glm-4.7", "opencode"), 204_800);
 assert.equal(inferOpenCodeContextWindow("minimax-m4", "opencode"), 204_800);
 assert.equal(inferOpenCodeContextWindow("longcat-3.0", "opencode"), 262_144);
 assert.equal(inferOpenCodeContextWindow("mimo-v2-omni", "opencode"), 262_144);

 // Over-estimate fixes: these must never exceed the published window.
 assert.equal(inferOpenCodeContextWindow("glm-5", "opencode"), 204_800);
 assert.equal(inferOpenCodeContextWindow("minimax-m2.5", "opencode"), 204_800);
 assert.equal(
  inferOpenCodeContextWindow("minimax-m3-free", "opencode"),
  200_000,
 );
 assert.equal(inferOpenCodeContextWindow("hy3-free", "opencode"), 190_000);

 // Free variants cap to the low tier.
 assert.equal(
  inferOpenCodeContextWindow("deepseek-v4-flash-free", "opencode"),
  200_000,
 );
 assert.equal(
  inferOpenCodeContextWindow("mimo-v2.5-free", "opencode"),
  200_000,
 );
 assert.equal(
  inferOpenCodeContextWindow("claude-sonnet-6-free", "opencode"),
  200_000,
 );

 // Case-insensitive matching.
 assert.equal(
  inferOpenCodeContextWindow("Claude-Opus-4-6", "opencode"),
  1_000_000,
 );
});

void test("inferOpenCodeContextWindow falls back to the default for unknown families", () => {
 assert.equal(
  inferOpenCodeContextWindow("brand-new-model", "opencode"),
  256_000,
 );
 assert.equal(inferOpenCodeContextWindow("acme-ultra-9", "opencode"), 256_000);
 assert.equal(
  inferOpenCodeContextWindow("brand-new-model", "opencode-go"),
  256_000,
 );
});
