import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentState } from "@earendil-works/pi-agent-core";
import type {
  Api,
  AssistantMessage,
  Model,
  Usage,
} from "@earendil-works/pi-ai";

import {
  classifyRunFailure,
  findTrailingContextOverflowError,
  isPrematureStreamEnd,
  isRecoverableLength,
} from "../src/compaction/failure-classification.ts";

const EMPTY_USAGE: Usage = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
};

// Real error shape from #558 (LiteLLM custom gateway, 65k model).
const LITELLM_OVERFLOW_ERROR =
  "400 litellm.ContextWindowExceededError: litellm.BadRequestError: ContextWindowExceededError: " +
  "Hosted_vllmException - This model's maximum context length is 65536 tokens. However, you requested " +
  "4096 output tokens and your prompt contains at least 61441 input tokens.";

const STREAM_END_ERROR = "Stream ended without finish_reason";

function createModel(
  contextWindow: number,
  id = "deepseek-r1-32b",
): Model<Api> {
  return {
    id,
    name: id,
    api: "openai-completions",
    provider: "custom-gateway",
    baseUrl: "https://gateway.example.invalid",
    reasoning: false,
    input: ["text"],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow,
    maxTokens: 4096,
  };
}

function createUser(text: string, timestamp: number) {
  return { role: "user" as const, content: text, timestamp };
}

function createAssistantError(
  model: Model<Api>,
  errorMessage: string,
  timestamp: number,
  extra: Partial<AssistantMessage> = {},
): AssistantMessage {
  return {
    role: "assistant",
    content: [{ type: "text", text: "" }],
    api: model.api,
    provider: model.provider,
    model: model.id,
    usage: EMPTY_USAGE,
    stopReason: "error",
    errorMessage,
    timestamp,
    ...extra,
  };
}

function createUsage(input: number, output: number): Usage {
  return {
    input,
    output,
    cacheRead: 0,
    cacheWrite: 0,
    totalTokens: input + output,
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
  };
}

function stateFor(
  model: Model<Api>,
  messages: Array<{ role: string; content: string; timestamp: number }>,
): Pick<AgentState, "messages" | "model"> {
  return { model, messages: messages as AgentState["messages"] };
}

void test("classifyRunFailure detects overflow from error text (no stopReason gate)", () => {
  const model = createModel(65_536);
  const failure = createAssistantError(model, LITELLM_OVERFLOW_ERROR, 3);

  const result = classifyRunFailure({
    state: stateFor(model, [createUser("hi", 1), failure]),
  });
  assert.equal(result.kind, "overflow");
  assert.equal(result.message, failure);
});

void test("classifyRunFailure ignores stale failures from a different model", () => {
  const oldModel = createModel(65_536, "small-model");
  const newModel = createModel(200_000, "big-model");
  const failure = createAssistantError(oldModel, LITELLM_OVERFLOW_ERROR, 2);

  const result = classifyRunFailure({
    state: stateFor(newModel, [createUser("hi", 1), failure]),
  });
  assert.equal(result.kind, "none");
});

void test("classifyRunFailure detects silent overflow via usage (z.ai style)", () => {
  const model = createModel(200_000);
  // stopReason "stop" but usage.input exceeds the window — silent overflow.
  const silent = {
    ...createAssistantError(model, "", 3, {
      stopReason: "stop" as const,
      content: [{ type: "text", text: "done" }],
      usage: createUsage(210_000, 100),
    }),
  };

  const result = classifyRunFailure({
    state: stateFor(model, [createUser("hi", 1), silent]),
    contextWindow: model.contextWindow,
  });
  assert.equal(result.kind, "overflow");
});

void test("classifyRunFailure detects MiMo length-stop overflow", () => {
  const model = createModel(200_000);
  const mimo = {
    role: "assistant" as const,
    content: [{ type: "text", text: "" }],
    api: model.api,
    provider: model.provider,
    model: model.id,
    usage: createUsage(199_000, 0), // input fills ~99.5% of the window, zero output
    stopReason: "length" as const,
    timestamp: 3,
  };

  const result = classifyRunFailure({
    state: stateFor(model, [createUser("hi", 1), mimo]),
    contextWindow: model.contextWindow,
  });
  assert.equal(result.kind, "overflow");
});

void test("classifyRunFailure detects recoverable length stops", () => {
  const model = createModel(200_000);
  const truncated = {
    role: "assistant" as const,
    content: [{ type: "text", text: "partial" }],
    api: model.api,
    provider: model.provider,
    model: model.id,
    usage: createUsage(20_000, 512), // output well below 4096 desired
    stopReason: "length" as const,
    timestamp: 3,
  };

  const result = classifyRunFailure({
    state: stateFor(model, [createUser("hi", 1), truncated]),
    maxOutputTokens: 4096,
  });
  assert.equal(result.kind, "recoverable-length");
});

void test("classifyRunFailure treats stream-end as transient at low usage", () => {
  const model = createModel(262_144);
  const failure = createAssistantError(model, STREAM_END_ERROR, 3);

  const result = classifyRunFailure({
    state: stateFor(model, [createUser("hi", 1), failure]),
    contextWindow: model.contextWindow,
    effectiveTokens: 50_000, // well below soft warning
  });
  assert.equal(result.kind, "transient");
});

void test("classifyRunFailure treats stream-end as overflow near the limit", () => {
  const model = createModel(262_144);
  const failure = createAssistantError(model, STREAM_END_ERROR, 3);

  const result = classifyRunFailure({
    state: stateFor(model, [createUser("hi", 1), failure]),
    contextWindow: model.contextWindow,
    effectiveTokens: 230_000, // past the 262k soft warning (~217k)
  });
  assert.equal(result.kind, "overflow");
});

void test("classifyRunFailure keeps rate limits transient (no compact-first)", () => {
  const model = createModel(262_144);
  const rateLimited = createAssistantError(
    model,
    "429 rate limit exceeded, too many requests",
    3,
  );

  const result = classifyRunFailure({
    state: stateFor(model, [createUser("hi", 1), rateLimited]),
    contextWindow: model.contextWindow,
    effectiveTokens: 250_000,
  });
  // Not overflow: no premature-stream-end pattern, so it stays transient/retryable.
  assert.equal(result.kind, "transient");
});

void test("classifyRunFailure marks quota/billing errors fatal (non-retryable)", () => {
  const model = createModel(262_144);
  const quota = createAssistantError(
    model,
    "GoUsageLimitError: monthly usage limit reached",
    3,
  );

  const result = classifyRunFailure({
    state: stateFor(model, [createUser("hi", 1), quota]),
    contextWindow: model.contextWindow,
    effectiveTokens: 250_000,
  });
  assert.equal(result.kind, "fatal");
});

void test("classifyRunFailure returns none for clean tails", () => {
  const model = createModel(200_000);
  const result = classifyRunFailure({
    state: stateFor(model, [
      createUser("hi", 1),
      {
        role: "toolResult",
        toolCallId: "c1",
        toolName: "read_range",
        content: [{ type: "text", text: "rows" }],
        isError: false,
        timestamp: 2,
      },
    ]),
  });
  assert.equal(result.kind, "none");
});

void test("findTrailingContextOverflowError is the overflow-only wrapper", () => {
  const model = createModel(65_536);
  const overflow = createAssistantError(model, LITELLM_OVERFLOW_ERROR, 3);
  assert.equal(
    findTrailingContextOverflowError(
      stateFor(model, [createUser("hi", 1), overflow]),
    ),
    overflow,
  );

  const transient = createAssistantError(model, STREAM_END_ERROR, 3);
  assert.equal(
    findTrailingContextOverflowError(
      stateFor(model, [createUser("hi", 1), transient]),
    ),
    null,
  );
});

void test("isPrematureStreamEnd matches stream-truncation signatures", () => {
  assert.equal(
    isPrematureStreamEnd("Stream ended without finish_reason"),
    true,
  );
  assert.equal(isPrematureStreamEnd("stream ended before message_stop"), true);
  assert.equal(isPrematureStreamEnd("400 something else"), false);
});

void test("isRecoverableLength flags short length stops", () => {
  const model = createModel(200_000);
  const short: AssistantMessage = {
    role: "assistant",
    content: [{ type: "text", text: "partial" }],
    api: model.api,
    provider: model.provider,
    model: model.id,
    usage: createUsage(20_000, 100),
    stopReason: "length",
    timestamp: 3,
  };
  assert.equal(isRecoverableLength(short, 4096), true);
  assert.equal(isRecoverableLength(short, 100), false);

  const stopped: AssistantMessage = {
    ...short,
    stopReason: "stop",
  };
  assert.equal(isRecoverableLength(stopped, 4096), false);
});
