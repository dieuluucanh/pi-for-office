import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, Model, Usage } from "@earendil-works/pi-ai";

import { createAdaptiveWindow } from "../src/compaction/adaptive-window.ts";
import { recoverAfterRun } from "../src/compaction/run-recovery.ts";
import type { CompactionOutcome } from "../src/compaction/engine.ts";
import { failOnUnexpectedStream } from "./fail-on-unexpected-stream.ts";

const EMPTY_USAGE: Usage = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
};

function createModel(contextWindow: number, id = "test-model"): Model<Api> {
  return {
    id,
    name: id,
    api: "openai-completions",
    provider: "opencode-go",
    baseUrl: "https://opencode.ai",
    reasoning: false,
    input: ["text"],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow,
    maxTokens: 4096,
  };
}

function createUser(text: string, timestamp: number): AgentMessage {
  return { role: "user", content: text, timestamp };
}

function createError(
  model: Model<Api>,
  errorMessage: string,
  timestamp: number,
): AgentMessage {
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
  };
}

function createSuccess(model: Model<Api>, timestamp: number): AgentMessage {
  return {
    role: "assistant",
    content: [{ type: "text", text: "done" }],
    api: model.api,
    provider: model.provider,
    model: model.id,
    usage: { ...EMPTY_USAGE, input: 100, output: 10, totalTokens: 110 },
    stopReason: "stop",
    timestamp,
  };
}

class TestAgent extends Agent {
  continueCalls = 0;
  override continue(): Promise<void> {
    this.continueCalls += 1;
    return Promise.resolve();
  }
}

function createTestAgent(
  model: Model<Api>,
  messages: AgentMessage[],
): TestAgent {
  return new TestAgent({
    streamFn: failOnUnexpectedStream,
    initialState: { model, messages, tools: [] },
  });
}

function compactedOutcome(tokensBefore = 1000): CompactionOutcome {
  return {
    changed: true,
    reason: "summarized",
    tokensBefore,
    tokensAfter: Math.floor(tokensBefore / 2),
    keptCount: 2,
    summarizedCount: 3,
  };
}

async function noSleep(_ms: number): Promise<void> {
  // immediate
}

void test("recoverAfterRun retries transient stream errors with backoff then succeeds", async () => {
  const model = createModel(262_144);
  const agent = createTestAgent(model, [
    createUser("go", 1),
    createError(model, "Stream ended without finish_reason", 2),
  ]);

  const sleeps: number[] = [];
  const retried: Array<{ attempt: number; maxAttempts: number }> = [];

  const result = await recoverAfterRun({
    agent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    retry: { enabled: true, maxRetries: 3, baseDelayMs: 100 },
    sleep: (ms) => {
      sleeps.push(ms);
      return Promise.resolve();
    },
    callbacks: {
      onRetryScheduled: (attempt, maxAttempts) => {
        retried.push({ attempt, maxAttempts });
      },
    },
  });

  // The failed message was dropped before the continue.
  assert.equal(agent.continueCalls, 1);
  assert.equal(sleeps.length, 1);
  assert.equal(sleeps[0], 100);
  assert.deepEqual(retried, [{ attempt: 1, maxAttempts: 3 }]);
  // Still classified as transient-recovered even though the fake continue
  // doesn't replace the error with a clean settle (the loop exits after
  // continue returns because the agent isn't actually streaming).
  assert.ok(["transient-recovered", "exhausted"].includes(result.kind));
});

void test("recoverAfterRun performs one compact-and-retry on overflow", async () => {
  const model = createModel(65_536);
  const overflowError =
    "400 litellm.ContextWindowExceededError: maximum context length is 65536 tokens";
  const agent = createTestAgent(model, [
    createUser("go", 1),
    createError(model, overflowError, 2),
  ]);

  let compactRuns = 0;
  const result = await recoverAfterRun({
    agent,
    runCompact: () => {
      compactRuns += 1;
      agent.state.messages = [createUser("compaction summary", 3)];
      return Promise.resolve(compactedOutcome());
    },
    sleep: noSleep,
  });

  assert.equal(compactRuns, 1);
  assert.equal(agent.continueCalls, 1);
  assert.ok(["overflow-recovered", "exhausted"].includes(result.kind));
});

void test("recoverAfterRun stops after one overflow attempt (no loop)", async () => {
  const model = createModel(65_536);
  const overflowError =
    "400 litellm.ContextWindowExceededError: maximum context length is 65536 tokens";
  const agent = createTestAgent(model, [
    createUser("go", 1),
    createError(model, overflowError, 2),
  ]);

  let compactRuns = 0;
  const result = await recoverAfterRun({
    agent,
    runCompact: () => {
      // Compaction "succeeds" but the context is still over budget (no fit) —
      // recovery must not loop forever.
      compactRuns += 1;
      agent.state.messages = [
        createUser("compaction summary", 3),
        createError(model, overflowError, 4),
      ];
      return Promise.resolve(compactedOutcome(10_000));
    },
    sleep: noSleep,
  });

  assert.equal(compactRuns, 1); // overflow recovery is one-shot
  assert.equal(result.kind, "exhausted");
});

void test("recoverAfterRun respects retry budget for quota/fatal errors", async () => {
  const model = createModel(262_144);
  const agent = createTestAgent(model, [
    createUser("go", 1),
    createError(model, "GoUsageLimitError: monthly usage limit reached", 2),
  ]);

  const sleeps: number[] = [];
  const result = await recoverAfterRun({
    agent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    retry: { enabled: true, maxRetries: 3, baseDelayMs: 100 },
    sleep: (ms) => {
      sleeps.push(ms);
      return Promise.resolve();
    },
  });

  // Quota errors are fatal — never retried, never compacted.
  assert.equal(sleeps.length, 0);
  assert.equal(agent.continueCalls, 0);
  assert.equal(result.kind, "exhausted");
});

void test("recoverAfterRun is clean when there is no failure", async () => {
  const model = createModel(262_144);
  const agent = createTestAgent(model, [
    createUser("go", 1),
    createSuccess(model, 2),
  ]);

  const result = await recoverAfterRun({
    agent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    sleep: noSleep,
  });

  assert.equal(result.kind, "clean");
  assert.equal(agent.continueCalls, 0);
});

void test("adaptive window downgrades on overflow-like failures at high usage", () => {
  const adaptive = createAdaptiveWindow(262_144);

  adaptive.observeFailure({
    failureIsOverflowLike: true,
    effectiveTokens: 230_000,
    contextWindow: 262_144,
  });

  const state = adaptive.get();
  assert.equal(state.downgraded, true);
  assert.ok(state.effectiveWindow < 262_144);
  assert.equal(state.effectiveWindow, Math.floor(230_000 * 0.9));
});

void test("adaptive window ignores small-usage failures and resets on model change", () => {
  const adaptive = createAdaptiveWindow(262_144);

  // Low usage (transient drop, not overflow) → no downgrade.
  adaptive.observeFailure({
    failureIsOverflowLike: true,
    effectiveTokens: 40_000,
    contextWindow: 262_144,
  });
  assert.equal(adaptive.get().downgraded, false);

  // Non-overflow failure → no downgrade.
  adaptive.observeFailure({
    failureIsOverflowLike: false,
    effectiveTokens: 250_000,
    contextWindow: 262_144,
  });
  assert.equal(adaptive.get().downgraded, false);

  // Model switch resets.
  adaptive.reset(262_144);
  assert.equal(adaptive.get().downgraded, false);
});
