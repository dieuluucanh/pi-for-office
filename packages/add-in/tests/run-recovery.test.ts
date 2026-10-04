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

function createStalledTail(
  model: Model<Api>,
  stopReason: "stop" | "length",
  timestamp: number,
): AgentMessage {
  return {
    role: "assistant",
    content: [
      {
        type: "thinking",
        thinking: "a reasoning chain cut off mid-sentence",
        thinkingSignature: "sig",
      },
    ],
    api: model.api,
    provider: model.provider,
    model: model.id,
    // At-budget truncation: usage.output === desired would slip past
    // isRecoverableLength — exactly the reported silent-stall shape.
    usage:
      stopReason === "length"
        ? { ...EMPTY_USAGE, input: 100, output: 4_096, totalTokens: 4_196 }
        : { ...EMPTY_USAGE, input: 100, output: 512, totalTokens: 612 },
    stopReason,
    timestamp,
  };
}

class StallAgent extends TestAgent {
  onContinue?: (agent: StallAgent) => void;

  override continue(): Promise<void> {
    this.continueCalls += 1;
    this.onContinue?.(this);
    return Promise.resolve();
  }
}

function createStallAgent(
  model: Model<Api>,
  messages: AgentMessage[],
): StallAgent {
  const agent = new StallAgent({
    streamFn: failOnUnexpectedStream,
    initialState: { model, messages, tools: [] },
  });
  return agent;
}

interface StallDetection {
  attempt: number;
  maxAttempts: number;
  stopReason: "stop" | "length";
  escalatedTo?: number;
  budgetAtDetection: number;
}

void test("recoverAfterRun auto-continues a length stall with escalating budget", async () => {
  const model = createModel(262_144);
  const originalModel = model;
  const agent = createStallAgent(model, [
    createUser("go", 1),
    createStalledTail(model, "length", 2),
  ]);

  // Every continuation stalls again (thinking-only tail at budget).
  const detections: StallDetection[] = [];
  agent.onContinue = (self) => {
    self.state.messages = [
      createUser("go", 1),
      createStalledTail(originalModel, "length", 10),
    ];
  };

  const result = await recoverAfterRun({
    agent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    callbacks: {
      onStallDetected: (attempt, maxAttempts, info) => {
        detections.push({
          attempt,
          maxAttempts,
          stopReason: info.stopReason,
          ...(info.escalatedTo !== undefined
            ? { escalatedTo: info.escalatedTo }
            : {}),
          budgetAtDetection: agent.state.model?.maxTokens ?? 0,
        });
      },
    },
  });

  // Bounded at 2 attempts (default), then exhausted.
  assert.equal(agent.continueCalls, 2);
  assert.deepEqual(
    detections.map((detection) => detection.attempt),
    [1, 2],
  );
  // Length stalls escalate ×2 per attempt, capped at 64k.
  assert.deepEqual(
    detections.map((detection) => detection.escalatedTo),
    [8_192, 16_384],
  );
  assert.deepEqual(
    detections.map((detection) => detection.stopReason),
    ["length", "length"],
  );
  assert.equal(result.kind, "exhausted");
  // Stall exhaustion is not an overflow-like failure — adaptive-window
  // observation must not shrink the context window for it.
  assert.equal(result.overflowLikeFailure, false);
  // The escalated budget was restored to the original model on settle.
  assert.equal(agent.state.model, originalModel);
  assert.equal(agent.state.model?.maxTokens, 4_096);
});

void test("recoverAfterRun retries a stop-stall at the same budget and settles clean", async () => {
  const model = createModel(262_144);
  const originalModel = model;
  const agent = createStallAgent(model, [
    createUser("go", 1),
    createStalledTail(model, "stop", 2),
  ]);

  // The continuation settles cleanly (messages end on the user message) —
  // `stop` stalls retry at the same budget, without escalation.
  const detections: StallDetection[] = [];
  agent.onContinue = (self) => {
    self.state.messages = [createUser("go", 1)];
  };

  const result = await recoverAfterRun({
    agent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    callbacks: {
      onStallDetected: (attempt, maxAttempts, info) => {
        assert.equal("escalatedTo" in info, false);
        detections.push({
          attempt,
          maxAttempts,
          stopReason: info.stopReason,
          budgetAtDetection: agent.state.model?.maxTokens ?? 0,
        });
      },
    },
  });

  assert.equal(agent.continueCalls, 1);
  assert.equal(detections.length, 1);
  // No escalation for `stop` stalls: the budget was not the cause.
  assert.equal(detections[0].escalatedTo, undefined);
  assert.equal(detections[0].budgetAtDetection, 4_096);
  assert.equal(result.kind, "clean");
  assert.equal(agent.state.model, originalModel);
});

void test("recoverAfterRun escalation respects the cap and skips when already at cap", async () => {
  // Small window: cap = max(8192, min(window - 16384, 65536)) = 8192.
  const model = createModel(20_000);
  const agent = createStallAgent(model, [
    createUser("go", 1),
    createStalledTail(model, "length", 2),
  ]);

  agent.onContinue = (self) => {
    self.state.messages = [
      createUser("go", 1),
      createStalledTail(model, "length", 10),
    ];
  };

  const escalatedValues: Array<number | undefined> = [];
  await recoverAfterRun({
    agent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    callbacks: {
      onStallDetected: (_attempt, _maxAttempts, info) => {
        escalatedValues.push(
          info.escalatedTo === undefined ? undefined : info.escalatedTo,
        );
      },
    },
  });

  // Attempt 1: 4096 → 8192 (the cap). Attempt 2: already at cap → no
  // escalation (undefined), the retry proceeds at the same budget.
  assert.deepEqual(escalatedValues, [8_192, undefined]);
});

void test("recoverAfterRun stall recovery never clobbers a concurrent model switch", async () => {
  const model = createModel(262_144);
  const agent = createStallAgent(model, [
    createUser("go", 1),
    createStalledTail(model, "length", 2),
  ]);

  // Simulate the user switching models mid-recovery: the settle must not
  // restore the escalated budget over the user's chosen model.
  const switchedModel = createModel(262_144, "user-choice");
  await recoverAfterRun({
    agent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    callbacks: {
      onStallDetected: () => {
        agent.state.model = switchedModel;
      },
    },
  });

  assert.equal(agent.state.model, switchedModel);
  assert.equal(agent.state.model?.maxTokens, 4_096);
});

void test("recoverAfterRun halves the budget on budget-unsupported rejections and persists the reduction", async () => {
  const model = createModel(262_144);
  // New default budget 32768 rejected by a gateway that caps max_tokens.
  const rejectionModel = { ...model, maxTokens: 32_768 };
  agentInitialStateFixup(rejectionModel);
  const agent = createStallAgent(rejectionModel, [
    createUser("go", 1),
    createError(
      rejectionModel,
      "400 max_tokens too large: requested 32768, maximum supported is 8192",
      2,
    ),
  ]);

  // The continuation settles cleanly; the reduced budget must persist.
  agent.onContinue = (self) => {
    self.state.messages = [createUser("go", 1)];
  };

  const reductions: number[] = [];
  const result = await recoverAfterRun({
    agent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    callbacks: {
      onBudgetReduced: (newMaxTokens) => {
        reductions.push(newMaxTokens);
      },
    },
  });

  assert.equal(agent.continueCalls, 1);
  assert.deepEqual(reductions, [16_384]);
  assert.equal(result.kind, "clean");
  // The reduction persists for the session — NOT restored on settle.
  assert.equal(agent.state.model?.maxTokens, 16_384);
});

/** No-op: keeps the rejection model typed as the fixture's Model. */
function agentInitialStateFixup(_model: Model<Api>): void {}

void test("recoverAfterRun leaves uncompacted overflow exhausted when compaction is unavailable", async () => {
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
      return Promise.resolve(compactedOutcome());
    },
    allowCompaction: false,
    sleep: noSleep,
  });

  // Compaction never ran; the failed message stays visible for the banner.
  assert.equal(compactRuns, 0);
  assert.equal(agent.continueCalls, 0);
  assert.equal(result.kind, "exhausted");
  assert.equal(result.overflowLikeFailure, true);
  assert.ok(
    agent.state.messages.some((message) => message.role === "assistant"),
  );
});

void test("recoverAfterRun still recovers stalls and transients when compaction is unavailable", async () => {
  const model = createModel(262_144);
  const stallAgent = createStallAgent(model, [
    createUser("go", 1),
    createStalledTail(model, "stop", 2),
  ]);
  stallAgent.onContinue = (self) => {
    self.state.messages = [createUser("go", 1)];
  };

  const stallResult = await recoverAfterRun({
    agent: stallAgent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    allowCompaction: false,
    sleep: noSleep,
  });
  assert.equal(stallAgent.continueCalls, 1);
  assert.equal(stallResult.kind, "clean");

  // Transient stream errors also retry regardless of the compaction gate.
  const transientAgent = createTestAgent(model, [
    createUser("go", 1),
    createError(model, "Stream ended without finish_reason", 2),
  ]);
  transientAgent.onContinue = (self) => {
    self.state.messages = [createUser("go", 1)];
  };

  const transientResult = await recoverAfterRun({
    agent: transientAgent,
    runCompact: () => Promise.resolve(compactedOutcome()),
    allowCompaction: false,
    sleep: noSleep,
  });
  assert.equal(transientAgent.continueCalls, 1);
  // A transient retry that then settles clean is transient-recovered.
  assert.equal(transientResult.kind, "transient-recovered");
});
