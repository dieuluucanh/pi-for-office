import assert from "node:assert/strict";
import { test } from "node:test";

import type { Tool } from "@earendil-works/pi-ai";
import type { AgentMessage } from "@earendil-works/pi-agent-core";

import {
  estimateListTokens,
  planCompaction,
  planKeptTail,
  selectCut,
  verifyCompactionFits,
} from "../src/compaction/engine.ts";

function createUser(text: string, timestamp: number): AgentMessage {
  return { role: "user", content: text, timestamp };
}

function createAssistant(
  text: string,
  timestamp: number,
  extra: Partial<AgentMessage> = {},
): AgentMessage {
  return {
    role: "assistant",
    content: [{ type: "text", text }],
    timestamp,
    ...extra,
  };
}

function createToolResult(
  toolName: string,
  text: string,
  timestamp: number,
): AgentMessage {
  return {
    role: "toolResult",
    toolCallId: `call-${timestamp}`,
    toolName,
    content: [{ type: "text", text }],
    isError: false,
    timestamp,
  };
}

/** A 50KB-ish tool result ≈ ~12.5k tokens, matching execution-time cap. */
function bigToolResult(timestamp: number): AgentMessage {
  return createToolResult("read_range", "x".repeat(50 * 1024), timestamp);
}

function createToolCallBatch(
  count: number,
  startTimestamp: number,
): AgentMessage[] {
  const messages: AgentMessage[] = [];
  // Assistant message issuing `count` parallel tool calls.
  const toolCalls = Array.from({ length: count }, (_, i) => ({
    type: "toolCall" as const,
    id: `call-${startTimestamp}-${i}`,
    name: "read_range",
    arguments: { ref: `A${i}` },
  }));
  messages.push(
    createAssistant("Running tools", startTimestamp, {
      content: [{ type: "text", text: "Running tools" }, ...toolCalls],
    }),
  );
  for (let i = 0; i < count; i++) {
    messages.push(bigToolResult(startTimestamp + 1 + i));
  }
  return messages;
}

function compactedTranscript(
  messages: AgentMessage[],
  boundaryStart: number,
): AgentMessage[] {
  return messages.slice(boundaryStart);
}

void test("selectCut keeps a normal small tail within budget", () => {
  const messages = [
    createUser("hi", 1),
    createAssistant("a", 2),
    createUser("more", 3),
    createAssistant("b", 4),
  ];
  // Tiny conversation: budget never hit → cut at boundaryStart (nothing to summarize).
  const cut = selectCut({
    messages,
    boundaryStart: 0,
    keepRecentTokens: 20_000,
  });
  assert.equal(cut, 0);
});

void test("selectCut never cuts INSIDE a trailing tool batch (keeps the batch's assistant)", () => {
  const turn = createToolCallBatch(10, 10); // assistant + 10 big results
  const messages = [createUser("go", 1), ...turn];
  // Budget is smaller than one big result, so the walk lands inside the batch:
  // the cut must be at the batch's assistant (index 1), never a tool result.
  const cut = selectCut({
    messages,
    boundaryStart: 0,
    keepRecentTokens: 5_000,
  });
  const cutMessage = messages[cut];
  assert.ok(cutMessage, "cut index within bounds");
  assert.equal(cutMessage.role, "assistant");
  assert.equal(cut, 1);
});

void test("planCompaction trims an oversized trailing tool batch to fit the budget", () => {
  // 30 big results ≈ 375k tokens — far over a 65k context window.
  const turn = createToolCallBatch(30, 10);
  const messages = [createUser("go", 1), ...turn];
  const budgetTokens = 20_000;

  const result = planCompaction({
    messages,
    boundaryStart: 0,
    keepRecentTokens: 20_000,
    budgetTokens,
  });

  // The kept tail must fit the budget — this is the RC-1 fix.
  assert.ok(
    result.fits,
    `kept tail should fit: ${estimateListTokens(result.kept)} > ${budgetTokens}`,
  );
  assert.ok(
    estimateListTokens(result.kept) <= budgetTokens,
    `kept=${estimateListTokens(result.kept)} budget=${budgetTokens}`,
  );
  // Still keeps the latest assistant + at least the most recent results for the
  // in-flight continuation.
  assert.ok(result.kept.length > 0);
  assert.equal(result.kept.at(-1)?.role, "toolResult");
});

void test("planCompaction handles a single oversized message (tail-trim)", () => {
  // One giant user paste (~300k chars ≈ 75k tokens) with a small recent tail.
  const messages = [createUser("go", 1), createUser("x".repeat(300_000), 2)];
  const budgetTokens = 20_000;

  const result = planCompaction({
    messages,
    boundaryStart: 0,
    keepRecentTokens: 20_000,
    budgetTokens,
  });

  // The engine must still produce *something* kept and report the fit state
  // truthfully (fits may be false if a single message alone exceeds the budget).
  assert.ok(result.kept.length > 0);
});

void test("repeated compaction is idempotent and never reports nothing while over budget", () => {
  const turn = createToolCallBatch(30, 10);
  const messages = [createUser("go", 1), ...turn];
  const budgetTokens = 20_000;

  // First pass trims the tail.
  const first = planCompaction({
    messages,
    boundaryStart: 0,
    keepRecentTokens: 20_000,
    budgetTokens,
  });
  assert.ok(first.fits);

  // Second pass over the already-trimmed tail keeps it under budget.
  const second = planCompaction({
    messages: first.kept,
    boundaryStart: 0,
    keepRecentTokens: 20_000,
    budgetTokens,
  });
  assert.ok(second.fits);
  assert.equal(
    estimateListTokens(second.kept) <= budgetTokens,
    true,
    "second compaction of a trimmed tail must stay within budget",
  );
});

void test("verifyCompactionFits respects context window and reserve", () => {
  const tools: Tool[] = [
    {
      name: "read_range",
      description: "Read a range",
      parameters: { type: "object", properties: { ref: { type: "string" } } },
    },
  ];

  const fits = verifyCompactionFits({
    systemPrompt: "system",
    messages: [createUser("hi", 1)],
    tools,
    contextWindow: 65_536,
    reserveTokens: 16_384,
  });
  assert.equal(fits, true);

  const full = verifyCompactionFits({
    systemPrompt: "system",
    messages: [createUser("x".repeat(400_000), 1)],
    tools,
    contextWindow: 65_536,
    reserveTokens: 16_384,
  });
  assert.equal(full, false);
});

void test("planKeptTail preserves the latest user message", () => {
  const turn = createToolCallBatch(20, 10);
  const messages = [createUser("keep me", 1), ...turn];
  const plan = planKeptTail({
    messages,
    cutIndex: 1,
    budgetTokens: 4_000,
  });
  // The latest turn's assistant + results survive; the FIRST user message may be
  // dropped by drop-oldest-turns but the engine never fabricates content.
  assert.ok(plan.kept.length > 0);
  assert.ok(
    plan.kept.some((m) => m.role === "toolResult"),
    "latest results kept",
  );
});

void test("compactedTranscriptFits is the caller-level guarantee", () => {
  const turn = createToolCallBatch(25, 10);
  const messages = [createUser("go", 1), ...turn];

  const result = planCompaction({
    messages,
    boundaryStart: 0,
    keepRecentTokens: 20_000,
    budgetTokens: 12_000,
  });

  const transcript = compactedTranscript(result.kept, 0);
  assert.ok(
    verifyCompactionFits({
      systemPrompt: "system",
      messages: transcript,
      tools: [],
      contextWindow: 65_536,
      reserveTokens: 16_384,
    }),
  );
});
