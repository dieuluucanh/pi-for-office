import assert from "node:assert/strict";
import { test } from "node:test";

import type { Tool } from "@earendil-works/pi-ai";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { AgentState } from "@earendil-works/pi-agent-core";
import type { ArchivedMessagesMessage } from "../src/messages/archived-history.ts";
import type {
  ArtifactMessage,
  UserMessageWithAttachments,
} from "../src/messages/attachments.ts";

import {
  estimateContextTokens,
  estimateEffectiveRequestTokens,
  estimateMessageTokens,
  estimateRequestTokens,
  estimateTextTokens,
  estimateToolsTokens,
} from "../src/utils/context-tokens.ts";

function createUserMessage(text: string, timestamp: number): AgentMessage {
  return { role: "user", content: text, timestamp };
}

function createToolResult(text: string, timestamp: number): AgentMessage {
  return {
    role: "toolResult",
    toolCallId: `call-${timestamp}`,
    toolName: "read_range",
    content: [{ type: "text", text }],
    isError: false,
    // ToolResultMessage's `details` is optional; absence is fine.
    timestamp,
  };
}

function createAssistantWithUsage(
  timestamp: number,
  input: number,
): AgentMessage {
  return {
    role: "assistant",
    content: [{ type: "text", text: "ok" }],
    api: "openai-completions",
    provider: "opencode-go",
    model: "muse-spark-1.3-contributor",
    usage: {
      input,
      output: 10,
      cacheRead: 0,
      cacheWrite: 0,
      totalTokens: input + 10,
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
    },
    stopReason: "stop",
    timestamp,
  };
}

function createCompactionSummary(
  summary: string,
  timestamp: number,
): AgentMessage {
  return {
    role: "compactionSummary",
    summary,
    tokensBefore: 100,
    timestamp,
  };
}

const MINIMAL_STATE = (
  messages: AgentMessage[],
): Pick<AgentState, "systemPrompt" | "messages" | "model"> => {
  const model: NonNullable<AgentState["model"]> = {
    id: "muse-spark-1.3-contributor",
    name: "Muse Spark 1.3 Contributor",
    api: "openai-completions",
    provider: "opencode-go",
    baseUrl: "https://opencode.ai",
    reasoning: false,
    input: ["text"],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow: 262_144,
    maxTokens: 4096,
  };
  return { systemPrompt: "system", messages, model };
};

void test("estimateTextTokens counts CJK/Hangul/Kana chars as ~1 token each", () => {
  // 4 CJK chars ≈ 4 tokens (not 4/4 = 1).
  const cjk = estimateTextTokens("中文中文");
  assert.ok(cjk >= 4, `expected >=4 tokens for 4 CJK chars, got ${cjk}`);

  // Hangul syllable block.
  const hangul = estimateTextTokens("한글테스트");
  assert.ok(
    hangul >= 5,
    `expected >=5 tokens for 5 Hangul chars, got ${hangul}`,
  );

  // Kana.
  const kana = estimateTextTokens("ひらがな");
  assert.ok(kana >= 4, `expected >=4 tokens for 4 kana chars, got ${kana}`);
});

void test("estimateTextTokens treats ASCII with chars/4 and emoji conservatively", () => {
  // "hello" = 5 chars → ceil(5/4) = 2.
  assert.equal(estimateTextTokens("hello"), 2);

  // Emoji (astral plane) costs at least 2 tokens each, never 0.25/char.
  const withEmoji = estimateTextTokens("hi 😀");
  assert.ok(withEmoji > estimateTextTokens("hi "), "emoji should add tokens");
});

void test("estimateMessageTokens ignores UI-only roles", () => {
  const archived: ArchivedMessagesMessage = {
    role: "archivedMessages",
    archivedMessages: [createUserMessage("X".repeat(10_000), 1)],
    archivedChatMessageCount: 1,
    timestamp: 2,
  };

  const artifact: ArtifactMessage = {
    role: "artifact",
    action: "create",
    filename: "a.txt",
    content: "X".repeat(10_000),
    timestamp: "2026-01-01T00:00:00.000Z",
  };

  assert.equal(estimateMessageTokens(archived), 0);
  assert.equal(estimateMessageTokens(artifact), 0);
});

void test("estimateMessageTokens counts attachments and images by data", () => {
  const withAttachment: UserMessageWithAttachments = {
    role: "user-with-attachments",
    content: [{ type: "text", text: "look" }],
    attachments: [
      {
        id: "a1",
        type: "image",
        fileName: "shot.png",
        mimeType: "image/png",
        size: 1000,
        content: "x".repeat(1024),
      },
    ],
    timestamp: 1,
  };

  const plain = estimateMessageTokens(createUserMessage("look", 1));
  const withImage = estimateMessageTokens(withAttachment);
  assert.ok(withImage > plain, "image attachment should add tokens");
});

void test("estimateRequestTokens totals system + messages + tools + output", () => {
  const messages = [createUserMessage("hello", 1)];
  const tools: Tool[] = [
    {
      name: "read_range",
      description: "Read a range",
      parameters: { type: "object", properties: { ref: { type: "string" } } },
    },
  ];

  const withoutOutput = estimateRequestTokens({
    systemPrompt: "system",
    messages,
    tools,
  });
  const withOutput = estimateRequestTokens({
    systemPrompt: "system",
    messages,
    tools,
    maxOutputTokens: 4096,
  });

  assert.ok(withoutOutput > 0);
  assert.equal(withOutput, withoutOutput + 4096);
  assert.equal(estimateToolsTokens(undefined), 0);
  assert.equal(estimateToolsTokens([]), 0);
});

void test("estimateContextTokens uses provider usage anchor and invalidates on model switch", () => {
  const anchor = createAssistantWithUsage(10, 1000);
  const trailing = createUserMessage("more".repeat(200), 11); // ~1k chars

  const anchored = estimateContextTokens(
    MINIMAL_STATE([createUserMessage("hi", 1), anchor, trailing]),
  );
  // Anchor (1010) + trailing estimate (>=200/4=50) — NOT system + all messages from scratch.
  assert.ok(
    anchored.totalTokens > 1000,
    "anchor + trailing should exceed anchor alone",
  );

  // Model switch invalidates the anchor: from-scratch estimate includes system
  // prompt + every message — far smaller than the provider-reported 1010.
  const switchedState = MINIMAL_STATE([
    createUserMessage("hi", 1),
    anchor,
    trailing,
  ]);
  switchedState.model = {
    ...switchedState.model,
    id: "different-model",
  };
  const afterSwitch = estimateContextTokens(switchedState);
  assert.ok(
    afterSwitch.totalTokens < 1000,
    `model switch should invalidate anchor (drop the 1010 usage), got ${afterSwitch.totalTokens}`,
  );
});

void test("estimateContextTokens invalidates stale usage after compaction", () => {
  const anchor = createAssistantWithUsage(10, 1000);
  const compaction = createCompactionSummary("summary here", 20); // newer than anchor
  const trailing = createUserMessage("more", 21);

  const stale = estimateContextTokens(
    MINIMAL_STATE([createUserMessage("hi", 1), anchor, compaction, trailing]),
  );
  // From-scratch estimate of system + 4 messages; the stale 1010 usage must not
  // be included.
  assert.ok(
    stale.totalTokens < 1000,
    `stale anchor should drop to from-scratch (exclude 1010 usage), got ${stale.totalTokens}`,
  );
});

void test("estimateEffectiveRequestTokens reflects shaping (raw > effective for long tool batches)", () => {
  // A batch of many large tool results: shaping keeps the last 6 verbatim and
  // previews older ones, so the request-facing estimate must be well below raw.
  const messages: AgentMessage[] = [createUserMessage("go", 1)];
  const bigText = "d".repeat(2_000); // > maxCharsBeforeCompaction (1200)
  for (let i = 0; i < 12; i++) {
    messages.push(createToolResult(bigText, 10 + i));
  }

  const result = estimateEffectiveRequestTokens({
    state: MINIMAL_STATE(messages),
    contextWindow: 262_144,
  });

  assert.ok(result.rawTokens > 0);
  assert.ok(result.effectiveTokens > 0);
  assert.ok(
    result.effectiveTokens < result.rawTokens,
    `shaping should reduce the estimate: raw=${result.rawTokens} effective=${result.effectiveTokens}`,
  );
  assert.ok(result.lastUsage === null, "no usage anchor in this state");
});

void test("estimateEffectiveRequestTokens accepts output budget", () => {
  const messages = [createUserMessage("hello", 1)];
  const base = estimateEffectiveRequestTokens({
    state: MINIMAL_STATE(messages),
    contextWindow: 262_144,
  });
  const withOutput = estimateEffectiveRequestTokens({
    state: MINIMAL_STATE(messages),
    contextWindow: 262_144,
    maxOutputTokens: 4096,
  });
  assert.equal(withOutput.effectiveTokens, base.effectiveTokens + 4096);
});
