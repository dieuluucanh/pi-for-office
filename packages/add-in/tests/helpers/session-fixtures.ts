/** Shared session fixtures (typed, so tests need no object-literal casts). */

import type { SessionData, SessionMetadata } from "../../src/storage/local/types.ts";

export const TEST_MODEL: SessionData["model"] = {
  id: "test-model",
  name: "Test Model",
  api: "openai-completions",
  provider: "test",
  baseUrl: "https://example.invalid/v1",
  reasoning: false,
  input: ["text"],
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
  contextWindow: 200_000,
  maxTokens: 4096,
};

const EMPTY_USAGE: SessionMetadata["usage"] = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
};

export function userMessage(text: string): SessionData["messages"][number] {
  return {
    role: "user",
    content: [{ type: "text", text }],
    timestamp: 1,
  };
}

export function assistantMessage(text: string): SessionData["messages"][number] {
  return {
    role: "assistant",
    api: "openai-completions",
    provider: "test",
    model: "test-model",
    content: [{ type: "text", text }],
    usage: {
      ...EMPTY_USAGE,
      input: 1,
      output: 1,
      totalTokens: 2,
      cost: { ...EMPTY_USAGE.cost },
    },
    stopReason: "stop",
    timestamp: 2,
  };
}

export function sessionRecord(args: {
  id: string;
  lastModified: string;
  title?: string;
  messages?: SessionData["messages"];
}): SessionData {
  return {
    id: args.id,
    title: args.title ?? args.id,
    model: TEST_MODEL,
    thinkingLevel: "off",
    messages: args.messages ?? [userMessage(`${args.title ?? args.id} prompt`)],
    createdAt: args.lastModified,
    lastModified: args.lastModified,
  };
}

export function sessionMetadata(args: {
  id: string;
  lastModified: string;
  title?: string;
  messageCount?: number;
}): SessionMetadata {
  return {
    id: args.id,
    title: args.title ?? args.id,
    createdAt: args.lastModified,
    lastModified: args.lastModified,
    messageCount: args.messageCount ?? 1,
    usage: { ...EMPTY_USAGE, cost: { ...EMPTY_USAGE.cost } },
    thinkingLevel: "off",
    preview: "",
  };
}
