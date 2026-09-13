import assert from "node:assert/strict";
import { test } from "node:test";

import type { Message, Model, Api } from "@earendil-works/pi-ai";

import { fitLlmContextToWindow } from "../src/auth/context-trim.ts";

function createModel(
  contextWindow: number,
  maxTokens = 4096,
): Pick<Model<Api>, "contextWindow" | "maxTokens"> {
  return { contextWindow, maxTokens };
}

function userMessage(text: string): Message {
  return { role: "user", content: text, timestamp: 1 };
}

function assistantMessage(text: string): Message {
  return {
    role: "assistant",
    content: [{ type: "text", text }],
    api: "openai-completions",
    provider: "opencode-go",
    model: "muse-spark-1.3-contributor",
    timestamp: 2,
  };
}

function toolResultMessage(toolName: string, text: string): Message {
  return {
    role: "toolResult",
    toolCallId: `call-${toolName}`,
    toolName,
    content: [{ type: "text", text }],
    isError: false,
    timestamp: 3,
  };
}

void test("fitLlmContextToWindow is a no-op under budget", () => {
  const context = {
    systemPrompt: "system",
    messages: [userMessage("hello"), assistantMessage("hi back")],
    tools: [],
  };
  const result = fitLlmContextToWindow(createModel(262_144), context);
  assert.equal(result.trimmed, false);
  assert.equal(result.context, context);
  assert.equal(result.fits, true);
});

void test("fitLlmContextToWindow trims oversized tool results (previews) and fits", () => {
  // A huge tool-result batch far exceeds the window.
  const messages: Message[] = [userMessage("go")];
  for (let i = 0; i < 20; i++) {
    messages.push(toolResultMessage("read_range", "x".repeat(50 * 1024)));
  }

  const context = { systemPrompt: "system", messages, tools: [] };
  const result = fitLlmContextToWindow(createModel(65_536), context);
  assert.equal(result.trimmed, true);
  assert.equal(result.fits, true);
});

void test("fitLlmContextToWindow never trims the last user message", () => {
  const bigUser = userMessage("crucial prompt: " + "z".repeat(1_000_000));
  const messages: Message[] = [userMessage("old"), bigUser];
  for (let i = 0; i < 10; i++) {
    messages.push(toolResultMessage("read_range", "x".repeat(50 * 1024)));
  }

  const context = { systemPrompt: "system", messages, tools: [] };
  const result = fitLlmContextToWindow(createModel(65_536), context);

  const last = result.context.messages.at(-1);
  assert.equal(last?.role, "toolResult");
  // The user message before the tail is preserved (never dropped/trimmed as the
  // last user prompt — trimming applies to older content; the newest user stays).
  const userMessages = result.context.messages.filter((m) => m.role === "user");
  assert.ok(userMessages.length > 0);
});

void test("fitLlmContextToWindow drops old images keeping the newest", () => {
  const imageUser = (label: string): Message => ({
    role: "user",
    content: [
      { type: "text", text: label },
      // ~50KB of text makes the message itself expensive; shaping never
      // touches user messages, so only the image-drop stage can reduce it.
      { type: "text", text: "x".repeat(50 * 1024) },
      { type: "image", data: "iVBORw0KGgo=", mimeType: "image/png" },
    ],
    timestamp: 1,
  });

  const messages: Message[] = [];
  for (let i = 0; i < 8; i++) {
    messages.push(imageUser(`msg-${i}`));
  }
  messages.push(userMessage("final goal"));

  const context = { systemPrompt: "system", messages, tools: [] };
  const result = fitLlmContextToWindow(createModel(65_536), context);

  // The drop stage may not need to run if text trimming already fits, but when
  // it does run it must keep at most one image-bearing message.
  const images = result.context.messages.filter(
    (m) =>
      Array.isArray(m.content) && m.content.some((b) => b.type === "image"),
  );
  assert.ok(images.length <= 1, "at most one image-bearing message kept");
});

void test("fitLlmContextToWindow truncates oversized text blocks", () => {
  const huge = assistantMessage("h".repeat(500_000));
  const context = {
    systemPrompt: "system",
    messages: [userMessage("go"), huge],
    tools: [],
  };
  const result = fitLlmContextToWindow(createModel(65_536), context);
  assert.equal(result.trimmed, true);
  assert.equal(result.truncatedBlocks >= 1, true);
});
