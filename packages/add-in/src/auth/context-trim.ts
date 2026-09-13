/**
 * Provider-boundary context trim (defense in depth).
 *
 * Runs just before dispatching a request: if the shaped LLM context still
 * exceeds the model window (minus a response margin), defensively trim it so an
 * over-budget request is never sent — even if state-level guards missed
 * something. Also protects the compaction summarizer call, which flows through
 * the same stream function.
 *
 * Rules:
 * - never removes the last user message;
 * - oldest tool results become previews (reuses the request-time shaping);
 * - images beyond the most recent one are dropped;
 * - the largest text blocks are truncated;
 * - an explicit `[context trimmed to fit the model window]` note is inserted.
 */

import type {
  Api,
  AssistantMessage,
  Context,
  Message,
  Model,
  ToolResultMessage,
  UserMessage,
} from "@earendil-works/pi-ai";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import { shapeToolResultsForLlm } from "../messages/tool-result-shaping.js";
import { estimateRequestTokens } from "../utils/context-tokens.js";

export interface ContextTrimResult {
  context: Context;
  trimmed: boolean;
  /** Sanity: whether the trimmed context now fits. */
  fits: boolean;
  droppedImages: number;
  truncatedBlocks: number;
}

const RESPONSE_MARGIN_TOKENS = 2048;
const TEXT_TRIM_MAX_CHARS = 120_000;

export function fitLlmContextToWindow(
  model: Pick<Model<Api>, "contextWindow" | "maxTokens">,
  context: Context,
): ContextTrimResult {
  const contextWindow = model.contextWindow || 200_000;
  if (!Number.isFinite(contextWindow) || contextWindow <= 0) {
    return {
      context,
      trimmed: false,
      fits: true,
      droppedImages: 0,
      truncatedBlocks: 0,
    };
  }

  const maxOutput =
    model.maxTokens && model.maxTokens > 0
      ? model.maxTokens
      : RESPONSE_MARGIN_TOKENS;
  const budget = Math.max(
    RESPONSE_MARGIN_TOKENS,
    contextWindow - maxOutput - RESPONSE_MARGIN_TOKENS,
  );

  const countTokens = (c: Context): number => {
    const requestArgs: Parameters<typeof estimateRequestTokens>[0] = {
      messages: c.messages as AgentMessage[],
    };
    if (c.systemPrompt !== undefined) requestArgs.systemPrompt = c.systemPrompt;
    if (c.tools !== undefined) requestArgs.tools = c.tools;
    return estimateRequestTokens(requestArgs);
  };

  let trimmedCtx: Context = context;
  if (countTokens(trimmedCtx) <= budget) {
    return {
      context,
      trimmed: false,
      fits: true,
      droppedImages: 0,
      truncatedBlocks: 0,
    };
  }

  trimmedCtx = shapenToolResults(context);

  let droppedImages = 0;
  if (countTokens(trimmedCtx) > budget) {
    const dropped = dropImagesKeepNewest(trimmedCtx);
    droppedImages = dropped.droppedImages;
    trimmedCtx = dropped.context;
  }

  let truncatedBlocks = 0;
  if (countTokens(trimmedCtx) > budget) {
    const truncated = truncateLargestTextBlocks(trimmedCtx);
    truncatedBlocks = truncated.truncatedBlocks;
    trimmedCtx = truncated.context;
  }

  const fits = countTokens(trimmedCtx) <= budget;
  const withNote: Context = fits
    ? trimmedCtx
    : {
        ...trimmedCtx,
        systemPrompt: `${trimmedCtx.systemPrompt ?? ""}\n\n[context trimmed to fit the model window]`,
      };

  return {
    context: withNote,
    trimmed: true,
    fits,
    droppedImages,
    truncatedBlocks,
  };
}

/** Preview older tool results using the standard request-time shaping. */
function shapenToolResults(context: Context): Context {
  const shaped = shapeToolResultsForLlm(context.messages as AgentMessage[], {
    recentToolResultsToKeep: 4,
  }) as Message[];
  if (shaped === context.messages) return context;
  return { ...context, messages: shaped };
}

function isImageMessage(msg: Message): boolean {
  return (
    (msg.role === "user" || msg.role === "toolResult") &&
    Array.isArray(msg.content) &&
    msg.content.some((block) => block.type === "image")
  );
}

/** Drop image blocks from all but the most recent image-bearing message. */
function dropImagesKeepNewest(context: Context): {
  context: Context;
  droppedImages: number;
} {
  const messages = context.messages;
  const imageIndexes = messages
    .map((m, i) => (isImageMessage(m) ? i : -1))
    .filter((i) => i >= 0);
  if (imageIndexes.length === 0) {
    return { context, droppedImages: 0 };
  }

  const keepIndex = imageIndexes.at(-1) ?? -1;
  let droppedImages = 0;
  const next: Message[] = messages.map((m, i) => {
    if (i === keepIndex || !isImageMessage(m)) return m;
    if (m.role === "toolResult") {
      const content = m.content.filter((block) => {
        if (block.type === "image") {
          droppedImages += 1;
          return false;
        }
        return true;
      });
      const trimmed: ToolResultMessage = { ...m, content };
      return trimmed;
    }
    if (m.role === "user") {
      const content = Array.isArray(m.content)
        ? m.content.filter((block) => {
            if (block.type === "image") {
              droppedImages += 1;
              return false;
            }
            return true;
          })
        : m.content;
      const trimmed: UserMessage = { ...m, content };
      return trimmed;
    }
    return m;
  });

  return { context: { ...context, messages: next }, droppedImages };
}

function trimTextBlock<TBlock extends { type: string; text: string }>(
  block: TBlock,
  state: { changed: boolean; truncatedBlocks: number },
): TBlock {
  if (block.text.length <= TEXT_TRIM_MAX_CHARS) return block;
  state.changed = true;
  state.truncatedBlocks += 1;
  return {
    ...block,
    text: block.text.slice(0, TEXT_TRIM_MAX_CHARS) + "\n…[context trimmed]…",
  };
}

/** Truncate the largest text blocks (never the last user message) to fit. */
function truncateLargestTextBlocks(context: Context): {
  context: Context;
  truncatedBlocks: number;
} {
  const messages = context.messages;
  const lastUserIndex = findLastUserIndex(messages);

  let truncatedBlocks = 0;
  let changed = false;
  const state = {
    get changed(): boolean {
      return changed;
    },
    set changed(value: boolean) {
      changed = value;
    },
    get truncatedBlocks(): number {
      return truncatedBlocks;
    },
    set truncatedBlocks(value: number) {
      truncatedBlocks = value;
    },
  };

  const next: Message[] = messages.map((m, i) => {
    if (i === lastUserIndex) return m; // never trim the latest user prompt

    if (m.role === "user") {
      const content = m.content;
      if (typeof content === "string") {
        if (content.length <= TEXT_TRIM_MAX_CHARS) return m;
        const trimmed = trimTextBlock({ type: "text", text: content }, state);
        const trimmedUser: UserMessage = { ...m, content: trimmed.text };
        void trimmed;
        return trimmedUser;
      }
      if (!Array.isArray(content)) return m;
      const blocks = content.map((block) =>
        trimTextBlock(block as { type: string; text: string }, state),
      );
      const trimmedUser: UserMessage = {
        ...m,
        content: blocks as UserMessage["content"],
      };
      return trimmedUser;
    }

    if (m.role === "assistant") {
      const blocks = m.content.map((block) =>
        trimTextBlock(block as { type: string; text: string }, state),
      );
      const trimmedAssistant: AssistantMessage = {
        ...m,
        content: blocks as AssistantMessage["content"],
      };
      return trimmedAssistant;
    }

    if (m.role === "toolResult") {
      const blocks = m.content.map((block) => {
        if (block.type === "text" && block.text.length > TEXT_TRIM_MAX_CHARS) {
          state.changed = true;
          state.truncatedBlocks += 1;
          return {
            ...block,
            text:
              block.text.slice(0, TEXT_TRIM_MAX_CHARS) +
              "\n…[context trimmed]…",
          };
        }
        return block;
      });
      const trimmedResult: ToolResultMessage = { ...m, content: blocks };
      return trimmedResult;
    }

    return m;
  });

  return {
    context: changed ? { ...context, messages: next } : context,
    truncatedBlocks,
  };
}

function findLastUserIndex(messages: readonly Message[]): number {
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i]?.role === "user") return i;
  }
  return -1;
}
