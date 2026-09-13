/**
 * Context window token estimation utilities.
 *
 * We use the same conservative heuristic as pi-coding-agent: tokens ≈ chars / 4,
 * with script-aware adjustments so CJK/Hangul/Kana text (1 char ≈ 1 token) and
 * emoji (astral plane) are not undercounted ~4-8×. When we have
 * provider-reported `usage` for the last assistant turn, we use it as an anchor
 * because it already reflects prompt caching and provider-side tokenization.
 */

import type { Tool, Usage } from "@earendil-works/pi-ai";
import type { AgentMessage, AgentState } from "@earendil-works/pi-agent-core";
import { createConvertToLlm } from "../messages/convert-to-llm.js";

const CHARS_PER_TOKEN = 4;
/** Per-message framing/model-format overhead (role markers, delimiters, etc.). */
const MESSAGE_FRAMING_TOKENS = 4;
/** Default image token cost when no data is available (~1200 tokens). */
const IMAGE_FALLBACK_TOKENS = 1200;
/** Upper bound for a single image token estimate (very large screenshots). */
const IMAGE_MAX_TOKENS = 4096;

/** Is this code point a CJK ideograph / Hangul / Kana (≈1 token per char)? */
function isCjkLike(codePoint: number): boolean {
  // CJK Unified Ideographs + Ext A
  if (codePoint >= 0x3400 && codePoint <= 0x4dbf) return true;
  if (codePoint >= 0x4e00 && codePoint <= 0x9fff) return true;
  // CJK Compatibility Ideographs
  if (codePoint >= 0xf900 && codePoint <= 0xfaff) return true;
  // CJK Ext B–H
  if (codePoint >= 0x20000 && codePoint <= 0x2fa1f) return true;
  // Hiragana / Katakana (incl. halfwidth)
  if (codePoint >= 0x3040 && codePoint <= 0x30ff) return true;
  if (codePoint >= 0xff66 && codePoint <= 0xff9f) return true;
  // Hangul syllables + Jamo + compatibility Jamo
  if (codePoint >= 0x1100 && codePoint <= 0x11ff) return true;
  if (codePoint >= 0x3130 && codePoint <= 0x318f) return true;
  if (codePoint >= 0xac00 && codePoint <= 0xd7af) return true;
  return false;
}

/**
 * Script-aware token estimate:
 * - CJK-like chars ≈ 1 token each;
 * - astral-plane (emoji, rare scripts) ≈ 2 tokens each;
 * - everything else ≈ 1 token per 4 chars (Pi's chars/4 heuristic).
 */
export function estimateTextTokens(text: string): number {
  let tokens = 0;
  let latinChars = 0;

  for (const ch of text) {
    const codePoint = ch.codePointAt(0) ?? 0;
    if (codePoint > 0xffff) {
      // Astral plane: emoji mostly; CJK Ext B+ handled in isCjkLike.
      tokens += isCjkLike(codePoint) ? 1 : 2;
    } else if (isCjkLike(codePoint)) {
      tokens += 1;
    } else {
      latinChars += 1;
    }
  }

  return tokens + Math.ceil(latinChars / CHARS_PER_TOKEN);
}

/** Size-aware image token estimate, floored at the default, capped for sanity. */
function estimateImageTokens(block: { data?: string }): number {
  const data = block.data;
  if (typeof data !== "string" || data.length === 0) {
    return IMAGE_FALLBACK_TOKENS;
  }
  // Base64 inflates raw bytes ~1.33×. Vision APIs bill by pixel area (~85–1105
  // tokens for typical screenshots). Use a conservative size proxy so a few
  // large captures can't dominate the estimate, but small icons don't vanish.
  return Math.min(
    IMAGE_MAX_TOKENS,
    Math.max(IMAGE_FALLBACK_TOKENS, Math.ceil(data.length / 32)),
  );
}

/**
 * Total tokens that count against the model's context window.
 *
 * For providers with prompt caching (e.g. Anthropic), `usage.input` may exclude
 * cached prompt tokens. `cacheRead`/`cacheWrite` still count towards context.
 */
export function calculateContextTokens(u: Usage): number {
  return u.totalTokens || u.input + u.output + u.cacheRead + u.cacheWrite;
}

export function estimateMessageTokens(message: AgentMessage): number {
  if (message.role === "artifact" || message.role === "archivedMessages") {
    // UI-only, never sent to the model.
    return 0;
  }

  if (message.role === "compactionSummary") {
    return estimateTextTokens(message.summary) + MESSAGE_FRAMING_TOKENS;
  }

  if (message.role === "user" || message.role === "user-with-attachments") {
    let tokens = MESSAGE_FRAMING_TOKENS;
    const content = message.content;
    if (typeof content === "string") {
      tokens += estimateTextTokens(content);
    } else if (Array.isArray(content)) {
      for (const block of content) {
        if (block.type === "text") tokens += estimateTextTokens(block.text);
        if (block.type === "image") tokens += estimateImageTokens(block);
      }
    }
    // Attachments are forwarded as image/text content by convertToLlm.
    if (message.role === "user-with-attachments" && message.attachments) {
      for (const attachment of message.attachments) {
        if (attachment.type === "image") {
          tokens += estimateImageTokens({ data: attachment.content });
        } else if (attachment.extractedText) {
          tokens += estimateTextTokens(attachment.extractedText);
        }
      }
    }
    return tokens;
  }

  if (message.role === "assistant") {
    let tokens = MESSAGE_FRAMING_TOKENS;
    for (const block of message.content) {
      if (block.type === "text") tokens += estimateTextTokens(block.text);
      else if (block.type === "thinking")
        tokens += estimateTextTokens(block.thinking);
      else if (block.type === "toolCall") {
        tokens += estimateTextTokens(block.name);
        try {
          tokens += estimateTextTokens(JSON.stringify(block.arguments));
        } catch {
          // ignore unstringifiable args
        }
      }
    }
    return tokens;
  }

  if (message.role === "toolResult") {
    let tokens = MESSAGE_FRAMING_TOKENS;
    for (const block of message.content) {
      if (block.type === "text") tokens += estimateTextTokens(block.text);
      if (block.type === "image") tokens += estimateImageTokens(block);
    }
    return tokens;
  }

  // Unknown custom message types: ignore.
  return 0;
}

/** Token cost of tool schemas (name + description + parameters JSON). */
export function estimateToolsTokens(tools?: readonly Tool[]): number {
  if (!tools || tools.length === 0) return 0;

  let chars = 0;
  for (const tool of tools) {
    chars += tool.name.length + tool.description.length;
    try {
      chars += JSON.stringify(tool.parameters ?? {}).length;
    } catch {
      // ignore unstringifiable schema
    }
  }
  return Math.ceil(chars / CHARS_PER_TOKEN) + tools.length * 4;
}

/**
 * Estimated tokens for one full model request: system prompt + messages +
 * tool schemas + (optionally) reserved output tokens. Used by pre-flight
 * guards to decide whether a request can fit before dispatching it.
 */
export function estimateRequestTokens(args: {
  systemPrompt?: string;
  messages: readonly AgentMessage[];
  tools?: readonly Tool[];
  maxOutputTokens?: number;
}): number {
  let total = estimateTextTokens(args.systemPrompt ?? "");
  for (const message of args.messages) {
    total += estimateMessageTokens(message);
  }
  total += estimateToolsTokens(args.tools);
  if (args.maxOutputTokens && args.maxOutputTokens > 0) {
    total += args.maxOutputTokens;
  }
  return total;
}

export type ContextTokenEstimate = {
  /** Estimated total tokens in context (system prompt + messages) */
  totalTokens: number;
  /** The last provider usage we anchored on (for debug display) */
  lastUsage: Usage | null;
};
export type EffectiveContextTokenEstimate = {
  /** Raw estimate over the full persisted message list (unshaped). */
  rawTokens: number;
  /** Request-facing estimate after the same shaping convertToLlm applies. */
  effectiveTokens: number;
  /** The last provider usage we anchored on (for debug display). */
  lastUsage: Usage | null;
};

/**
 * Estimate the token cost of the *next* request, mirroring the exact shaping
 * `createConvertToLlm` applies (older large tool results → previews, archived
 * messages dropped, compaction summaries → user messages) plus tool schemas.
 *
 * `effectiveTokens` is what actually gets sent to the provider (before framing
 * differences), so pre-flight guards and the status meter should use it;
 * `rawTokens` reflects the full persisted history for diagnostics.
 */
export function estimateEffectiveRequestTokens(args: {
  state: Pick<AgentState, "systemPrompt" | "messages" | "model">;
  tools?: readonly Tool[];
  contextWindow?: number;
  maxOutputTokens?: number;
}): EffectiveContextTokenEstimate {
  const { state, tools, contextWindow, maxOutputTokens } = args;

  const raw = estimateContextTokens(state);

  // Same normalization + shaping path as convertToLlm (createConvertToLlm is
  // built with the same getContextWindow wiring the Agent uses).
  const convertToLlm = createConvertToLlm({
    getContextWindow: () => contextWindow,
  });
  const llmMessages = convertToLlm(state.messages);

  let effectiveTokens = estimateTextTokens(state.systemPrompt);
  for (const message of llmMessages) {
    effectiveTokens += estimateMessageTokens(message as AgentMessage);
  }
  effectiveTokens += estimateToolsTokens(tools);
  if (maxOutputTokens && maxOutputTokens > 0) {
    effectiveTokens += maxOutputTokens;
  }

  return {
    rawTokens: raw.totalTokens,
    effectiveTokens,
    lastUsage: raw.lastUsage,
  };
}

export function estimateContextTokens(
  state: Pick<AgentState, "systemPrompt" | "messages" | "model">,
): ContextTokenEstimate {
  const messages = state.messages;

  let lastUsage: Usage | null = null;
  let lastUsageIndex: number | null = null;
  let lastUsageTimestamp = 0;
  let anchorProvider: string | undefined;
  let anchorModel: string | undefined;

  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg === undefined) continue;
    if (msg.role !== "assistant") continue;
    if (msg.stopReason === "error" || msg.stopReason === "aborted") continue;

    const t = calculateContextTokens(msg.usage);
    if (t > 0) {
      lastUsage = msg.usage;
      lastUsageIndex = i;
      lastUsageTimestamp = msg.timestamp;
      anchorProvider = msg.provider;
      anchorModel = msg.model;
      break;
    }
  }

  let lastCompactionTimestamp = 0;
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (msg === undefined) continue;
    if (msg.role === "compactionSummary") {
      lastCompactionTimestamp = msg.timestamp;
      break;
    }
  }

  // A usage anchor from a different model must not be trusted: switching the
  // active model resets how the provider charges/measures context.
  const activeModel = state.model;
  const modelChanged =
    lastUsage !== null &&
    activeModel != null &&
    (anchorProvider !== activeModel.provider || anchorModel !== activeModel.id);

  const usageIsStale =
    (lastUsage !== null && lastCompactionTimestamp > lastUsageTimestamp) ||
    modelChanged;

  let totalTokens = 0;
  if (lastUsage && lastUsageIndex !== null && !usageIsStale) {
    totalTokens = calculateContextTokens(lastUsage);
    for (let i = lastUsageIndex + 1; i < messages.length; i++) {
      const msg = messages[i];
      if (msg !== undefined) {
        totalTokens += estimateMessageTokens(msg);
      }
    }
  } else {
    // No reliable usage signal (or it became stale after /compact or a model
    // switch). Estimate from scratch.
    totalTokens = estimateTextTokens(state.systemPrompt);
    for (const m of messages) {
      totalTokens += estimateMessageTokens(m);
    }
  }

  return { totalTokens, lastUsage };
}
