/**
 * Pure compaction engine — no DOM, no LLM calls.
 *
 * Guarantees: compaction always makes progress. The kept tail is *bounded* to a
 * token budget (never the whole trailing tool batch), oversized content in the
 * kept tail is trimmed, and the caller verifies the assembled transcript fits
 * before committing.
 *
 * Key difference from the old `findCutIndex`: cuts only at valid turn
 * boundaries (user/assistant), never backs off across an entire tool batch.
 * When a huge trailing batch would be kept verbatim, the trim stage previews
 * its oversized results instead of keeping them whole.
 */

import type {
  AssistantMessage,
  Tool,
  ToolResultMessage,
  UserMessage,
} from "@earendil-works/pi-ai";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import { shapeToolResultsForLlm } from "../messages/tool-result-shaping.js";
import {
  estimateMessageTokens,
  estimateRequestTokens,
} from "../utils/context-tokens.js";

const DEFAULT_KEEP_VERBOSE_PAIRS = 3;
const TOOL_CALL_ARGS_MAX_CHARS = 1500;
const USER_TEXT_MAX_CHARS = 300_000;

export interface KeptTailPlan {
  kept: AgentMessage[];
  trimmedToolResults: number;
  truncatedToolCallArgs: number;
  droppedImages: number;
  droppedThinking: number;
  droppedTurns: number;
}

export interface CompactionFitArgs {
  systemPrompt?: string;
  messages: readonly AgentMessage[];
  tools?: readonly Tool[];
  contextWindow: number;
  reserveTokens: number;
}

/** Result of running compaction, so auto-compaction/recovery can verify progress. */
export interface CompactionOutcome {
  changed: boolean;
  reason: "summarized" | "tail-trimmed" | "nothing-to-compact" | "failed";
  tokensBefore: number;
  tokensAfter: number;
  keptCount: number;
  summarizedCount: number;
  droppedTurns?: number;
  trimmedToolResults?: number;
  truncatedToolCallArgs?: number;
  droppedImages?: number;
  droppedThinking?: number;
  errorMessage?: string;
}

export function estimateListTokens(messages: readonly AgentMessage[]): number {
  let total = 0;
  for (const message of messages) {
    total += estimateMessageTokens(message);
  }
  return total;
}

/** Valid positions to cut: user/assistant turns. Never a tool result. */
function isValidCutPointRole(role: AgentMessage["role"]): boolean {
  return (
    role === "user" || role === "user-with-attachments" || role === "assistant"
  );
}

/**
 * Select the oldest kept-tail boundary.
 *
 * Walk backwards from the newest message, accumulating estimated tokens. When
 * the budget is reached at index `i`, cut at the nearest valid cut point AT or
 * AFTER `i` (keeps less, summarizes the oversized batch) if one exists;
 * otherwise cut at the nearest valid point BEFORE `i` (the assistant that owns
 * the trailing tool batch) — the trim stage shrinks that batch afterwards.
 *
 * Never backs off across an entire tool batch to "keep it whole": a huge
 * trailing batch ends up either summarized (cut at next turn) or trimmed.
 */
export function selectCut(args: {
  messages: readonly AgentMessage[];
  boundaryStart: number;
  keepRecentTokens: number;
}): number {
  const { messages, boundaryStart, keepRecentTokens } = args;

  let accumulated = 0;
  for (let i = messages.length - 1; i >= boundaryStart; i--) {
    const message = messages[i];
    if (!message) continue;

    accumulated += estimateMessageTokens(message);
    if (accumulated < keepRecentTokens) continue;

    // Budget reached at i: prefer a valid cut at/after i (summarizes the batch).
    for (let j = i; j < messages.length; j++) {
      const candidate = messages[j];
      if (candidate && isValidCutPointRole(candidate.role)) {
        return j;
      }
    }
    // No valid point after i: i is inside the trailing tool batch of the last
    // assistant. Cut at the batch's assistant (nearest valid point before i).
    for (let j = i; j >= boundaryStart; j--) {
      const candidate = messages[j];
      if (candidate && isValidCutPointRole(candidate.role)) {
        return j;
      }
    }
    return boundaryStart;
  }

  // Total content is under the budget: nothing worth summarizing.
  return boundaryStart;
}

function isToolResult(message: AgentMessage): message is ToolResultMessage {
  return message.role === "toolResult";
}

function isAssistant(message: AgentMessage): message is AssistantMessage {
  return message.role === "assistant";
}

function isUserMessage(
  message: AgentMessage,
): message is
  | UserMessage
  | Extract<AgentMessage, { role: "user-with-attachments" }> {
  return message.role === "user" || message.role === "user-with-attachments";
}

/**
 * Replace oversized tool-call argument JSON with a short preview, so a single
 * assistant tool call can't blow past the kept-tail budget. Keeps the payload
 * small (truncated preview object) while marking that truncation happened.
 */
function truncateToolCallArgs(message: AgentMessage): AgentMessage {
  if (!isAssistant(message)) return message;

  let changed = false;
  const content = message.content.map((block) => {
    if (block.type !== "toolCall") return block;
    try {
      const serialized = JSON.stringify(block.arguments);
      if (serialized.length <= TOOL_CALL_ARGS_MAX_CHARS) return block;
      changed = true;
      const head = serialized.slice(0, TOOL_CALL_ARGS_MAX_CHARS);
      return {
        ...block,
        arguments: {
          __truncated: true,
          originalChars: serialized.length,
          preview: head,
        },
      };
    } catch {
      return block;
    }
  });

  return changed ? { ...message, content } : message;
}

/**
 * Drop old images (tool results) and old thinking blocks (assistant messages)
 * from the kept tail — they are the least useful for continuing the work and
 * the most expensive per token.
 */
function dropOldImagesAndThinking(messages: readonly AgentMessage[]): {
  messages: AgentMessage[];
  droppedImages: number;
  droppedThinking: number;
} {
  let droppedImages = 0;
  let droppedThinking = 0;

  const cleaned: AgentMessage[] = messages.map((message) => {
    if (isToolResult(message)) {
      const blocks = message.content.filter((block) => {
        if (block.type === "image") {
          droppedImages += 1;
          return false;
        }
        return true;
      });
      return { ...message, content: blocks };
    }
    if (isAssistant(message)) {
      const blocks = message.content.filter((block) => {
        if (block.type === "thinking") {
          droppedThinking += 1;
          return false;
        }
        return true;
      });
      return { ...message, content: blocks };
    }
    return message;
  });

  return { messages: cleaned, droppedImages, droppedThinking };
}

/**
 * Drop oldest kept turns until the tail fits the budget. Preserves tool-call/
 * result pairing and always keeps the latest user message and the latest
 * assistant+results cycle for the in-flight continuation.
 */
function dropOldestTurnsToFit(
  messages: readonly AgentMessage[],
  budgetTokens: number,
): { messages: AgentMessage[]; droppedTurns: number } {
  let droppedTurns = 0;
  let current = [...messages];

  while (current.length > 0 && estimateListTokens(current) > budgetTokens) {
    // Find the first turn group: user message, or the assistant+results cycle
    // that owns the leading tool results.
    const firstCut = current.findIndex(
      (m) => m !== undefined && isValidCutPointRole(m.role),
    );
    if (firstCut < 0 || firstCut >= current.length - 1) {
      break; // nothing to drop (only one turn left)
    }
    const start = firstCut;

    const end = findTurnGroupEnd(current, start);
    const next = current.slice(end + 1);
    if (next.length === 0) break; // dropping would empty the tail
    current = next;
    droppedTurns += 1;
  }

  return { messages: current, droppedTurns };
}

/** End index (inclusive) of the turn group that starts at `start`. */
function findTurnGroupEnd(
  messages: readonly AgentMessage[],
  start: number,
): number {
  // Group = one user/assistant message plus any tool results that follow it.
  let end = start;
  while (end + 1 < messages.length) {
    const next = messages[end + 1];
    if (next !== undefined && isToolResult(next)) {
      end += 1;
      continue;
    }
    break;
  }
  return end;
}

/**
 * If the kept tail is still over budget after all trims, truncate oversized
 * user/assistant text payloads with an explicit marker. The latest user message
 * is never dropped entirely, but its content can be cut to fit.
 */
function truncateOversizedText(
  messages: readonly AgentMessage[],
): AgentMessage[] {
  return messages.map((message) => {
    if (isUserMessage(message)) {
      const content = message.content;
      if (typeof content === "string" && content.length > USER_TEXT_MAX_CHARS) {
        return {
          ...message,
          content:
            content.slice(0, USER_TEXT_MAX_CHARS) +
            "\n[text truncated for context]",
        };
      }
      if (Array.isArray(content)) {
        let changed = false;
        const blocks = content.map((block) => {
          if (block.type !== "text") return block;
          if (block.text.length <= USER_TEXT_MAX_CHARS) return block;
          changed = true;
          return {
            ...block,
            text:
              block.text.slice(0, USER_TEXT_MAX_CHARS) +
              "\n[text truncated for context]",
          };
        });
        return changed ? { ...message, content: blocks } : message;
      }
      return message;
    }
    if (isAssistant(message)) {
      let changed = false;
      const blocks = message.content.map((block) => {
        if (block.type !== "text") return block;
        if (block.text.length <= USER_TEXT_MAX_CHARS) return block;
        changed = true;
        return {
          ...block,
          text:
            block.text.slice(0, USER_TEXT_MAX_CHARS) +
            "\n[text truncated for context]",
        };
      });
      return changed ? { ...message, content: blocks } : message;
    }
    return message;
  });
}

/**
 * Plan the kept tail: start from the cut index, then trim until it fits the
 * budget. The escalation ladder:
 *   1. preview oversized tool results (keep the last N pairs verbatim)
 *   2. truncate oversized tool-call argument JSON
 *   3. drop old images/thinking
 *   4. drop oldest turns (preserving the latest user + latest assistant cycle)
 *   5. truncate oversized text payloads with an explicit marker
 */
export function planKeptTail(args: {
  messages: readonly AgentMessage[];
  cutIndex: number;
  budgetTokens: number;
  keepVerbosePairs?: number | undefined;
}): KeptTailPlan {
  const {
    messages,
    cutIndex,
    budgetTokens,
    keepVerbosePairs = DEFAULT_KEEP_VERBOSE_PAIRS,
  } = args;

  const stats: KeptTailPlan = {
    kept: [],
    trimmedToolResults: 0,
    truncatedToolCallArgs: 0,
    droppedImages: 0,
    droppedThinking: 0,
    droppedTurns: 0,
  };

  let kept = messages.slice(cutIndex);
  if (kept.length === 0) return stats;

  // Level 1: preview oversized/older tool results. When even the default number
  // of verbatim pairs exceeds the budget, keep progressively fewer verbatim
  // (down to the single most recent result) so the tail always fits.
  if (estimateListTokens(kept) > budgetTokens) {
    let verbose = keepVerbosePairs;
    for (;;) {
      kept = shapeToolResultsForLlm(kept, {
        recentToolResultsToKeep: verbose,
      });
      stats.trimmedToolResults = kept.filter(
        (m) =>
          isToolResult(m) &&
          m.content.some(
            (b) =>
              b.type === "text" && b.text.startsWith("[Compacted tool result]"),
          ),
      ).length;
      if (estimateListTokens(kept) <= budgetTokens || verbose <= 1) break;
      verbose = Math.max(1, verbose - 1);
    }
  }
  if (estimateListTokens(kept) <= budgetTokens) {
    stats.kept = kept;
    return stats;
  }

  // Level 2: truncate oversized tool-call arguments.
  let changedArgs = 0;
  kept = kept.map((m) => {
    const next = truncateToolCallArgs(m);
    if (next !== m && isAssistant(m)) {
      changedArgs += m.content.filter((b) => b.type === "toolCall").length;
    }
    return next;
  });
  stats.truncatedToolCallArgs = changedArgs;
  if (estimateListTokens(kept) <= budgetTokens) {
    stats.kept = kept;
    return stats;
  }

  // Level 3: drop old images + thinking.
  const cleaned = dropOldImagesAndThinking(kept);
  kept = cleaned.messages;
  stats.droppedImages = cleaned.droppedImages;
  stats.droppedThinking = cleaned.droppedThinking;
  if (estimateListTokens(kept) <= budgetTokens) {
    stats.kept = kept;
    return stats;
  }

  // Level 4: drop oldest turns.
  const dropped = dropOldestTurnsToFit(kept, budgetTokens);
  stats.droppedTurns = dropped.droppedTurns;
  kept = dropped.messages;
  if (estimateListTokens(kept) <= budgetTokens) {
    stats.kept = kept;
    return stats;
  }

  // Level 5: truncate oversized text payloads (last resort, explicit markers).
  stats.kept = truncateOversizedText(kept);
  return stats;
}

/** Floor for keepRecentTokens halving escalation. */
export const MIN_KEEP_RECENT_TOKENS = 2_000;

export interface CompactionPlanResult {
  /** Cut index (index in `messages`) where the kept tail starts, or -1 when nothing to summarize. */
  cutIndex: number;
  /** The kept tail (already trimmed to fit `budgetTokens`). */
  kept: AgentMessage[];
  /** The plan details from the final trim pass. */
  plan: KeptTailPlan;
  /** keepRecentTokens that produced this plan. */
  keepRecentTokens: number;
  /** True when the kept tail fits the budget without further escalation. */
  fits: boolean;
}

/**
 * Combine cut selection with the trim ladder, escalating `keepRecentTokens`
 * (halving from the start value down to a floor) until the kept tail fits the
 * budget. Pure and deterministic — used by runCompactCommand and tested here.
 */
export function planCompaction(args: {
  messages: readonly AgentMessage[];
  boundaryStart: number;
  keepRecentTokens: number;
  budgetTokens: number;
  keepVerbosePairs?: number;
}): CompactionPlanResult {
  const {
    messages,
    boundaryStart,
    keepRecentTokens,
    budgetTokens,
    keepVerbosePairs,
  } = args;

  let keepRecent = Math.max(MIN_KEEP_RECENT_TOKENS, keepRecentTokens);
  let lastResult: CompactionPlanResult | null = null;

  while (keepRecent >= MIN_KEEP_RECENT_TOKENS) {
    const cutIndex = selectCut({
      messages,
      boundaryStart,
      keepRecentTokens: keepRecent,
    });
    const plan = planKeptTail({
      messages,
      cutIndex,
      budgetTokens,
      keepVerbosePairs,
    });

    const fits = estimateListTokens(plan.kept) <= budgetTokens;
    lastResult = {
      cutIndex,
      kept: plan.kept,
      plan,
      keepRecentTokens: keepRecent,
      fits,
    };
    if (fits) break;

    // Escalate: halve the recent budget and re-cut.
    if (keepRecent <= MIN_KEEP_RECENT_TOKENS) break;
    keepRecent = Math.max(MIN_KEEP_RECENT_TOKENS, Math.floor(keepRecent / 2));
  }

  return (
    lastResult ?? {
      cutIndex: boundaryStart,
      kept: messages.slice(boundaryStart),
      plan: {
        kept: messages.slice(boundaryStart),
        trimmedToolResults: 0,
        truncatedToolCallArgs: 0,
        droppedImages: 0,
        droppedThinking: 0,
        droppedTurns: 0,
      },
      keepRecentTokens,
      fits: false,
    }
  );
}

/**
 * True when the assembled transcript fits the model window:
 * input tokens (system + messages + tools) <= contextWindow - reserveTokens.
 */
export function verifyCompactionFits(args: CompactionFitArgs): boolean {
  const { systemPrompt, messages, tools, contextWindow, reserveTokens } = args;

  if (!Number.isFinite(contextWindow) || contextWindow <= 0) return false;

  const requestArgs: Parameters<typeof estimateRequestTokens>[0] = {
    messages: messages as AgentMessage[],
  };
  if (systemPrompt !== undefined) requestArgs.systemPrompt = systemPrompt;
  if (tools !== undefined && tools.length > 0) requestArgs.tools = tools;

  const inputTokens = estimateRequestTokens(requestArgs);
  const maxInput = Math.max(0, contextWindow - reserveTokens);
  return inputTokens <= maxInput;
}
