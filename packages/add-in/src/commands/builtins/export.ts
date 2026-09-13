function isCommandsBuiltinsExportPayloadShape(
  value: DynamicValue,
): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Builtin export/compaction commands.
 */

import type {
  Api,
  AssistantMessage,
  Model,
  StopReason,
  Usage,
} from "@earendil-works/pi-ai";
import {
  isContextOverflow,
  retryAssistantCall,
  type RetryPolicy,
} from "@earendil-works/pi-ai";
import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";

import type { SlashCommand } from "../types.js";
import type { ActiveAgentProvider } from "./model.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";
import { createCompactionSummaryMessage } from "../../messages/compaction.js";
import {
  createArchivedMessagesMessage,
  splitArchivedMessages,
} from "../../messages/archived-history.js";
import { getErrorMessage } from "../../utils/errors.js";
import {
  extractTextBlocks,
  summarizeContentForTranscript,
} from "../../utils/content.js";
import {
  estimateMessageTokens,
  estimateRequestTokens,
  estimateTextTokens,
  estimateToolsTokens,
} from "../../utils/context-tokens.js";
import type { PiSidebar } from "../../ui/pi-sidebar.js";
import { getWorkbookChangeAuditLog } from "../../audit/workbook-change-audit.js";
import {
  effectiveKeepRecentTokens,
  effectiveReserveTokens,
} from "../../compaction/defaults.js";
import {
  MIN_KEEP_RECENT_TOKENS,
  planCompaction,
  verifyCompactionFits,
  type CompactionOutcome,
} from "../../compaction/engine.js";
import {
  buildCompactionMemoryFocusInstruction,
  collectCompactionMemoryCues,
  mergeCompactionAdditionalFocus,
} from "../../compaction/memory-nudge.js";

type TranscriptEntry = {
  role: AgentMessage["role"];
  text: string;
  usage?: Usage;
  stopReason?: StopReason;
};

function isApiModel(model: DynamicValue): model is Model<Api> {
  if (!isCommandsBuiltinsExportPayloadShape(model)) return false;

  return (
    typeof model.id === "string" &&
    typeof model.name === "string" &&
    typeof model.provider === "string" &&
    typeof model.api === "string"
  );
}

function hasContent(
  message: AgentMessage,
): message is AgentMessage & { content: DynamicValue } {
  return isCommandsBuiltinsExportPayloadShape(message) && "content" in message;
}

function messageToTranscriptText(message: AgentMessage): string {
  if (message.role === "archivedMessages") {
    return `[archived history: ${message.archivedChatMessageCount} chat messages]`;
  }

  if (message.role === "compactionSummary") return message.summary;
  if (hasContent(message))
    return summarizeContentForTranscript(message.content);
  return "";
}

function countChatMessages(messages: AgentMessage[]): number {
  let count = 0;
  for (const m of messages) {
    const role = m.role;
    if (
      role === "user" ||
      role === "assistant" ||
      role === "user-with-attachments"
    ) {
      count += 1;
    }
  }
  return count;
}

type ExportDestination = "clipboard" | "download";

function parseExportDestination(
  raw: string,
  fallback: ExportDestination,
): ExportDestination {
  const normalized = raw.trim().toLowerCase();
  if (normalized === "clipboard") return "clipboard";
  if (normalized === "file" || normalized === "download") return "download";
  return fallback;
}

function triggerJsonDownload(fileName: string, content: string): void {
  const blob = new Blob([content], { type: "application/json" });
  const url = URL.createObjectURL(blob);

  // Use a programmatic `<a download>` click (hidden, rel=noopener). This works
  // in Office WebViews (WKWebView included) without touching window.open.
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.rel = "noopener";
  anchor.hidden = true;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

async function exportWorkbookAuditLog(rawArgs: string): Promise<void> {
  const destination = parseExportDestination(rawArgs, "download");

  const entries = await getWorkbookChangeAuditLog().list(500);
  const payload = {
    exported: new Date().toISOString(),
    count: entries.length,
    entries,
  };

  const json = JSON.stringify(payload, null, 2);

  if (destination === "clipboard") {
    await navigator.clipboard.writeText(json);
    showToast(
      t("export.toast.audit_copied", {
        count: String(entries.length),
        size: (json.length / 1024).toFixed(0),
      }),
    );
    return;
  }

  triggerJsonDownload(
    `pi-audit-log-${new Date().toISOString().slice(0, 10)}.json`,
    json,
  );
  showToast(
    t("export.toast.audit_downloaded", { count: String(entries.length) }),
  );
}

// =============================================================================
// Compaction helpers
// =============================================================================

// Mirrors pi-coding-agent defaults (see docs/compaction.md in pi-coding-agent).

const SUMMARIZATION_SYSTEM_PROMPT =
  "You are a context summarization assistant. Your task is to read a conversation between a user and an AI assistant, then produce a structured summary following the exact format specified.\n\n" +
  "Do NOT continue the conversation. Do NOT respond to any questions in the conversation. ONLY output the structured summary.";

const SUMMARIZATION_PROMPT = `The messages above are a conversation to summarize. Create a structured context checkpoint summary that another LLM will use to continue the work.

Use this EXACT format:

## Goal
[What is the user trying to accomplish? Can be multiple items if the session covers different tasks.]

## Constraints & Preferences
- [Any constraints, preferences, or requirements mentioned by user]
- [Or "(none)" if none were mentioned]

## Progress
### Done
- [x] [Completed tasks/changes]

### In Progress
- [ ] [Current work]

### Blocked
- [Issues preventing progress, if any]

## Key Decisions
- **[Decision]**: [Brief rationale]

## Next Steps
1. [Ordered list of what should happen next]

## Critical Context
- [Any data, examples, or references needed to continue]
- [Or "(none)" if not applicable]

Keep each section concise. Preserve exact identifiers (sheet names, cell addresses, tool names, error messages).`;

const UPDATE_SUMMARIZATION_PROMPT = `The messages above are NEW conversation messages to incorporate into the existing summary provided in <previous-summary> tags.

Update the existing structured summary with new information. RULES:
- PRESERVE all existing information from the previous summary
- ADD new progress, decisions, and context from the new messages
- UPDATE the Progress section: move items from "In Progress" to "Done" when completed
- UPDATE "Next Steps" based on what was accomplished
- PRESERVE exact identifiers (sheet names, cell addresses, tool names, error messages)
- If something is no longer relevant, you may remove it

Use this EXACT format:

## Goal
[Preserve existing goals, add new ones if the task expanded]

## Constraints & Preferences
- [Preserve existing, add new ones discovered]

## Progress
### Done
- [x] [Include previously done items AND newly completed items]

### In Progress
- [ ] [Current work - update based on progress]

### Blocked
- [Current blockers - remove if resolved]

## Key Decisions
- **[Decision]**: [Brief rationale] (preserve all previous, add new)

## Next Steps
1. [Update based on current state]

## Critical Context
- [Preserve important context, add new if needed]

Keep each section concise. Preserve exact identifiers (sheet names, cell addresses, tool names, error messages).`;

type SerializeLimits = {
  maxUserChars: number;
  maxAssistantChars: number;
  maxToolResultChars: number;
};

function truncateMiddle(text: string, maxChars: number): string {
  if (maxChars <= 0) return "";
  if (text.length <= maxChars) return text;

  const marker = "\n…[truncated]…\n";
  const keep = Math.max(0, maxChars - marker.length);
  const head = Math.floor(keep / 2);
  const tail = keep - head;

  return text.slice(0, head) + marker + text.slice(text.length - tail);
}

function getPreviousCompaction(messages: AgentMessage[]): {
  boundaryStart: number;
  previousSummary?: string;
} {
  for (let i = messages.length - 1; i >= 0; i--) {
    const m = messages[i];
    if (m?.role === "compactionSummary") {
      return { boundaryStart: i + 1, previousSummary: m.summary };
    }
  }
  return { boundaryStart: 0 };
}

function serializeConversation(
  messages: AgentMessage[],
  limits: SerializeLimits,
  maxTotalChars?: number,
): string {
  const parts: string[] = [];

  for (const msg of messages) {
    if (msg.role === "artifact") continue;

    if (msg.role === "user" || msg.role === "user-with-attachments") {
      const raw =
        typeof msg.content === "string"
          ? msg.content
          : extractTextBlocks(msg.content);
      const text = truncateMiddle(raw, limits.maxUserChars);
      if (text.trim().length > 0) parts.push(`[User]: ${text}`);
      continue;
    }

    if (msg.role === "assistant") {
      const textParts: string[] = [];
      const thinkingParts: string[] = [];
      const toolCalls: string[] = [];

      for (const block of msg.content) {
        if (block.type === "text") {
          textParts.push(block.text);
        } else if (block.type === "thinking") {
          thinkingParts.push(block.thinking);
        } else if (block.type === "toolCall") {
          let args = "";
          try {
            args = JSON.stringify(block.arguments);
          } catch {
            args = "{}";
          }
          toolCalls.push(`${block.name}(${args})`);
        }
      }

      if (thinkingParts.length > 0) {
        const t = truncateMiddle(
          thinkingParts.join("\n"),
          limits.maxAssistantChars,
        );
        parts.push(`[Assistant thinking]: ${t}`);
      }

      if (textParts.length > 0) {
        const t = truncateMiddle(
          textParts.join("\n"),
          limits.maxAssistantChars,
        );
        parts.push(`[Assistant]: ${t}`);
      }

      if (toolCalls.length > 0) {
        const t = truncateMiddle(
          toolCalls.join("; "),
          limits.maxAssistantChars,
        );
        parts.push(`[Assistant tool calls]: ${t}`);
      }

      continue;
    }

    if (msg.role === "toolResult") {
      const raw = extractTextBlocks(msg.content);
      const text = truncateMiddle(raw, limits.maxToolResultChars);
      const label = `${msg.toolName}${msg.isError ? " (error)" : ""}`;
      if (text.trim().length > 0) {
        parts.push(`[Tool result ${label}]: ${text}`);
      }
      continue;
    }

    // Ignore other message types.
  }

  if (maxTotalChars === undefined || maxTotalChars <= 0) {
    return parts.join("\n\n");
  }

  // Bounded serialization: join newest-first, dropping oldest parts until the
  // prompt fits `maxTotalChars`. Oldest content is least relevant to summarize.
  let joined = "";
  let dropped = 0;
  for (let i = parts.length - 1; i >= 0; i--) {
    const part = parts[i];
    if (!part) continue;
    const candidate = joined.length === 0 ? part : `${part}\n\n${joined}`;
    if (candidate.length > maxTotalChars && joined.length > 0) {
      dropped += 1;
      continue;
    }
    if (candidate.length > maxTotalChars) {
      // A single oldest message is too big even alone: hard-truncate it.
      joined = truncateMiddle(candidate, maxTotalChars);
      break;
    }
    joined = candidate;
  }
  if (dropped > 0) {
    const marker = `\n\n[${dropped} earlier message${dropped === 1 ? "" : "s"} omitted to fit the summarization budget]`;
    joined = marker + (joined.length > 0 ? `\n\n${joined}` : "");
  }
  return joined;
}

function buildSummarizationPrompt(args: {
  conversationText: string;
  previousSummary?: string;
  customInstructions?: string;
}): string {
  const base = args.previousSummary
    ? UPDATE_SUMMARIZATION_PROMPT
    : SUMMARIZATION_PROMPT;
  const withFocus = args.customInstructions
    ? `${base}\n\nAdditional focus: ${args.customInstructions}`
    : base;

  let prompt = `<conversation>\n${args.conversationText}\n</conversation>\n\n`;

  if (args.previousSummary) {
    prompt += `<previous-summary>\n${args.previousSummary}\n</previous-summary>\n\n`;
  }

  prompt += withFocus;
  return prompt;
}

function isPromptTooLongError(err: DynamicValue): boolean {
  const msg = getErrorMessage(err).toLowerCase();
  return (
    msg.includes("prompt is too long") ||
    msg.includes("context_length_exceeded") ||
    (msg.includes("maximum") && msg.includes("tokens"))
  );
}

export function createExportCommands(
  getActiveAgent: ActiveAgentProvider,
): SlashCommand[] {
  return [
    {
      name: "export",
      description: t("command.export.json"),
      source: "builtin",
      execute: async (args: string) => {
        const parts = args
          .trim()
          .split(/\s+/u)
          .filter((part) => part.length > 0);
        const mode = parts[0]?.toLowerCase();

        if (mode === "audit" || mode === "audit-log") {
          try {
            await exportWorkbookAuditLog(parts.slice(1).join(" "));
          } catch (error) {
            showToast(
              t("export.toast.audit_export_failed", {
                error: getErrorMessage(error),
              }),
            );
          }
          return;
        }

        const agent = getActiveAgent();
        if (!agent) {
          showToast(t("export.toast.no_session"));
          return;
        }

        const msgs = agent.state.messages;
        if (msgs.length === 0) {
          showToast(t("export.toast.no_messages"));
          return;
        }

        const transcript: TranscriptEntry[] = msgs.map((m) => {
          const text = messageToTranscriptText(m);
          if (m.role === "assistant") {
            return {
              role: m.role,
              text,
              usage: m.usage,
              stopReason: m.stopReason,
            };
          }
          return { role: m.role, text };
        });

        const exportData = {
          exported: new Date().toISOString(),
          model: agent.state.model
            ? {
                id: agent.state.model.id,
                name: agent.state.model.name,
                provider: agent.state.model.provider,
              }
            : null,
          thinkingLevel: agent.state.thinkingLevel,
          messageCount: msgs.length,
          transcript,
          // Also include raw messages for full fidelity debugging
          raw: msgs,
        };

        const json = JSON.stringify(exportData, null, 2);
        const destination = parseExportDestination(args, "clipboard");

        if (destination === "clipboard") {
          try {
            await navigator.clipboard.writeText(json);
            showToast(
              t("export.toast.transcript_copied", {
                count: String(msgs.length),
                size: (json.length / 1024).toFixed(0),
              }),
            );
          } catch (error) {
            showToast(
              t("export.toast.copy_failed", { error: getErrorMessage(error) }),
            );
          }
          return;
        }

        triggerJsonDownload(
          `pi-session-${new Date().toISOString().slice(0, 10)}.json`,
          json,
        );
        showToast(
          t("export.toast.transcript_downloaded", {
            count: String(msgs.length),
          }),
        );
      },
    },
  ];
}

/**
 * Run compaction for a specific agent.
 *
 * Used by the `/compact` slash command (with the active agent) and by
 * auto-compaction / overflow recovery (with the agent that owns the run, which
 * may not be the active tab).
 */
/** Default summarizer retry policy (matches pi coding-agent defaults). */
const DEFAULT_SUMMARIZER_RETRY: RetryPolicy = {
  enabled: true,
  maxRetries: 2,
  baseDelayMs: 2000,
};

/**
 * Run compaction for a specific agent.
 *
 * Used by the `/compact` slash command (with the active agent) and by
 * auto-compaction / overflow recovery (with the agent that owns the run, which
 * may not be the active tab).
 *
 * New engine behavior:
 * - kept tail is bounded + trimmed by `planCompaction` (never a whole trailing
 *   tool batch), so “Nothing to compact” cannot happen while over budget;
 * - the summarization prompt is bounded with oldest-message omission;
 * - the summarizer stream is retryable (`retryAssistantCall`);
 * - the assembled transcript is verified to fit before committing;
 * - the caller can persist via `onCommitted`.
 */
export async function runCompactCommand(
  agent: Agent,
  args: string,
  opts?: {
    /** Persist the committed transcript (e.g. session autosave). */
    onCommitted?: () => void | Promise<void>;
    /** Override summarizer retry policy. */
    retry?: RetryPolicy;
    /** Abort signal for the summarizer call. */
    signal?: AbortSignal;
  },
): Promise<CompactionOutcome> {
  const allMessages = agent.state.messages;
  const {
    archivedMessages: existingArchivedMessages,
    messagesWithoutArchived,
  } = splitArchivedMessages(allMessages);

  const tokensBefore = estimateRequestTokens({
    systemPrompt: agent.state.systemPrompt,
    messages: messagesWithoutArchived,
    tools: agent.state.tools as never,
  });

  const failure = (errorMessage: string): CompactionOutcome => ({
    changed: false,
    reason: "failed",
    tokensBefore,
    tokensAfter: tokensBefore,
    keptCount: messagesWithoutArchived.length,
    summarizedCount: 0,
    errorMessage,
  });

  if (messagesWithoutArchived.length < 4) {
    showToast(t("export.toast.compact.few_messages"));
    return failure(t("export.toast.compact.few_messages"));
  }

  showToast(t("export.toast.compact.compacting"), 60000);

  const now = Date.now();
  const model = agent.state.model;
  if (!isApiModel(model)) {
    showToast(t("export.toast.compact.no_model"));
    return failure(t("export.toast.compact.no_model"));
  }

  // IMPORTANT: use the agent's configured stream function + API key resolver.
  // Calling pi-ai's completeSimple() directly bypasses:
  // - our CORS proxy logic (streamFunction)
  // - our API key/OAuth resolution (agent.getApiKey)
  // and can crash in browser WebViews due to env key fallbacks using `process`.
  const apiKey = agent.getApiKey
    ? await agent.getApiKey(model.provider)
    : undefined;
  if (!apiKey) {
    showToast(
      t("export.toast.compact.no_api_key", { provider: model.provider }),
    );
    return failure(
      t("export.toast.compact.no_api_key", { provider: model.provider }),
    );
  }

  const contextWindow = model.contextWindow || 200000;

  // Pi uses reserveTokens to ensure we don't run out of room for the model's response.
  const reserveTokens = effectiveReserveTokens(contextWindow);
  const keepRecentTokens = effectiveKeepRecentTokens(
    contextWindow,
    reserveTokens,
  );

  const maxTokens = Math.max(
    256,
    Math.min(model.maxTokens, Math.floor(0.8 * reserveTokens)),
  );

  // Budget for the kept tail: window - reserve - summary output - system/tool
  // overhead - margin. Everything in the compacted transcript must fit.
  const systemTokens = estimateTextTokens(agent.state.systemPrompt);
  const toolsTokens = estimateToolsTokens(agent.state.tools as never);
  const margin = 1024;
  const keptBudgetTokens = Math.max(
    MIN_KEEP_RECENT_TOKENS,
    contextWindow -
      reserveTokens -
      maxTokens -
      systemTokens -
      toolsTokens -
      margin,
  );

  // Summarization prompt budget (chars) so the request itself never overflows.
  const serializedBudgetChars = Math.max(
    4000,
    (contextWindow - reserveTokens - systemTokens - maxTokens - margin) * 4,
  );

  const { boundaryStart, previousSummary } = getPreviousCompaction(
    messagesWithoutArchived,
  );
  const userCompactionFocus = args.trim() || undefined;
  let memoryNudgeShown = false;

  const buildPromptText = (
    messagesToSummarize: AgentMessage[],
    limits: SerializeLimits,
  ): string => {
    const memoryCues = collectCompactionMemoryCues(messagesToSummarize);
    if (memoryCues.cueCount > 0 && !memoryNudgeShown) {
      const cueLabel = memoryCues.cueCount === 1 ? "cue" : "cues";
      showToast(
        t("export.toast.compact.memory_nudge", {
          count: memoryCues.cueCount,
          cue: cueLabel,
        }),
        12000,
      );
      memoryNudgeShown = true;
    }

    const conversationText = serializeConversation(
      messagesToSummarize,
      limits,
      serializedBudgetChars,
    );
    const memoryFocus = buildCompactionMemoryFocusInstruction(memoryCues);
    const customInstructions = mergeCompactionAdditionalFocus(
      userCompactionFocus,
      memoryFocus,
    );
    return buildSummarizationPrompt({
      conversationText,
      ...(previousSummary !== undefined ? { previousSummary } : {}),
      ...(customInstructions !== undefined ? { customInstructions } : {}),
    });
  };

  const produceSummarizer = async (
    messagesToSummarize: AgentMessage[],
    limits: SerializeLimits,
  ): Promise<AssistantMessage> => {
    const promptText = buildPromptText(messagesToSummarize, limits);
    const stream = await agent.streamFunction(
      model,
      {
        systemPrompt: SUMMARIZATION_SYSTEM_PROMPT,
        messages: [
          {
            role: "user",
            content: [{ type: "text", text: promptText }],
            timestamp: Date.now(),
          },
        ],
      },
      {
        apiKey,
        ...(agent.sessionId !== undefined
          ? { sessionId: agent.sessionId }
          : {}),
        maxTokens,
        // Match pi-coding-agent: don't force temperature when using reasoning,
        // since Anthropic requires temperature=1 when thinking is enabled.
        reasoning: "high",
      },
    );
    return stream.result();
  };

  const runOnce = async (
    producer: () => Promise<AssistantMessage>,
  ): Promise<{ summary: string; result: AssistantMessage }> => {
    const result = await retryAssistantCall(
      producer,
      opts?.retry ?? DEFAULT_SUMMARIZER_RETRY,
      opts?.signal,
      {
        onRetryScheduled: (_attempt, _maxAttempts, _delayMs) => {
          showToast(t("export.toast.compact.retrying"), 60000);
        },
      },
    );

    if (result.stopReason === "error") {
      throw new Error(
        result.errorMessage || t("export.toast.compact.failed_error"),
      );
    }
    if (result.stopReason === "length") {
      // Partial summaries must never become a checkpoint (mirrors pi's
      // getSummarizationFailure).
      throw new Error(t("export.toast.compact.failed_length"));
    }

    const summary =
      extractTextBlocks(result.content).trim() ||
      t("export.toast.compact.summary_unavailable");
    return { summary, result };
  };

  const defaultLimits: SerializeLimits = {
    maxUserChars: 4000,
    maxAssistantChars: 8000,
    maxToolResultChars: 8000,
  };

  const aggressiveLimits: SerializeLimits = {
    maxUserChars: 1200,
    maxAssistantChars: 2500,
    maxToolResultChars: 2500,
  };

  // Plan the kept tail (engine): bounded, trimmed, never a whole tool batch.
  const planResult = planCompaction({
    messages: messagesWithoutArchived,
    boundaryStart,
    keepRecentTokens,
    budgetTokens: keptBudgetTokens,
  });
  const messagesToSummarize = messagesWithoutArchived.slice(
    boundaryStart,
    planResult.cutIndex,
  );
  const kept = planResult.kept;

  const summarizedCount = countChatMessages(messagesToSummarize);
  const summarizedTokens = messagesToSummarize.reduce(
    (total, message) => total + estimateMessageTokens(message),
    0,
  );

  const commit = async (newMessages: AgentMessage[]): Promise<void> => {
    agent.state.messages = newMessages;
    const iface = document.querySelector<PiSidebar>("pi-sidebar");
    iface?.requestUpdate();
    await opts?.onCommitted?.();
  };

  const tokensAfter = (newMessages: AgentMessage[]): number =>
    estimateRequestTokens({
      systemPrompt: agent.state.systemPrompt,
      messages: newMessages,
      tools: agent.state.tools as never,
    });

  try {
    if (messagesToSummarize.length === 0) {
      // Nothing older to summarize: either the transcript already fits (a real
      // no-op) or the kept tail was trimmed to fit (tail-trimmed progress). If
      // the tail after trimming is still over budget, report why.
      const fits = verifyCompactionFits({
        systemPrompt: agent.state.systemPrompt,
        messages: [archivedOnly(existingArchivedMessages, now), ...kept],
        tools: agent.state.tools as never,
        contextWindow,
        reserveTokens,
      });
      if (fits) {
        if (
          planResult.plan.droppedTurns === 0 &&
          planResult.plan.trimmedToolResults === 0 &&
          planResult.fits
        ) {
          showToast(t("export.toast.compact.nothing"));
          return {
            changed: false,
            reason: "nothing-to-compact",
            tokensBefore,
            tokensAfter: tokensBefore,
            keptCount: kept.length,
            summarizedCount: 0,
          };
        }
        const nextMessages = [
          archivedOnly(existingArchivedMessages, now),
          ...kept,
        ];
        await commit(nextMessages);
        showToast(t("export.toast.compact.summarized", { count: 0 }));
        return {
          changed: nextMessages !== allMessages,
          reason: "tail-trimmed",
          tokensBefore,
          tokensAfter: tokensAfter(nextMessages),
          keptCount: kept.length,
          summarizedCount: 0,
          ...planResult.plan,
        };
      }
      showToast(
        t("export.toast.compact.failed", {
          msg: t("export.toast.compact.exhausted"),
        }),
      );
      return failure(t("export.toast.compact.exhausted"));
    }

    let out: { summary: string; result: AssistantMessage };

    try {
      out = await runOnce(() =>
        produceSummarizer(messagesToSummarize, defaultLimits),
      );
    } catch (e) {
      const msg = getErrorMessage(e);
      const isOverflowLike =
        isPromptTooLongError(e) ||
        (e instanceof Error &&
          isContextOverflow(
            {
              role: "assistant",
              content: [],
              api: model.api,
              provider: model.provider,
              model: model.id,
              usage: {
                input: 0,
                output: 0,
                cacheRead: 0,
                cacheWrite: 0,
                totalTokens: 0,
                cost: {
                  input: 0,
                  output: 0,
                  cacheRead: 0,
                  cacheWrite: 0,
                  total: 0,
                },
              },
              stopReason: "error",
              errorMessage: msg,
              timestamp: Date.now(),
            },
            contextWindow,
          ));
      if (!isOverflowLike) throw e;

      // Retry once with stronger truncation + a smaller recent tail.
      showToast(t("export.toast.compact.retrying"), 60000);
      out = await runOnce(() =>
        produceSummarizer(messagesToSummarize, aggressiveLimits),
      );
    }

    const archived = createArchivedMessagesMessage({
      existingArchivedMessages,
      newlyArchivedMessages: messagesToSummarize,
      timestamp: now,
    });

    const compacted = createCompactionSummaryMessage({
      summary: out.summary,
      tokensBefore: summarizedTokens,
      timestamp: now,
    });

    const nextMessages: AgentMessage[] = [archived, compacted, ...kept];

    // Post-compaction fit verification: escalate (smaller tail) until it fits.
    const fits = verifyCompactionFits({
      systemPrompt: agent.state.systemPrompt,
      messages: nextMessages,
      tools: agent.state.tools as never,
      contextWindow,
      reserveTokens,
    });
    if (!fits) {
      // Escalate: drop the tail harder with a smaller budget, keeping the summary.
      const escalated = planCompaction({
        messages: messagesWithoutArchived,
        boundaryStart,
        keepRecentTokens: Math.max(
          MIN_KEEP_RECENT_TOKENS,
          keepRecentTokens / 2,
        ),
        budgetTokens: keptBudgetTokens,
      });
      const escalatedMessages: AgentMessage[] = [
        archived,
        compacted,
        ...escalated.kept,
      ];
      if (
        verifyCompactionFits({
          systemPrompt: agent.state.systemPrompt,
          messages: escalatedMessages,
          tools: agent.state.tools as never,
          contextWindow,
          reserveTokens,
        })
      ) {
        await commit(escalatedMessages);
        showToast(
          t("export.toast.compact.summarized", { count: summarizedCount }),
        );
        return {
          changed: true,
          reason: "summarized",
          tokensBefore,
          tokensAfter: tokensAfter(escalatedMessages),
          keptCount: escalated.kept.length,
          summarizedCount,
          ...escalated.plan,
        };
      }
      showToast(
        t("export.toast.compact.failed", {
          msg: t("export.toast.compact.exhausted"),
        }),
      );
      return failure(t("export.toast.compact.exhausted"));
    }

    await commit(nextMessages);
    showToast(t("export.toast.compact.summarized", { count: summarizedCount }));
    return {
      changed: true,
      reason: "summarized",
      tokensBefore,
      tokensAfter: tokensAfter(nextMessages),
      keptCount: kept.length,
      summarizedCount,
      ...planResult.plan,
    };
  } catch (e) {
    const msg = getErrorMessage(e);
    if (msg === "Nothing to compact") {
      // Should be unreachable with the engine, but keep a safe fallback.
      showToast(t("export.toast.compact.nothing"));
      return {
        changed: false,
        reason: "nothing-to-compact",
        tokensBefore,
        tokensAfter: tokensBefore,
        keptCount: messagesWithoutArchived.length,
        summarizedCount: 0,
      };
    }
    showToast(t("export.toast.compact.failed", { msg }));
    return failure(msg);
  }
}

function archivedOnly(
  existingArchivedMessages: AgentMessage[],
  timestamp: number,
): AgentMessage {
  return createArchivedMessagesMessage({
    existingArchivedMessages,
    newlyArchivedMessages: [],
    timestamp,
  });
}

export function createCompactCommands(
  getActiveAgent: ActiveAgentProvider,
): SlashCommand[] {
  return [
    {
      name: "compact",
      description: t("command.export.summarize"),
      source: "builtin",
      execute: async (args: string) => {
        const agent = getActiveAgent();
        if (!agent) {
          showToast(t("export.toast.no_session"));
          return;
        }

        await runCompactCommand(agent, args);
      },
    },
  ];
}
