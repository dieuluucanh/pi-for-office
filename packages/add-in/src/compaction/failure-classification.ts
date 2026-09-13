/**
 * Terminal-run failure classification.
 *
 * Decides what a settled agent run needs next after it ends in an assistant
 * error/length tail:
 * - `overflow`           → drop the failed message, compact, retry once
 * - `recoverable-length` → same compact-and-retry path (output cut short)
 * - `transient`          → drop the failed message, backoff, retry (Pi parity)
 * - `fatal`              → non-retryable (quota/billing/unknown) — surface error
 * - `none`               → clean settle, nothing to do
 *
 * This replaces the stopReason-gated overflow check (`findTrailingContextOverflowError`)
 * so silent-overflow (z.ai) and length-stop (MiMo) branches actually fire.
 */

import type { AgentState } from "@earendil-works/pi-agent-core";
import {
 isContextOverflow,
 isRetryableAssistantError,
 type AssistantMessage,
} from "@earendil-works/pi-ai";

import {
 getCompactionThresholds,
 type CompactionThresholds,
} from "./defaults.js";

export type RunFailureKind =
 | "overflow"
 | "recoverable-length"
 | "transient"
 | "fatal"
 | "none";

export interface RunFailureClassification {
 kind: RunFailureKind;
 /** The trailing assistant message that failed, when one exists. */
 message: AssistantMessage | null;
}

/**
 * Locally mirrors pi-ai's `isRecoverableLength` (not exported by the add-in's
 * pi-ai build): a `length` stop that ended below the intended output budget is
 * recoverable — it may be caused by context pressure or provider truncation.
 */
export function isRecoverableLength(
 message: AssistantMessage,
 desiredMaxOutput: number,
): boolean {
 return (
  message.stopReason === "length" &&
  desiredMaxOutput > 0 &&
  message.usage.output < desiredMaxOutput
 );
}

/** Premature stream ending patterns (from pi-ai's retryable error catalog). */
const PREMATURE_STREAM_END_PATTERNS = [
 /ended without/i,
 /stream ended before message_stop/i,
 /stream ended before a terminal response event/i,
 /http2 request did not get a response/i,
];

export function isPrematureStreamEnd(errorMessage: string): boolean {
 return PREMATURE_STREAM_END_PATTERNS.some((pattern) =>
  pattern.test(errorMessage),
 );
}

/**
 * Classify the terminal assistant message of a settled run.
 *
 * Uses pi-ai's `isContextOverflow` (all three branches — error text, silent
 * usage overflow, MiMo length-stop) WITHOUT a `stopReason === "error"` pre-gate,
 * then recoverable length, then retryability. A premature stream end (e.g.
 * "Stream ended without finish_reason") is `transient` at low context usage,
 * but `overflow` (compact-first) when the request-facing estimate is already past
 * the soft-warning threshold — gateways often close the stream silently on
 * overflow instead of returning a proper error.
 */
export function classifyRunFailure(args: {
 state: Pick<AgentState, "messages" | "model">;
 contextWindow?: number;
 maxOutputTokens?: number;
 /** Request-facing (shaping-aware) estimate at failure time. */
 effectiveTokens?: number;
 thresholds?: CompactionThresholds;
}): RunFailureClassification {
 const { state, contextWindow, maxOutputTokens, effectiveTokens, thresholds } =
  args;

 const messages = state.messages;
 const last = messages.at(-1);
 if (!last || last.role !== "assistant") {
  return { kind: "none", message: null };
 }

 // Only failures from the currently active model count: after switching to a
 // larger-context model, a stale overflow error must not trigger recovery.
 const model = state.model;
 if (model && (last.provider !== model.provider || last.model !== model.id)) {
  return { kind: "none", message: null };
 }

 // 1. Context overflow (all branches).
 if (isContextOverflow(last, contextWindow)) {
  return { kind: "overflow", message: last };
 }

 // 2. Length-stop ending below the intended output budget.
 if (
  maxOutputTokens !== undefined &&
  isRecoverableLength(last, maxOutputTokens)
 ) {
  return { kind: "recoverable-length", message: last };
 }

 // 3. Retryable provider/transport error.
 if (last.stopReason === "error" && last.errorMessage) {
  const resolvedThresholds =
   thresholds ??
   (contextWindow !== undefined
    ? getCompactionThresholds(contextWindow)
    : undefined);

  const prematureStreamEnd = isPrematureStreamEnd(last.errorMessage);
  const softWarning = resolvedThresholds?.softWarningTokens;

  if (
   prematureStreamEnd &&
   effectiveTokens !== undefined &&
   softWarning !== undefined &&
   effectiveTokens >= softWarning
  ) {
   // Likely a silent overflow: the gateway closed the stream instead of
   // returning a proper context-length error. Compact before retrying.
   return { kind: "overflow", message: last };
  }

  if (isRetryableAssistantError(last)) {
   return { kind: "transient", message: last };
  }

  return { kind: "fatal", message: last };
 }

 return { kind: "none", message: null };
}

/**
 * Thin wrapper for existing callers/error banner: returns the trailing overflow
 * assistant message, or null. Equivalent to `classifyRunFailure().kind === "overflow"`.
 */
export function findTrailingContextOverflowError(
 state: Pick<AgentState, "messages" | "model">,
): AssistantMessage | null {
 const classification = classifyRunFailure({ state });
 return classification.kind === "overflow" ? classification.message : null;
}
