/**
 * Unified post-run recovery (Pi parity).
 *
 * Mirrors pi-coding-agent's `_handlePostAgentRun` / `_prepareRetry` semantics as
 * an explicit settle→classify→act loop:
 * - overflow / recoverable-length → drop the failed assistant message, compact,
 *   then continue — at most once per failure (gated by `allowCompaction`);
 * - transient (e.g. "Stream ended without finish_reason") → drop the failed
 *   message, backoff, continue — up to `policy.maxRetries` with exponential
 *   backoff;
 * - stalled (thinking-only tail) → drop the tail, continue — bounded by
 *   `maxStallRecovery`. When the stall ended on `length` (the output budget was
 *   the cause) the budget escalates ×2 per attempt — capped, and restored to
 *   the original model when the loop settles;
 * - budget-unsupported (provider rejected the output limit) → halve the
 *   model's output budget and continue — bounded; the reduction persists for
 *   the session (a gateway that caps `max_tokens` lower stays capped);
 * - counters reset on a clean settle; aborts are terminal and never retried.
 *
 * Recovery is decoupled from the auto-compact setting: stall/stream-error
 * recovery always runs; only the compaction step is gated by
 * `allowCompaction`.
 */

import type { Agent } from "@earendil-works/pi-agent-core";
import type { RetryPolicy } from "@earendil-works/pi-ai";

import { classifyRunFailure } from "./failure-classification.js";
import type { CompactionOutcome } from "./engine.js";
import { estimateEffectiveRequestTokens } from "../utils/context-tokens.js";

export interface RunRecoveryResult {
  kind: "clean" | "overflow-recovered" | "transient-recovered" | "exhausted";
  retryAttempts: number;
  /**
   * True when the settle was caused by an overflow-like final failure
   * (overflow / recoverable-length). Callers use this to decide adaptive-window
   * observation — stall exhaustion must not shrink the effective window.
   */
  overflowLikeFailure: boolean;
}

export interface RunRecoveryCallbacks {
  /** Emitted before the backoff sleep of a transient retry. */
  onRetryScheduled?: (
    attempt: number,
    maxAttempts: number,
    delayMs: number,
    errorMessage: string,
  ) => void | Promise<void>;
  /** Emitted right before a compact-and-retry starts. */
  onCompactStart?: () => void | Promise<void>;
  /**
   * Emitted when a stalled run is auto-continued (bounded). `escalatedTo` is
   * set when the output budget was escalated for this attempt (`length` stall).
   */
  onStallDetected?: (
    attempt: number,
    maxAttempts: number,
    info: { stopReason: "stop" | "length"; escalatedTo?: number },
  ) => void | Promise<void>;
  /** Emitted when the output budget is reduced after a provider rejection. */
  onBudgetReduced?: (newMaxTokens: number) => void | Promise<void>;
  /** Emitted once the loop settles (success, exhausted, or aborted). */
  onSettled?: (result: RunRecoveryResult) => void | Promise<void>;
}

export interface RunRecoveryOptions {
  agent: Agent;
  runCompact: () => Promise<CompactionOutcome>;
  /** Transient retry policy (defaults mirror pi: true / 3 / 2000). */
  retry?: RetryPolicy | undefined;
  callbacks?: RunRecoveryCallbacks;
  /** Injectable sleep for tests. */
  sleep?: (ms: number, signal?: AbortSignal) => Promise<void>;
  signal?: AbortSignal;
  /**
   * Compaction is available for overflow-like recovery (auto-compact setting).
   * Stall/stream-error recovery runs regardless — only the compaction step is
   * gated. Defaults to true.
   */
  allowCompaction?: boolean;
  /** Max drop-the-tail-and-continue attempts for stalled runs (default 2). */
  maxStallRecovery?: number;
}
const DEFAULT_RETRY_POLICY: RetryPolicy = {
  enabled: true,
  maxRetries: 3,
  baseDelayMs: 2000,
};

function defaultSleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new Error("Aborted"));
      return;
    }
    const timer = setTimeout(resolve, ms);
    signal?.addEventListener(
      "abort",
      () => {
        clearTimeout(timer);
        reject(new Error("Aborted"));
      },
      { once: true },
    );
  });
}

/** Max attempts for the overflow compact-and-retry path (Pi: one shot). */
const MAX_OVERFLOW_RECOVERY = 1;

/** Max halve-and-retry attempts when a provider rejects the output budget. */
const MAX_BUDGET_REDUCTION_ATTEMPTS = 2;

/** Floor for session-persistent output-budget reductions. */
const MIN_OUTPUT_BUDGET = 1_024;

/** Max drop-the-tail-and-continue attempts for stalled runs. */
const DEFAULT_MAX_STALL_RECOVERY = 2;

/**
 * Run the recovery loop until the agent settles cleanly or recovery budgets are
 * exhausted. Assumes the caller has already awaited the prompt run.
 *
 * Returns how the loop ended:
 * - `clean`: nothing to recover (or already fully recovered);
 * - `overflow-recovered`: a compact-and-retry succeeded;
 * - `transient-recovered`: a transient retry succeeded;
 * - `exhausted`: gave up after bounded retries (error stays in transcript).
 *
 * The result also carries `overflowLikeFailure` — true only when the final
 * failure was overflow-like — so callers can decide adaptive-window observation.
 */
export async function recoverAfterRun(options: RunRecoveryOptions): Promise<RunRecoveryResult> {
  const {
    agent,
    runCompact,
    retry: retryPolicy = DEFAULT_RETRY_POLICY,
    callbacks,
    signal,
  } = options;

  const sleep = options.sleep ?? defaultSleep;
  const allowCompaction = options.allowCompaction ?? true;
  const maxStallRecovery =
    options.maxStallRecovery ?? DEFAULT_MAX_STALL_RECOVERY;

  const originalModel = agent.state.model;
  const contextWindow = originalModel?.contextWindow;
  const state = agent.state;

  /** Model object recovery has swapped in (escalation), or null. */
  let workingModel: typeof originalModel | null = null;
  /** Set once a budget-unsupported reduction persists for the session. */
  let budgetReducedPersistently = false;

  let retryAttempts = 0;
  let overflowAttempts = 0;
  let stallAttempts = 0;
  let budgetReductionAttempts = 0;

  /** Restore the run's model when the loop settles (unless reduced persistently). */
  const restoreModelOnSettle = (): void => {
    if (budgetReducedPersistently || !workingModel) return;
    // Reference equality: only restore when the user has not switched models
    // mid-recovery.
    if (agent.state.model === workingModel) {
      agent.state.model = originalModel;
    }
    workingModel = null;
  };

  /** Settle: restore any escalated budget, then return the result. */
  const settle = (
    kind: RunRecoveryResult["kind"],
    overflowLike: boolean,
  ): RunRecoveryResult => {
    restoreModelOnSettle();
    return { kind, retryAttempts, overflowLikeFailure: overflowLike };
  };

  /** Swap in an escalated/reduced output budget. No-op on a concurrent switch. */
  const applyBudget = (maxTokens: number): boolean => {
    const expected = workingModel ?? originalModel;
    if (!expected || agent.state.model !== expected) {
      // The user switched models mid-recovery — never clobber their choice.
      return false;
    }
    workingModel = { ...expected, maxTokens };
    agent.state.model = workingModel;
    return true;
  };

  /** Current working output budget (escalated/reduced, or the model's own). */
  const currentBudget = (): number =>
    workingModel?.maxTokens ?? originalModel?.maxTokens ?? 0;

  /**
   * Cap for stall escalation: keep the doubled budget inside the provider-safe
   * range — never above 64k, never above the context window minus the
   * compaction reserve, and at least 8k.
   */
  const escalationCap = (): number => {
    const window = contextWindow ?? 0;
    const windowBound = window > 0 ? Math.max(0, window - 16_384) : 65_536;
    return Math.max(8_192, Math.min(windowBound, 65_536));
  };

  for (;;) {
    const maxOutputTokens = currentBudget();
    const { effectiveTokens } = estimateEffectiveRequestTokens({
      state,
      tools: state.tools,
      contextWindow,
      maxOutputTokens,
    });

    const classification = classifyRunFailure({
      state,
      contextWindow,
      ...(maxOutputTokens > 0 ? { maxOutputTokens } : {}),
      effectiveTokens,
    });

    // Clean settle (or empty/non-assistant tail): done.
    if (classification.kind === "none" || classification.message === null) {
      let kind: RunRecoveryResult["kind"] = "clean";
      if (retryAttempts > 0) {
        kind = "transient-recovered";
      } else if (overflowAttempts > 0) {
        kind = "overflow-recovered";
      }
      const result = settle(kind, false);
      await callbacks?.onSettled?.(result);
      return result;
    }

    // Overflow / recoverable-length: one compact-and-retry per failure.
    if (
      classification.kind === "overflow" ||
      classification.kind === "recoverable-length"
    ) {
      if (overflowAttempts >= MAX_OVERFLOW_RECOVERY) {
        await callbacks?.onSettled?.(settle("exhausted", true));
        return settle("exhausted", true);
      }
      if (!allowCompaction) {
        // Compaction is unavailable (auto-compact disabled) — keep the failure
        // visible for the banner/retry path instead of retrying blind.
        await callbacks?.onSettled?.(settle("exhausted", true));
        return settle("exhausted", true);
      }
      overflowAttempts += 1;

      // Drop the failed assistant message so the retry context is clean.
      const messages = agent.state.messages;
      if (messages.at(-1) === classification.message) {
        agent.state.messages = messages.slice(0, -1);
      }

      await callbacks?.onCompactStart?.();
      let compacted = false;
      try {
        const outcome = await runCompact();
        compacted =
          outcome.changed &&
          outcome.tokensAfter < outcome.tokensBefore &&
          agent.state.messages.at(-1)?.role !== "assistant";
      } catch (err) {
        console.warn("[pi] Recovery compaction failed:", err);
      }
      if (!compacted) {
        // Compaction freed nothing — restore the failure so it stays visible.
        if (!agent.state.messages.includes(classification.message)) {
          agent.state.messages = [
            ...agent.state.messages,
            classification.message,
          ];
        }
        await callbacks?.onSettled?.(settle("exhausted", true));
        return settle("exhausted", true);
      }

      await agent.continue();
      continue;
    }

    // Stalled run: thinking-only tail (no text, no tool calls) — drop the tail
    // and continue. When the stall ended on `length`, the output budget was the
    // cause: escalate ×2 per attempt, capped and restored on settle. Aborts
    // during recovery restore the dropped tail.
    if (classification.kind === "stalled") {
      if (stallAttempts >= maxStallRecovery) {
        await callbacks?.onSettled?.(settle("exhausted", false));
        return settle("exhausted", false);
      }
      stallAttempts += 1;

      const stallStopReason =
        classification.message.stopReason === "length" ? "length" : "stop";

      // Drop the thinking-only tail so the retry context is clean (required:
      // `agent.continue()` rejects an assistant tail; convertMessages drops it
      // from the request anyway).
      const messages = agent.state.messages;
      if (messages.at(-1) === classification.message) {
        agent.state.messages = messages.slice(0, -1);
      }

      let escalatedTo: number | undefined;
      if (stallStopReason === "length") {
        const cap = escalationCap();
        // Double the *current* working budget (it already reflects prior
        // escalations), clamped to the provider-safe cap.
        const nextBudget = Math.min(currentBudget() * 2, cap);
        if (currentBudget() < nextBudget && applyBudget(nextBudget)) {
          escalatedTo = nextBudget;
        }
      }

      await callbacks?.onStallDetected?.(
        stallAttempts,
        maxStallRecovery,
        escalatedTo === undefined
          ? { stopReason: stallStopReason }
          : { stopReason: stallStopReason, escalatedTo },
      );

      await agent.continue();
      continue;
    }

    // Budget-unsupported: the provider rejected the output limit itself.
    // Halve the budget and retry (bounded); the reduction persists for the
    // session — a gateway that caps `max_tokens` lower stays capped.
    if (classification.kind === "budget-unsupported") {
      if (
        budgetReductionAttempts >= MAX_BUDGET_REDUCTION_ATTEMPTS ||
        currentBudget() <= MIN_OUTPUT_BUDGET ||
        !originalModel
      ) {
        await callbacks?.onSettled?.(settle("exhausted", false));
        return settle("exhausted", false);
      }
      budgetReductionAttempts += 1;

      const reduced = Math.max(
        MIN_OUTPUT_BUDGET,
        Math.floor(currentBudget() / 2),
      );
      if (!applyBudget(reduced)) {
        // Concurrent model switch — keep the failure visible.
        await callbacks?.onSettled?.(settle("exhausted", false));
        return settle("exhausted", false);
      }
      budgetReducedPersistently = true;

      await callbacks?.onBudgetReduced?.(reduced);

      await agent.continue();
      continue;
    }

    // Fatal: non-retryable, keep the error.
    if (classification.kind === "fatal") {
      await callbacks?.onSettled?.(settle("exhausted", false));
      return settle("exhausted", false);
    }

    // Transient: bounded exponential-backoff retry, dropped failed message.
    if (classification.kind === "transient") {
      if (!retryPolicy.enabled || retryAttempts >= retryPolicy.maxRetries) {
        await callbacks?.onSettled?.(settle("exhausted", false));
        return settle("exhausted", false);
      }

      retryAttempts += 1;
      const delayMs = retryPolicy.baseDelayMs * 2 ** (retryAttempts - 1);
      const errorMessage =
        classification.message.errorMessage ?? "Unknown error";

      await callbacks?.onRetryScheduled?.(
        retryAttempts,
        retryPolicy.maxRetries,
        delayMs,
        errorMessage,
      );

      const messages = agent.state.messages;
      if (messages.at(-1) === classification.message) {
        agent.state.messages = messages.slice(0, -1);
      }

      try {
        await sleep(delayMs, signal);
      } catch {
        // Aborted during backoff: restore the failure and stop.
        const current = agent.state.messages;
        if (!current.includes(classification.message)) {
          agent.state.messages = [...current, classification.message];
        }
        await callbacks?.onSettled?.(settle("exhausted", false));
        return settle("exhausted", false);
      }

      await agent.continue();
      continue;
    }
  }
}

// Re-export the classification helpers for UI wiring.
export { classifyRunFailure } from "./failure-classification.js";
