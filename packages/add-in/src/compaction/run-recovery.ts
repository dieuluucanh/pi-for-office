/**
 * Unified post-run recovery (Pi parity).
 *
 * Mirrors pi-coding-agent's `_handlePostAgentRun` / `_prepareRetry` semantics as
 * an explicit settle→classify→act loop:
 * - overflow / recoverable-length → drop the failed assistant message, compact,
 *   then continue — at most once per failure;
 * - transient (e.g. "Stream ended without finish_reason") → drop the failed
 *   message, backoff, continue — up to `policy.maxRetries` with exponential
 *   backoff;
 * - counters reset on a clean settle; aborts are terminal and never retried.
 */

import type { Agent } from "@earendil-works/pi-agent-core";
import type { RetryPolicy } from "@earendil-works/pi-ai";

import { classifyRunFailure } from "./failure-classification.js";
import type { CompactionOutcome } from "./engine.js";
import { estimateEffectiveRequestTokens } from "../utils/context-tokens.js";

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
  /** Emitted once the loop settles (success, exhausted, or aborted). */
  onSettled?: (result: {
    kind: "clean" | "overflow-recovered" | "transient-recovered" | "exhausted";
    retryAttempts: number;
  }) => void | Promise<void>;
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

/**
 * Run the recovery loop until the agent settles cleanly or recovery budgets are
 * exhausted. Assumes the caller has already awaited the prompt run.
 *
 * Returns how the loop ended:
 * - `clean`: nothing to recover (or already fully recovered);
 * - `overflow-recovered`: a compact-and-retry succeeded;
 * - `transient-recovered`: a transient retry succeeded;
 * - `exhausted`: gave up after bounded retries (error stays in transcript).
 */
export async function recoverAfterRun(options: RunRecoveryOptions): Promise<{
  kind: "clean" | "overflow-recovered" | "transient-recovered" | "exhausted";
  retryAttempts: number;
}> {
  const {
    agent,
    runCompact,
    retry: retryPolicy = DEFAULT_RETRY_POLICY,
    callbacks,
    signal,
  } = options;

  const sleep = options.sleep ?? defaultSleep;

  const model = agent.state.model;
  const contextWindow = model?.contextWindow;
  const maxOutputTokens = model?.maxTokens;
  const state = agent.state;

  let retryAttempts = 0;
  let overflowAttempts = 0;

  for (;;) {
    const { effectiveTokens } = estimateEffectiveRequestTokens({
      state,
      tools: state.tools,
      contextWindow,
      maxOutputTokens,
    });

    const classification = classifyRunFailure({
      state,
      contextWindow,
      maxOutputTokens,
      effectiveTokens,
    });

    // Clean settle (or empty/non-assistant tail): done.
    if (classification.kind === "none" || classification.message === null) {
      await callbacks?.onSettled?.({ kind: "clean", retryAttempts });

      let kind:
        | "clean"
        | "overflow-recovered"
        | "transient-recovered"
        | "exhausted" = "clean";
      if (retryAttempts > 0) {
        kind = "transient-recovered";
      } else if (overflowAttempts > 0) {
        kind = "overflow-recovered";
      }
      return { kind, retryAttempts };
    }

    // Overflow / recoverable-length: one compact-and-retry per failure.
    if (
      classification.kind === "overflow" ||
      classification.kind === "recoverable-length"
    ) {
      if (overflowAttempts >= MAX_OVERFLOW_RECOVERY) {
        await callbacks?.onSettled?.({ kind: "exhausted", retryAttempts });
        return { kind: "exhausted", retryAttempts };
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
        await callbacks?.onSettled?.({ kind: "exhausted", retryAttempts });
        return { kind: "exhausted", retryAttempts };
      }

      await agent.continue();
      continue;
    }

    // Fatal: non-retryable, keep the error.
    if (classification.kind === "fatal") {
      await callbacks?.onSettled?.({ kind: "exhausted", retryAttempts });
      return { kind: "exhausted", retryAttempts };
    }

    // Transient: bounded exponential-backoff retry, dropped failed message.
    if (classification.kind === "transient") {
      if (!retryPolicy.enabled || retryAttempts >= retryPolicy.maxRetries) {
        await callbacks?.onSettled?.({ kind: "exhausted", retryAttempts });
        return { kind: "exhausted", retryAttempts };
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
        await callbacks?.onSettled?.({ kind: "exhausted", retryAttempts });
        return { kind: "exhausted", retryAttempts };
      }

      await agent.continue();
      continue;
    }
  }
}

// Re-export the classification helpers for UI wiring.
export { classifyRunFailure } from "./failure-classification.js";
