/**
 * UI-level ordered action queue.
 *
 * Needed because some actions (notably `/compact`) run outside the Agent loop
 * (they call `agent.streamFunction(...)` directly) and therefore don't set
 * `agent.state.isStreaming`. Without a queue, user input can be lost when
 * compaction rewrites the message list.
 *
 * The queue also owns automatic context protection for its agent (#566):
 * - pre-prompt auto-compaction (existing behavior)
 * - mid-turn auto-compaction between tool-loop continuations (via
 *   `agent.prepareNextTurn`)
 * - one compact-and-retry recovery when a run ends in a context-overflow error
 */

import type { Agent } from "@earendil-works/pi-agent-core";
import type { RetryPolicy } from "@earendil-works/pi-ai";
import { t } from "../language/index.js";

import { commandRegistry } from "../commands/types.js";
import {
  maybeAutoCompactBeforeContinuation,
  maybeAutoCompactBeforePrompt,
} from "../compaction/auto-compaction.js";
import { recoverAfterRun } from "../compaction/run-recovery.js";
import { createAdaptiveWindow } from "../compaction/adaptive-window.js";
import type { CompactionOutcome } from "../compaction/engine.js";
import { estimateEffectiveRequestTokens } from "../utils/context-tokens.js";

export type QueuedAction =
  | { type: "prompt"; text: string }
  | { type: "command"; name: string; args: string };

export interface ActionQueue {
  enqueuePrompt: (text: string) => void;
  enqueueCommand: (name: string, args: string) => void;
  drainQueuedActions: () => QueuedAction[];
  isBusy: () => boolean;
  shutdown: () => void;
}

interface ActionQueueDisplay {
  setActionQueue: (
    items: Array<{ type: "prompt" | "command"; label: string; text: string }>,
  ) => void;
}

interface BusyIndicatorHost {
  setBusyIndicator: (label: string | null, hint?: string | null) => void;
}

export function createActionQueue(opts: {
  agent: Agent;
  sidebar: BusyIndicatorHost;
  queueDisplay: ActionQueueDisplay;
  autoCompactEnabled: boolean;
  /** Runs compaction for this queue's agent (not the active tab's agent). */
  runCompact: () => Promise<CompactionOutcome>;
  /** Transient retry policy for post-run recovery (mirrors pi settings). */
  retry?: RetryPolicy;
  /** Optional retry-state callback for UI. */
  onRetryState?: (state: {
    active: boolean;
    attempt?: number;
    maxAttempts?: number;
    label?: string;
  }) => void;
}): ActionQueue {
  const {
    agent,
    sidebar,
    queueDisplay,
    autoCompactEnabled,
    retry,
    onRetryState,
  } = opts;

  const actions: QueuedAction[] = [];
  let running = false;
  let closed = false;
  /** Set after a silent-overflow settle — the next turn must compact first. */
  let pendingCompact = false;

  const syncDisplay = () => {
    queueDisplay.setActionQueue(
      actions.map((a) => {
        if (a.type === "prompt")
          return {
            type: "prompt",
            label: t("action-queue.queued"),
            text: a.text,
          };
        return {
          type: "command",
          label: `/${a.name}`,
          text: a.args ? a.args : "",
        };
      }),
    );
  };

  const isBusy = () => running || agent.state.isStreaming;

  const runCompactWithIndicator = async (): Promise<CompactionOutcome> => {
    sidebar.setBusyIndicator(
      t("action-queue.compacting"),
      t("action-queue.compactingHint"),
    );
    try {
      return await opts.runCompact();
    } finally {
      sidebar.setBusyIndicator(null);
    }
  };

  // Session-scoped adaptive window: after an overflow-like failure, guards
  // assume a smaller effective window so they compact earlier (provider caps
  // can be lower than the catalog's `model.contextWindow`).
  const claimedWindow = () => agent.state.model?.contextWindow || 200000;
  const adaptive = createAdaptiveWindow(claimedWindow());

  // Reset the adaptive window when the model changes.
  let lastModelId: string | undefined;
  const effectiveContextWindow = (): number => {
    const model = agent.state.model;
    if (model && model.id !== lastModelId) {
      lastModelId = model.id;
      adaptive.reset(model.contextWindow || 200000);
    }
    return adaptive.get().effectiveWindow;
  };

  // Mid-turn auto-compaction: checked between tool-loop continuations so a
  // single tool-heavy turn can't overflow a small context window.
  agent.prepareNextTurn = async () => {
    if (closed) return undefined;
    return maybeAutoCompactBeforeContinuation({
      agent,
      enabled: autoCompactEnabled,
      contextWindow: effectiveContextWindow(),
      runCompact: runCompactWithIndicator,
    });
  };

  const shutdown = () => {
    closed = true;
    delete agent.prepareNextTurn;
    actions.length = 0;
    syncDisplay();
  };

  async function runCommand(name: string, args: string): Promise<void> {
    const cmd = commandRegistry.get(name);
    if (!cmd) throw new Error(`Unknown command: /${name}`);

    // Special-case: show an explicit non-streaming indicator for compaction.
    if (name === "compact") {
      sidebar.setBusyIndicator(
        t("action-queue.compacting"),
        t("action-queue.compactingHint"),
      );
      try {
        await cmd.execute(args);
      } finally {
        sidebar.setBusyIndicator(null);
      }
      return;
    }

    await cmd.execute(args);
  }

  async function process(): Promise<void> {
    if (running || closed) return;
    running = true;

    try {
      // Drain sequentially.
      while (!closed && actions.length > 0) {
        // Never start queued actions while the agent is still streaming.
        await agent.waitForIdle();

        if (closed) break;

        const next = actions.shift();
        if (!next) break;
        syncDisplay();

        if (closed) break;

        if (next.type === "command") {
          await runCommand(next.name, next.args);
          continue;
        }

        // next.type === "prompt"
        await maybeAutoCompactBeforePrompt({
          agent,
          nextUserText: next.text,
          enabled: autoCompactEnabled || pendingCompact,
          contextWindow: effectiveContextWindow(),
          runCompact: runCompactWithIndicator,
        });
        pendingCompact = false;

        if (closed) break;
        await agent.prompt(next.text);

        if (closed) break;
        if (autoCompactEnabled) {
          // Pi-parity recovery loop: transient stream errors (e.g. "Stream
          // ended without finish_reason") retry with backoff; overflow ends in
          // one compact-and-retry. Never loops forever — budgets are bounded.
          const initialWindow = effectiveContextWindow();
          const recovery = await recoverAfterRun({
            agent,
            runCompact: runCompactWithIndicator,
            retry,
            callbacks: {
              onRetryScheduled: (
                attempt,
                maxAttempts,
                delayMs,
                errorMessage,
              ) => {
                sidebar.setBusyIndicator(
                  t("action-queue.retrying", {
                    attempt,
                    maxAttempts,
                    delayMs: Math.round(delayMs / 1000),
                  }),
                  errorMessage,
                );
                onRetryState?.({
                  active: true,
                  attempt,
                  maxAttempts,
                  label: t("action-queue.retrying", {
                    attempt,
                    maxAttempts,
                    delayMs: Math.round(delayMs / 1000),
                  }),
                });
              },
              onCompactStart: () => {
                sidebar.setBusyIndicator(
                  t("action-queue.compacting"),
                  t("action-queue.compactingHint"),
                );
              },
              onSettled: () => {
                sidebar.setBusyIndicator(null);
                onRetryState?.({ active: false });
              },
            },
          });

          // Adaptive window: an exhausted overflow/stream-end failure at high
          // usage suggests the real provider cap is below `model.contextWindow`.
          if (recovery.kind === "exhausted") {
            const state = agent.state;
            const { effectiveTokens } = estimateEffectiveRequestTokens({
              state,
              tools: state.tools,
              contextWindow: claimedWindow(),
              maxOutputTokens: state.model?.maxTokens ?? 0,
            });
            adaptive.observeFailure({
              failureIsOverflowLike: true,
              effectiveTokens,
              contextWindow: claimedWindow(),
            });
          }

          // Post-success silent-overflow check: if the provider accepted a
          // request whose input exceeded the (effective) window, flag the next
          // turn to compact first.
          const last = agent.state.messages.at(-1);
          if (last && last.role === "assistant" && last.stopReason === "stop") {
            const inputTokens =
              (last.usage?.input ?? 0) + (last.usage?.cacheRead ?? 0);
            if (inputTokens > initialWindow * 0.99) {
              pendingCompact = true;
              onRetryState?.({
                active: true,
                label: t("status.context.full"),
              });
            }
          }
        }
      }
    } finally {
      running = false;
      syncDisplay();
    }
  }

  const enqueuePrompt = (text: string) => {
    if (closed) return;

    const trimmed = text.trim();
    if (!trimmed) return;

    actions.push({ type: "prompt", text: trimmed });
    syncDisplay();
    void process();
  };

  const enqueueCommand = (name: string, args: string) => {
    if (closed) return;

    const cmdName = name.trim();
    if (!cmdName) return;

    actions.push({ type: "command", name: cmdName, args: args.trim() });
    syncDisplay();
    void process();
  };

  const drainQueuedActions = (): QueuedAction[] => {
    if (actions.length === 0) {
      return [];
    }

    const drained = [...actions];
    actions.length = 0;
    syncDisplay();
    return drained;
  };

  return {
    enqueuePrompt,
    enqueueCommand,
    drainQueuedActions,
    isBusy,
    shutdown,
  };
}
