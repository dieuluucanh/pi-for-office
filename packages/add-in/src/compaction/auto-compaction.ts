/**
 * Auto-compaction.
 *
 * Hard trigger:
 *   projectedContextTokens > hardTriggerTokens
 *
 * where hardTriggerTokens is derived from model context window and compaction
 * defaults (see `getCompactionThresholds`).
 *
 * Two trigger points share the same budgets:
 * - before a queued user prompt (`maybeAutoCompactBeforePrompt`)
 * - mid-turn, between tool-loop continuations (`maybeAutoCompactBeforeContinuation`)
 *
 * Both use the request-facing (shaping-aware) estimate and verify that
 * compaction actually reduced the transcript below the trigger before allowing
 * the request to proceed — the guard never dispatches an over-budget request.
 */

import type { Agent, AgentLoopTurnUpdate } from "@earendil-works/pi-agent-core";

import {
  estimateEffectiveRequestTokens,
  estimateTextTokens,
} from "../utils/context-tokens.js";

import { getCompactionThresholds } from "./defaults.js";
import type { CompactionOutcome } from "./engine.js";

export function shouldAutoCompactForProjectedTokens(args: {
  projectedTokens: number;
  contextWindow: number;
}): boolean {
  const { projectedTokens, contextWindow } = args;
  const { hardTriggerTokens } = getCompactionThresholds(contextWindow);
  return projectedTokens > hardTriggerTokens;
}

/** Context needed for the pre-flight guard. */
export interface AutoCompactContext {
  agent: Agent;
  nextUserText: string;
  enabled: boolean;
  /** Effective window (may be smaller than model.contextWindow after a failure). */
  contextWindow: number;
  /** When true, trust the runtime to have compacted (verify reduction). */
  runCompact: () => Promise<CompactionOutcome>;
}

function projectedTokensFor(
  agent: Agent,
  nextUserText: string,
  contextWindow: number,
): number {
  const state = agent.state;
  const maxOutput = state.model?.maxTokens ?? 0;
  const { effectiveTokens } = estimateEffectiveRequestTokens({
    state,
    tools: state.tools,
    contextWindow,
    maxOutputTokens: maxOutput,
  });
  return effectiveTokens + estimateTextTokens(nextUserText);
}

export async function maybeAutoCompactBeforePrompt(
  args: AutoCompactContext,
): Promise<boolean> {
  const { agent, nextUserText, enabled, contextWindow, runCompact } = args;

  if (!enabled) return false;
  if (agent.state.isStreaming) return false;

  const model = agent.state.model;
  if (!model) return false;

  const projectedTokens = projectedTokensFor(
    agent,
    nextUserText,
    contextWindow,
  );
  if (
    !shouldAutoCompactForProjectedTokens({ projectedTokens, contextWindow })
  ) {
    return false;
  }

  // Nothing to summarize / no room to improve.
  if (agent.state.messages.length < 4) return false;

  const before = estimateEffectiveRequestTokens({
    state: agent.state,
    tools: agent.state.tools,
    contextWindow,
    maxOutputTokens: model.maxTokens ?? 0,
  }).effectiveTokens;

  const outcome = await runCompact();

  // Verify compaction actually reduced the transcript below the trigger. If
  // not (compaction failed or the tail is dominated by one huge message), the
  // caller must not proceed with an over-budget request.
  const after = estimateEffectiveRequestTokens({
    state: agent.state,
    tools: agent.state.tools,
    contextWindow,
    maxOutputTokens: model.maxTokens ?? 0,
  }).effectiveTokens;

  // Verify compaction actually reduced the transcript below the trigger. If
  // not (compaction failed or the tail is dominated by one huge message), the
  // caller must not proceed with an over-budget request.
  return (
    outcome.changed &&
    outcome.reason !== "failed" &&
    after < before &&
    projectedTokensFor(agent, nextUserText, contextWindow) <=
      getCompactionThresholds(contextWindow).hardTriggerTokens
  );
}

/**
 * Mid-turn compaction check, run from `Agent.prepareNextTurn` after each
 * completed tool batch. A single tool-heavy turn can overflow a small context
 * window before the next between-prompt check would ever fire (#566).
 *
 * Returns a replacement loop context when compaction rewrote the transcript,
 * so the in-flight run continues from the compacted history.
 */
export async function maybeAutoCompactBeforeContinuation(args: {
  agent: Agent;
  enabled: boolean;
  contextWindow: number;
  runCompact: () => Promise<CompactionOutcome>;
}): Promise<AgentLoopTurnUpdate | undefined> {
  const { agent, enabled, contextWindow, runCompact } = args;

  if (!enabled) return undefined;

  const messages = agent.state.messages;
  const last = messages.at(-1);
  // Only act when another continuation request is coming (tool loop).
  if (!last || last.role !== "toolResult") return undefined;

  const model = agent.state.model;
  if (!model) return undefined;

  const { effectiveTokens } = estimateEffectiveRequestTokens({
    state: agent.state,
    tools: agent.state.tools,
    contextWindow,
    maxOutputTokens: model.maxTokens ?? 0,
  });
  if (
    !shouldAutoCompactForProjectedTokens({
      projectedTokens: effectiveTokens,
      contextWindow,
    })
  ) {
    return undefined;
  }

  if (messages.length < 4) return undefined;

  const before = agent.state.messages;
  const beforeTokens = estimateEffectiveRequestTokens({
    state: agent.state,
    tools: agent.state.tools,
    contextWindow,
    maxOutputTokens: model.maxTokens ?? 0,
  }).effectiveTokens;

  const outcome = await runCompact();
  if (agent.state.messages === before) return undefined;

  const afterTokens = estimateEffectiveRequestTokens({
    state: agent.state,
    tools: agent.state.tools,
    contextWindow,
    maxOutputTokens: model.maxTokens ?? 0,
  }).effectiveTokens;

  // Only continue the loop from the compacted context when it actually fit.
  if (
    !outcome.changed ||
    outcome.reason === "failed" ||
    afterTokens >= beforeTokens
  ) {
    return undefined;
  }

  return {
    context: {
      systemPrompt: agent.state.systemPrompt,
      messages: [...agent.state.messages],
      tools: [...agent.state.tools],
    },
  };
}
