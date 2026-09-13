# Compaction (`/compact`)

Pi for Office runs each chat inside the selected model’s **context window** (e.g. Claude Opus 4.6: 200k tokens). When the conversation grows too large, requests will fail with errors like **“prompt is too long”**.

`/compact` is the manual escape hatch: it **replaces older history with a structured summary**, while keeping the most recent work verbatim.

## Automatic triggers

Auto-compaction (enabled by default, `compaction.enabled`) uses the shared hard budgets from `getCompactionThresholds` and fires at three points:

1. **Before a queued prompt** — projected context (current estimate + the new prompt) exceeds the hard trigger. The projection uses the **request-facing (shaping-aware)** estimate, and the request is only dispatched after compaction verifiably reduced the transcript below the trigger.
2. **Mid-turn, between tool-loop continuations** — after each completed tool batch, so a single tool-heavy turn can’t overflow a small context window before the next between-prompt check. The in-flight run continues from the compacted history.
3. **Post-run recovery (Pi parity)** — when a run settles in an error tail:
   - **context overflow** (error text, silent usage overflow like z.ai, MiMo length-stop, or a stream-end at high usage) → drop the failed assistant message, compact, retry **once**;
   - **transient stream failures** (e.g. `Stream ended without finish_reason`, which pi-ai classifies as retryable via the `ended without` pattern) → drop the failed message, exponential-backoff retry (`retry.enabled`, `retry.maxRetries`, `retry.baseDelayMs`; defaults true / 3 / 2000, mirroring pi coding-agent).

Deterministic quota/billing errors are never retried (`isRetryableAssistantError` already excludes them). When budgets are exhausted, the error stays in the transcript with an actionable banner offering **Retry / Compact & retry / New session**.

When auto-compaction is disabled, overflow errors surface a banner suggesting `/compact`, scoping the request, or a larger-context model — instead of the raw provider error.

> Note: Compaction permanently drops older messages from the session (except what’s captured in the summary). If you need a full transcript, run `/export` **before** compaction.

## When to use `/compact`

- Context usage is trending high (see the status bar).
- You hit a hard failure like `prompt is too long` / context window exceeded.
- The model starts “forgetting” early decisions.

## What `/compact` does

At a high level, compaction produces a new message list:

1. A single **compaction summary** message (structured markdown)
2. A **recent tail** of messages kept as-is

Everything older than the kept tail is removed.

### 1) Find the compaction boundary

If the session already contains a `compactionSummary` message, we treat it as the boundary:

- we summarize only messages **after** the last summary
- and we **update** the existing summary instead of stacking multiple summaries

### 2) Choose what to keep vs summarize

The compaction engine (`src/compaction/engine.ts`) selects a **valid turn boundary** (user/assistant messages only — never inside a tool batch) so the kept tail stays roughly the last **`keepRecentTokens` (~20,000)** tokens, and then **bounds the kept tail to a hard budget** with an escalation ladder:

1. reduce effective `keepRecentTokens` (halved toward a 2k floor);
2. preview oversized kept tool results (keep the most recent few verbatim), truncate oversized tool-call argument JSON;
3. drop old images/thinking blocks;
4. drop oldest kept turns, preserving tool-call/result pairing and always keeping the latest user message and latest assistant+results cycle;
5. truncate oversized text payloads with an explicit marker.

This is the fix for the historical stuck-loop: an oversized **trailing tool batch** (dozens of parallel results × 50KB) can no longer be kept wholesale, and `/compact` can never report “Nothing to compact” while the transcript is still over budget — it trims instead.

### 3) Generate the structured summary

We serialize the to-be-summarized messages into a plain transcript:

- `[User]: ...`
- `[Assistant]: ...`
- `[Assistant thinking]: ...` (when present)
- `[Tool result <name>]: ...`

Then we ask the current model to produce a structured checkpoint (or update the previous summary).

`/compact` supports optional arguments:

- `/compact focus on formulas and sheet names`

Those arguments are appended to the prompt as an “Additional focus”.

Compaction also runs a lightweight **memory nudge** on the messages being summarized:

- if older user messages include explicit memory cues (for example, "remember this" / "don't forget"), Pi shows a reminder toast before summarization
- the summarizer gets extra focus instructions to call out durable memory in **Critical Context** and distinguish:
  - behavioral preferences/rules → `instructions`
  - factual memory → `notes/` or workbook-scoped notes

### 4) Replace the session messages

After summarization succeeds, we **verify the assembled transcript fits** (`contextWindow - reserveTokens`), escalating the tail trim if needed, then atomically replace the in-memory session with:

- `compactionSummary` (new/updated)
- `...keptTail`

The committed transcript is **persisted immediately** (`persistence.saveSession({ force: true })` via the `onCommitted` hook), so a pane reload can’t resurrect the oversized pre-compaction history. In the UI, the summary is rendered as a collapsible “compact” card.

## What the model sees after compaction

`compactionSummary` is a custom UI message type, but it *is* included in LLM context.

Internally it’s converted into a `user` message like:

```text
The conversation history before this point was compacted into the following summary:

<summary>
...
</summary>
```

So the next turn’s prompt contains:

- the summary (as a single user message)
- plus the kept recent tail

## Token budgeting (implementation details)

We mirror Pi’s compaction defaults:

- `reserveTokens`: **16,384** (clamped for smaller context windows)
- `keepRecentTokens`: **20,000** (also clamped)
- summary generation `maxTokens`: `floor(0.8 * reserveTokens)` (then clamped to `model.maxTokens`)

Estimates are **script-aware** (CJK/Hangul/Kana count ~1 token per char instead of chars/4), include tool schemas, per-message framing, and size-aware image costs. The status bar uses the request-facing (shaping-aware) effective estimate as the primary meter, with the raw persisted-history count shown when it diverges.

The summarization prompt has a **total serialized budget** (chars), so the summarizer itself can never overflow: per-message limits are applied first, then oldest messages are omitted with an explicit marker. The summarizer call is wrapped in `retryAssistantCall` (transient retries), and a `length` stop is treated as failure (partial summaries never become checkpoints). If the request is still “too long”, one retry with aggressive truncation runs.

## What happens when context is >100%

If the status bar shows **>100%** context usage, normal chat turns are likely to fail.

The engine keeps the transcript under budget by trimming (see above), so `/compact` always makes progress — it never reports “Nothing to compact” while over budget. The boundary safety net (`src/auth/context-trim.ts`) additionally trims the outgoing request (previews old tool results, drops old images, truncates text) so **no over-budget request is ever dispatched** — including the summarizer call.

If compaction is exhausted even after escalation (e.g. a single message alone exceeds the window), an actionable banner offers **Retry / Compact & retry / New session** and explains the dominating item instead of silently failing.

## Small context windows (custom gateways)

Models behind custom gateways often have much smaller windows (32k–65k) than the 128k–200k mainstream models Pi’s defaults are tuned for. Recommendations:

- **Set “Max context tokens” accurately** in the gateway settings (`/settings` → custom gateway). This single value drives all context budgets: compaction thresholds, tool-output caps, and how many recent tool results are kept verbatim. Overstating it causes hard 400s; understating it wastes capacity.
- **Budgets scale automatically** below a 128k window: tool-output truncation caps shrink linearly (e.g. ~25KB instead of 50KB at 65k, floor 8KB / 200 lines), and history shaping keeps fewer verbatim tool results (3 at 65k, floor 2).
- **Scope your prompts.** Select the relevant range or name the sheet you care about instead of asking for whole-workbook analysis; large multi-sheet reads consume a small window very quickly.
- **Start new chats per task** (`/new`) rather than carrying long histories across unrelated tasks.

## Status bar interaction

The status bar context % is computed from:

- the **last successful assistant usage** (includes cached tokens like `cacheRead/cacheWrite`), plus
- an estimate for any messages after that usage

After `/compact`, last usage becomes stale (because the message list is rewritten). The UI detects this and temporarily estimates context usage from scratch until a new assistant response provides fresh usage.

## Adaptive effective window

Provider catalogs can overstate the real cap (e.g. an OpenCode Go contributor tier). After an overflow-like failure with a request estimate ≥50% of the claimed window, the runtime records the failure estimate and uses a **reduced effective window** (`×0.9`) for subsequent compaction triggers and guards — visible in the status bar. It resets on model switch.

## Where this is implemented

- `/compact` implementation: `src/commands/builtins/export.ts`
- Compaction engine (cut selection, tail trim ladder, fit verification): `src/compaction/engine.ts`
- Failure classification (overflow/transient/fatal): `src/compaction/failure-classification.ts`
- Pi-parity post-run recovery loop: `src/compaction/run-recovery.ts`
- Adaptive effective window: `src/compaction/adaptive-window.ts`
- Provider-boundary context trim: `src/auth/context-trim.ts`
- Summary message type: `src/messages/compaction.ts`
- Injecting summary into LLM context: `src/messages/convert-to-llm.ts`
- UI rendering of the summary card: `src/ui/message-renderers.ts`
- Context % display + stale-usage fallback: `src/taskpane/status-bar.ts`
