# Agent stops after a thinking block (reasoning gateways)

Symptoms, root cause, and fixes for runs that end silently after a long
reasoning stretch — the transcript shows "thought", the input button flips
from Stop back to Enter, and no error banner appears.

## Symptoms

- The agent finishes (or is cut off mid-way through) a thinking block, then
  the run ends without doing any work.
- The visible thinking chain is interrupted mid-sentence.
- Asking it to continue makes it think and stall again in a loop.
- Happens most with verbose reasoning models on custom OpenAI-compatible
  gateways (synthetic-style reasoners); less with providers that interleave
  reasoning with visible text.

## Root cause

Three things line up:

1. **Output budget.** Models without registry metadata get a default output
   budget of 32 768 tokens (raised from 4 096 in this fix; older catalogs and
   custom providers restore with the raised budget automatically). Reasoning
   models spend the same budget on thinking alone — the `openai-completions`
   path applies no separate thinking headroom.
2. **No thinking headroom on the OpenAI path.** pi-ai adds thinking headroom
   for Anthropic-style APIs, but OpenAI-compatible gateways consume the same
   `max_completion_tokens` budget for reasoning and answer.
3. **Silent turn end.** When reasoning truncates with `finish_reason: "length"`
   and the message carries no tool calls, the agent loop treats it as a normal
   turn end — no error, no retry. The recovery layer now detects this shape
   (see below).

A secondary vector: proxy-mediated streams that idle >5 minutes between chunks
(long thinking stretches) used to be killed by undici's default idle timeouts
and ended with a clean connection close. The production proxy now disables
those limits for proxied streams.

## What the add-in does about it

- **Stall auto-continue.** A run that ends with only thinking (no text, no tool
  calls) is detected by shape, dropped, and continued automatically — up to 2
  attempts. When the stall ended on `length`, the output budget escalates ×2
  per attempt (capped at 64k and by the context window; the escalation is
  restored when the loop settles). A busy indicator shows
  "Output truncated — continuing automatically (attempt/max)…".
- **Budget-unsupported fallback.** If a provider rejects the output limit
  itself (400-style "max_tokens too large"), the budget is halved (floor 1024)
  and the reduced budget persists for the session.
- **Recovery decoupled from auto-compact.** Stall and stream-error recovery run
  regardless of the auto-compact setting; only the compaction step of
  overflow-like failures stays gated.
- **Diagnostics.** Every settled run logs one line to the taskpane console
  (`[pi] Run settled: stopReason=… outputTokens=… outputBudget=… model=…`) so
  stalls are diagnosable from the console.

## How to verify

1. Sideload the add-in and select a reasoning model on a custom gateway.
2. Give it a task that previously stalled.
3. Expected: when thinking truncates, a busy indicator appears and the run
   continues with a doubled budget — no manual "continue" prompts.

If stalls persist: check the console line for `outputTokens` — a value pinned
at the budget means the provider caps output lower than requested; the
budget-unsupported fallback handles the 400-style rejection, but silent caps
are provider behavior (try a smaller request or a different model).

## Related

- [docs/compaction.md](compaction.md) — context protection and recovery loop.
- [docs/central-proxy.md](central-proxy.md) — the production proxy helper.
