/**
 * Pane-prompt ↔ reply correlation — pure, Pi-free bookkeeping.
 *
 * Bridge mode runs ONE shared Pi conversation, so a pane prompt must be
 * correlated with the assistant reply that answers it. Verified against the Pi
 * runtime: the `input` event fires *before* a message is queued (so it cannot
 * identify the turn), `before_agent_start` does not re-fire for queued
 * follow-ups, and `message_start` with `role: "user"` fires when the message is
 * actually delivered. Correlation therefore happens at delivery time by
 * matching the delivered user text against the prompts this bridge injected.
 *
 * Keeping this module free of Pi and WebSocket imports makes the matching and
 * TTL rules unit-testable without a live session.
 */

/** How long an injected prompt may wait for delivery before it is dropped. */
export const DEFAULT_PROMPT_TTL_MS = 5 * 60_000;

/** One pane-originated prompt injected into the shared Pi conversation. */
export interface PanePromptEntry<TPane> {
 /** The pane that sent it; replies are routed only to this pane. */
 pane: TPane;
 /** Correlation id echoed on the final `agent_message`. */
 promptId: string;
 /** The exact text injected into Pi (document tag included). */
 sentText: string;
}

interface PendingEntry<TPane> {
 entry: PanePromptEntry<TPane>;
 enqueuedAt: number;
}

/**
 * FIFO of injected pane prompts plus the prompt currently bound to a delivered
 * user message (i.e. the turn being answered).
 */
export class PanePromptTracker<TPane> {
 private readonly pending: PendingEntry<TPane>[] = [];
 private current: PanePromptEntry<TPane> | null = null;
 private readonly ttlMs: number;

 constructor(ttlMs: number = DEFAULT_PROMPT_TTL_MS) {
  this.ttlMs = ttlMs;
 }

 /** Queue a prompt that was just injected into the session. */
 enqueue(entry: PanePromptEntry<TPane>, now: number = Date.now()): void {
  this.prune(now);
  this.pending.push({ entry, enqueuedAt: now });
 }

 /**
  * Bind a delivered user message to the oldest pending pane prompt.
  *
  * Returns the bound entry when `text` is the prompt this bridge injected
  * (exact match, or a prefix match if another extension transformed the tail),
  * otherwise null — a TUI/RPC prompt, which must not receive a pane reply.
  */
 bindDeliveredUserMessage(
  text: string,
  now: number = Date.now(),
 ): PanePromptEntry<TPane> | null {
  this.prune(now);
  const head = this.pending[0];
  if (head === undefined) return null;
  if (head.entry.sentText !== text && !text.startsWith(head.entry.sentText)) {
   return null;
  }
  this.pending.shift();
  this.current = head.entry;
  return head.entry;
 }

 /** The pane prompt the current turn is answering, if any. */
 active(): PanePromptEntry<TPane> | null {
  return this.current;
 }

 /** Finish the active prompt and return it. */
 complete(): PanePromptEntry<TPane> | null {
  const done = this.current;
  this.current = null;
  return done;
 }

 /** Drop every queued and active prompt (session shutdown / abort). */
 clear(): void {
  this.pending.length = 0;
  this.current = null;
 }

 /** Number of prompts still waiting to be delivered. */
 get pendingCount(): number {
  return this.pending.length;
 }

 /** Drop queued prompts that were never delivered within the TTL. */
 private prune(now: number): void {
  while (this.pending.length > 0) {
   const head = this.pending[0];
   if (head === undefined || now - head.enqueuedAt <= this.ttlMs) return;
   this.pending.shift();
  }
 }
}

/**
 * Leading line prepended to every pane prompt so the shared conversation shows
 * which document a message came from (and the model can reuse the id as a
 * `document` target).
 */
export function formatDocumentTag(target: {
 host: string;
 documentName: string | null;
 documentId: string | null;
}): string {
 const label = target.documentName ?? "untitled";
 const id = target.documentId === null ? "" : ` id=${target.documentId}`;
 return `[office:${target.host} "${label}"${id}]`;
}
