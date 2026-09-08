/**
 * Bridge protocol — shared wire format between the pi-for-office task pane
 * (client) and this Pi extension (server).
 *
 * The server lives inside a local Pi process. The task pane connects to it
 * over WebSocket at `ws://127.0.0.1:<port>`.
 *
 * Two flows are supported:
 *
 * 1. **Tool proxy (Pi → pane)** — the Pi agent calls an `office_*` tool; the
 *    extension forwards a `tool_call` to the attached pane for the matching
 *    host app; the pane executes the Office.js operation and answers with a
 *    `tool_result`.
 *
 * 2. **Pane-driven chat (pane → Pi)** — the user types in the add-in sidebar;
 *    the pane forwards a `user_message`; the extension injects it into the Pi
 *    session and streams the assistant's reply back as `agent_message`.
 *
 * Every connection starts with a `hello` (client) / `welcome` (server) pair so
 * both sides know which Office host app is attached.
 */

/** Default TCP port the bridge server listens on (loopback only). */
export const BRIDGE_DEFAULT_PORT = 38617;

/** Protocol version. Bump on breaking message-shape changes. */
export const BRIDGE_PROTOCOL_VERSION = 1;

/** Office host applications the add-in can attach. */
export type OfficeHostApp = "excel" | "word" | "powerpoint";

/** Every office tool name is namespaced by host: "excel.read_range" etc. */
export type OfficeToolId = string;

/* ── Client → Server ─────────────────────────────────────────────────── */

export interface HelloMessage {
  type: "hello";
  protocolVersion: number;
  host: OfficeHostApp;
  clientName: string;
  paneId: string;
}

export interface PingMessage {
  type: "ping";
  ts: number;
}

/** Result of a proxied Office.js tool call executed by the pane. */
export interface ToolResultMessage {
  type: "tool_result";
  id: string;
  ok: boolean;
  /** Markdown/text shown to the LLM. */
  text: string;
  /** Structured details (may be large; used by the agent tool result). */
  details?: unknown;
  error?: string;
}

/** User typed a prompt in the add-in sidebar; ask Pi to answer it. */
export interface UserMessageMessage {
  type: "user_message";
  text: string;
}

/** Pane notifies the server it finished applying a message (optional). */
export interface ClientStatusMessage {
  type: "status";
  model?: string;
  provider?: string;
}

export type ClientMessage =
  | HelloMessage
  | PingMessage
  | ToolResultMessage
  | UserMessageMessage
  | ClientStatusMessage;

/* ── Server → Client ─────────────────────────────────────────────────── */

export interface WelcomeMessage {
  type: "welcome";
  protocolVersion: number;
  piVersion: string | null;
  serverName: string;
}

export interface PongMessage {
  type: "pong";
  ts: number;
}

/** Pi agent asks the pane to execute one Office.js operation. */
export interface ToolCallMessage {
  type: "tool_call";
  id: string;
  tool: OfficeToolId;
  args: Record<string, unknown>;
}

/** A chunk of the assistant's reply (streamed, display-only). */
export interface AgentDeltaMessage {
  type: "agent_message";
  kind: "delta";
  text: string;
}

/** Final assistant reply for a pane-initiated prompt. */
export interface AgentFinalMessage {
  type: "agent_message";
  kind: "final";
  text: string;
  messageId?: string;
}

/** Non-office tool activity (bash, read, …) so the pane can show progress. */
export interface ToolActivityMessage {
  type: "tool_activity";
  tool: string;
  status: "start" | "end" | "error";
  summary?: string;
}

export interface ServerErrorMessage {
  type: "error";
  code: string;
  message: string;
}

export type ServerMessage =
  | WelcomeMessage
  | PongMessage
  | ToolCallMessage
  | AgentDeltaMessage
  | AgentFinalMessage
  | ToolActivityMessage
  | ServerErrorMessage;

/* ── Utilities ───────────────────────────────────────────────────────── */

let callCounter = 0;

/** Monotonic id generator for tool_call / tool_result correlation. */
export function nextCallId(prefix = "c"): string {
  callCounter += 1;
  return `${prefix}-${Date.now().toString(36)}-${callCounter.toString(36)}`;
}

/** Validate an incoming client frame. Returns the parsed message or throws. */
export function parseClientMessage(raw: string): ClientMessage {
  const data: unknown = JSON.parse(raw);
  if (typeof data !== "object" || data === null) {
    throw new Error("bridge: expected a JSON object message");
  }
  const msg = data as Record<string, unknown>;
  if (typeof msg.type !== "string") {
    throw new Error("bridge: message missing 'type'");
  }
  switch (msg.type) {
    case "hello": {
      const m = data as HelloMessage;
      if (!["excel", "word", "powerpoint"].includes(m.host)) {
        throw new Error(`bridge: hello with unknown host '${String(m.host)}'`);
      }
      return m;
    }
    case "ping":
    case "tool_result":
    case "user_message":
    case "status":
      return data as ClientMessage;
    default:
      throw new Error(`bridge: unknown client message type '${String(msg.type)}'`);
  }
}
