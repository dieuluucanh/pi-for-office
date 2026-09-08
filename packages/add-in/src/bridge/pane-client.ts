/**
 * Pane bridge client — runs inside the Office task pane (or any WebSocket
 * capable runtime) and connects to the local Pi bridge server.
 *
 * Responsibilities:
 *  - `hello`/`welcome` handshake so the Pi side knows which host app is open.
 *  - Execute incoming `tool_call`s via the host op registry and answer with
 *    `tool_result`.
 *  - Forward user prompts as `user_message` and surface the assistant's reply
 *    (`agent_message` final) and Pi tool activity through callbacks.
 *  - Keep the connection alive with `ping`s.
 *
 * The Office.js executors are intentionally injected (see registry.ts) so this
 * class is testable outside Office.
 */

import {
  BRIDGE_DEFAULT_PORT,
  BRIDGE_PROTOCOL_VERSION,
  type ClientMessage,
  type OfficeHostApp,
  type ServerMessage,
} from "@dieulc/pi-office-protocol";
import type { OfficeOpExecutor } from "./ops.js";

const WELCOME_TIMEOUT_MS = 5_000;
const PING_INTERVAL_MS = 25_000;

export interface PaneBridgeClientOptions {
  host: OfficeHostApp;
  /** Op registry used to execute incoming tool calls. */
  registry: ReadonlyMap<string, OfficeOpExecutor>;
  url?: string;
  clientName?: string;
  paneId?: string;
}

export interface PaneBridgeClientCallbacks {
  /** The Pi agent finished answering a prompt this pane sent. */
  onAssistantFinal?(text: string, messageId?: string): void;
  /** Non-office tool activity from the Pi agent (bash, read, …). */
  onActivity?(activity: { tool: string; status: string; summary?: string }): void;
  onServerError?(error: { code: string; message: string }): void;
  onStatusChange?(connected: boolean): void;
}

function defaultPaneId(): string {
  const cryptoObj = globalThis.crypto;
  if (cryptoObj && typeof cryptoObj.randomUUID === "function") {
    return cryptoObj.randomUUID();
  }
  return `pane-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

export class PaneBridgeClient {
  readonly host: OfficeHostApp;
  readonly paneId: string;
  readonly clientName: string;

  private readonly registry: ReadonlyMap<string, OfficeOpExecutor>;
  private readonly callbacks: PaneBridgeClientCallbacks;
  private readonly url: string;

  private ws: WebSocket | null = null;
  private pingTimer: ReturnType<typeof setInterval> | null = null;
  private connected = false;
  private connectResolve: (() => void) | null = null;
  private connectReject: ((error: Error) => void) | null = null;
  private welcomeTimer: ReturnType<typeof setTimeout> | null = null;
  private connectionGen = 0;

  constructor(options: PaneBridgeClientOptions, callbacks: PaneBridgeClientCallbacks = {}) {
    this.host = options.host;
    this.clientName = options.clientName ?? "pi-for-office";
    this.paneId = options.paneId ?? defaultPaneId();
    this.registry = options.registry;
    this.callbacks = callbacks;
    const port = extractPort(options.url) ?? BRIDGE_DEFAULT_PORT;
    this.url = options.url ?? `ws://127.0.0.1:${port}`;
  }

  get isConnected(): boolean {
    return this.connected;
  }

  /** Connect, handshake, and start heartbeats. Resolves after `welcome`. */
  connect(): Promise<void> {
    if (this.ws && this.connected) return Promise.resolve();
    this.disconnect();

    const gen = ++this.connectionGen;

    return new Promise<void>((resolve, reject) => {
      this.connectResolve = resolve;
      this.connectReject = reject;

      const ws = new WebSocket(this.url);
      this.ws = ws;

      this.welcomeTimer = setTimeout(() => {
        if (gen === this.connectionGen) {
          this.failConnect(new Error(`bridge: no welcome from ${this.url} within ${WELCOME_TIMEOUT_MS / 1000}s`));
        }
      }, WELCOME_TIMEOUT_MS);

      ws.addEventListener("error", () => {
        if (gen === this.connectionGen) {
          this.failConnect(new Error(`bridge: connection to ${this.url} failed`));
        }
      });
      ws.addEventListener("close", () => {
        if (gen === this.connectionGen) {
          this.failConnect(new Error("bridge: server closed before welcome"));
        }
      });
      ws.addEventListener("open", () => {
        this.send({
          type: "hello",
          protocolVersion: BRIDGE_PROTOCOL_VERSION,
          host: this.host,
          clientName: this.clientName,
          paneId: this.paneId,
        });
      });
      ws.addEventListener("message", (event) => {
        if (gen === this.connectionGen) {
          this.handleFrame(String(event.data));
        }
      });
    });
  }

  /** Close the connection and stop heartbeats. */
  disconnect(): void {
    if (this.pingTimer !== null) {
      clearInterval(this.pingTimer);
      this.pingTimer = null;
    }
    if (this.welcomeTimer !== null) {
      clearTimeout(this.welcomeTimer);
      this.welcomeTimer = null;
    }
    if (this.ws) {
      this.ws.close(1000, "pane closing");
      this.ws = null;
    }
    this.connectReject = null;
    this.connectResolve = null;
    this.setConnected(false);
  }

  /** Send a user prompt to the Pi agent. */
  sendPrompt(text: string): boolean {
    return this.send({ type: "user_message", text });
  }

  /** Report the model/provider the pane's local agent is using (informational). */
  reportStatus(model?: string, provider?: string): boolean {
    const payload: { type: "status"; model?: string; provider?: string } = { type: "status" };
    if (model !== undefined) payload.model = model;
    if (provider !== undefined) payload.provider = provider;
    return this.send(payload);
  }

  /* ── internals ─────────────────────────────────────────────────────── */

  private handleFrame(raw: string): void {
    let msg: ServerMessage;
    try {
      msg = JSON.parse(raw) as ServerMessage;
    } catch {
      return;
    }

    if (msg.type === "welcome") {
      this.connected = true;
      if (this.welcomeTimer !== null) {
        clearTimeout(this.welcomeTimer);
        this.welcomeTimer = null;
      }
      const resolve = this.connectResolve;
      this.connectResolve = null;
      this.connectReject = null;
      this.callbacks.onStatusChange?.(true);
      this.startHeartbeat();
      resolve?.();
      return;
    }

    if (msg.type === "error") {
      if (msg.code === "protocol_mismatch") {
        this.failConnect(new Error(msg.message));
        return;
      }
      this.callbacks.onServerError?.({ code: msg.code, message: msg.message });
      return;
    }

    switch (msg.type) {
      case "tool_call": {
        this.handleToolCall(msg.id, msg.tool, msg.args ?? {});
        break;
      }
      case "agent_message":
        if (msg.kind === "final") {
          this.callbacks.onAssistantFinal?.(msg.text, msg.messageId);
        }
        break;
      case "tool_activity":
        this.callbacks.onActivity?.({
          tool: msg.tool,
          status: msg.status,
          ...(msg.summary === undefined ? {} : { summary: msg.summary }),
        });
        break;
      case "pong":
        break;
    }
  }

  private handleToolCall(id: string, tool: string, args: Record<string, unknown>): void {
    const executor = this.registry.get(tool);
    if (!executor) {
      this.send({
        type: "tool_result",
        id,
        ok: false,
        text: "",
        error: `Unknown op "${tool}" on this ${this.host} pane.`,
      } satisfies ClientMessage);
      return;
    }
    void executor(args).then((outcome) => {
      const message: ClientMessage = {
        type: "tool_result",
        id,
        ok: !outcome.isError,
        text: outcome.text,
        ...(outcome.isError
          ? { error: outcome.text }
          : outcome.details === undefined
            ? {}
            : { details: outcome.details }),
      };
      this.send(message);
    });
  }

  private failConnect(error: Error): void {
    const reject = this.connectReject;
    this.connectResolve = null;
    this.connectReject = null;
    if (this.welcomeTimer !== null) {
      clearTimeout(this.welcomeTimer);
      this.welcomeTimer = null;
    }
    this.setConnected(false);
    // Keep the socket for the server-close case; the caller decides to retry.
    if (reject) reject(error);
  }

  private setConnected(value: boolean): void {
    if (this.connected === value) return;
    this.connected = value;
    this.callbacks.onStatusChange?.(value);
  }

  private startHeartbeat(): void {
    this.pingTimer = setInterval(() => {
      this.send({ type: "ping", ts: Date.now() });
    }, PING_INTERVAL_MS);
  }

  private send(message: object): boolean {
    const ws = this.ws;
    if (!ws || ws.readyState !== WebSocket.OPEN) return false;
    try {
      ws.send(JSON.stringify(message));
      return true;
    } catch {
      return false;
    }
  }
}

function extractPort(url: string | undefined): number | null {
  if (!url) return null;
  try {
    const parsed = new URL(url);
    const port = Number.parseInt(parsed.port, 10);
    return Number.isFinite(port) && port > 0 ? port : null;
  } catch {
    return null;
  }
}
