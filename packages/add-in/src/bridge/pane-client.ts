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
  BRIDGE_PROTOCOL_VERSION,
  type ClientMessage,
  type HelloMessage,
  type OfficeHostApp,
  type ServerMessage,
  type WelcomeMessage,
} from "@dieulc/pi-office-protocol";
import { defaultBridgePort } from "../config/local-service-defaults.js";
import type { OfficeOpExecutor } from "./ops.js";

const WELCOME_TIMEOUT_MS = 5_000;
const PING_INTERVAL_MS = 25_000;

/** Why a connect attempt failed — the UI turns this into an actionable line. */
export type BridgeConnectFailureReason =
  | "refused" // nothing listening (Pi not running, wrong port/URL)
  | "timeout" // no welcome within the handshake budget
  | "protocol-mismatch" // hello/welcome protocol versions differ
  | "closed"; // server closed before the handshake finished

/** Typed connect failure so callers can distinguish cause, not just message. */
export class BridgeConnectError extends Error {
  readonly reason: BridgeConnectFailureReason;

  constructor(reason: BridgeConnectFailureReason, message: string) {
    super(message);
    this.name = "BridgeConnectError";
    this.reason = reason;
  }
}

export interface PaneBridgeClientOptions {
  host: OfficeHostApp;
  /** Op registry used to execute incoming tool calls. */
  registry: ReadonlyMap<string, OfficeOpExecutor>;
  url?: string;
  clientName?: string;
  paneId?: string;
  /**
   * Op ids this pane can execute, advertised in `hello` so the Pi server can
   * gate tool calls and expose only the supported tools. When omitted the
   * server treats this pane as a legacy 0.2.x client (v1 op set only).
   */
  ops?: readonly string[];
  /** Catalog version the ops were derived from (documented in the shared catalog). */
  catalogVersion?: number;
}

export interface PaneBridgeClientCallbacks {
  /** The Pi agent finished answering a prompt this pane sent. */
  onAssistantFinal?(text: string, messageId?: string): void;
  /** Non-office tool activity from the Pi agent (bash, read, …). */
  onActivity?(activity: {
    tool: string;
    status: string;
    summary?: string;
  }): void;
  onServerError?(error: { code: string; message: string }): void;
  onStatusChange?(connected: boolean): void;
  /** Server metadata from the welcome frame (version/capabilities). */
  onWelcome?(welcome: WelcomeMessage): void;
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
  private readonly ops: readonly string[] | undefined;
  private readonly catalogVersion: number | undefined;

  private ws: WebSocket | null = null;
  private pingTimer: ReturnType<typeof setInterval> | null = null;
  private connected = false;
  private connectResolve: (() => void) | null = null;
  private connectReject: ((error: Error) => void) | null = null;
  private welcomeTimer: ReturnType<typeof setTimeout> | null = null;
  private connectionGen = 0;
  /** Last pong timestamp — the watchdog uses this to spot half-open sockets. */
  private lastPongAt = 0;

  constructor(
    options: PaneBridgeClientOptions,
    callbacks: PaneBridgeClientCallbacks = {},
  ) {
    this.host = options.host;
    this.clientName = options.clientName ?? "pi-for-office";
    this.paneId = options.paneId ?? defaultPaneId();
    this.registry = options.registry;
    this.callbacks = callbacks;
    this.ops = options.ops;
    this.catalogVersion = options.catalogVersion;
    const port = extractPort(options.url) ?? defaultBridgePort();
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
          this.failConnect(
            new BridgeConnectError(
              "timeout",
              `bridge: no welcome from ${this.url} within ${WELCOME_TIMEOUT_MS / 1000}s`,
            ),
          );
        }
      }, WELCOME_TIMEOUT_MS);

      ws.addEventListener("error", () => {
        if (gen === this.connectionGen) {
          this.failConnect(
            new BridgeConnectError(
              "refused",
              `bridge: connection to ${this.url} failed (is the Pi process running?)`,
            ),
          );
        }
      });
      ws.addEventListener("close", () => {
        if (gen === this.connectionGen) {
          this.failConnect(
            new BridgeConnectError(
              "closed",
              "bridge: server closed before welcome",
            ),
          );
        }
      });
      ws.addEventListener("open", () => {
        const hello: HelloMessage = {
          type: "hello",
          protocolVersion: BRIDGE_PROTOCOL_VERSION,
          host: this.host,
          clientName: this.clientName,
          paneId: this.paneId,
        };
        if (this.ops !== undefined) hello.ops = [...this.ops];
        if (this.catalogVersion !== undefined)
          hello.catalogVersion = this.catalogVersion;
        this.send(hello);
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
    const payload: { type: "status"; model?: string; provider?: string } = {
      type: "status",
    };
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
      this.lastPongAt = Date.now();
      if (this.welcomeTimer !== null) {
        clearTimeout(this.welcomeTimer);
        this.welcomeTimer = null;
      }
      const resolve = this.connectResolve;
      this.connectResolve = null;
      this.connectReject = null;
      this.callbacks.onWelcome?.(msg);
      this.callbacks.onStatusChange?.(true);
      this.startHeartbeat();
      resolve?.();
      return;
    }

    if (msg.type === "error") {
      if (msg.code === "protocol_mismatch") {
        this.failConnect(
          new BridgeConnectError("protocol-mismatch", msg.message),
        );
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
        // Liveness proof: a half-open socket that eats pings is dead weight.
        this.lastPongAt = Date.now();
        break;
    }
  }

  private handleToolCall(
    id: string,
    tool: string,
    args: Record<string, DynamicValue>,
  ): void {
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
    this.lastPongAt = Date.now();
    this.pingTimer = setInterval(() => {
      if (!this.send({ type: "ping", ts: Date.now() })) {
        // Socket is gone but the close event may never come — call it dead.
        this.killConnection("ping send failed");
        return;
      }
      if (Date.now() - this.lastPongAt > PING_INTERVAL_MS * 2) {
        // No pong for two full intervals → half-open socket (e.g. laptop
        // hibernated). Kill it so the manager reconnects instead of showing
        // a stale "Connected".
        this.killConnection(
          `no pong for ${Math.round((Date.now() - this.lastPongAt) / 1000)}s`,
        );
      }
    }, PING_INTERVAL_MS);
  }

  /** Force-close a stale socket and report the disconnect to the manager. */
  private killConnection(_why: string): void {
    if (this.pingTimer !== null) {
      clearInterval(this.pingTimer);
      this.pingTimer = null;
    }
    const ws = this.ws;
    this.ws = null;
    if (ws) {
      try {
        ws.close(1001, "pane: stale connection");
      } catch {
        // Closing an already-dead socket can throw in some runtimes.
      }
      // Node's `ws` has terminate(); browsers don't — best effort.
      (ws as WebSocket & { terminate?: () => void }).terminate?.();
    }
    this.setConnected(false);
  }

  private send(message: ClientMessage): boolean {
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
