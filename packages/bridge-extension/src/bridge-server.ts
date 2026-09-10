/**
 * Bridge server — a loopback-only WebSocket server inside the Pi process.
 *
 * The pi-for-office task pane connects here. This module owns the socket
 * lifecycle, the attached-pane registry, and the pending tool-call correlation
 * map. It has no Pi knowledge; the extension entry point (`index.ts`) wires it
 * to the Pi session.
 */

import { WebSocketServer, WebSocket } from "ws";
import type {
  IncomingMessage,
  Server as HttpServer,
  ServerResponse,
} from "node:http";
import { createServer } from "node:http";
import type { AddressInfo } from "node:net";

import {
  BRIDGE_PROTOCOL_VERSION,
  nextCallId,
  parseClientMessage,
  type ClientMessage,
  type OfficeHostApp,
  type ServerMessage,
  type ToolResultMessage,
} from "./protocol.js";

export interface AttachedPane {
  ws: WebSocket;
  host: OfficeHostApp;
  paneId: string;
  clientName: string;
  connectedAt: number;
  lastSeen: number;
  model?: string;
  provider?: string;
}

export interface BridgeServerHandlers {
  /** A user typed a prompt in the add-in sidebar. */
  onUserMessage(text: string, pane: AttachedPane): void;
}

export interface CallOfficeToolResult {
  text: string;
  details?: unknown;
}

const HELLO_TIMEOUT_MS = 10_000;
const TOOL_CALL_TIMEOUT_MS = 120_000;
const HEARTBEAT_INTERVAL_MS = 30_000;
const MAX_TEXT_CHARS = 50_000;
const MAX_DETAILS_BYTES = 1_000_000;

/** Extra browser origins for the /health CORS gate, from env (comma-separated). */
function parseExtraOrigins(): string[] {
  const raw = process.env.PI_OFFICE_BRIDGE_ALLOWED_ORIGINS;
  if (!raw) return [];
  return raw
    .split(/[,\s]+/u)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

interface PendingCall {
  resolve(result: CallOfficeToolResult): void;
  reject(error: Error): void;
  timer: ReturnType<typeof setTimeout>;
}

export class OfficeBridgeServer {
  private readonly port: number;
  private readonly handlers: BridgeServerHandlers;
  private readonly serverName: string;
  private readonly piVersion: string | null;
  private readonly startedAt = Date.now();

  private httpServer: HttpServer | null = null;
  private wss: WebSocketServer | null = null;
  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;

  private readonly panes: AttachedPane[] = [];
  private readonly pending = new Map<string, PendingCall>();

  constructor(options: {
    port: number;
    serverName?: string;
    piVersion?: string | null;
    handlers: BridgeServerHandlers;
  }) {
    this.port = options.port;
    this.serverName = options.serverName ?? "pi-office-bridge";
    this.piVersion = options.piVersion ?? null;
    this.handlers = options.handlers;
  }

  get isRunning(): boolean {
    return this.wss !== null;
  }

  get actualPort(): number | null {
    const addr = this.httpServer?.address();
    return typeof addr === "object" && addr !== null
      ? (addr as AddressInfo).port
      : null;
  }

  /** Pane list copy (ordered by most recent connection first). */
  attachedPanes(): readonly AttachedPane[] {
    return [...this.panes].sort((a, b) => b.connectedAt - a.connectedAt);
  }

  /** Start listening. Idempotent. Resolves once the socket is bound. */
  start(): Promise<void> {
    if (this.wss) return Promise.resolve();

    const httpServer = createServer((req, res) =>
      this.handleHttpRequest(req, res),
    );
    const wss = new WebSocketServer({
      server: httpServer,
      maxPayload: 16 * 1024 * 1024,
    });

    this.httpServer = httpServer;
    this.wss = wss;

    wss.on("connection", (ws) => this.handleConnection(ws));
    wss.on("error", (error) => {
      console.error(`[office-bridge] server error: ${String(error)}`);
    });

    return new Promise((resolve, reject) => {
      const onError = (error: Error) => {
        cleanup();
        reject(error);
      };
      const cleanup = () => {
        httpServer.off("listening", onListening);
        httpServer.off("error", onError);
      };
      const onListening = () => {
        cleanup();
        this.startHeartbeat();
        resolve();
      };
      httpServer.once("listening", onListening);
      httpServer.once("error", onError);
      httpServer.listen(this.port, "127.0.0.1");
    });
  }

  /** Close the server and drop all panes. Idempotent. */
  stop(): Promise<void> {
    if (!this.wss) return Promise.resolve();

    const wss = this.wss;
    const httpServer = this.httpServer;
    this.wss = null;
    this.httpServer = null;

    if (this.heartbeatTimer !== null) {
      clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }

    for (const call of this.pending.values()) {
      clearTimeout(call.timer);
      call.reject(new Error("office-bridge: server stopped"));
    }
    this.pending.clear();

    for (const pane of this.panes) {
      pane.ws.close(1001, "bridge shutting down");
    }
    this.panes.length = 0;

    return new Promise((resolve) => {
      wss.close(() => resolve());
      if (httpServer) {
        httpServer.close(() => resolve());
      }
    });
  }

  /**
   * Proxy one Office.js operation to the most recent pane attached for `host`.
   * Resolves with the pane's tool_result, rejects when no pane is attached, the
   * call times out, the pane disconnects, or the caller's signal aborts.
   */
  callOfficeTool(
    host: OfficeHostApp,
    op: string,
    args: Record<string, unknown>,
    signal?: AbortSignal,
    timeoutMs: number = TOOL_CALL_TIMEOUT_MS,
  ): Promise<CallOfficeToolResult> {
    const pane = this.findPane(host);
    if (!pane) {
      return Promise.reject(
        new Error(
          `office-bridge: no ${host} workbook/document is attached. ` +
            "Open the document in the Office add-in (pi-for-office) to enable this tool.",
        ),
      );
    }

    const id = nextCallId("tool");
    const message: ServerMessage = {
      type: "tool_call",
      id,
      tool: `${host}.${op}`,
      args,
    };

    return new Promise<CallOfficeToolResult>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(
          new Error(
            `office-bridge: ${op} timed out after ${timeoutMs / 1000}s`,
          ),
        );
      }, timeoutMs);

      this.pending.set(id, { resolve, reject, timer });

      const onAbort = () => {
        const call = this.pending.get(id);
        if (!call) return;
        clearTimeout(call.timer);
        this.pending.delete(id);
        reject(new Error("office-bridge: tool call aborted by the agent"));
      };

      if (signal?.aborted) {
        onAbort();
        return;
      }
      signal?.addEventListener("abort", onAbort, { once: true });

      if (!this.sendToPane(pane, message)) {
        clearTimeout(timer);
        this.pending.delete(id);
        reject(
          new Error(
            "office-bridge: pane disconnected before the tool call was sent",
          ),
        );
      }
    });
  }

  /** Push a message to every attached pane. */
  broadcast(message: ServerMessage): void {
    for (const pane of this.panes) {
      this.sendToPane(pane, message);
    }
  }

  /* ── HTTP (health/diagnostics, loopback-only) ─────────────────────── */

  /**
   * Browser origins allowed to read this loopback HTTP surface. The add-in
   * task pane runs at these origins (dev Vite server + hosted GitHub Pages).
   * Extend with the PI_OFFICE_BRIDGE_ALLOWED_ORIGINS env var
   * (comma-separated) when the add-in is hosted elsewhere.
   */
  private static readonly ALLOWED_ORIGINS: ReadonlySet<string> = new Set([
    "https://localhost:3141",
    "https://pi-excel.localhost",
    "https://dieuluucanh.github.io",
    ...parseExtraOrigins(),
  ]);

  private static resolveAllowOrigin(req: IncomingMessage): string | null {
    const origin = req.headers.origin;
    if (typeof origin !== "string" || origin.trim().length === 0) {
      // Non-browser client (curl / tests / server tooling) — no CORS gate.
      return "*";
    }
    if (OfficeBridgeServer.ALLOWED_ORIGINS.has(origin)) return origin;
    return null; // Unknown browser origin → omit header; browser blocks.
  }

  /**
   * Minimal loopback HTTP surface used by the add-in's "Test connection"
   * probe and by curl. WebSocket upgrades are handled by `ws` at the server
   * level; ordinary requests (GET /health, OPTIONS preflight) land here.
   * The endpoint is unauthenticated but exposes only bridge metadata.
   */
  private handleHttpRequest(req: IncomingMessage, res: ServerResponse): void {
    const urlRaw = req.url ?? "/";
    let url: URL;
    try {
      url = new URL(urlRaw, "http://127.0.0.1");
    } catch {
      this.writeHttp(res, 400, { ok: false, error: "bad_request" }, req);
      return;
    }
    const allowOrigin = OfficeBridgeServer.resolveAllowOrigin(req);

    if (req.method === "OPTIONS") {
      this.writeHttp(res, 204, null, req, allowOrigin);
      return;
    }

    if (req.method === "GET" && url.pathname === "/health") {
      const panes = this.attachedPanes().map((p) => ({
        host: p.host,
        paneId: p.paneId,
        clientName: p.clientName,
        connectedAt: p.connectedAt,
        lastSeen: p.lastSeen,
        model: p.model,
        provider: p.provider,
      }));

      this.writeHttp(
        res,
        200,
        {
          ok: true,
          service: this.serverName,
          protocolVersion: BRIDGE_PROTOCOL_VERSION,
          piVersion: this.piVersion,
          port: this.actualPort,
          uptimeMs: Date.now() - this.startedAt,
          panes,
        },
        req,
        allowOrigin,
      );
      return;
    }

    this.writeHttp(
      res,
      404,
      { ok: false, error: "not_found" },
      req,
      allowOrigin,
    );
  }

  private writeHttp(
    res: ServerResponse,
    status: number,
    body: unknown,
    _req?: IncomingMessage,
    allowOrigin: string | null = "*",
  ): void {
    const headers: Record<string, string> = {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      // Loopback-only server. The Private-Network header keeps the probe
      // working from the hosted GitHub Pages origin (public → localhost).
      "Access-Control-Allow-Methods": "GET, OPTIONS",
      "Access-Control-Allow-Headers": "*",
      "Access-Control-Allow-Private-Network": "true",
    };
    if (allowOrigin !== null) {
      headers["Access-Control-Allow-Origin"] = allowOrigin;
      headers["Vary"] = "Origin";
    }
    res.writeHead(status, headers);
    if (status === 204 || body === null) {
      res.end();
      return;
    }
    res.end(JSON.stringify(body));
  }

  /* ── Internals ─────────────────────────────────────────────────────── */

  private findPane(host: OfficeHostApp): AttachedPane | null {
    const sorted = [...this.panes].sort(
      (a, b) => b.connectedAt - a.connectedAt,
    );
    return sorted.find((p) => p.host === host) ?? null;
  }

  private handleConnection(ws: WebSocket): void {
    let pane: AttachedPane | null = null;

    const helloTimer = setTimeout(() => {
      if (!pane) {
        ws.close(1008, "bridge: hello not received in time");
      }
    }, HELLO_TIMEOUT_MS);

    ws.on("message", (raw) => {
      let msg: ClientMessage;
      try {
        msg = parseClientMessage(raw.toString());
      } catch (error) {
        this.sendError(ws, "bad_message", String(error));
        return;
      }

      switch (msg.type) {
        case "hello": {
          if (msg.protocolVersion !== BRIDGE_PROTOCOL_VERSION) {
            this.sendError(
              ws,
              "protocol_mismatch",
              `bridge: client protocol v${msg.protocolVersion} does not match server v${BRIDGE_PROTOCOL_VERSION}`,
            );
            ws.close(1008, "protocol mismatch");
            return;
          }
          pane = {
            ws,
            host: msg.host,
            paneId: msg.paneId,
            clientName: msg.clientName,
            connectedAt: Date.now(),
            lastSeen: Date.now(),
          };
          // A pane reconnecting replaces any older pane with the same paneId.
          this.panes.splice(
            0,
            this.panes.length,
            ...this.panes.filter((p) => p.paneId !== pane!.paneId),
            pane,
          );
          clearTimeout(helloTimer);
          this.sendToPane(ws, {
            type: "welcome",
            protocolVersion: BRIDGE_PROTOCOL_VERSION,
            piVersion: this.piVersion,
            serverName: this.serverName,
          });
          break;
        }
        case "ping": {
          this.sendToPane(ws, { type: "pong", ts: msg.ts });
          break;
        }
        case "tool_result": {
          if (pane) pane.lastSeen = Date.now();
          this.handleToolResult(msg, ws);
          break;
        }
        case "user_message": {
          if (!pane) return;
          pane.lastSeen = Date.now();
          const text = msg.text?.trim();
          if (text) this.handlers.onUserMessage(text, pane);
          break;
        }
        case "status": {
          if (!pane) return;
          pane.lastSeen = Date.now();
          if (msg.model !== undefined) pane.model = msg.model;
          if (msg.provider !== undefined) pane.provider = msg.provider;
          break;
        }
      }
    });

    ws.on("close", () => {
      clearTimeout(helloTimer);
      if (pane) this.detachPane(pane);
    });

    ws.on("error", () => {
      clearTimeout(helloTimer);
      if (pane) this.detachPane(pane);
    });
  }

  private handleToolResult(msg: ToolResultMessage, ws: WebSocket): void {
    const call = this.pending.get(msg.id);
    if (!call) return;

    // Only the pane that received the call may answer it.
    const pane = this.panes.find((p) => p.ws === ws);
    if (!pane) {
      call.reject(
        new Error("office-bridge: pane disconnected before answering"),
      );
      return;
    }

    clearTimeout(call.timer);
    this.pending.delete(msg.id);

    if (!msg.ok) {
      call.reject(
        new Error(`office-bridge: ${msg.error ?? "office tool failed"}`),
      );
      return;
    }

    const text =
      msg.text.length > MAX_TEXT_CHARS
        ? `${msg.text.slice(0, MAX_TEXT_CHARS)}\n…[truncated: ${msg.text.length - MAX_TEXT_CHARS} chars]`
        : msg.text;

    let details: unknown = msg.details;
    if (details !== undefined) {
      const bytes = Buffer.byteLength(JSON.stringify(details));
      if (bytes > MAX_DETAILS_BYTES) {
        details = {
          truncated: true,
          note: `details exceeded ${MAX_DETAILS_BYTES} bytes`,
        };
      }
    }

    call.resolve({ text, details });
  }

  private detachPane(pane: AttachedPane): void {
    const idx = this.panes.findIndex((p) => p === pane);
    if (idx >= 0) this.panes.splice(idx, 1);

    // Reject any pending calls owned by this pane (best-effort: find calls and
    // reject them — tracked separately so we sweep all on disconnect).
    for (const [id, call] of this.pending) {
      clearTimeout(call.timer);
      call.reject(
        new Error(
          "office-bridge: pane disconnected while the tool was running",
        ),
      );
      this.pending.delete(id);
    }
  }

  private startHeartbeat(): void {
    this.heartbeatTimer = setInterval(() => {
      for (const pane of this.panes) {
        const alive = pane.ws.readyState === WebSocket.OPEN;
        if (!alive) {
          pane.ws.terminate();
          this.detachPane(pane);
        } else if (Date.now() - pane.lastSeen > HEARTBEAT_INTERVAL_MS * 3) {
          pane.ws.ping();
        }
      }
    }, HEARTBEAT_INTERVAL_MS);
  }

  private sendToPane(
    pane: WebSocket | AttachedPane,
    message: ServerMessage,
  ): boolean {
    const ws = pane instanceof WebSocket ? pane : pane.ws;
    if (ws.readyState !== WebSocket.OPEN) return false;
    try {
      ws.send(JSON.stringify(message));
      return true;
    } catch {
      return false;
    }
  }

  private sendError(ws: WebSocket, code: string, message: string): void {
    this.sendToPane(ws, { type: "error", code, message });
  }
}
