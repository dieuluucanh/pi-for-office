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
  MIN_SUPPORTED_PROTOCOL_VERSION,
  isProtocolVersionSupported,
  CATALOG_VERSION,
  LEGACY_V1_OPS,
  OFFICE_CATALOG_BY_OP,
  OFFICE_TARGET_PARAM,
  nextCallId,
  parseClientMessage,
  type BridgeCapability,
  type ClientMessage,
  type DocumentIdentitySource,
  type OfficeHostApp,
  type ServerMessage,
  type ToolResultMessage,
} from "./protocol.js";

export interface AttachedPane {
  ws: WebSocket;
  host: OfficeHostApp;
  paneId: string;
  clientName: string;
  /**
   * Protocol version the pane declared in `hello` (validated at the handshake).
   * v3+ panes receive streamed `agent_message` deltas; v2 panes get final
   * replies only, because they would drop a delta.
   */
  protocolVersion: number;
  connectedAt: number;
  lastSeen: number;
  model?: string;
  provider?: string;
  /**
   * Ops this pane advertised in `hello` (validated against the server catalog
   * and normalized to `<host>.<op>` ids). `null` means a legacy pane that did
   * not advertise an `ops` list — only the v1 op set is allowed for it.
   */
  ops: readonly string[] | null;
  /** Catalog version the pane derived its ops from, when advertised. */
  catalogVersion: number | null;
  /** Count of advertised ops dropped because they are unknown to this server. */
  opsIgnoredCount: number;
  /**
   * Stable local id for the pane's document (`url_sha256:…` /
   * `wps_path_sha256:…`), or null when the host cannot provide one (unsaved).
   */
  documentId: string | null;
  /** Best-effort display name for the document. */
  documentName: string | null;
  /** Where the document identity came from. */
  documentSource: DocumentIdentitySource;
  /**
   * Routing key for this pane: the document id when known, otherwise an
   * ephemeral per-connection id. Two windows showing the same file share a
   * `targetId`, which resolution treats as one logical document.
   */
  targetId: string;
  /** True when this pane's window/webview last reported focus. */
  focused: boolean;
  /** Timestamp of the last focus gain (ms since epoch; 0 = never reported). */
  focusUpdatedAt: number;
  /** Last reported visibility (`document.visibilityState === "visible"`). */
  visible: boolean;
}

export interface BridgeServerHandlers {
  /** A user typed a prompt in the add-in sidebar. */
  onUserMessage(text: string, pane: AttachedPane): void;
  /**
   * A pane asked to abort the run answering `promptId` (or whatever run is in
   * flight for that pane when it is omitted). Protocol v3; optional so a host
   * without an abort path still compiles.
   */
  onAbortRequest?(promptId: string | undefined, pane: AttachedPane): void;
  /**
   * A pane attached or detached. Lets the host (Pi) refresh its status line
   * immediately instead of waiting for the next session event.
   */
  onPanesChanged?(panes: readonly AttachedPane[]): void;
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

/**
 * Split the protocol v2 `document` selector out of a tool-call args record so
 * the pane never receives it. Non-string values are treated as "no target".
 */
function splitTarget(args: Record<string, unknown>): {
  target: string | undefined;
  rest: Record<string, unknown>;
} {
  const { [OFFICE_TARGET_PARAM]: raw, ...rest } = args;
  return {
    target: typeof raw === "string" ? raw : undefined,
    rest,
  };
}

interface PendingCall {
  resolve(result: CallOfficeToolResult): void;
  reject(error: Error): void;
  timer: ReturnType<typeof setTimeout>;
}

export class OfficeBridgeServer {
  /**
   * Capabilities advertised in `welcome` and `GET /health`. These describe the
   * **server**, so they are advertised to every client regardless of the
   * protocol version the pane declared.
   */
  static readonly CAPABILITIES: readonly BridgeCapability[] = [
    "http-health",
    "agent-deltas",
    "client-abort",
  ];

  private readonly port: number;
  private readonly handlers: BridgeServerHandlers;
  private readonly serverName: string;
  private readonly serverVersion: string;
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
    serverVersion?: string;
    piVersion?: string | null;
    handlers: BridgeServerHandlers;
  }) {
    this.port = options.port;
    this.serverName = options.serverName ?? "pi-office-bridge";
    this.serverVersion = options.serverVersion ?? "unknown";
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
    this.notifyPanesChanged();

    return new Promise((resolve) => {
      wss.close(() => resolve());
      if (httpServer) {
        httpServer.close(() => resolve());
      }
    });
  }

  /**
   * Proxy one Office.js operation to the pane that owns the requested document.
   *
   * The target comes from `args.document`: missing or `"active"` resolves to the
   * focused document, otherwise it must be an id, a unique document name, or a
   * TUI `#name` mention. The selector is stripped before the call is forwarded.
   * Rejects when no pane of `host` is attached, the target is ambiguous or
   * unknown, the call times out, the pane disconnects, or the signal aborts.
   */
  callOfficeTool(
    host: OfficeHostApp,
    op: string,
    args: Record<string, unknown>,
    signal?: AbortSignal,
    timeoutMs: number = TOOL_CALL_TIMEOUT_MS,
  ): Promise<CallOfficeToolResult> {
    const { target, rest } = splitTarget(args);

    let pane: AttachedPane;
    try {
      pane = this.resolvePane(host, target);
    } catch (error) {
      return Promise.reject(error);
    }

    const opId = `${host}.${op}`;
    // Capability gate: a pane may only execute ops it advertised (or, for
    // legacy panes that advertise nothing, only the v1 op set).
    if (pane.ops !== null) {
      if (!pane.ops.includes(opId)) {
        return Promise.reject(
          new Error(
            `office-bridge: the attached ${host} pane does not advertise "${opId}" ` +
              `(it supports ${pane.ops.length} ops). ` +
              "Reload the pi-for-office add-in to enable this tool.",
          ),
        );
      }
    } else if (!LEGACY_V1_OPS.includes(opId)) {
      return Promise.reject(
        new Error(
          `office-bridge: "${opId}" requires a newer add-in. ` +
            "The attached pane is a legacy client (no capability list); update " +
            "pi-for-office and reload the document to enable this tool.",
        ),
      );
    }

    const id = nextCallId("tool");
    const message: ServerMessage = {
      type: "tool_call",
      id,
      tool: opId,
      args: rest,
      ...(pane.documentId !== null ? { documentId: pane.documentId } : {}),
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
      const all = this.attachedPanes();
      const panes = all.map((p) => ({
        host: p.host,
        paneId: p.paneId,
        clientName: p.clientName,
        document: {
          id: p.documentId,
          name: p.documentName,
          source: p.documentSource,
        },
        targetId: p.targetId,
        focused: p.focused,
        visible: p.visible,
        // >1 means the same document is open in several windows.
        windows: all.filter(
          (q) => q.host === p.host && q.targetId === p.targetId,
        ).length,
        connectedAt: p.connectedAt,
        lastSeen: p.lastSeen,
        model: p.model,
        provider: p.provider,
        ops: p.ops,
        catalogVersion: p.catalogVersion,
        protocolVersion: p.protocolVersion,
        opsIgnoredCount: p.opsIgnoredCount,
      }));

      this.writeHttp(
        res,
        200,
        {
          ok: true,
          service: this.serverName,
          serverVersion: this.serverVersion,
          capabilities: OfficeBridgeServer.CAPABILITIES,
          protocolVersion: BRIDGE_PROTOCOL_VERSION,
          piVersion: this.piVersion,
          port: this.actualPort,
          uptimeMs: Date.now() - this.startedAt,
          catalogVersion: CATALOG_VERSION,
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

  /**
   * Resolve a `document` selector to exactly one pane of `host`.
   *
   * Order: focused document → the only attached document → exact id
   * (`documentId`/`targetId`) → case-insensitive exact name → unique name
   * substring. Ambiguity or no match throws an actionable error listing the
   * open documents instead of guessing.
   */
  private resolvePane(host: OfficeHostApp, target?: string): AttachedPane {
    const panes = this.panes.filter((p) => p.host === host);
    if (panes.length === 0) {
      throw new Error(
        `office-bridge: no ${host} workbook/document is attached. ` +
          "Open the document in the Office add-in (pi-for-office) to enable this tool.",
      );
    }

    // A TUI `#name` mention is a valid selector; normalize it away.
    const needle = (target ?? "").trim().replace(/^#/u, "").trim();
    if (needle === "" || needle.toLowerCase() === "active") {
      return this.pickFocusedOrOnly(panes, host, "active");
    }

    const exactId = panes.filter(
      (p) => p.targetId === needle || p.documentId === needle,
    );
    if (exactId.length > 0)
      return this.pickFocusedOrOnly(exactId, host, needle);

    const lower = needle.toLowerCase();
    const exactName = panes.filter(
      (p) => (p.documentName ?? "").toLowerCase() === lower,
    );
    if (exactName.length > 0) {
      // The same name on several *different* documents is ambiguous even when
      // one is focused; only multiple windows of one document may use focus.
      const distinct = new Set(exactName.map((p) => p.targetId)).size;
      if (distinct > 1) {
        throw new Error(
          `office-bridge: "${needle}" matches several ${host} documents with the same name. ` +
            `Open: ${this.describePanes(host)}. ` +
            `Pass ${OFFICE_TARGET_PARAM}=<id> to choose one.`,
        );
      }
      return this.pickFocusedOrOnly(exactName, host, needle);
    }

    const partial = panes.filter((p) =>
      (p.documentName ?? "").toLowerCase().includes(lower),
    );
    if (partial.length === 1) {
      const only = partial[0];
      if (only !== undefined) return only;
    }
    if (partial.length > 1) {
      throw new Error(
        `office-bridge: "${needle}" matches several ${host} documents. ` +
          `Open: ${this.describePanes(host)}. ` +
          `Pass ${OFFICE_TARGET_PARAM}=<exact name|id> or open the one you mean.`,
      );
    }

    throw new Error(
      `office-bridge: no ${host} document matches "${needle}". ` +
        `Open: ${this.describePanes(host)}. ` +
        `Pass ${OFFICE_TARGET_PARAM}=<name|id|"active"> (or mention #<name> in the prompt).`,
    );
  }

  /**
   * Pick the focused pane among candidates, else the only one. When several
   * candidates share a document id (the same file open in two windows) and none
   * is focused, fail instead of silently editing the wrong window.
   */
  private pickFocusedOrOnly(
    candidates: readonly AttachedPane[],
    host: OfficeHostApp,
    needle: string,
  ): AttachedPane {
    const focused = [...candidates]
      .filter((p) => p.focused)
      .sort((a, b) => b.focusUpdatedAt - a.focusUpdatedAt);
    const first = focused[0];
    if (first !== undefined) return first;
    if (candidates.length === 1) {
      const only = candidates[0];
      if (only !== undefined) return only;
    }
    throw new Error(
      `office-bridge: several ${host} documents are open and none is focused, ` +
        `so "${needle}" is ambiguous. Open: ${this.describePanes(host)}. ` +
        `Focus the document you mean, or pass ${OFFICE_TARGET_PARAM}=<name|id>.`,
    );
  }

  /** Human-readable list of the open documents of `host`, for error messages. */
  private describePanes(host: OfficeHostApp): string {
    const panes = this.panes.filter((p) => p.host === host);
    return panes
      .map((p) => {
        const windows = panes.filter((q) => q.targetId === p.targetId).length;
        const name = p.documentName ?? `Untitled (${p.targetId})`;
        const notes = [
          p.focused ? "focused" : null,
          windows > 1 ? `${windows} windows` : null,
        ].filter((note): note is string => note !== null);
        const suffix = notes.length > 0 ? `, ${notes.join(", ")}` : "";
        return `"${name}" (${p.documentId ?? p.targetId}${suffix})`;
      })
      .join(", ");
  }

  /**
   * Mark `pane` as focused, demoting every other pane of the same host so at
   * most one document per host app is "active" at a time.
   */
  private markPaneFocused(pane: AttachedPane, visible?: boolean): void {
    for (const other of this.panes) {
      if (other !== pane && other.host === pane.host && other.focused) {
        other.focused = false;
      }
    }
    pane.focused = true;
    pane.focusUpdatedAt = Date.now();
    if (visible !== undefined) pane.visible = visible;
  }

  /**
   * Normalize a pane's advertised ops against the server catalog: keep only
   * ops that exist in the catalog and belong to the pane's host; return null
   * when the pane advertised nothing (legacy client).
   */
  private normalizeAdvertisedOps(
    host: OfficeHostApp,
    advertised: readonly string[] | undefined,
  ): readonly string[] | null {
    if (advertised === undefined || advertised.length === 0) return null;
    const known = new Set<string>();
    for (const op of advertised) {
      const entry = OFFICE_CATALOG_BY_OP.get(op);
      if (entry && entry.host === host) known.add(op);
    }
    return known.size > 0 ? [...known] : null;
  }

  /** Count ops the pane advertised that this server dropped as unknown/host-mismatched. */
  private countIgnoredOps(
    host: OfficeHostApp,
    advertised: readonly string[] | undefined,
  ): number {
    if (advertised === undefined) return 0;
    return advertised.filter((op) => {
      const entry = OFFICE_CATALOG_BY_OP.get(op);
      return entry === undefined || entry.host !== host;
    }).length;
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
          if (!isProtocolVersionSupported(msg.protocolVersion)) {
            this.sendError(
              ws,
              "protocol_mismatch",
              `bridge: client protocol v${msg.protocolVersion} is not supported ` +
                `(server v${BRIDGE_PROTOCOL_VERSION}, accepting v${MIN_SUPPORTED_PROTOCOL_VERSION}–v${BRIDGE_PROTOCOL_VERSION})`,
            );
            ws.close(1008, "protocol mismatch");
            return;
          }
          pane = {
            ws,
            host: msg.host,
            paneId: msg.paneId,
            clientName: msg.clientName,
            protocolVersion: msg.protocolVersion,
            connectedAt: Date.now(),
            lastSeen: Date.now(),
            // Normalize advertised ops against the server catalog: entries the
            // server does not know (or that belong to a different host) are
            // dropped and counted so the operator can see the mismatch.
            ops: this.normalizeAdvertisedOps(msg.host, msg.ops),
            catalogVersion: msg.catalogVersion ?? null,
            opsIgnoredCount: this.countIgnoredOps(msg.host, msg.ops),
            documentId: msg.document.id,
            documentName: msg.document.name,
            documentSource: msg.document.source,
            targetId: msg.document.id ?? `ephemeral:${msg.paneId}`,
            focused: false,
            focusUpdatedAt: 0,
            visible: true,
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
            serverVersion: this.serverVersion,
            capabilities: [...OfficeBridgeServer.CAPABILITIES],
          });
          this.notifyPanesChanged();
          break;
        }
        case "ping": {
          this.sendToPane(ws, { type: "pong", ts: msg.ts });
          break;
        }
        case "abort": {
          // Protocol v3: the pane's Stop button. Only meaningful after a hello
          // (there is no pane to attribute the request to before that).
          if (!pane) {
            this.sendError(ws, "bad_message", "bridge: abort before hello");
            return;
          }
          pane.lastSeen = Date.now();
          this.handlers.onAbortRequest?.(msg.promptId, pane);
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
          // Typing in the pane is the strongest signal of which document the
          // user means, so it doubles as a focus report.
          this.markPaneFocused(pane);
          const text = msg.text?.trim();
          if (text) this.handlers.onUserMessage(text, pane);
          break;
        }
        case "pane_state": {
          if (!pane) return;
          pane.lastSeen = Date.now();
          if (msg.visible !== undefined) pane.visible = msg.visible;
          if (msg.focused === true) {
            this.markPaneFocused(pane, msg.visible);
          } else if (msg.focused === false) {
            pane.focused = false;
          }
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

  /** Notify the host that the attached-pane set changed. Never throws. */
  private notifyPanesChanged(): void {
    try {
      this.handlers.onPanesChanged?.(this.attachedPanes());
    } catch (error) {
      console.error(
        `[office-bridge] onPanesChanged handler failed: ${String(error)}`,
      );
    }
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

    this.notifyPanesChanged();
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

  sendToPane(pane: WebSocket | AttachedPane, message: ServerMessage): boolean {
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
