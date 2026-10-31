/**
 * pi-office-bridge — native Pi extension.
 *
 * Runs a loopback WebSocket server inside the local Pi process so the
 * pi-for-office task-pane add-in (Excel / Word / PowerPoint) can attach.
 *
 *  - Registers `office_<host>_<op>` tools the Pi agent can call; each call is
 *    proxied to the attached pane, which executes the Office.js operation.
 *  - Injects pane-typed prompts into the Pi session and streams the assistant
 *    reply back to the pane.
 *
 * The server is **off by default**: a session never claims the port until the
 * user runs `/office connect` (or an explicit `--office-bridge-autostart`).
 * That is what lets several terminal/Desktop sessions coexist without silently
 * fighting over 38617.
 *
 * Pure extension: uses only the public extension API — Pi core is never
 * touched, so Pi can be updated freely.
 */

import type {
  AgentToolResult,
  ExtensionAPI,
  ExtensionContext,
} from "@earendil-works/pi-coding-agent";
import { Type, type Static } from "typebox";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { CATALOG_VERSION, nextCallId } from "./protocol.js";
import type { AttachedPane, BridgeServerHandlers } from "./bridge-server.js";
import { OfficeBridgeServer } from "./bridge-server.js";
import { PanePromptTracker, formatDocumentTag } from "./pane-replies.js";
import {
  createDocumentMentionProvider,
  type DocumentMentionTarget,
  type MentionProvider,
} from "./document-mentions.js";
import {
  OFFICE_TOOL_DESCRIPTORS,
  hostForToolName,
  OFFICE_TOOL_NAMES,
  HOST_APP_LABEL,
} from "./office-tools.js";
import type { OfficeToolDescriptor } from "./office-tools.js";
import {
  OFFICE_LIST_TOOL,
  reconcileBridgeToolActivation,
  type PaneCapability,
} from "./active-tools.js";
import { createBridgeController, type BridgeController } from "./lifecycle.js";
import {
  officeCommandCompletions,
  officeCommandHelp,
  parseOfficeCommand,
} from "./commands.js";
import {
  FLAG_AUTOSTART,
  FLAG_PORT,
  resolveBridgeAutostart,
  resolveBridgePort,
} from "./config.js";

/** Which Pi version we run inside (shown in the pane's `welcome` frame). */
function piVersion(): string | null {
  return process.env.PI_VERSION ?? null;
}

/**
 * The bridge extension's own package version, read from the `package.json`
 * that ships next to this module. Falls back to `"unknown"` so a bundled or
 * relocated install never breaks the `welcome` frame.
 */
function serverVersion(): string {
  try {
    const here = dirname(fileURLToPath(import.meta.url));
    const raw = readFileSync(resolve(here, "..", "package.json"), "utf8");
    const pkg = JSON.parse(raw) as { version?: unknown };
    return typeof pkg.version === "string" && pkg.version.length > 0
      ? pkg.version
      : "unknown";
  } catch {
    return "unknown";
  }
}

export interface OfficeBridgeExtensionDeps {
  /** Test seam: build the server (the real factory is the default). */
  createServer?: (
    port: number,
    handlers: BridgeServerHandlers,
  ) => OfficeBridgeServer;
  /** Test seam: deterministic timestamps in the lifecycle state. */
  now?: () => number;
  /** Test seam: environment used for port/autostart precedence. */
  env?: Record<string, string | undefined>;
}

/**
 * Builds the extension. Exported so the test harness can drive a fake Pi API
 * without loading a real session.
 */
export function createOfficeBridgeExtension(
  pi: ExtensionAPI,
  deps: OfficeBridgeExtensionDeps = {},
): void {
  const env = deps.env ?? process.env;
  let currentCtx: ExtensionContext | null = null;
  let controller: BridgeController<OfficeBridgeServer> | null = null;

  const createServer =
    deps.createServer ??
    ((port, handlers) =>
      new OfficeBridgeServer({
        port,
        serverName: "pi-office-bridge",
        serverVersion: serverVersion(),
        piVersion: piVersion(),
        handlers,
      }));

  const activeServer = (): OfficeBridgeServer | null =>
    controller?.server() ?? null;
  const bridgeIsOn = (): boolean =>
    controller?.snapshot().state === "on" && activeServer() !== null;

  /**
   * Pane prompts injected into the shared conversation, bound to the turn
   * being answered at delivery time so replies go only to the originating pane.
   */
  const panePrompts = new PanePromptTracker<AttachedPane>();

  /** Short human label for a pane's document (status line, commands). */
  function paneDocumentLabel(pane: AttachedPane): string {
    const name = pane.documentName ?? `Untitled (${pane.targetId})`;
    return `${HOST_APP_LABEL[pane.host]} · "${name}"${pane.focused ? " (focused)" : ""}`;
  }

  /** Map an attached pane to the `#` mention dropdown's target shape. */
  function toMentionTarget(
    pane: AttachedPane,
    all: readonly AttachedPane[],
  ): DocumentMentionTarget {
    return {
      host: pane.host,
      hostLabel: HOST_APP_LABEL[pane.host],
      name: pane.documentName,
      documentId: pane.documentId,
      targetId: pane.targetId,
      focused: pane.focused,
      windows: all.filter(
        (q) => q.host === pane.host && q.targetId === pane.targetId,
      ).length,
    };
  }

  pi.registerFlag(FLAG_PORT, {
    description: `Port for the pi-office bridge server (default 38617)`,
    type: "string",
  });
  pi.registerFlag(FLAG_AUTOSTART, {
    description:
      "Start the pi-office bridge when the session opens (default: off; also PI_OFFICE_BRIDGE_AUTOSTART=1)",
    type: "boolean",
    default: false,
  });

  function resolvePort(): number {
    return resolveBridgePort((name) => pi.getFlag(name), env);
  }

  function resolveAutostart(): boolean {
    return resolveBridgeAutostart((name) => pi.getFlag(name), env);
  }

  function updateStatus(): void {
    const ui = currentCtx?.ui;
    if (!ui) return;
    const snapshot = controller?.snapshot();
    if (!snapshot || snapshot.state === "off") {
      ui.setStatus("office-bridge", undefined);
      return;
    }
    const port = snapshot.actualPort ?? snapshot.port;
    if (snapshot.state === "starting") {
      ui.setStatus("office-bridge", `office bridge starting on :${port}…`);
      return;
    }
    if (snapshot.state === "closing") {
      ui.setStatus("office-bridge", "office bridge closing…");
      return;
    }
    const panes = activeServer()?.attachedPanes() ?? [];
    if (panes.length === 0) {
      ui.setStatus(
        "office-bridge",
        `office bridge on :${port} — no app attached`,
      );
      return;
    }
    const labels = [...countByHost(panes)]
      .map(([label, count]) => (count > 1 ? `${label}×${count}` : label))
      .join(", ");
    ui.setStatus(
      "office-bridge",
      `office: ${labels} attached (bridge :${port})`,
    );
  }

  /** Count attached panes per host label (for the status line). */
  function countByHost(panes: readonly AttachedPane[]): Map<string, number> {
    const counts = new Map<string, number>();
    for (const pane of panes) {
      const label = HOST_APP_LABEL[pane.host];
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
    return counts;
  }

  let panesChangedTimer: ReturnType<typeof setTimeout> | null = null;

  /** Debounced status refresh triggered by pane attach/detach (~100 ms). */
  function scheduleStatusUpdate(): void {
    if (panesChangedTimer !== null) clearTimeout(panesChangedTimer);
    panesChangedTimer = setTimeout(() => {
      panesChangedTimer = null;
      updateStatus();
    }, 100);
  }

  /** Extract display text from a message content payload (user or assistant). */
  function flattenTextContent(content: unknown): string {
    if (typeof content === "string") return content;
    if (!Array.isArray(content)) return "";
    const parts: string[] = [];
    for (const part of content) {
      if (typeof part !== "object" || part === null) continue;
      const p = part as { type?: unknown; text?: unknown };
      if (p.type === "text" && typeof p.text === "string") parts.push(p.text);
      if (p.type === "thinking" && typeof p.text === "string")
        parts.push(p.text);
    }
    return parts.join("\n").trim();
  }

  function offlineToolError(): Error {
    return new Error(
      "office-bridge: server is not running. Run /office connect (or Connect in " +
        "the Documents panel), then try again.",
    );
  }

  function registerOfficeTool(descriptor: OfficeToolDescriptor): void {
    type Params = Static<typeof descriptor.parameters>;

    pi.registerTool({
      name: descriptor.name,
      label: descriptor.label,
      description: descriptor.description,
      promptSnippet: descriptor.promptSnippet,
      promptGuidelines: descriptor.promptGuidelines,
      parameters: descriptor.parameters,
      executionMode: "sequential",
      async execute(
        _toolCallId: string,
        params: Params,
        signal: AbortSignal | undefined,
        _onUpdate,
        _ctx,
      ): Promise<AgentToolResult<unknown>> {
        // SAFETY: `Params` is the TypeBox-derived shape of
        // `descriptor.parameters` (always an object); the bridge forwards it
        // verbatim as the Office.js args record, so this widening is sound.
        const args = params as unknown as Record<string, unknown>;
        const active = activeServer();
        if (!bridgeIsOn() || active === null) throw offlineToolError();
        const result = await active.callOfficeTool(
          descriptor.host,
          descriptor.op,
          args,
          signal,
        );
        return {
          content: [{ type: "text", text: result.text }],
          details: result.details,
        };
      },
    });
  }

  function registerAllOfficeTools(): void {
    for (const descriptor of OFFICE_TOOL_DESCRIPTORS) {
      registerOfficeTool(descriptor);
    }
  }

  /**
   * Discovery tool: lists the documents currently attached to the bridge so the
   * agent can choose a `document` target. It is active only while the bridge is
   * on, so an offline session does not advertise an Office capability.
   */
  function registerOfficeListDocumentsTool(): void {
    const parameters = Type.Object({
      host: Type.Optional(
        Type.String({
          description:
            'Only list documents of this host app ("excel", "word", or "powerpoint").',
        }),
      ),
    });
    type Params = Static<typeof parameters>;

    pi.registerTool({
      name: OFFICE_LIST_TOOL,
      label: "List Open Office Documents",
      description:
        "List every Office document attached to the bridge (Excel, Word, PowerPoint) " +
        "with its document id, name, host app, and whether it is focused. Use this to " +
        "choose the `document` target for the office_* tools.",
      promptSnippet: "List the Office documents attached to the bridge",
      parameters,
      executionMode: "sequential",
      async execute(
        _toolCallId: string,
        params: Params,
        _signal: AbortSignal | undefined,
        _onUpdate,
        _ctx,
      ): Promise<AgentToolResult<unknown>> {
        const active = activeServer();
        if (!bridgeIsOn() || active === null) throw offlineToolError();
        const all = active.attachedPanes();
        const host = typeof params.host === "string" ? params.host : undefined;
        const panes = all.filter((p) => host === undefined || p.host === host);
        if (panes.length === 0) {
          return {
            content: [
              {
                type: "text",
                text:
                  host === undefined
                    ? "No Office documents are attached to the bridge."
                    : `No ${host} documents are attached to the bridge.`,
              },
            ],
            details: undefined,
          };
        }
        const rows = panes.map((p) => {
          const windows = all.filter(
            (q) => q.host === p.host && q.targetId === p.targetId,
          ).length;
          const name = p.documentName ?? `Untitled (${p.targetId})`;
          return `| ${name} | \`${p.documentId ?? p.targetId}\` | ${HOST_APP_LABEL[p.host]} | ${p.focused ? "yes" : "no"} | ${windows} |`;
        });
        const table = [
          "| document | id | host | focused | windows |",
          "|---|---|---|---|---|",
          ...rows,
        ].join("\n");
        return { content: [{ type: "text", text: table }], details: undefined };
      },
    });
  }

  /**
   * Reconcile the Pi active tool set with the lifecycle and the attached panes.
   * Off removes the whole bridge family (including the discovery tool); on adds
   * the discovery tool and exactly the ops the attached panes advertise.
   */
  function reconcileActiveTools(): void {
    const panes: readonly AttachedPane[] = activeServer()?.attachedPanes() ?? [];
    const capabilities: PaneCapability[] = panes.map((pane) => ({
      host: pane.host,
      ops: pane.ops,
    }));
    const next = reconcileBridgeToolActivation(pi.getActiveTools(), {
      enabled: bridgeIsOn(),
      panes: capabilities,
    });
    pi.setActiveTools(next);
  }

  /** Stable per-turn pane-awareness block appended to the system prompt. */
  function paneContextBlock(): string | null {
    if (!bridgeIsOn()) return null;
    const active = activeServer();
    if (active === null) return null;
    const panes = active.attachedPanes();
    if (panes.length === 0) return null;

    const lines = panes.map((p) => {
      const name = p.documentName ?? `Untitled (${p.targetId})`;
      const id = p.documentId ?? p.targetId;
      const marks = [
        p.focused ? "focused" : null,
        p.ops === null ? "legacy client (v1 ops)" : null,
      ].filter((mark): mark is string => mark !== null);
      return `- ${HOST_APP_LABEL[p.host]} · "${name}" — id \`${id}\`${marks.length > 0 ? ` (${marks.join(", ")})` : ""}`;
    });
    return (
      "\n\n## Attached Office documents\n" +
      lines.join("\n") +
      "\n\nEvery office_* tool accepts an optional `document` argument: a document id from the list above, " +
      'a unique document name, or "active" (the default) for the focused document. When several documents ' +
      "of one host are open and none is focused, calls without a document target fail with the candidate " +
      "list — pass the target explicitly instead of guessing. A user's `#name` mention selects that " +
      "document. Call office_list_documents if this list may be stale.\n\n" +
      "These tools edit the attached documents directly — use them, and never emit HTML for Word. " +
      "Formatting support depends on the host and the APIs it exposes: each tool's schema lists exactly " +
      "the properties it accepts. If a property is not in the schema, say it is not available rather " +
      "than claiming support, and call office_word_get_capabilities when a request depends on a Word " +
      "feature you are unsure about."
    );
  }

  /* ── server handlers ───────────────────────────────────────────────── */

  const handlers: BridgeServerHandlers = {
    onUserMessage: (text, pane) => {
      if (!currentCtx) return;
      const promptId = nextCallId("prompt");
      const sentText = `${formatDocumentTag(pane)}\n${text}`;
      panePrompts.enqueue({ pane, promptId, sentText });
      updateStatus();
      // followUp: if Pi is idle the message is delivered immediately and
      // triggers a turn; if a turn is running it is queued until it settles.
      pi.sendUserMessage(sentText, { deliverAs: "followUp" });
    },
    // Protocol v3: a pane's Stop button asks us to abort the run that is
    // answering its prompt. Only honoured while a pane prompt is actually in
    // flight (never to kill a TUI-initiated turn) and, when the pane named a
    // promptId, only for that prompt.
    onAbortRequest: (promptId, pane) => {
      const prompt = panePrompts.active();
      if (!prompt || prompt.pane !== pane) return;
      if (promptId !== undefined && promptId !== prompt.promptId) return;
      clearDeltaState();
      currentCtx?.abort();
      updateStatus();
    },
    // Keep the TUI status line live and the active tool set in sync:
    // opening an app activates that host's office tools; closing it
    // deactivates them.
    onPanesChanged: () => {
      scheduleStatusUpdate();
      reconcileActiveTools();
    },
  };

  function ensureController(): BridgeController<OfficeBridgeServer> {
    if (controller) return controller;
    controller = createBridgeController({
      createServer: (port) => createServer(port, handlers),
      defaultPort: resolvePort(),
      ...(deps.now ? { now: deps.now } : {}),
      onChange: () => {
        scheduleStatusUpdate();
        reconcileActiveTools();
      },
    });
    return controller;
  }

  /** One-line command status, including the owner and the autostart fact. */
  function statusText(): string {
    const snapshot = controller?.snapshot();
    if (!snapshot || snapshot.state === "off") {
      const lines = [
        `Office bridge is off (port ${snapshot?.port ?? resolvePort()}). Run /office connect to start it.`,
      ];
      if (snapshot?.lastError) lines.push(`Last error: ${snapshot.lastError}`);
      if (resolveAutostart()) {
        lines.push("Autostart is enabled (--office-bridge-autostart).");
      }
      return lines.join("\n");
    }
    if (snapshot.state === "starting") {
      return `Office bridge is starting on :${snapshot.port}…`;
    }
    if (snapshot.state === "closing") return "Office bridge is closing…";

    const port = snapshot.actualPort ?? snapshot.port;
    const panes = activeServer()?.attachedPanes() ?? [];
    if (panes.length === 0) {
      return `Office bridge is listening on :${port} — no app attached yet. (catalog v${CATALOG_VERSION})`;
    }
    const lines = panes.map((p) => {
      const model = p.model ? `, model=${p.model}` : "";
      const provider = p.provider ? `, provider=${p.provider}` : "";
      const ago = Math.max(0, Math.round((Date.now() - p.lastSeen) / 1000));
      const ops =
        p.ops === null
          ? "legacy client (v1 ops)"
          : `${p.ops.length} ops${p.catalogVersion ? ` (catalog v${p.catalogVersion})` : ""}`;
      const ignored =
        p.opsIgnoredCount > 0
          ? `, ${p.opsIgnoredCount} advertised op(s) ignored (unknown to this server)`
          : "";
      return `- ${paneDocumentLabel(p)} (${p.clientName}, pane ${p.paneId.slice(0, 8)})${model}${provider}, seen ${ago}s ago — ${ops}${ignored}`;
    });
    return `Office bridge on :${port} (catalog v${CATALOG_VERSION})\n${lines.join("\n")}`;
  }

  /* ── lifecycle ─────────────────────────────────────────────────────── */

  pi.on("session_start", async (_event, ctx) => {
    currentCtx = ctx;
    registerAllOfficeTools();
    registerOfficeListDocumentsTool();

    // TUI `#` document mentions. Registered on every session_start so a session
    // rebind (which rebuilds Pi's autocomplete provider stack) re-registers it.
    // Read the panes live so the dropdown is never stale.
    ctx.ui.addAutocompleteProvider((current) =>
      createDocumentMentionProvider(current as MentionProvider, () => {
        const panes = activeServer()?.attachedPanes() ?? [];
        return panes.map((pane) => toMentionTarget(pane, panes));
      }),
    );

    const bridge = ensureController();
    reconcileActiveTools();
    updateStatus();

    if (!resolveAutostart()) return;
    if (bridge.snapshot().state !== "off") return;
    try {
      const snapshot = await bridge.connect();
      ctx.ui.notify(
        `Office bridge listening on ws://127.0.0.1:${snapshot.actualPort ?? snapshot.port}`,
        "info",
      );
    } catch (error) {
      ctx.ui.notify(
        error instanceof Error ? error.message : String(error),
        "error",
      );
    }
    reconcileActiveTools();
    updateStatus();
  });

  pi.on("session_shutdown", async () => {
    currentCtx = null;
    panePrompts.clear();
    clearDeltaState();
    if (panesChangedTimer !== null) {
      clearTimeout(panesChangedTimer);
      panesChangedTimer = null;
    }
    const active = controller;
    if (active) {
      await active.close().catch((error: unknown) => {
        console.error(
          `[office-bridge] shutdown close failed: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      });
    }
  });

  /* ── agent → pane forwarding ───────────────────────────────────────── */

  // Tell the agent which Office apps are attached before each turn so it uses
  // the office_* tools directly instead of claiming capabilities are missing.
  pi.on("before_agent_start", (event, _ctx) => {
    const block = paneContextBlock();
    if (block === null) return undefined;
    return { systemPrompt: event.systemPrompt + block };
  });

  // Protocol v3: stream partial assistant text to a pane that can render it.
  // Deltas are display-only (the final message replaces the live buffer), so
  // chunks are coalesced into ~frame-sized batches instead of one frame per
  // token. Only `text_delta` is forwarded — thinking and tool-call deltas are
  // not part of the pane's reply text.
  function deltaTextFromEvent(event: { assistantMessageEvent?: unknown }): string | null {
    const update = event.assistantMessageEvent as
      | { type?: unknown; delta?: unknown }
      | null
      | undefined;
    if (!update || typeof update !== "object") return null;
    if (update.type !== "text_delta") return null;
    return typeof update.delta === "string" && update.delta.length > 0
      ? update.delta
      : null;
  }

  function clearDeltaState(): void {
    deltaBuffer = "";
    deltaPane = null;
    deltaPromptId = null;
    if (deltaTimer !== null) {
      clearTimeout(deltaTimer);
      deltaTimer = null;
    }
  }

  /** Send whatever partial text is buffered, then drop it. Safe anytime. */
  function flushDeltas(): void {
    if (deltaTimer !== null) {
      clearTimeout(deltaTimer);
      deltaTimer = null;
    }
    const pane = deltaPane;
    const text = deltaBuffer;
    const promptId = deltaPromptId;
    deltaBuffer = "";
    if (pane === null || text.length === 0) return;
    const active = activeServer();
    if (!bridgeIsOn() || active === null) return;
    active.sendToPane(pane, {
      type: "agent_message",
      kind: "delta",
      text,
      promptId: promptId ?? undefined,
    });
  }

  const DELTA_FLUSH_MS = 60;
  let deltaBuffer = "";
  let deltaTimer: ReturnType<typeof setTimeout> | null = null;
  let deltaPane: AttachedPane | null = null;
  let deltaPromptId: string | null = null;

  function queueDelta(pane: AttachedPane, promptId: string, text: string): void {
    // A new pane prompt started: never leak the previous prompt's tail into it.
    if (deltaPane !== null && deltaPane !== pane) flushDeltas();
    deltaPane = pane;
    deltaPromptId = promptId;
    deltaBuffer += text;
    if (deltaTimer !== null) return;
    deltaTimer = setTimeout(() => {
      deltaTimer = null;
      flushDeltas();
    }, DELTA_FLUSH_MS);
  }

  // Bind a delivered user message to the pane prompt that produced it. The Pi
  // `input` event fires at enqueue time and queued follow-ups skip
  // `before_agent_start`, so delivery (`message_start`) is the reliable point.
  pi.on("message_start", (event, _ctx) => {
    if (event.message.role !== "user") return;
    const text = flattenTextContent(event.message.content);
    if (text) panePrompts.bindDeliveredUserMessage(text);
  });

  pi.on("message_update", (event, _ctx) => {
    if (!bridgeIsOn()) return;
    const prompt = panePrompts.active();
    if (!prompt) return;
    // A v2 pane cannot render a delta; the final reply is all it gets.
    if (prompt.pane.protocolVersion < 3) return;
    const text = deltaTextFromEvent(event);
    if (text === null) return;
    queueDelta(prompt.pane, prompt.promptId, text);
  });

  // Forward the final assistant reply to the pane that prompted it — never to
  // every pane, and never for a TUI-initiated turn.
  pi.on("message_end", async (event, _ctx) => {
    if (event.message.role !== "assistant") return;
    if (!bridgeIsOn()) return;
    const prompt = panePrompts.active();
    if (!prompt) return;

    const text = flattenTextContent(event.message.content);
    // Flush the buffered tail first so the pane shows the end of the reply
    // before the authoritative final replaces its live buffer.
    flushDeltas();
    panePrompts.complete();
    if (!text) return;
    // SAFETY: assistant messages may carry an id that the Pi event type does
    // not surface; read it through a narrow shape check before using it.
    const maybeId = (event.message as unknown as { id?: unknown }).id;
    activeServer()?.sendToPane(prompt.pane, {
      type: "agent_message",
      kind: "final",
      text,
      messageId: typeof maybeId === "string" ? maybeId : undefined,
      promptId: prompt.promptId,
    });
    updateStatus();
  });

  // An aborted/settled turn that produced no final must not leave a pane bound
  // to the next, unrelated turn.
  pi.on("agent_settled", () => {
    panePrompts.complete();
    clearDeltaState();
  });

  // Forward progress of non-office tools while a pane is awaiting a reply.
  pi.on("tool_execution_start", async (event) => {
    if (!bridgeIsOn()) return;
    const prompt = panePrompts.active();
    if (!prompt) return;
    if (hostForToolName(event.toolName) !== null) return;
    activeServer()?.sendToPane(prompt.pane, {
      type: "tool_activity",
      tool: event.toolName,
      status: "start",
      summary: summarizeArgs(event.toolName, event.args),
    });
  });

  pi.on("tool_execution_end", async (event) => {
    if (!bridgeIsOn()) return;
    const prompt = panePrompts.active();
    if (!prompt) return;
    if (hostForToolName(event.toolName) !== null) return;
    activeServer()?.sendToPane(prompt.pane, {
      type: "tool_activity",
      tool: event.toolName,
      status: event.isError ? "error" : "end",
      summary: event.isError
        ? String(event.result ?? "tool failed")
        : undefined,
    });
  });

  /* ── commands ──────────────────────────────────────────────────────── */

  pi.registerCommand("office", {
    description:
      "Control the pi-office bridge: status, connect, close or restart it.",
    getArgumentCompletions: (prefix) => officeCommandCompletions(prefix),
    handler: async (args, ctx) => {
      const command = parseOfficeCommand(args);
      if (command.kind === "help") {
        ctx.ui.notify(officeCommandHelp(), "info");
        return;
      }
      if (command.kind === "error") {
        throw new Error(command.message);
      }
      if (command.kind === "status") {
        ctx.ui.notify(statusText(), "info");
        return;
      }

      const bridge = ensureController();
      if (command.kind === "connect") {
        const snapshot = await bridge.connect(command.port);
        ctx.ui.notify(
          `Office bridge listening on ws://127.0.0.1:${snapshot.actualPort ?? snapshot.port}. Open the Pi for Office add-in, or run /office status.`,
          "info",
        );
      } else if (command.kind === "close") {
        await bridge.close();
        ctx.ui.notify(
          "Office bridge closed. The add-in is disconnected; run /office connect to resume.",
          "info",
        );
      } else if (command.kind === "restart") {
        const snapshot = await bridge.restart(command.port);
        ctx.ui.notify(
          `Office bridge restarted on ws://127.0.0.1:${snapshot.actualPort ?? snapshot.port}.`,
          "info",
        );
      }
      reconcileActiveTools();
      updateStatus();
    },
  });

  // Also list every office tool we exposed so users can confirm them.
  pi.registerCommand("office-tools", {
    description:
      "List the office tools registered by the pi-office bridge extension.",
    handler: async (_args, ctx) => {
      const active = new Set(pi.getActiveTools());
      const names = [...OFFICE_TOOL_NAMES, OFFICE_LIST_TOOL];
      const lines = names.map(
        (name) => `${active.has(name) ? "[on] " : "[off]"} ${name}`,
      );
      ctx.ui.notify(
        `Office tools (${OFFICE_TOOL_NAMES.length} catalogue + discovery, catalog v${CATALOG_VERSION}):\n${lines.join("\n")}`,
        "info",
      );
    },
  });
}

/** Compact one-line summary of a non-office tool call for the pane UI. */
function summarizeArgs(tool: string, args: unknown): string {
  if (tool === "bash" || tool === "powershell") {
    const command = (args as { command?: unknown })?.command;
    if (typeof command === "string") {
      const oneLine = command.replace(/\s+/g, " ").trim();
      return oneLine.length > 120 ? `${oneLine.slice(0, 120)}…` : oneLine;
    }
  }
  if (tool === "read" || tool === "write" || tool === "edit") {
    const path = (args as { path?: unknown })?.path;
    if (typeof path === "string") return path;
  }
  return tool;
}

export default function (pi: ExtensionAPI): void {
  createOfficeBridgeExtension(pi);
}
