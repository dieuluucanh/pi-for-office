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
 * Pure extension: uses only the public extension API — Pi core is never
 * touched, so Pi can be updated freely.
 */

import type {
  AgentToolResult,
  ExtensionAPI,
  ExtensionCommandContext,
  ExtensionContext,
  RegisteredCommand,
} from "@earendil-works/pi-coding-agent";
import { Type, type Static } from "typebox";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { CATALOG_VERSION, BRIDGE_DEFAULT_PORT } from "./protocol.js";
import type { AttachedPane } from "./bridge-server.js";
import { OfficeBridgeServer } from "./bridge-server.js";
import {
  OFFICE_TOOL_DESCRIPTORS,
  hostForToolName,
  OFFICE_TOOL_NAMES,
  HOST_APP_LABEL,
} from "./office-tools.js";
import type { OfficeToolDescriptor } from "./office-tools.js";
import {
  reconcileOfficeToolActivation,
  type PaneCapability,
} from "./active-tools.js";

const FLAG_PORT = "office-bridge-port";

/** Which Pi version we run inside (shown in the pane's welcome frame). */
function piVersion(): string | null {
  try {
    // The coding-agent package exposes its version via package.json at runtime.
    // Fall back to process env when unavailable.
    return process.env.PI_VERSION ?? null;
  } catch {
    return null;
  }
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

export default function (pi: ExtensionAPI): void {
  let server: OfficeBridgeServer | null = null;
  let currentCtx: ExtensionContext | null = null;

  /** Panes awaiting the assistant reply to their injected prompt. */
  const pendingReplyTargets: AttachedPane[] = [];

  pi.registerFlag(FLAG_PORT, {
    description: `Port for the pi-office bridge server (default ${BRIDGE_DEFAULT_PORT})`,
    type: "string",
  });

  function resolvePort(): number {
    const raw = pi.getFlag(FLAG_PORT);
    if (typeof raw === "string" && raw.trim() !== "") {
      const parsed = Number.parseInt(raw.trim(), 10);
      if (Number.isFinite(parsed) && parsed > 0 && parsed < 65536)
        return parsed;
    }
    const env = process.env.PI_OFFICE_BRIDGE_PORT;
    if (env) {
      const parsed = Number.parseInt(env, 10);
      if (Number.isFinite(parsed) && parsed > 0 && parsed < 65536)
        return parsed;
    }
    return BRIDGE_DEFAULT_PORT;
  }

  function updateStatus(): void {
    const ui = currentCtx?.ui;
    if (!ui) return;
    const panes = server?.attachedPanes() ?? [];
    if (!server?.isRunning) {
      ui.setStatus("office-bridge", undefined);
      return;
    }
    const port = server.actualPort ?? resolvePort();
    if (panes.length === 0) {
      ui.setStatus(
        "office-bridge",
        `office bridge on :${port} — no app attached`,
      );
      return;
    }
    const labels = panes.map((p) => HOST_APP_LABEL[p.host]).join(", ");
    ui.setStatus(
      "office-bridge",
      `office: ${labels} attached (bridge :${port})`,
    );
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

  /** Extract display text from an assistant message content payload. */
  function flattenAssistantText(content: unknown): string {
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
        const active = server;
        if (!active?.isRunning) {
          throw new Error(
            `office-bridge: server is not running. Check the Pi extension loaded, then open the add-in.`,
          );
        }
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
   * Reconcile the Pi active tool set with the currently attached panes: only
   * the office ops the attached panes advertise stay active; everything else
   * (user/other-extension tools) is preserved. Called on attach/detach.
   */
  function reconcileActiveTools(panes: readonly AttachedPane[]): void {
    const capabilities: PaneCapability[] = panes.map((pane) => ({
      host: pane.host,
      ops: pane.ops,
    }));
    const next = reconcileOfficeToolActivation(
      pi.getActiveTools(),
      capabilities,
    );
    pi.setActiveTools(next);
  }

  /** Stable per-turn pane-awareness block appended to the system prompt. */
  function paneContextBlock(): string | null {
    const active = server;
    if (!active?.isRunning) return null;
    const panes = active.attachedPanes();
    if (panes.length === 0) return null;

    const lines = panes.map((p) => {
      const legacy =
        p.ops === null
          ? " (legacy client: only v1 ops — update the add-in for newer tools)"
          : "";
      return `- ${HOST_APP_LABEL[p.host]} is attached via the office bridge${legacy}. The open document is live: the office_${p.host}_* tools edit it directly.`;
    });
    return (
      "\n\n## Attached Office documents\n" +
      lines.join("\n") +
      "\n\nEditing these documents with the office_* tools is fully supported, including formatting and " +
      "structured documents (titles, headings, lists, tables, alignment, indents). Never claim formatting is " +
      "unavailable, and never emit HTML for Word — use the office_word_* tools."
    );
  }

  /* ── lifecycle ─────────────────────────────────────────────────────── */

  pi.on("session_start", async (_event, ctx) => {
    currentCtx = ctx;
    registerAllOfficeTools();

    if (server?.isRunning) {
      updateStatus();
      return;
    }

    const port = resolvePort();
    const bridge = new OfficeBridgeServer({
      port,
      serverName: "pi-office-bridge",
      serverVersion: serverVersion(),
      piVersion: piVersion(),
      handlers: {
        onUserMessage: (text, pane) => {
          pendingReplyTargets.push(pane);
          updateStatus();
          // followUp: if Pi is idle the message is delivered immediately and
          // triggers a turn; if a turn is running it is queued until it settles.
          pi.sendUserMessage(text, { deliverAs: "followUp" });
        },
        // Keep the TUI status line live and the active tool set in sync:
        // opening an app activates that host's office tools; closing it
        // deactivates them.
        onPanesChanged: (panes) => {
          scheduleStatusUpdate();
          reconcileActiveTools(panes);
        },
      },
    });

    try {
      await bridge.start();
      server = bridge;
      ctx.ui.notify(
        `Office bridge listening on ws://127.0.0.1:${port}`,
        "info",
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const code = (error as NodeJS.ErrnoException | null)?.code;
      if (code === "EADDRINUSE") {
        ctx.ui.notify(
          `Office bridge: port ${port} is already in use — another Pi process ` +
            `is running the bridge. Start this one on a free port with ` +
            `--${FLAG_PORT} <port> or PI_OFFICE_BRIDGE_PORT=<port>.`,
          "error",
        );
      } else {
        ctx.ui.notify(`Office bridge failed to start: ${message}`, "error");
      }
      console.error(`[office-bridge] start failed: ${message}`);
    }
    updateStatus();
  });

  pi.on("session_shutdown", async () => {
    currentCtx = null;
    pendingReplyTargets.length = 0;
    if (panesChangedTimer !== null) {
      clearTimeout(panesChangedTimer);
      panesChangedTimer = null;
    }
    const active = server;
    server = null;
    if (active) {
      await active.stop();
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

  // Forward the final assistant reply to the pane that prompted it.
  pi.on("message_end", async (event, _ctx) => {
    if (event.message.role !== "assistant") return;
    if (!server?.isRunning) return;
    const target = pendingReplyTargets.shift();
    if (!target) return;

    const text = flattenAssistantText(event.message.content);
    if (!text) return;
    // SAFETY: assistant messages may carry an id that the Pi event type does
    // not surface; read it through a narrow shape check before using it.
    const maybeId = (event.message as unknown as { id?: unknown }).id;
    server.broadcast({
      type: "agent_message",
      kind: "final",
      text,
      messageId: typeof maybeId === "string" ? maybeId : undefined,
    });
    updateStatus();
  });

  // Forward progress of non-office tools while a pane is awaiting a reply.
  pi.on("tool_execution_start", async (event) => {
    if (!server?.isRunning) return;
    if (pendingReplyTargets.length === 0) return;
    if (hostForToolName(event.toolName) !== null) return;
    server.broadcast({
      type: "tool_activity",
      tool: event.toolName,
      status: "start",
      summary: summarizeArgs(event.toolName, event.args),
    });
  });

  pi.on("tool_execution_end", async (event) => {
    if (!server?.isRunning) return;
    if (pendingReplyTargets.length === 0) return;
    if (hostForToolName(event.toolName) !== null) return;
    server.broadcast({
      type: "tool_activity",
      tool: event.toolName,
      status: event.isError ? "error" : "end",
      summary: event.isError
        ? String(event.result ?? "tool failed")
        : undefined,
    });
  });

  /* ── commands ──────────────────────────────────────────────────────── */

  const officeCommand: Omit<RegisteredCommand, "name" | "sourceInfo"> = {
    description:
      "Show the pi-office bridge status: server port and attached Office apps (Excel/Word/PowerPoint).",
    handler: async (_args: string, ctx: ExtensionCommandContext) => {
      const active = server;
      if (!active?.isRunning) {
        ctx.ui.notify("Office bridge is not running.", "warning");
        return;
      }
      const port = active.actualPort ?? resolvePort();
      const panes = active.attachedPanes();
      if (panes.length === 0) {
        ctx.ui.notify(
          `Office bridge is listening on :${port} — no app attached yet.`,
          "info",
        );
        return;
      }
      const lines = panes.map((p) => {
        const model = p.model ? `, model=${p.model}` : "";
        const provider = p.provider ? `, provider=${p.provider}` : "";
        const ago = Math.max(0, Math.round((Date.now() - p.lastSeen) / 1000));
        const ops =
          p.ops !== null
            ? `${p.ops.length} ops${p.catalogVersion ? ` (catalog v${p.catalogVersion})` : ""}`
            : "legacy client (v1 ops)";
        const ignored =
          p.opsIgnoredCount > 0
            ? `, ${p.opsIgnoredCount} advertised op(s) ignored (unknown to this server)`
            : "";
        return `- ${HOST_APP_LABEL[p.host]} (${p.clientName}, pane ${p.paneId.slice(0, 8)})${model}${provider}, seen ${ago}s ago — ${ops}${ignored}`;
      });
      ctx.ui.notify(
        `Office bridge on :${port} (catalog v${CATALOG_VERSION})\n${lines.join("\n")}`,
        "info",
      );
    },
  };

  pi.registerCommand("office", officeCommand);

  // Also list every office tool we exposed so users can confirm them:
  pi.registerCommand("office-tools", {
    description:
      "List the office tools registered by the pi-office bridge extension.",
    handler: async (_args: string, ctx: ExtensionCommandContext) => {
      ctx.ui.notify(
        `Office tools (${OFFICE_TOOL_NAMES.length}, catalog v${CATALOG_VERSION}):\n${OFFICE_TOOL_NAMES.join("\n")}`,
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
