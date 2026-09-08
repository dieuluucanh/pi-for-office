/**
 * Bridge op contract — pane side of the office bridge.
 *
 * The Pi extension (`@dieulc/pi-office-bridge`) forwards `office_<host>_<op>`
 * tool calls here as `<host>.<op>` ids (see the shared catalog in
 * `packages/bridge-extension/src/office-tools.ts`). This module resolves each
 * op to a pane-side executor that performs the actual Office.js operation.
 *
 * Op ids MUST match the catalog in the Pi extension. When adding an op, add it
 * to both places (same repo — see README.md "Bridge contract").
 */

export interface OfficeOpOutcome {
  /** Markdown/text returned to the LLM. */
  text: string;
  /** Structured details (optional; persisted with the tool result). */
  details?: unknown;
  /** When true the outcome is an error description rather than a result. */
  isError?: boolean;
}

/** Executor signature. `args` is the tool-call parameter object. */
export type OfficeOpExecutor = (args: Record<string, unknown>) => Promise<OfficeOpOutcome>;

export type OfficeOpRegistry = ReadonlyMap<string, OfficeOpExecutor>;

/** Resolve the text of a delegated agent-tool result for bridge output. */
export function textOfAgentToolResult(result: {
  content?: Array<{ type?: string; text?: unknown }> | string;
}): string {
  const content = result.content;
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";
  const parts: string[] = [];
  for (const part of content) {
    if (typeof part === "object" && part !== null && typeof part.text === "string") {
      parts.push(part.text);
    }
  }
  return parts.join("\n").trim();
}

/** Wrap any executor so thrown errors become error outcomes (never throws). */
export function guardExecutor(executor: OfficeOpExecutor): OfficeOpExecutor {
  return async (args) => {
    try {
      return await executor(args);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return { text: `Office operation failed: ${message}`, isError: true };
    }
  };
}

/**
 * Execute a bridge op id (`excel.read_range`, `word.insert_text`, …).
 * Throws when the op is unknown; executors themselves never throw.
 */
export async function executeOfficeOp(
  registry: OfficeOpRegistry,
  op: string,
  args: Record<string, unknown>,
): Promise<OfficeOpOutcome> {
  const executor = registry.get(op);
  if (!executor) {
    return {
      text: `Unknown office op "${op}". The add-in and the Pi bridge extension are out of sync — update the add-in (ops: ${[...registry.keys()].join(", ")}).`,
      isError: true,
    };
  }
  return executor(args);
}
