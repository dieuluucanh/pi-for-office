/**
 * Bridge tool delegation — run a local agent tool through the bridge and
 * unwrap its AgentToolResult into the OfficeOpOutcome shape.
 *
 * Bridge executors become thin one-line mappings over the SAME factory the
 * browser-only path uses, so bridge and local modes share one implementation
 * by construction.
 */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { Static, TSchema } from "typebox";
import type { OfficeOpOutcome } from "./ops.js";

/** Extract the text payload of an AgentToolResult. */
export function textOfAgentToolResult(result: {
  content?: Array<{ type?: string; text?: DynamicValue }> | string;
}): string {
  const content = result.content;
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";
  const parts: string[] = [];
  for (const part of content) {
    if (
      typeof part === "object" &&
      part !== null &&
      typeof part.text === "string"
    ) {
      parts.push(part.text);
    }
  }
  return parts.join("\n").trim();
}

/**
 * Wrap a local tool factory as a bridge op executor. The synthetic call id
 * "bridge" is never surfaced; args are forwarded verbatim.
 */
export function delegateTool(
  tool: AgentTool<TSchema>,
): (args: Record<string, DynamicValue>) => Promise<OfficeOpOutcome> {
  return async (args) => {
    // SAFETY: bridge args are validated against the shared catalog schema by
    // the local tool itself at execution time; forwarding the dynamic record
    // as the tool's params is intentional (same widening as the excel-ops path).
    // A cast through `never` is required because AgentTool's params type is
    // derived from the schema and is wider than Object-record.
    const castArgs = args as Static<typeof tool.parameters>;
    const result = await tool.execute("bridge", castArgs);
    return { text: textOfAgentToolResult(result), details: result.details };
  };
}
