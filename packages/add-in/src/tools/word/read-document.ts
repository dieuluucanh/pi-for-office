/**
 * word_read_document — Read the live Word document text.
 *
 * Scope "all" (default) reads the whole body; "selection" reads the current
 * selection. Output is truncated to maxChars (default 20k).
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const MAX_READ_CHARS = 200_000;
const DEFAULT_READ_CHARS = 20_000;

const schema = Type.Object({
  scope: Type.Optional(
    Type.Union([Type.Literal("all"), Type.Literal("selection")], {
      description:
        '"all" (default): read the whole document body. ' +
        '"selection": read only the currently selected text.',
    }),
  ),
  maxChars: Type.Optional(
    Type.Number({
      description: `Maximum characters to return (default ${DEFAULT_READ_CHARS}, max ${MAX_READ_CHARS}).`,
    }),
  ),
});

type Params = Static<typeof schema>;

export function createWordReadDocumentTool(): AgentTool<typeof schema> {
  return {
    name: "word_read_document",
    label: t("tools.wordReadDocument"),
    description:
      "Read text from the live Word document: the whole body by default, or just the current selection. " +
      "Always read before modifying — never guess what's in the document.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      const scope = params.scope === "selection" ? "selection" : "all";
      const maxChars =
        params.maxChars === undefined
          ? DEFAULT_READ_CHARS
          : Math.min(Math.max(params.maxChars, 100), MAX_READ_CHARS);

      try {
        const text = await wordRun(async (context) => {
          if (scope === "selection") {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            return selection.text ?? "";
          }
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return body.text ?? "";
        });

        const truncated = text.length > maxChars;
        const shown = truncated ? text.slice(0, maxChars) : text;

        const lines: string[] = [];
        if (scope === "selection") {
          lines.push(`**Selection (${shown.length} chars shown)**`);
        } else {
          const truncationNote = truncated ? `, showing first ${maxChars}` : "";
          lines.push(
            `**Document text** (${text.length} chars total${truncationNote})`,
          );
        }
        if (shown.trim().length === 0) {
          lines.push("");
          lines.push("_Empty._");
        } else {
          lines.push("");
          lines.push(shown);
          if (truncated) {
            lines.push("");
            lines.push(
              `_…truncated (${text.length - maxChars} chars remaining). Pass maxChars to read more._`,
            );
          }
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error reading document: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
