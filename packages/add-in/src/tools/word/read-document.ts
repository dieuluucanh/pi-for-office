/**
 * word_read_document — Read the live Word document text.
 *
 * Scope "all" (default) reads the whole body; "selection" reads the current
 * selection. Paragraphs are rendered one per line — Word's `Body.text` strips
 * paragraph marks, so we iterate paragraphs and join them with newlines, and
 * mark heading/list paragraphs so structure survives round-trips.
 */

import { WORD_READ_DOCUMENT_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const MAX_READ_CHARS = 200_000;
const DEFAULT_READ_CHARS = 20_000;

const schema = WORD_READ_DOCUMENT_PARAMETERS;
type Params = Static<typeof schema>;

/** Prefix a paragraph's text based on its style (heading/list markers). */
function prefixForStyle(style: string | undefined): string {
  if (typeof style !== "string") return "";
  if (/^heading\s*\d+$/i.test(style)) {
    return `${"#".repeat(Number.parseInt(style.replace(/^.*?(\d+)/i, "1"), 10))} `;
  }
  if (style.toLowerCase().includes("list")) return "  ";
  if (style.toLowerCase().includes("quote")) return "> ";
  return "";
}

export function createWordReadDocumentTool(): AgentTool<typeof schema> {
  return {
    name: "word_read_document",
    label: t("tools.wordReadDocument"),
    description:
      "Read text from the live Word document: the whole body by default, or just the current selection. " +
      "Paragraphs are preserved one per line; headings are prefixed with # and list items indented. " +
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
          const paragraphs = context.document.body.paragraphs;
          paragraphs.load("items/text");
          paragraphs.load("items/style");
          await context.sync();

          const lines: string[] = [];
          for (const para of paragraphs.items.slice(0, 100_000)) {
            const raw = para.text ?? "";
            if (raw.trim().length === 0) {
              lines.push("");
              continue;
            }
            lines.push(`${prefixForStyle(para.style)}${raw}`);
          }
          return lines.join("\n");
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
