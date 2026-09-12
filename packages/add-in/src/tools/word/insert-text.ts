/**
 * word_insert_text — Insert text into the live Word document.
 *
 * Locations: "end" (default), "start", or "replace_selection".
 * Multi-line text becomes separate paragraphs. Supports optional formatting
 * (bold/italic/underline/size/name/color), paragraph props (style, alignment,
 * spacing, indents), and `format: "markdown"` for a markdown subset
 * (headings, bold/italic, bullets, page-break rules).
 */

import { WORD_INSERT_TEXT_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  insertWordBlocks,
  wordRun,
  type WordInsertBlock,
  type WordInsertLocation,
} from "../../word/helpers.js";
import { parseMarkdownToBlocks } from "../../word/markdown.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = WORD_INSERT_TEXT_PARAMETERS;
type Params = Static<typeof schema>;

/** Build blocks from plain text: newlines → separate paragraphs. */
function plainTextBlocks(text: string): WordInsertBlock[] {
  const lines = text.split(/\r?\n/u);
  return lines.map((line) => ({
    text: line.trimEnd(),
    type: "paragraph" as const,
  }));
}

export function createWordInsertTextTool(): AgentTool<typeof schema> {
  return {
    name: "word_insert_text",
    label: t("tools.wordInsertText"),
    description:
      "Insert text into the live Word document — at the start, at the end (default), " +
      "or replacing the current selection. Multi-line text becomes separate paragraphs. " +
      "Optionally format the inserted text in the same call: bold, italic, underline, size (points), " +
      'font name, color (#RRGGBB), alignment ("Left"/"Centered"/"Right"/"Justified"), style names ' +
      '(e.g. "Heading 1", "Title"), spacing, and indents. Pass format: "markdown" to insert a markdown ' +
      "document (headings, bold/italic, bullets, page-break rules). " +
      "Use this to add new content to the opened document.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      const location: WordInsertLocation =
        params.location === "start" || params.location === "replace_selection"
          ? params.location
          : "end";

      if (params.text.length === 0) {
        return {
          content: [{ type: "text", text: "Nothing to insert (empty text)." }],
          details: undefined,
        };
      }

      try {
        const format = params.format === "markdown" ? "markdown" : "text";
        const blocks =
          format === "markdown"
            ? parseMarkdownToBlocks(params.text)
            : plainTextBlocks(params.text);

        const fontFormat = {
          ...(params.bold === undefined ? {} : { bold: params.bold }),
          ...(params.italic === undefined ? {} : { italic: params.italic }),
          ...(params.underline === undefined
            ? {}
            : { underline: params.underline }),
          ...(params.size === undefined ? {} : { size: params.size }),
          ...(params.name === undefined ? {} : { name: params.name }),
          ...(params.color === undefined ? {} : { color: params.color }),
        };
        const paragraphFormat = {
          ...(params.style === undefined ? {} : { style: params.style }),
          ...(params.alignment === undefined
            ? {}
            : { alignment: params.alignment }),
          ...(params.spaceBefore === undefined
            ? {}
            : { spaceBefore: params.spaceBefore }),
          ...(params.spaceAfter === undefined
            ? {}
            : { spaceAfter: params.spaceAfter }),
          ...(params.lineSpacing === undefined
            ? {}
            : { lineSpacing: params.lineSpacing }),
          ...(params.firstLineIndent === undefined
            ? {}
            : { firstLineIndent: params.firstLineIndent }),
          ...(params.leftIndent === undefined
            ? {}
            : { leftIndent: params.leftIndent }),
        };

        const hasFont = Object.keys(fontFormat).length > 0;
        const hasParagraph = Object.keys(paragraphFormat).length > 0;
        const blocksWithFormat: WordInsertBlock[] =
          hasFont || hasParagraph
            ? blocks.map((block) => ({
                ...block,
                format: { ...fontFormat, ...paragraphFormat },
              }))
            : blocks;

        await wordRun(async (context) => {
          insertWordBlocks(context, blocksWithFormat, location);
          await context.sync();
        });

        const formatNote = hasFont || hasParagraph ? " with formatting" : "";
        const summary =
          location === "replace_selection"
            ? `Replaced the current selection with ${params.text.length} chars${formatNote}.`
            : `Inserted ${params.text.length} chars at the ${location} of the document${formatNote}.`;
        return {
          content: [{ type: "text", text: summary }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error inserting text: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
