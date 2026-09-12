/**
 * word_insert_blocks — Assemble a structured formatted document from blocks.
 *
 * Each block can be a paragraph, heading (level 1-6), bullet/numbered list
 * item, or page break, with per-block font/paragraph formatting.
 */

import { WORD_INSERT_BLOCKS_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  insertWordBlocks,
  wordRun,
  type WordInsertBlock,
  type WordInsertLocation,
} from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = WORD_INSERT_BLOCKS_PARAMETERS;
type Params = Static<typeof schema>;

/** Convert one block param object to the engine's WordInsertBlock. */
function toWordBlock(param: Params["blocks"][number]): WordInsertBlock {
  const fontFormat = {
    ...(param.bold === undefined ? {} : { bold: param.bold }),
    ...(param.italic === undefined ? {} : { italic: param.italic }),
    ...(param.underline === undefined ? {} : { underline: param.underline }),
    ...(param.size === undefined ? {} : { size: param.size }),
    ...(param.name === undefined ? {} : { name: param.name }),
    ...(param.color === undefined ? {} : { color: param.color }),
  };
  const paragraphFormat = {
    ...(param.style === undefined ? {} : { style: param.style }),
    ...(param.alignment === undefined ? {} : { alignment: param.alignment }),
    ...(param.spaceBefore === undefined
      ? {}
      : { spaceBefore: param.spaceBefore }),
    ...(param.spaceAfter === undefined ? {} : { spaceAfter: param.spaceAfter }),
    ...(param.lineSpacing === undefined
      ? {}
      : { lineSpacing: param.lineSpacing }),
    ...(param.firstLineIndent === undefined
      ? {}
      : { firstLineIndent: param.firstLineIndent }),
    ...(param.leftIndent === undefined ? {} : { leftIndent: param.leftIndent }),
  };

  return {
    text: param.text,
    type: param.type ?? "paragraph",
    ...(param.level === undefined ? {} : { level: param.level }),
    ...(param.listType === undefined ? {} : { listType: param.listType }),
    format: { ...fontFormat, ...paragraphFormat },
  };
}

export function createWordInsertBlocksTool(): AgentTool<typeof schema> {
  return {
    name: "word_insert_blocks",
    label: t("tools.wordInsertBlocks"),
    description:
      "Assemble a structured formatted document in the live Word document from blocks: " +
      "paragraphs, headings (level 1-6), bullet/numbered list items, and page breaks. " +
      "Each block can carry its own font (bold/italic/underline/size/name/color), " +
      "paragraph alignment, style name, spacing and indents. Use this for a complete formatted " +
      "document in one call — e.g. a title, then body paragraphs, then a right-aligned signature block.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      if (params.blocks.length === 0) {
        return {
          content: [{ type: "text", text: "No blocks provided." }],
          details: undefined,
        };
      }

      const location: WordInsertLocation =
        params.location === "start" || params.location === "replace_selection"
          ? params.location
          : "end";

      try {
        const blocks = params.blocks.map(toWordBlock);
        await wordRun(async (context) => {
          insertWordBlocks(context, blocks, location);
          await context.sync();
        });

        const blockCount = blocks.length;
        const headingCount = blocks.filter((b) => b.type === "heading").length;
        const listCount = blocks.filter((b) => b.type === "listItem").length;
        return {
          content: [
            {
              type: "text",
              text:
                `Inserted ${blockCount} block${blockCount === 1 ? "" : "s"} at the ${location} of the document` +
                (headingCount > 0
                  ? ` (${headingCount} heading${headingCount === 1 ? "" : "s"})`
                  : "") +
                (listCount > 0
                  ? ` (${listCount} list item${listCount === 1 ? "" : "s"})`
                  : "") +
                " with formatting.",
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error inserting blocks: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
