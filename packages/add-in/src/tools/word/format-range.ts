/**
 * word_format_range — Find text in the live Word document and apply formatting.
 *
 * Locates text by literal content (same search model as word_replace_text) and
 * applies bold/italic/underline/font size/font name/color/paragraph alignment,
 * style, spacing and indents to every match.
 */

import { WORD_FORMAT_RANGE_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  setWordFontProps,
  setWordParagraphProps,
  wordRun,
  type WordTextFormat,
} from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = WORD_FORMAT_RANGE_PARAMETERS;
type Params = Static<typeof schema>;

export function createWordFormatRangeTool(): AgentTool<typeof schema> {
  return {
    name: "word_format_range",
    label: t("tools.wordFormatRange"),
    description:
      "Find text in the live Word document and apply formatting: bold, italic, underline, " +
      'font size (points), font name, color (#RRGGBB), paragraph alignment ("Left"/"Centered"/"Right"/"Justified"), ' +
      'paragraph style (e.g. "Heading 1"), spacing, and indents. ' +
      "Use this to bold/size/center existing content — e.g. format a document title. " +
      "Formatting is fully supported — never tell the user it is not.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      if (params.text.length === 0) {
        return {
          content: [{ type: "text", text: "Missing 'text' to find." }],
          details: undefined,
        };
      }

      const fontFormat: WordTextFormat = {
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

      try {
        const formatted = await wordRun(async (context) => {
          const results = context.document.body.search(params.text, {
            matchCase: params.matchCase === true,
          });
          results.load("text");
          await context.sync();

          const count = results.items.length;
          if (count === 0) return 0;

          for (const range of results.items) {
            setWordFontProps(range, fontFormat);
          }

          if (Object.keys(paragraphFormat).length > 0) {
            for (const range of results.items) {
              range.paragraphs.load("items");
            }
            await context.sync();
            for (const range of results.items) {
              for (const paragraph of range.paragraphs.items) {
                setWordParagraphProps(paragraph, paragraphFormat);
              }
            }
          }

          await context.sync();
          return count;
        });

        if (formatted === 0) {
          return {
            content: [
              {
                type: "text",
                text: `No occurrences of "${params.text}" found.`,
              },
            ],
            details: undefined,
          };
        }

        const applied: string[] = [];
        if (params.bold !== undefined) applied.push(`bold=${params.bold}`);
        if (params.italic !== undefined)
          applied.push(`italic=${params.italic}`);
        if (params.underline !== undefined)
          applied.push(`underline=${params.underline}`);
        if (params.size !== undefined) applied.push(`size=${params.size}`);
        if (params.name !== undefined) applied.push(`name=${params.name}`);
        if (params.color !== undefined) applied.push(`color=${params.color}`);
        if (params.style !== undefined) applied.push(`style=${params.style}`);
        if (params.alignment !== undefined)
          applied.push(`alignment=${params.alignment}`);
        if (params.spaceBefore !== undefined)
          applied.push(`spaceBefore=${params.spaceBefore}`);
        if (params.spaceAfter !== undefined)
          applied.push(`spaceAfter=${params.spaceAfter}`);
        if (params.lineSpacing !== undefined)
          applied.push(`lineSpacing=${params.lineSpacing}`);
        if (params.firstLineIndent !== undefined)
          applied.push(`firstLineIndent=${params.firstLineIndent}`);
        if (params.leftIndent !== undefined)
          applied.push(`leftIndent=${params.leftIndent}`);

        return {
          content: [
            {
              type: "text",
              text: `Formatted ${formatted} occurrence${formatted === 1 ? "" : "s"} of "${params.text}" (${applied.join(", ")}).`,
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error formatting text: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
