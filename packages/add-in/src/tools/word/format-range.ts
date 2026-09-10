/**
 * word_format_range — Find text in the live Word document and apply formatting.
 *
 * Locates text by literal content (same search model as word_replace_text) and
 * applies bold/italic/underline/font size/font name/color/paragraph alignment
 * to every match.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  wordRun,
  setWordFontProps,
  type WordTextFormat,
} from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = Type.Object({
  text: Type.String({
    description:
      "Literal text to find and format. Every occurrence is formatted.",
  }),
  matchCase: Type.Optional(
    Type.Boolean({
      description: "Case-sensitive search (default false).",
    }),
  ),
  bold: Type.Optional(
    Type.Boolean({
      description: "Set bold (true) or unbold (false).",
    }),
  ),
  italic: Type.Optional(
    Type.Boolean({
      description: "Set italic (true) or unitalicize (false).",
    }),
  ),
  underline: Type.Optional(
    Type.Boolean({
      description: "Underline the text (single underline).",
    }),
  ),
  size: Type.Optional(
    Type.Number({
      description: "Font size in points (e.g. 16).",
    }),
  ),
  name: Type.Optional(
    Type.String({
      description: 'Font name (e.g. "Times New Roman").',
    }),
  ),
  color: Type.Optional(
    Type.String({
      description: 'Font color as #RRGGBB (e.g. "#000000").',
    }),
  ),
  alignment: Type.Optional(
    Type.Union(
      [
        Type.Literal("Left"),
        Type.Literal("Centered"),
        Type.Literal("Right"),
        Type.Literal("Justified"),
      ],
      {
        description:
          'Paragraph alignment. Word values: "Left", "Centered" (note the capital C), "Right", "Justified".',
      },
    ),
  ),
});

type Params = Static<typeof schema>;

export function createWordFormatRangeTool(): AgentTool<typeof schema> {
  return {
    name: "word_format_range",
    label: t("tools.wordFormatRange"),
    description:
      "Find text in the live Word document and apply formatting: bold, italic, underline, " +
      'font size (points), font name, color (#RRGGBB), and paragraph alignment ("Left"/"Centered"/"Right"/"Justified"). ' +
      "Use this to bold/size/center existing content — e.g. format a document title.",
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

      const format: WordTextFormat = params;

      try {
        const formatted = await wordRun(async (context) => {
          const results = context.document.body.search(params.text, {
            matchCase: params.matchCase === true,
          });
          results.load("text");
          await context.sync();

          const count = results.items.length;
          if (count === 0) return 0;

          // Font props are writable on the proxy — queue them for all ranges
          // without extra syncs.
          for (const range of results.items) {
            setWordFontProps(range, format);
          }

          if (params.alignment !== undefined) {
            // Alignment lives on Word.Paragraph; paragraph proxies must be
            // loaded (and synced) before they can be written.
            for (const range of results.items) {
              range.paragraphs.load("items");
            }
            await context.sync();
            for (const range of results.items) {
              for (const paragraph of range.paragraphs.items) {
                paragraph.alignment = params.alignment;
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
        if (params.alignment !== undefined)
          applied.push(`alignment=${params.alignment}`);

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
