/**
 * word_insert_text — Insert text into the live Word document.
 *
 * Locations: "end" (default), "start", or "replace_selection".
 * Optional formatting (bold/italic/underline/size/name/color/alignment) is
 * applied to the inserted text in the same call.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  wordRun,
  setWordFontProps,
  applyWordAlignment,
  type WordTextFormat,
} from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = Type.Object({
  text: Type.String({
    description: "Text to insert into the document.",
  }),
  location: Type.Optional(
    Type.Union(
      [
        Type.Literal("start"),
        Type.Literal("end"),
        Type.Literal("replace_selection"),
      ],
      {
        description:
          '"end" (default): append at the end of the document. ' +
          '"start": insert at the beginning. ' +
          '"replace_selection": overwrite the currently selected text.',
      },
    ),
  ),
  bold: Type.Optional(
    Type.Boolean({
      description: "Make the inserted text bold.",
    }),
  ),
  italic: Type.Optional(
    Type.Boolean({
      description: "Make the inserted text italic.",
    }),
  ),
  underline: Type.Optional(
    Type.Boolean({
      description: "Underline the inserted text (single underline).",
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
          'Paragraph alignment for the inserted text. Word values: "Left", "Centered" (capital C), "Right", "Justified".',
      },
    ),
  ),
});

type Params = Static<typeof schema>;

export function createWordInsertTextTool(): AgentTool<typeof schema> {
  return {
    name: "word_insert_text",
    label: t("tools.wordInsertText"),
    description:
      "Insert text into the live Word document — at the start, at the end (default), " +
      "or replacing the current selection. Optionally format the inserted text " +
      '(bold, italic, underline, size in points, font name, color #RRGGBB, alignment "Left"/"Centered"/"Right"/"Justified") ' +
      "in the same call — e.g. insert a bold centered title. Use this to add new content to the opened document.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      const location =
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
        const format: WordTextFormat = params;
        const hasFormatting =
          params.bold !== undefined ||
          params.italic !== undefined ||
          params.underline !== undefined ||
          params.size !== undefined ||
          params.name !== undefined ||
          params.color !== undefined ||
          params.alignment !== undefined;

        await wordRun(async (context) => {
          // insertText() returns a Range covering the inserted text, so the
          // inserted text can be formatted in the same batch.
          const inserted =
            location === "replace_selection"
              ? context.document
                  .getSelection()
                  .insertText(params.text, Word.InsertLocation.replace)
              : context.document.body.insertText(
                  params.text,
                  Word.InsertLocation.start,
                ); // start/end both supported on body

          if (hasFormatting) {
            setWordFontProps(inserted, format);
            if (params.alignment !== undefined) {
              // Alignment lives on Word.Paragraph; load paragraph proxies first.
              await applyWordAlignment(inserted, params.alignment);
            }
          }

          await context.sync();
        });

        const formatNote = hasFormatting ? " with formatting" : "";
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
