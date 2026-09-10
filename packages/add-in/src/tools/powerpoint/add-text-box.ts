/**
 * powerpoint_add_text_box — Add a text box with content to a slide.
 *
 * Optional formatting (bold/italic/underline/fontSize/fontName/fontColor/
 * alignment) is applied to the inserted text in the same batch.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  powerPointRun,
  setPowerPointTextFormat,
  type PowerPointTextFormat,
} from "../../powerpoint/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = Type.Object({
  slideIndex: Type.Number({
    description: "1-based slide number to add the text box to.",
  }),
  text: Type.String({
    description: "Text content for the text box.",
  }),
  x: Type.Optional(
    Type.Number({
      description: "Left edge in points (default 72).",
    }),
  ),
  y: Type.Optional(
    Type.Number({
      description: "Top edge in points (default 72).",
    }),
  ),
  width: Type.Optional(
    Type.Number({
      description: "Width in points (default 400).",
    }),
  ),
  height: Type.Optional(
    Type.Number({
      description: "Height in points (default 60).",
    }),
  ),
  bold: Type.Optional(
    Type.Boolean({
      description: "Make the text bold.",
    }),
  ),
  italic: Type.Optional(
    Type.Boolean({
      description: "Make the text italic.",
    }),
  ),
  underline: Type.Optional(
    Type.Boolean({
      description: "Underline the text (single underline).",
    }),
  ),
  fontSize: Type.Optional(
    Type.Number({
      description: "Font size in points (e.g. 24 for titles).",
    }),
  ),
  fontName: Type.Optional(
    Type.String({
      description: 'Font name (e.g. "Arial").',
    }),
  ),
  fontColor: Type.Optional(
    Type.String({
      description: 'Font color as #RRGGBB (e.g. "#333333").',
    }),
  ),
  alignment: Type.Optional(
    Type.Union(
      [
        Type.Literal("Left"),
        Type.Literal("Center"),
        Type.Literal("Right"),
        Type.Literal("Justify"),
      ],
      {
        description:
          'Horizontal alignment for the text. PowerPoint values: "Left", "Center", "Right", "Justify".',
      },
    ),
  ),
});

type Params = Static<typeof schema>;

export function createPowerPointAddTextBoxTool(): AgentTool<typeof schema> {
  return {
    name: "powerpoint_add_text_box",
    label: t("tools.powerPointAddTextBox"),
    description:
      "Add a text box with the given text to a slide of the live PowerPoint presentation. " +
      "Coordinates and geometry are in points. Optionally format the text in the same call " +
      '(bold, italic, underline, fontSize/fontName/fontColor, alignment "Left"/"Center"/"Right"/"Justify") — ' +
      "use formatting to produce well-styled slides instead of plain unformatted text.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      const slideIndex = Math.floor(params.slideIndex);
      if (!Number.isFinite(slideIndex) || slideIndex < 1) {
        return {
          content: [
            {
              type: "text",
              text: "slideIndex must be a 1-based slide number.",
            },
          ],
          details: undefined,
        };
      }
      if (params.text.length === 0) {
        return {
          content: [{ type: "text", text: "Nothing to add (empty text)." }],
          details: undefined,
        };
      }

      try {
        const format: PowerPointTextFormat = params;
        const hasFormatting =
          params.bold !== undefined ||
          params.italic !== undefined ||
          params.underline !== undefined ||
          params.fontSize !== undefined ||
          params.fontName !== undefined ||
          params.fontColor !== undefined ||
          params.alignment !== undefined;

        const out = await powerPointRun(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();

          if (slides.items.length < slideIndex) {
            return {
              error: `Presentation has only ${slides.items.length} slide(s).`,
            };
          }
          const slide = slides.items[slideIndex - 1];
          if (slide === undefined) return { error: "Slide not found." };

          const options: PowerPoint.ShapeAddOptions = {
            left: params.x ?? 72,
            top: params.y ?? 72,
            width: params.width ?? 400,
            height: params.height ?? 60,
          };
          // addTextBox() returns a Shape whose text can be formatted in the
          // same batch — create + format + id with a single sync.
          const shape = slide.shapes.addTextBox(params.text, options);
          if (hasFormatting) {
            setPowerPointTextFormat(shape.textFrame.textRange, format);
          }
          shape.load("id");
          await context.sync();
          return { id: shape.id };
        });

        if ("error" in out) {
          return {
            content: [{ type: "text", text: out.error }],
            details: undefined,
          };
        }
        const formatNote = hasFormatting ? " (formatted)" : "";
        return {
          content: [
            {
              type: "text",
              text: `Added a text box (${params.text.length} chars)${formatNote} to slide ${slideIndex}.`,
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error adding text box: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
