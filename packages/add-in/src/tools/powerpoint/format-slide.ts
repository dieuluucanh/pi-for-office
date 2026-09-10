/**
 * powerpoint_format_slide — Format all text shapes on a slide.
 *
 * PowerPoint has no text-search API (unlike Word), so existing content is
 * formatted by iterating every text-bearing shape on the slide and applying
 * the requested font/alignment properties.
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
    description: "1-based slide number to format.",
  }),
  bold: Type.Optional(
    Type.Boolean({
      description: "Set bold (true) or unbold (false) on every text box.",
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
  fontSize: Type.Optional(
    Type.Number({
      description: "Font size in points (e.g. 24).",
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
          'Horizontal alignment. PowerPoint values: "Left", "Center", "Right", "Justify".',
      },
    ),
  ),
});

type Params = Static<typeof schema>;

export function createPowerPointFormatSlideTool(): AgentTool<typeof schema> {
  return {
    name: "powerpoint_format_slide",
    label: t("tools.powerPointFormatSlide"),
    description:
      "Apply formatting (bold, italic, underline, fontSize in points, fontName, fontColor #RRGGBB, " +
      'alignment "Left"/"Center"/"Right"/"Justify") to every text box on a slide of the live ' +
      "PowerPoint presentation. Use this to restyle existing content, e.g. make all text on a slide bold and centered.",
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

      const format: PowerPointTextFormat = params;
      const hasFormatting =
        params.bold !== undefined ||
        params.italic !== undefined ||
        params.underline !== undefined ||
        params.fontSize !== undefined ||
        params.fontName !== undefined ||
        params.fontColor !== undefined ||
        params.alignment !== undefined;

      if (!hasFormatting) {
        return {
          content: [
            {
              type: "text",
              text: "No formatting options provided (nothing to apply).",
            },
          ],
          details: undefined,
        };
      }

      try {
        const formattedCount = await powerPointRun(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();

          if (slides.items.length < slideIndex) {
            return -1;
          }
          const slide = slides.items[slideIndex - 1];
          if (slide === undefined) return -1;

          // Two-phase load (mirrors bridge readSlideTexts): shape proxies +
          // which shapes carry text, then apply formatting to text-bearing ones.
          const shapes = slide.shapes;
          shapes.load("items/type/name");
          shapes.load("items/textFrame/hasText");
          await context.sync();

          let count = 0;
          for (const shape of shapes.items) {
            if (!shape.textFrame?.hasText) continue;
            setPowerPointTextFormat(shape.textFrame.textRange, format);
            count += 1;
          }
          await context.sync();
          return count;
        });

        if (formattedCount < 0) {
          return {
            content: [{ type: "text", text: `Slide ${slideIndex} not found.` }],
            details: undefined,
          };
        }

        const applied: string[] = [];
        if (params.bold !== undefined) applied.push(`bold=${params.bold}`);
        if (params.italic !== undefined)
          applied.push(`italic=${params.italic}`);
        if (params.underline !== undefined)
          applied.push(`underline=${params.underline}`);
        if (params.fontSize !== undefined)
          applied.push(`fontSize=${params.fontSize}`);
        if (params.fontName !== undefined)
          applied.push(`fontName=${params.fontName}`);
        if (params.fontColor !== undefined)
          applied.push(`fontColor=${params.fontColor}`);
        if (params.alignment !== undefined)
          applied.push(`alignment=${params.alignment}`);

        return {
          content: [
            {
              type: "text",
              text: `Formatted ${formattedCount} text box${formattedCount === 1 ? "" : "es"} on slide ${slideIndex} (${applied.join(", ")}).`,
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error formatting slide: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
