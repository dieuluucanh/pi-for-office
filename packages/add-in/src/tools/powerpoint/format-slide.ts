/**
 * powerpoint_format_slide — Format all text shapes on a slide.
 *
 * PowerPoint has no text-search API (unlike Word), so existing content is
 * formatted by iterating every text-bearing shape on the slide and applying
 * the requested font/alignment/fill properties. Slide background fill is a
 * slide-level property (PowerPointApi 1.10) applied once.
 */

import {
  POWERPOINT_FORMAT_SLIDE_PARAMETERS,
  officeToolDescription,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  hasPowerPointShapeFill,
  hasPowerPointShapeLine,
  hasPowerPointSlideBackground,
  hasPowerPointTextFormat,
  powerPointRun,
  powerPointShapesGateMessage,
  powerPointTextFormatGap,
  setPowerPointShapeFill,
  setPowerPointShapeLine,
  setPowerPointSlideBackground,
  setPowerPointTextFormat,
  type PowerPointShapeFill,
  type PowerPointShapeLine,
  type PowerPointSlideBackground,
  type PowerPointTextFormat,
} from "../../powerpoint/helpers.js";
import { requireCapability } from "../../office/capabilities.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = POWERPOINT_FORMAT_SLIDE_PARAMETERS;

type Params = Static<typeof schema>;

export function createPowerPointFormatSlideTool(): AgentTool<typeof schema> {
  return {
    name: "powerpoint_format_slide",
    label: t("tools.powerPointFormatSlide"),
    description: officeToolDescription("powerpoint.format_slide"),
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
      const fill: PowerPointShapeFill = params;
      const line: PowerPointShapeLine = params;
      const background: PowerPointSlideBackground = params;
      const hasFormatting =
        hasPowerPointTextFormat(format) ||
        hasPowerPointShapeFill(fill) ||
        hasPowerPointShapeLine(line);
      const hasBackground = hasPowerPointSlideBackground(background);

      if (!hasFormatting && !hasBackground) {
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
        // `slide.shapes` is PowerPointApi 1.3.
        const shapesGap = powerPointShapesGateMessage();
        if (shapesGap !== null) throw new Error(shapesGap);

        // Shape fill is PowerPointApi 1.4; slide background fill is 1.10.
        // Gate each group separately so a 1.4 host still gets shape fill.
        if (hasPowerPointTextFormat(format)) {
          const gap = powerPointTextFormatGap(format);
          if (gap !== null) throw new Error(gap);
        }
        if (hasPowerPointShapeFill(fill)) {
          const gap = requireCapability("powerpoint.shape.fill");
          if (gap !== null) throw new Error(gap);
        }
        if (hasPowerPointShapeLine(line)) {
          const gap = requireCapability("powerpoint.shape.line");
          if (gap !== null) throw new Error(gap);
        }
        if (hasBackground) {
          const gap = requireCapability("powerpoint.slide.background");
          if (gap !== null) throw new Error(gap);
        }

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
            if (hasPowerPointTextFormat(format)) {
              setPowerPointTextFormat(shape.textFrame.textRange, format);
            }
            if (hasPowerPointShapeFill(fill)) {
              setPowerPointShapeFill(shape, fill);
            }
            if (hasPowerPointShapeLine(line)) {
              setPowerPointShapeLine(shape, line);
            }
            count += 1;
          }

          if (hasBackground) {
            setPowerPointSlideBackground(slide, background);
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

        const applied = describeAppliedFormat(params);
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

/** Human-readable list of the properties that will be applied. */
function describeAppliedFormat(params: Params): string[] {
  const applied: string[] = [];
  if (params.bold !== undefined) applied.push(`bold=${params.bold}`);
  if (params.italic !== undefined) applied.push(`italic=${params.italic}`);
  if (params.underline !== undefined)
    applied.push(`underline=${String(params.underline)}`);
  if (params.fontSize !== undefined)
    applied.push(`fontSize=${params.fontSize}`);
  if (params.fontName !== undefined)
    applied.push(`fontName=${params.fontName}`);
  if (params.fontColor !== undefined)
    applied.push(`fontColor=${params.fontColor}`);
  if (params.alignment !== undefined)
    applied.push(`alignment=${params.alignment}`);
  if (params.bulletVisible !== undefined)
    applied.push(`bulletVisible=${params.bulletVisible}`);
  if (params.bulletStyle !== undefined)
    applied.push(`bulletStyle=${params.bulletStyle}`);
  if (params.bulletType !== undefined)
    applied.push(`bulletType=${params.bulletType}`);
  if (params.indentLevel !== undefined)
    applied.push(`indentLevel=${params.indentLevel}`);
  if (params.strikethrough !== undefined)
    applied.push(`strikethrough=${params.strikethrough}`);
  if (params.doubleStrikethrough !== undefined)
    applied.push(`doubleStrikethrough=${params.doubleStrikethrough}`);
  if (params.smallCaps !== undefined)
    applied.push(`smallCaps=${params.smallCaps}`);
  if (params.allCaps !== undefined) applied.push(`allCaps=${params.allCaps}`);
  if (params.subscript !== undefined)
    applied.push(`subscript=${params.subscript}`);
  if (params.superscript !== undefined)
    applied.push(`superscript=${params.superscript}`);
  if (params.fillColor !== undefined) applied.push(`fill=${params.fillColor}`);
  if (params.fillTransparency !== undefined)
    applied.push(`fillTransparency=${params.fillTransparency}`);
  if (params.fillClear === true) applied.push("fill cleared");
  if (params.lineColor !== undefined) applied.push(`line=${params.lineColor}`);
  if (params.lineWeight !== undefined)
    applied.push(`lineWeight=${params.lineWeight}`);
  if (params.lineStyle !== undefined)
    applied.push(`lineStyle=${params.lineStyle}`);
  if (params.lineDashStyle !== undefined)
    applied.push(`lineDashStyle=${params.lineDashStyle}`);
  if (params.lineVisible !== undefined)
    applied.push(`lineVisible=${params.lineVisible}`);
  if (params.lineTransparency !== undefined)
    applied.push(`lineTransparency=${params.lineTransparency}`);
  if (params.slideBackgroundColor !== undefined)
    applied.push(`background=${params.slideBackgroundColor}`);
  if (params.slideBackgroundTransparency !== undefined)
    applied.push(
      `backgroundTransparency=${params.slideBackgroundTransparency}`,
    );
  if (params.slideBackgroundClear === true) applied.push("background cleared");
  return applied;
}
