/**
 * powerpoint_add_text_box — Add a text box with content to a slide.
 *
 * Optional formatting (text props plus shape fill) is applied to the inserted
 * shape in the same batch. Shape fill needs PowerPointApi 1.4 and fails loudly
 * on an older host.
 */

import {
  POWERPOINT_ADD_TEXT_BOX_PARAMETERS,
  officeToolDescription,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  hasPowerPointShapeFill,
  hasPowerPointShapeLine,
  hasPowerPointTextFormat,
  powerPointRun,
  powerPointShapesGateMessage,
  powerPointTextFormatGap,
  setPowerPointShapeFill,
  setPowerPointShapeLine,
  setPowerPointTextFormat,
  type PowerPointShapeFill,
  type PowerPointShapeLine,
  type PowerPointTextFormat,
} from "../../powerpoint/helpers.js";
import { requireCapability } from "../../office/capabilities.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = POWERPOINT_ADD_TEXT_BOX_PARAMETERS;

type Params = Static<typeof schema>;

export function createPowerPointAddTextBoxTool(): AgentTool<typeof schema> {
  return {
    name: "powerpoint_add_text_box",
    label: t("tools.powerPointAddTextBox"),
    description: officeToolDescription("powerpoint.add_text_box"),
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
        // `slide.shapes` is PowerPointApi 1.3.
        const shapesGap = powerPointShapesGateMessage();
        if (shapesGap !== null) throw new Error(shapesGap);

        const format: PowerPointTextFormat = params;
        const fill: PowerPointShapeFill = params;
        const line: PowerPointShapeLine = params;
        const hasFormatting = hasPowerPointTextFormat(format);
        const hasFill = hasPowerPointShapeFill(fill);
        const hasLine = hasPowerPointShapeLine(line);

        if (hasFormatting) {
          const gap = powerPointTextFormatGap(format);
          if (gap !== null) throw new Error(gap);
        }
        if (hasFill) {
          const gap = requireCapability("powerpoint.shape.fill");
          if (gap !== null) throw new Error(gap);
        }
        if (hasLine) {
          const gap = requireCapability("powerpoint.shape.line");
          if (gap !== null) throw new Error(gap);
        }

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
          if (hasFill) {
            setPowerPointShapeFill(shape, fill);
          }
          if (hasLine) {
            setPowerPointShapeLine(shape, line);
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
        const formatNote =
          hasFormatting || hasFill || hasLine ? " (formatted)" : "";
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
