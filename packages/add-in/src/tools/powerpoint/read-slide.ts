/**
 * powerpoint_read_slide — Read all text on a slide of the live presentation.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { powerPointRun } from "../../powerpoint/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = Type.Object({
  slideIndex: Type.Number({
    description: "1-based slide number to read.",
  }),
});

type Params = Static<typeof schema>;

interface SlideTextEntry {
  name: string;
  type: string;
  text: string;
}

/** Read the text of every text-bearing shape on a slide. Mirrors ppt-ops. */
async function readSlideTexts(
  slide: PowerPoint.Slide,
): Promise<SlideTextEntry[]> {
  // Phase 1: shape proxies + which ones carry text.
  const shapes = slide.shapes;
  shapes.load("items/type/name");
  // textFrame presence must be probed separately to avoid touching text on
  // non-text shapes.
  shapes.load("items/textFrame/hasText");
  await shapes.context.sync();

  // Phase 2: load text for text-bearing shapes.
  const texts = shapes.items.map((shape) => {
    if (shape.textFrame?.hasText) {
      const range = shape.textFrame.textRange;
      range.load("text");
      return { shape, range };
    }
    return null;
  });
  await shapes.context.sync();

  const out: SlideTextEntry[] = [];
  for (const entry of texts) {
    if (entry === null) continue;
    const text = entry.range.text ?? "";
    if (text.trim().length === 0) continue;
    out.push({ name: entry.shape.name, type: entry.shape.type, text });
  }
  return out;
}

export function createPowerPointReadSlideTool(): AgentTool<typeof schema> {
  return {
    name: "powerpoint_read_slide",
    label: t("tools.powerPointReadSlide"),
    description:
      "Read all text content on one slide of the live PowerPoint presentation " +
      "(shapes, text frames, notes). Call powerpoint_get_overview first to learn the structure.",
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

      try {
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
          if (slide === undefined) {
            return { error: "Slide not found." };
          }
          const slideTexts = await readSlideTexts(slide);
          return { slideTexts, id: slide.id };
        });

        if ("error" in out) {
          return {
            content: [{ type: "text", text: out.error }],
            details: undefined,
          };
        }

        const lines: string[] = [`**Slide ${slideIndex}** (id: ${out.id})`];
        if (out.slideTexts.length === 0) {
          lines.push("_No text content on this slide._");
        } else {
          for (const item of out.slideTexts) {
            lines.push("");
            lines.push(`**${item.name || item.type}:**`);
            lines.push(item.text);
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
              text: `Error reading slide: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
