/**
 * powerpoint_get_overview — Structural overview of the live presentation.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { powerPointRun } from "../../powerpoint/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = Type.Object({});

type Params = Static<typeof schema>;

export function createPowerPointGetOverviewTool(): AgentTool<typeof schema> {
  return {
    name: "powerpoint_get_overview",
    label: t("tools.powerPointGetOverview"),
    description:
      "Get a structural overview of the live PowerPoint presentation: slide count and " +
      "shape count per slide. Call this before editing to learn the presentation structure.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await powerPointRun(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();

          const perSlide = slides.items.map((slide) => {
            slide.load("shapes");
            return slide;
          });
          await context.sync();

          return perSlide.map((slide) => ({
            id: slide.id,
            shapeCount: slide.shapes.items.length,
          }));
        });

        const lines: string[] = ["**Presentation overview**"];
        lines.push(`- Slides: ${result.length}`);
        for (let i = 0; i < result.length; i += 1) {
          const slide = result[i];
          if (slide === undefined) continue;
          lines.push(`  - Slide ${i + 1}: ${slide.shapeCount} shapes`);
        }
        lines.push(
          "Use powerpoint_read_slide to read a slide's text; powerpoint_add_slide / powerpoint_add_text_box to edit.",
        );

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error getting presentation overview: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
