/**
 * powerpoint_add_slide — Append a new slide to the live presentation.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { powerPointRun } from "../../powerpoint/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = Type.Object({});

type Params = Static<typeof schema>;

export function createPowerPointAddSlideTool(): AgentTool<typeof schema> {
  return {
    name: "powerpoint_add_slide",
    label: t("tools.powerPointAddSlide"),
    description:
      "Append a new slide to the live PowerPoint presentation using the default layout " +
      "and navigate to it. Use powerpoint_add_text_box afterwards to add content to it.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const newId = await powerPointRun(async (context) => {
          const slides = context.presentation.slides;
          const countResult = slides.getCount();
          await context.sync();
          slides.add();
          await context.sync();
          // The appended slide sits at index countResult.value.
          const added = slides.getItemAt(countResult.value);
          added.load("id");
          await context.sync();
          context.presentation.setSelectedSlides([added.id]);
          await context.sync();
          return added.id;
        });
        return {
          content: [
            {
              type: "text",
              text: `Added a new slide (id: ${newId}) and navigated to it.`,
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            { type: "text", text: `Error adding slide: ${getErrorMessage(e)}` },
          ],
          details: undefined,
        };
      }
    },
  };
}
