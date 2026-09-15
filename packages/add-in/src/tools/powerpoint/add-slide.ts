/**
 * powerpoint_add_slide — Append a new slide to the live presentation.
 *
 * Navigating the editing surface (`setSelectedSlides`) is PowerPointApi 1.5.
 * On a host known to lack it the slide is still added; the result says the
 * navigation was skipped rather than pretending it happened. An "unknown"
 * verdict (no requirements API) still attempts the call.
 */

import {
  officeToolDescription,
  POWERPOINT_ADD_SLIDE_PARAMETERS,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { powerPointRun } from "../../powerpoint/helpers.js";
import { capabilityUnsupported } from "../../office/capabilities.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = POWERPOINT_ADD_SLIDE_PARAMETERS;

type Params = Static<typeof schema>;

export function createPowerPointAddSlideTool(): AgentTool<typeof schema> {
  return {
    name: "powerpoint_add_slide",
    label: t("tools.powerPointAddSlide"),
    description: officeToolDescription("powerpoint.add_slide"),
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        // `setSelectedSlides` needs PowerPointApi 1.5. Skip it only when the
        // host is *known* to lack it.
        const canNavigate = !capabilityUnsupported(
          "powerpoint.slide.selection",
        );

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
          if (canNavigate) {
            context.presentation.setSelectedSlides([added.id]);
            await context.sync();
          }
          return added.id;
        });

        const navigationNote = canNavigate
          ? " and navigated to it"
          : " (navigation skipped: selecting slides needs PowerPointApi 1.5)";
        return {
          content: [
            {
              type: "text",
              text: `Added a new slide (id: ${newId})${navigationNote}.`,
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
