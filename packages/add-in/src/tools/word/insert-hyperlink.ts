/**
 * word_insert_hyperlink — Insert a clickable hyperlink with display text into
 * the live Word document.
 */

import { WORD_INSERT_HYPERLINK_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = WORD_INSERT_HYPERLINK_PARAMETERS;
type Params = Static<typeof schema>;

export function createWordInsertHyperlinkTool(): AgentTool<typeof schema> {
  return {
    name: "word_insert_hyperlink",
    label: t("tools.wordInsertHyperlink"),
    description:
      "Insert a clickable hyperlink with display text into the live Word document.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      if (params.text.length === 0) {
        return {
          content: [{ type: "text", text: "Missing display text." }],
          details: undefined,
        };
      }
      if (
        !/^https?:\/\//i.test(params.url) &&
        !params.url.startsWith("mailto:")
      ) {
        return {
          content: [
            {
              type: "text",
              text: "url must start with http://, https://, or mailto:.",
            },
          ],
          details: undefined,
        };
      }

      try {
        const location = params.location === "start" ? "Start" : "End";
        await wordRun(async (context) => {
          const body = context.document.body;
          const range = body.insertText(params.text, location);
          range.hyperlinks.add(range, {
            address: params.url,
            ...(params.screenTip === undefined
              ? {}
              : { screenTip: params.screenTip }),
          });
          await context.sync();
        });

        return {
          content: [
            {
              type: "text",
              text: `Inserted a hyperlink "${params.text}" → ${params.url} at the ${params.location ?? "end"} of the document.`,
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error inserting hyperlink: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
