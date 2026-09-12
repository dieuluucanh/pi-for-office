/**
 * word_insert_page_break — Insert a page break at the start or end of the
 * live Word document.
 */

import { WORD_INSERT_PAGE_BREAK_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = WORD_INSERT_PAGE_BREAK_PARAMETERS;
type Params = Static<typeof schema>;

export function createWordInsertPageBreakTool(): AgentTool<typeof schema> {
  return {
    name: "word_insert_page_break",
    label: t("tools.wordInsertPageBreak"),
    description:
      "Insert a page break at the start or end of the live Word document.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const location = params.location === "start" ? "Start" : "End";
        await wordRun(async (context) => {
          context.document.body.insertBreak("Page", location);
          await context.sync();
        });
        return {
          content: [
            {
              type: "text",
              text: `Inserted a page break at the ${params.location ?? "end"} of the document.`,
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error inserting page break: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
