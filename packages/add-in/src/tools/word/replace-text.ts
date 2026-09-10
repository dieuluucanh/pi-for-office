/**
 * word_replace_text — Find and replace literal text in the live Word document.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = Type.Object({
  find: Type.String({
    description:
      "Literal text to search for. Case-sensitive only when matchCase is true.",
  }),
  replace: Type.String({
    description: "Replacement text.",
  }),
  matchCase: Type.Optional(
    Type.Boolean({
      description: "Case-sensitive matching (default false).",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createWordReplaceTextTool(): AgentTool<typeof schema> {
  return {
    name: "word_replace_text",
    label: t("tools.wordReplaceText"),
    description:
      "Find and replace literal text across the live Word document. " +
      "Prefer this for targeted corrections; use word_insert_text for additions.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      if (params.find.length === 0) {
        return {
          content: [{ type: "text", text: "Missing 'find' text." }],
          details: undefined,
        };
      }

      try {
        const replaced = await wordRun(async (context) => {
          const results = context.document.body.search(params.find, {
            matchCase: params.matchCase === true,
          });
          results.load("length");
          await context.sync();
          const count = results.items.length;
          for (const item of results.items) {
            item.insertText(params.replace, Word.InsertLocation.replace);
          }
          await context.sync();
          return count;
        });

        const summary =
          replaced === 0
            ? `No occurrences of "${params.find}" found.`
            : `Replaced ${replaced} occurrence${replaced === 1 ? "" : "s"} of "${params.find}".`;
        return {
          content: [{ type: "text", text: summary }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error replacing text: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
