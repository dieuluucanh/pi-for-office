/**
 * word_get_overview — Structural overview of the live Word document.
 *
 * Paragraph/table counts plus a heading outline that includes heading TEXT
 * (bounded), so the agent sees both the skeleton and where content lives.
 */

import { WORD_GET_OVERVIEW_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordCount, wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const MAX_HEADINGS = 200;

const schema = WORD_GET_OVERVIEW_PARAMETERS;
type Params = Static<typeof schema>;

export function createWordGetOverviewTool(): AgentTool<typeof schema> {
  return {
    name: "word_get_overview",
    label: t("tools.wordGetOverview"),
    description:
      "Get a structural overview of the live Word document: paragraph/table counts, " +
      "word count, and the heading outline WITH heading text. Call this before editing " +
      "to learn the document structure and find where sections live.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const overview = await wordRun(async (context) => {
          const body = context.document.body;

          const paragraphs = body.paragraphs;
          paragraphs.load("items/style");
          paragraphs.load("items/text");
          const tables = body.tables;
          tables.load("items");
          body.load("text");
          await context.sync();

          const headings: Array<{
            level: number;
            text: string;
            index: number;
          }> = [];
          let index = 0;
          for (const para of paragraphs.items) {
            index += 1;
            const style = para.style;
            if (
              typeof style === "string" &&
              style.toLowerCase().startsWith("heading")
            ) {
              const level = Number.parseInt(
                style.replace(/^Heading\s*/i, ""),
                10,
              );
              headings.push({
                level: Number.isFinite(level) ? level : 0,
                text: (para.text ?? "").trim().slice(0, 120),
                index,
              });
              if (headings.length >= MAX_HEADINGS) break;
            }
          }

          const tableDims: string[] = [];
          for (const table of tables.items) {
            const rows = table.rows;
            const cols = table.columns;
            rows.load("items");
            cols.load("items");
            await context.sync();
            tableDims.push(`${rows.items.length}×${cols.items.length}`);
          }

          return {
            paragraphCount: paragraphs.items.length,
            tableCount: tables.items.length,
            tableDims,
            wordCount: wordCount(body.text ?? ""),
            headings,
          };
        });

        const lines: string[] = ["**Document overview**"];
        lines.push(`- Paragraphs: ${overview.paragraphCount}`);
        lines.push(`- Words: ${overview.wordCount}`);
        if (overview.tableCount > 0) {
          lines.push(
            `- Tables: ${overview.tableDims.join(", ") || overview.tableCount}`,
          );
        } else {
          lines.push("- Tables: 0");
        }
        if (overview.headings.length > 0) {
          lines.push("");
          lines.push("- Heading outline:");
          for (const h of overview.headings) {
            const indent = "  ".repeat(Math.max(0, h.level - 1));
            lines.push(`  ${indent}#${h.level} ${h.text || "(empty heading)"}`);
          }
        } else {
          lines.push("- Headings: none detected");
        }
        lines.push(
          "Use word_read_document to read text; word_insert_blocks / word_insert_text / word_replace_text to edit.",
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
              text: `Error getting document overview: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
