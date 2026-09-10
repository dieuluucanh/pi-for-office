/**
 * word_get_overview — Structural overview of the live Word document.
 *
 * Paragraph/table counts plus a heading-level outline (style names only —
 * extracting heading text for the whole doc would require loading every
 * paragraph's text; use word_read_document for content).
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = Type.Object({});

type Params = Static<typeof schema>;

export function createWordGetOverviewTool(): AgentTool<typeof schema> {
  return {
    name: "word_get_overview",
    label: t("tools.wordGetOverview"),
    description:
      "Get a structural overview of the live Word document: paragraph/table counts " +
      "and the heading-level outline. Call this before editing to learn the document structure.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const overview = await wordRun(async (context) => {
          const body = context.document.body;

          // Paragraph proxies: style only (no text) to build the heading outline.
          const paragraphs = body.paragraphs;
          paragraphs.load("items/style");
          const tables = body.tables;
          tables.load("items");

          await context.sync();

          const headings: string[] = [];
          let paragraphCount = 0;
          for (const para of paragraphs.items) {
            paragraphCount += 1;
            const style = para.style;
            if (
              typeof style === "string" &&
              style.toLowerCase().startsWith("heading")
            ) {
              // Keep only the style name — extracting heading text for the whole
              // doc would require loading every paragraph's text (see read_document).
              const level = style.replace(/^Heading\s*/i, "");
              headings.push(`H${level || "?"}`);
            }
            if (headings.length >= 200 || paragraphCount >= 200_000) break;
          }

          return {
            paragraphCount: paragraphs.items.length,
            tableCount: tables.items.length,
            headingCountByLevel: headings.reduce<Record<string, number>>(
              (acc, h) => {
                acc[h] = (acc[h] ?? 0) + 1;
                return acc;
              },
              {},
            ),
          };
        });

        const lines: string[] = ["**Document overview**"];
        lines.push(`- Paragraphs: ${overview.paragraphCount}`);
        lines.push(`- Tables: ${overview.tableCount}`);
        const headingEntries = Object.entries(overview.headingCountByLevel);
        if (headingEntries.length > 0) {
          lines.push(
            `- Headings: ${headingEntries.map(([level, n]) => `${level}×${n}`).join(", ")}`,
          );
        } else {
          lines.push("- Headings: none detected");
        }
        lines.push(
          "Use word_read_document to read text; word_insert_text / word_replace_text to edit.",
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
