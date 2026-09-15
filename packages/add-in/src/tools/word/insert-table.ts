/**
 * word_insert_table — Insert a values[row][col] table into the live Word
 * document, with optional header-row bolding, table style and alignment.
 */

import {
  WORD_INSERT_TABLE_PARAMETERS,
  officeToolDescription,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun, wordTableInsertGateMessage } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const MAX_TABLE_CELLS = 25_000;

const schema = WORD_INSERT_TABLE_PARAMETERS;
type Params = Static<typeof schema>;

export function createWordInsertTableTool(): AgentTool<typeof schema> {
  return {
    name: "word_insert_table",
    label: t("tools.wordInsertTable"),
    description: officeToolDescription("word.insert_table"),
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      const rows = params.cells.length;
      const firstRow = params.cells[0];
      const cols = firstRow?.length ?? 0;
      if (rows === 0 || cols === 0) {
        return {
          content: [
            { type: "text", text: "cells must be a non-empty 2D array." },
          ],
          details: undefined,
        };
      }
      for (const row of params.cells) {
        if (row.length === cols) continue;
        return {
          content: [
            {
              type: "text",
              text: "cells must be rectangular (every row the same length).",
            },
          ],
          details: undefined,
        };
      }
      if (rows * cols > MAX_TABLE_CELLS) {
        return {
          content: [
            {
              type: "text",
              text: `Table too large (${rows}×${cols} = ${rows * cols} cells; max ${MAX_TABLE_CELLS}).`,
            },
          ],
          details: undefined,
        };
      }

      try {
        // `Body.insertTable` is WordApi 1.3 — fail loud on an older host.
        const gap = wordTableInsertGateMessage();
        if (gap !== null) throw new Error(gap);

        const location = params.location === "start" ? "Start" : "End";
        await wordRun(async (context) => {
          const body = context.document.body;
          const table = body.insertTable(rows, cols, location, params.cells);
          if (params.headerRow && rows > 0) {
            table.rows.getFirst().font.bold = true;
          }
          if (params.style !== undefined) {
            table.style = params.style;
          }
          if (params.alignment !== undefined) {
            table.alignment = params.alignment;
          }
          await context.sync();
        });

        return {
          content: [
            {
              type: "text",
              text:
                `Inserted a ${rows}×${cols} table at the ${params.location ?? "end"} of the document` +
                (params.headerRow ? " (header row bolded)" : "") +
                (params.style !== undefined
                  ? ` with style "${params.style}"`
                  : "") +
                ".",
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error inserting table: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
