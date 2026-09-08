/**
 * Excel bridge executors.
 *
 * These delegate to the existing, battle-tested pi-for-office tools so the
 * bridge path behaves exactly like the in-browser agent path.
 */

import { createReadRangeTool } from "../tools/read-range.js";
import { createWriteCellsTool } from "../tools/write-cells.js";
import { createFillFormulaTool } from "../tools/fill-formula.js";
import { createGetWorkbookOverviewTool } from "../tools/get-workbook-overview.js";
import { createSearchWorkbookTool } from "../tools/search-workbook.js";
import {
  guardExecutor,
  textOfAgentToolResult,
  type OfficeOpExecutor,
} from "./ops.js";

/** Run a native tool with a synthetic bridge call id and unwrap its result. */
async function runTool(
  execute: (args: Record<string, unknown>) => Promise<{
    content?: Array<{ type?: string; text?: unknown }> | string;
    details?: unknown;
  }>,
  args: Record<string, unknown>,
) {
  const result = await execute(args);
  return { text: textOfAgentToolResult(result), details: result.details };
}

export const EXCEL_OPS: ReadonlyMap<string, OfficeOpExecutor> = new Map<
  string,
  OfficeOpExecutor
>([
  [
    "excel.get_overview",
    guardExecutor(async (args) => {
      const overview = createGetWorkbookOverviewTool();
      // The overview tool accepts an optional `sheet` argument.
      const sheet = typeof args.sheet === "string" ? { sheet: args.sheet } : {};
      return runTool((a) => overview.execute("bridge", a as never), sheet);
    }),
  ],
  [
    "excel.read_range",
    guardExecutor(async (args) => {
      const tool = createReadRangeTool();
      const params = {
        range: typeof args.range === "string" ? args.range : "",
        ...(typeof args.mode === "string" ? { mode: args.mode } : {}),
      };
      return runTool((a) => tool.execute("bridge", a as never), params);
    }),
  ],
  [
    "excel.write_cells",
    guardExecutor(async (args) => {
      const tool = createWriteCellsTool();
      const params = {
        start_cell: typeof args.start_cell === "string" ? args.start_cell : "",
        values: Array.isArray(args.values) ? (args.values as unknown[]) : [],
      };
      return runTool((a) => tool.execute("bridge", a as never), params);
    }),
  ],
  [
    "excel.fill_formula",
    guardExecutor(async (args) => {
      const tool = createFillFormulaTool();
      const params = {
        range: typeof args.range === "string" ? args.range : "",
        formula: typeof args.formula === "string" ? args.formula : "",
      };
      return runTool((a) => tool.execute("bridge", a as never), params);
    }),
  ],
  [
    "excel.search_workbook",
    guardExecutor(async (args) => {
      const tool = createSearchWorkbookTool();
      return runTool((a) => tool.execute("bridge", a as never), args);
    }),
  ],
]);

/** Registry of every bridge op this pane can execute (excel subset). */
export function buildExcelOps(): ReadonlyMap<string, OfficeOpExecutor> {
  return EXCEL_OPS;
}
