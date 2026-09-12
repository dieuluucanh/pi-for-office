/**
 * Excel bridge executors — thin delegates over the local Excel core tools.
 *
 * These delegate to the existing, battle-tested pi-for-office tools so the
 * bridge path behaves exactly like the in-browser agent path. The set matches
 * the shared catalog (`@dieulc/pi-office-protocol`): the 14 document-facing
 * core ops (instructions/conventions/skills stay local-only).
 */

import { createGetWorkbookOverviewTool } from "../tools/get-workbook-overview.js";
import { createReadRangeTool } from "../tools/read-range.js";
import { createWriteCellsTool } from "../tools/write-cells.js";
import { createFillFormulaTool } from "../tools/fill-formula.js";
import { createSearchWorkbookTool } from "../tools/search-workbook.js";
import { createModifyStructureTool } from "../tools/modify-structure.js";
import { createFormatCellsTool } from "../tools/format-cells.js";
import { createConditionalFormatTool } from "../tools/conditional-format.js";
import { createChartsTool } from "../tools/charts.js";
import { createTraceDependenciesTool } from "../tools/trace-dependencies.js";
import { createExplainFormulaTool } from "../tools/explain-formula.js";
import { createViewSettingsTool } from "../tools/view-settings.js";
import { createCommentsTool } from "../tools/comments.js";
import { createWorkbookHistoryTool } from "../tools/workbook-history.js";
import { guardExecutor, type OfficeOpExecutor } from "./ops.js";
import { delegateTool } from "./delegate.js";

export const EXCEL_OPS: ReadonlyMap<string, OfficeOpExecutor> = new Map<
 string,
 OfficeOpExecutor
>([
 [
  "excel.get_overview",
  guardExecutor(delegateTool(createGetWorkbookOverviewTool())),
 ],
 ["excel.read_range", guardExecutor(delegateTool(createReadRangeTool()))],
 ["excel.write_cells", guardExecutor(delegateTool(createWriteCellsTool()))],
 ["excel.fill_formula", guardExecutor(delegateTool(createFillFormulaTool()))],
 [
  "excel.search_workbook",
  guardExecutor(delegateTool(createSearchWorkbookTool())),
 ],
 [
  "excel.modify_structure",
  guardExecutor(delegateTool(createModifyStructureTool())),
 ],
 ["excel.format_cells", guardExecutor(delegateTool(createFormatCellsTool()))],
 [
  "excel.conditional_format",
  guardExecutor(delegateTool(createConditionalFormatTool())),
 ],
 ["excel.charts", guardExecutor(delegateTool(createChartsTool()))],
 [
  "excel.trace_dependencies",
  guardExecutor(delegateTool(createTraceDependenciesTool())),
 ],
 [
  "excel.explain_formula",
  guardExecutor(delegateTool(createExplainFormulaTool())),
 ],
 ["excel.view_settings", guardExecutor(delegateTool(createViewSettingsTool()))],
 ["excel.comments", guardExecutor(delegateTool(createCommentsTool()))],
 [
  "excel.workbook_history",
  guardExecutor(delegateTool(createWorkbookHistoryTool())),
 ],
]);

/** Registry of every bridge op this pane can execute (excel subset). */
export function buildExcelOps(): ReadonlyMap<string, OfficeOpExecutor> {
 return EXCEL_OPS;
}
