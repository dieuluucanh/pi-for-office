/**
 * Word bridge executors — thin delegates over the local Word tools.
 *
 * Each op maps to the SAME tool factory the browser-only path uses
 * (`src/tools/word/*`), so bridge and local modes share one implementation
 * and one catalog schema by construction.
 */

import { createWordGetOverviewTool } from "../tools/word/get-overview.js";
import { createWordReadDocumentTool } from "../tools/word/read-document.js";
import { createWordInsertTextTool } from "../tools/word/insert-text.js";
import { createWordReplaceTextTool } from "../tools/word/replace-text.js";
import { createWordFormatRangeTool } from "../tools/word/format-range.js";
import { createWordInsertBlocksTool } from "../tools/word/insert-blocks.js";
import { createWordInsertTableTool } from "../tools/word/insert-table.js";
import { createWordInsertPageBreakTool } from "../tools/word/insert-page-break.js";
import { createWordInsertImageTool } from "../tools/word/insert-image.js";
import { createWordInsertHyperlinkTool } from "../tools/word/insert-hyperlink.js";
import { guardExecutor, type OfficeOpExecutor } from "./ops.js";
import { delegateTool } from "./delegate.js";

export const WORD_OPS: ReadonlyMap<string, OfficeOpExecutor> = new Map<
 string,
 OfficeOpExecutor
>([
 [
  "word.get_overview",
  guardExecutor(delegateTool(createWordGetOverviewTool())),
 ],
 [
  "word.read_document",
  guardExecutor(delegateTool(createWordReadDocumentTool())),
 ],
 ["word.insert_text", guardExecutor(delegateTool(createWordInsertTextTool()))],
 [
  "word.replace_text",
  guardExecutor(delegateTool(createWordReplaceTextTool())),
 ],
 [
  "word.format_range",
  guardExecutor(delegateTool(createWordFormatRangeTool())),
 ],
 [
  "word.insert_blocks",
  guardExecutor(delegateTool(createWordInsertBlocksTool())),
 ],
 [
  "word.insert_table",
  guardExecutor(delegateTool(createWordInsertTableTool())),
 ],
 [
  "word.insert_page_break",
  guardExecutor(delegateTool(createWordInsertPageBreakTool())),
 ],
 [
  "word.insert_image",
  guardExecutor(delegateTool(createWordInsertImageTool())),
 ],
 [
  "word.insert_hyperlink",
  guardExecutor(delegateTool(createWordInsertHyperlinkTool())),
 ],
]);
