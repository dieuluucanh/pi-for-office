/**
 * PowerPoint bridge executors — thin delegates over the local PowerPoint
 * tools, so bridge and browser-only modes share one implementation.
 */

import { createPowerPointGetOverviewTool } from "../tools/powerpoint/get-overview.js";
import { createPowerPointReadSlideTool } from "../tools/powerpoint/read-slide.js";
import { createPowerPointAddSlideTool } from "../tools/powerpoint/add-slide.js";
import { createPowerPointAddTextBoxTool } from "../tools/powerpoint/add-text-box.js";
import { createPowerPointFormatSlideTool } from "../tools/powerpoint/format-slide.js";
import { guardExecutor, type OfficeOpExecutor } from "./ops.js";
import { delegateTool } from "./delegate.js";

export const POWERPOINT_OPS: ReadonlyMap<string, OfficeOpExecutor> = new Map<
 string,
 OfficeOpExecutor
>([
 [
  "powerpoint.get_overview",
  guardExecutor(delegateTool(createPowerPointGetOverviewTool())),
 ],
 [
  "powerpoint.read_slide",
  guardExecutor(delegateTool(createPowerPointReadSlideTool())),
 ],
 [
  "powerpoint.add_slide",
  guardExecutor(delegateTool(createPowerPointAddSlideTool())),
 ],
 [
  "powerpoint.add_text_box",
  guardExecutor(delegateTool(createPowerPointAddTextBoxTool())),
 ],
 [
  "powerpoint.format_slide",
  guardExecutor(delegateTool(createPowerPointFormatSlideTool())),
 ],
]);
