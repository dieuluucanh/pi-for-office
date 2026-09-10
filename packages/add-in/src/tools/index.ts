/**
 * Tool registry — creates all built-in tools for the agent.
 *
 * Canonical source of truth for core tools lives in `src/tools/registry.ts`.
 * Experimental/non-core tools are appended here.
 */

import type { OfficeApp, SpreadsheetHostKind } from "../host/index.js";
import { createCoreTools, type AnyCoreTool } from "./registry.js";
import { selectOfficeCoupledToolForHost } from "./host-selection.js";
import type { SkillReadCache } from "../skills/read-cache.js";
import { createTmuxTool } from "./tmux.js";
import { createPythonRunTool } from "./python-run.js";
import { createLibreOfficeConvertTool } from "./libreoffice-convert.js";
import { createPythonTransformRangeTool } from "./python-transform-range.js";
import { createFilesTool } from "./files.js";
import { createExecuteOfficeJsTool } from "./execute-office-js.js";
import { createExecuteWpsJsTool } from "./execute-wps-js.js";
import {
  createExtensionsManagerTool,
  type ExtensionsManagerToolRuntime,
} from "./extensions-manager.js";
import { createWordReadDocumentTool } from "./word/read-document.js";
import { createWordInsertTextTool } from "./word/insert-text.js";
import { createWordReplaceTextTool } from "./word/replace-text.js";
import { createWordGetOverviewTool } from "./word/get-overview.js";
import { createWordFormatRangeTool } from "./word/format-range.js";
import { createPowerPointReadSlideTool } from "./powerpoint/read-slide.js";
import { createPowerPointAddSlideTool } from "./powerpoint/add-slide.js";
import { createPowerPointAddTextBoxTool } from "./powerpoint/add-text-box.js";
import { createPowerPointGetOverviewTool } from "./powerpoint/get-overview.js";
import { createPowerPointFormatSlideTool } from "./powerpoint/format-slide.js";

export interface CreateAllToolsOptions {
  hostKind?: SpreadsheetHostKind;
  /** Which Office app hosts the sidebar. Excel (default) gets the full core set. */
  hostApp?: OfficeApp | null;
  getExtensionManager?: () => ExtensionsManagerToolRuntime | null;
  getSessionId?: () => string | null;
  skillReadCache?: SkillReadCache;
}

/** Core tools that work identically on every Office host (Excel/Word/PowerPoint). */
const HOST_AGNOSTIC_CORE_TOOL_NAMES = new Set<string>([
  "instructions",
  "conventions",
  "skills",
]);

function resolvePromptHostApp(
  hostApp: OfficeApp | null | undefined,
): "excel" | "word" | "powerpoint" {
  if (hostApp === "word") return "word";
  if (hostApp === "powerpoint") return "powerpoint";
  return "excel";
}

export function createAllTools(
  options: CreateAllToolsOptions = {},
): AnyCoreTool[] {
  const getExtensionManager = options.getExtensionManager ?? (() => null);
  const hostKind = options.hostKind ?? "office";
  const hostApp = resolvePromptHostApp(options.hostApp);

  const skills = {
    ...(options.getSessionId !== undefined
      ? { getSessionId: options.getSessionId }
      : {}),
    ...(options.skillReadCache !== undefined
      ? { readCache: options.skillReadCache }
      : {}),
  };

  // Excel-only tools are backed by the Excel JavaScript API (Excel.run) and the
  // active workbook grid. On Word/PowerPoint they would fail at runtime; the
  // Pi bridge supplies the host-appropriate document/presentation tools there.
  const isExcelHost = hostApp === "excel";

  return [
    ...createCoreTools({
      hostKind,
      skills,
    }).filter(
      (tool) => isExcelHost || HOST_AGNOSTIC_CORE_TOOL_NAMES.has(tool.name),
    ),
    createTmuxTool(),
    createPythonRunTool(),
    ...(isExcelHost ? [createLibreOfficeConvertTool()] : []),
    ...(isExcelHost
      ? [
          selectOfficeCoupledToolForHost(
            createPythonTransformRangeTool(),
            hostKind,
          ),
        ]
      : []),
    createFilesTool(),
    ...(isExcelHost
      ? [selectOfficeCoupledToolForHost(createExecuteOfficeJsTool(), hostKind)]
      : []),
    ...(hostKind === "wps" ? [createExecuteWpsJsTool()] : []),
    createExtensionsManagerTool({ getManager: getExtensionManager }),
    // Local document/presentation tools — the browser-only (no-bridge) path
    // for Word/PowerPoint hosts, mirroring how Excel gets local workbook tools.
    ...(hostApp === "word"
      ? [
          createWordGetOverviewTool(),
          createWordReadDocumentTool(),
          createWordInsertTextTool(),
          createWordReplaceTextTool(),
          createWordFormatRangeTool(),
        ]
      : []),
    ...(hostApp === "powerpoint"
      ? [
          createPowerPointGetOverviewTool(),
          createPowerPointReadSlideTool(),
          createPowerPointAddSlideTool(),
          createPowerPointAddTextBoxTool(),
          createPowerPointFormatSlideTool(),
        ]
      : []),
  ];
}
