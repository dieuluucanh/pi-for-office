/**
 * conditional_format — Add or clear conditional formatting rules.
 *
 * Supports the eight `Excel.ConditionalFormatType` families (**all ExcelApi
 * 1.6**, so one gate covers every family):
 * - `formula` (custom formula)
 * - `cell_value` (greater/less/equal/between)
 * - `data_bar`, `color_scale`, `icon_set`
 * - `top_bottom`, `preset_criteria` (blanks/average/dates/unique)
 * - `contains_text`
 */

import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { EXCEL_CONDITIONAL_FORMAT_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { dispatchWorkbookSnapshotCreated } from "../workbook/recovery-events.js";
import { captureConditionalFormatState } from "../workbook/recovery-states.js";
import { getWorkbookRecoveryLog } from "../workbook/recovery-log.js";
import { getErrorMessage } from "../utils/errors.js";
import { requireCapability } from "../office/capabilities.js";
import type { ConditionalFormatDetails } from "./tool-details.js";
import {
  CHECKPOINT_SKIPPED_NOTE,
  CHECKPOINT_SKIPPED_REASON,
} from "./recovery-metadata.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";

const schema = EXCEL_CONDITIONAL_FORMAT_PARAMETERS;

type Params = Static<typeof schema>;

/** The capability that supplies every conditional-format family. */
const CONDITIONAL_FORMAT_CAPABILITY = "excel.conditionalFormat.types";

const mutationFinalizeDependencies: MutationFinalizeDependencies = {
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
};

export function createConditionalFormatTool(): AgentTool<
  typeof schema,
  ConditionalFormatDetails
> {
  return {
    name: "conditional_format",
    label: "Conditional Format",
    description:
      "Add or clear conditional formatting rules. Rule families: custom formula, cell value, " +
      "data bar, color scale, icon set, top/bottom, preset criteria (blanks, averages, dates, " +
      "unique/duplicate), and text rules. All families need ExcelApi 1.6.",
    parameters: schema,
    execute: async (
      toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ConditionalFormatDetails>> => {
      try {
        // Every rule family — and `clearAll` — is ExcelApi 1.6. Fail loudly
        // instead of letting the host throw an opaque error mid-batch.
        const gap = requireCapability(CONDITIONAL_FORMAT_CAPABILITY);
        if (gap !== null) {
          return errorResult(toolCallId, params, gap);
        }

        if (params.action === "clear") {
          return await clearFormats(toolCallId, params);
        }

        if (!params.type) {
          return errorResult(
            toolCallId,
            params,
            'type is required when action is "add".',
          );
        }

        return await addFormat(toolCallId, params);
      } catch (e) {
        return errorResult(toolCallId, params, getErrorMessage(e));
      }
    },
  };
}

async function errorResult(
  toolCallId: string,
  params: Params,
  message: string,
): Promise<AgentToolResult<ConditionalFormatDetails>> {
  await finalizeMutationOperation(mutationFinalizeDependencies, {
    auditEntry: {
      toolName: "conditional_format",
      toolCallId,
      blocked: true,
      outputAddress: params.range,
      changedCount: 0,
      changes: [],
      summary: `error: ${message}`,
    },
  });

  return {
    content: [{ type: "text", text: `Error: ${message}` }],
    details: {
      kind: "conditional_format",
      action: params.action,
    },
  };
}

async function clearFormats(
  toolCallId: string,
  params: Params,
): Promise<AgentToolResult<ConditionalFormatDetails>> {
  const checkpointState = await captureConditionalFormatState(params.range);

  const result = await excelRun(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address,rowCount,columnCount");
    const formats = range.conditionalFormats;
    const countResult = formats.getCount();
    await context.sync();

    const existing = countResult.value;
    const cellCount = range.rowCount * range.columnCount;
    formats.clearAll();
    await context.sync();

    return {
      sheetName: sheet.name,
      address: range.address,
      existing,
      cellCount,
    };
  });

  const fullAddr = qualifiedAddress(result.sheetName, result.address);

  const output: AgentToolResult<ConditionalFormatDetails> = {
    content: [
      {
        type: "text",
        text: `Cleared ${result.existing} conditional format(s) from **${fullAddr}**.`,
      },
    ],
    details: {
      kind: "conditional_format",
      action: "clear",
      address: fullAddr,
    },
  };

  const recoveryUnavailableReason = checkpointState.supported
    ? CHECKPOINT_SKIPPED_REASON
    : (checkpointState.reason ?? CHECKPOINT_SKIPPED_REASON);

  await finalizeMutationOperation(mutationFinalizeDependencies, {
    auditEntry: {
      toolName: "conditional_format",
      toolCallId,
      blocked: false,
      outputAddress: fullAddr,
      changedCount: result.cellCount,
      changes: [],
      summary: `cleared ${result.existing} rule(s) across ${result.cellCount} cell(s)`,
    },
    recovery: {
      result: output,
      appendRecoverySnapshot: () => {
        if (!checkpointState.supported) {
          return Promise.resolve(null);
        }

        return getWorkbookRecoveryLog().appendConditionalFormat({
          toolName: "conditional_format",
          toolCallId,
          address: fullAddr,
          changedCount: result.cellCount,
          cellCount: result.cellCount,
          conditionalFormatRules: checkpointState.rules,
        });
      },
      appendResultNote: appendMutationResultNote,
      unavailableReason: recoveryUnavailableReason,
      unavailableNote: CHECKPOINT_SKIPPED_NOTE,
      dispatchSnapshotCreated: (checkpoint) => {
        dispatchWorkbookSnapshotCreated({
          snapshotId: checkpoint.id,
          toolName: checkpoint.toolName,
          address: checkpoint.address,
          changedCount: checkpoint.changedCount,
        });
      },
    },
  });

  return output;
}

async function addFormat(
  toolCallId: string,
  params: Params,
): Promise<AgentToolResult<ConditionalFormatDetails>> {
  validateAddParams(params);

  const checkpointState = await captureConditionalFormatState(params.range);

  const result = await excelRun(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address,rowCount,columnCount");
    await context.sync();

    const cf = applyRule(range, params);

    if (params.stop_if_true !== undefined) {
      cf.stopIfTrue = params.stop_if_true;
    }

    await context.sync();

    return {
      sheetName: sheet.name,
      address: range.address,
      cellCount: range.rowCount * range.columnCount,
    };
  });

  const fullAddr = qualifiedAddress(result.sheetName, result.address);
  const details = describeRule(params);

  const output: AgentToolResult<ConditionalFormatDetails> = {
    content: [
      {
        type: "text",
        text: `Added conditional format to **${fullAddr}** — ${details}.`,
      },
    ],
    details: {
      kind: "conditional_format",
      action: "add",
      address: fullAddr,
    },
  };

  const recoveryUnavailableReason = checkpointState.supported
    ? CHECKPOINT_SKIPPED_REASON
    : (checkpointState.reason ?? CHECKPOINT_SKIPPED_REASON);

  await finalizeMutationOperation(mutationFinalizeDependencies, {
    auditEntry: {
      toolName: "conditional_format",
      toolCallId,
      blocked: false,
      outputAddress: fullAddr,
      changedCount: result.cellCount,
      changes: [],
      summary: `added ${params.type ?? "rule"} across ${result.cellCount} cell(s)`,
    },
    recovery: {
      result: output,
      appendRecoverySnapshot: () => {
        if (!checkpointState.supported) {
          return Promise.resolve(null);
        }

        return getWorkbookRecoveryLog().appendConditionalFormat({
          toolName: "conditional_format",
          toolCallId,
          address: fullAddr,
          changedCount: result.cellCount,
          cellCount: result.cellCount,
          conditionalFormatRules: checkpointState.rules,
        });
      },
      appendResultNote: appendMutationResultNote,
      unavailableReason: recoveryUnavailableReason,
      unavailableNote: CHECKPOINT_SKIPPED_NOTE,
      dispatchSnapshotCreated: (checkpoint) => {
        dispatchWorkbookSnapshotCreated({
          snapshotId: checkpoint.id,
          toolName: checkpoint.toolName,
          address: checkpoint.address,
          changedCount: checkpoint.changedCount,
        });
      },
    },
  });

  return output;
}

/**
 * Add the rule for the requested family and return the created conditional
 * format so the caller can still apply `stop_if_true` to it.
 */
function applyRule(
  range: Excel.Range,
  params: Params,
): Excel.ConditionalFormat {
  const type = params.type;
  switch (type) {
    case "formula": {
      const cf = range.conditionalFormats.add(
        Excel.ConditionalFormatType.custom,
      );
      cf.custom.rule.formula = params.formula ?? "";
      applyFormat(cf.custom.format, params);
      return cf;
    }
    case "cell_value": {
      if (params.operator === undefined) {
        throw new Error('operator is required for type="cell_value".');
      }
      const cf = range.conditionalFormats.add(
        Excel.ConditionalFormatType.cellValue,
      );
      const rule: Excel.ConditionalCellValueRule = {
        formula1: stringifyValue(params.value),
        operator: params.operator,
      };
      if (params.value2 !== undefined) {
        rule.formula2 = stringifyValue(params.value2);
      }
      cf.cellValue.rule = rule;
      applyFormat(cf.cellValue.format, params);
      return cf;
    }
    case "data_bar": {
      const cf = range.conditionalFormats.add(
        Excel.ConditionalFormatType.dataBar,
      );
      const dataBar = cf.dataBar;
      dataBar.barDirection = params.data_bar_direction ?? "Context";
      if (params.data_bar_fill_color !== undefined) {
        dataBar.positiveFormat.fillColor = params.data_bar_fill_color;
      }
      if (params.data_bar_negative_fill_color !== undefined) {
        dataBar.negativeFormat.fillColor = params.data_bar_negative_fill_color;
      }
      if (params.data_bar_show_only !== undefined) {
        dataBar.showDataBarOnly = params.data_bar_show_only;
      }
      if (params.data_bar_axis_color !== undefined) {
        dataBar.axisColor = params.data_bar_axis_color;
      }
      return cf;
    }
    case "color_scale": {
      const cf = range.conditionalFormats.add(
        Excel.ConditionalFormatType.colorScale,
      );
      const criteria: Excel.ConditionalColorScaleCriteria = {
        minimum: {
          type: "LowestValue",
          color: params.color_scale_min_color ?? "#F8696B",
        },
        maximum: {
          type: "HighestValue",
          color: params.color_scale_max_color ?? "#63BE7B",
        },
      };
      if (params.color_scale_mid_color !== undefined) {
        criteria.midpoint = {
          type: "Percentile",
          formula: "50",
          color: params.color_scale_mid_color,
        };
      }
      cf.colorScale.criteria = criteria;
      return cf;
    }
    case "icon_set": {
      const cf = range.conditionalFormats.add(
        Excel.ConditionalFormatType.iconSet,
      );
      cf.iconSet.style = params.icon_set_style ?? "ThreeTrafficLights1";
      if (params.icon_set_reverse_order !== undefined) {
        cf.iconSet.reverseIconOrder = params.icon_set_reverse_order;
      }
      if (params.icon_set_show_icon_only !== undefined) {
        cf.iconSet.showIconOnly = params.icon_set_show_icon_only;
      }
      return cf;
    }
    case "top_bottom": {
      const cf = range.conditionalFormats.add(
        Excel.ConditionalFormatType.topBottom,
      );
      cf.topBottom.rule = {
        type: params.top_bottom_type ?? "TopItems",
        rank: params.top_bottom_rank ?? 10,
      };
      applyFormat(cf.topBottom.format, params);
      return cf;
    }
    case "preset_criteria": {
      const cf = range.conditionalFormats.add(
        Excel.ConditionalFormatType.presetCriteria,
      );
      cf.preset.rule = {
        criterion: params.preset_criterion ?? "AboveAverage",
      };
      applyFormat(cf.preset.format, params);
      return cf;
    }
    case "contains_text": {
      const cf = range.conditionalFormats.add(
        Excel.ConditionalFormatType.containsText,
      );
      cf.textComparison.rule = {
        operator: params.text_operator ?? "Contains",
        text: params.text ?? "",
      };
      applyFormat(cf.textComparison.format, params);
      return cf;
    }
    default: {
      // Exhaustive: the schema union is closed, so this is unreachable.
      throw new Error(`Unsupported rule type "${String(type)}".`);
    }
  }
}

function validateAddParams(params: Params): void {
  switch (params.type) {
    case "formula":
      if (!params.formula) {
        throw new Error('formula is required for type="formula".');
      }
      return;
    case "cell_value":
      if (!params.operator) {
        throw new Error('operator is required for type="cell_value".');
      }
      if (
        params.value === undefined ||
        params.value === null ||
        params.value === ""
      ) {
        throw new Error('value is required for type="cell_value".');
      }
      if (
        (params.operator === "Between" || params.operator === "NotBetween") &&
        (params.value2 === undefined ||
          params.value2 === null ||
          params.value2 === "")
      ) {
        throw new Error("value2 is required for Between/NotBetween operators.");
      }
      return;
    case "contains_text":
      if (!params.text) {
        throw new Error('text is required for type="contains_text".');
      }
      return;
    case "top_bottom":
      if (
        params.top_bottom_rank !== undefined &&
        (!Number.isFinite(params.top_bottom_rank) || params.top_bottom_rank < 1)
      ) {
        throw new Error("top_bottom_rank must be a number >= 1.");
      }
      return;
    // `data_bar`, `color_scale`, `icon_set` and `preset_criteria` all have
    // host-side defaults, so they need no extra parameters.
    default:
      return;
  }
}

function stringifyValue(value: string | number | undefined): string {
  if (value === undefined || value === null) return "";
  return typeof value === "number" ? value.toString() : value;
}

/** Describe the applied family for the tool result. */
function describeRule(params: Params): string {
  switch (params.type) {
    case "formula":
      return `formula rule (${params.formula ?? ""})`;
    case "cell_value":
      return `cell value rule (${params.operator ?? ""} ${params.value ?? ""}${
        params.value2 !== undefined ? ` and ${String(params.value2)}` : ""
      })`;
    case "data_bar":
      return `data bar (${params.data_bar_direction ?? "Context"}${describeFormat(params)})`;
    case "color_scale":
      return `color scale (${params.color_scale_min_color ?? "#F8696B"} → ${
        params.color_scale_mid_color !== undefined
          ? `${params.color_scale_mid_color} → `
          : ""
      }${params.color_scale_max_color ?? "#63BE7B"})`;
    case "icon_set":
      return `icon set (${params.icon_set_style ?? "ThreeTrafficLights1"})`;
    case "top_bottom":
      return `top/bottom rule (${params.top_bottom_type ?? "TopItems"} ${
        params.top_bottom_rank ?? 10
      }${describeFormat(params)})`;
    case "preset_criteria":
      return `preset rule (${params.preset_criterion ?? "AboveAverage"}${describeFormat(params)})`;
    case "contains_text":
      return `text rule (${params.text_operator ?? "Contains"} "${
        params.text ?? ""
      }"${describeFormat(params)})`;
    default:
      return "rule";
  }
}

function describeFormat(params: Params): string {
  const parts: string[] = [];
  if (params.fill_color) parts.push(`fill ${params.fill_color}`);
  if (params.font_color) parts.push(`font ${params.font_color}`);
  if (params.bold !== undefined) parts.push(params.bold ? "bold" : "not bold");
  if (params.italic !== undefined)
    parts.push(params.italic ? "italic" : "not italic");
  if (params.underline !== undefined)
    parts.push(params.underline ? "underline" : "no underline");
  return parts.length === 0 ? "" : `, ${parts.join(", ")}`;
}

function applyFormat(
  format: Excel.ConditionalRangeFormat,
  params: Params,
): void {
  if (params.fill_color) {
    format.fill.color = params.fill_color;
  }
  if (params.font_color) {
    format.font.color = params.font_color;
  }
  if (params.bold !== undefined) {
    format.font.bold = params.bold;
  }
  if (params.italic !== undefined) {
    format.font.italic = params.italic;
  }
  if (params.underline !== undefined) {
    format.font.underline = params.underline ? "Single" : "None";
  }
}
