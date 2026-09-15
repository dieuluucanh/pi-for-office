/**
 * charts — list, create, update, delete, and capture Excel charts.
 */

import { EXCEL_CHARTS_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { excelRun, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import {
  getWorkbookChangeAuditLog,
  type AppendWorkbookChangeAuditEntryArgs,
} from "../audit/workbook-change-audit.js";
import { dispatchWorkbookSnapshotCreated } from "../workbook/recovery-events.js";
import {
  getWorkbookRecoveryLog,
  type AppendChartRecoverySnapshotArgs,
  type WorkbookRecoverySnapshot,
} from "../workbook/recovery-log.js";
import {
  captureChartPresentState,
  type RecoveryChartAbsentState,
  type RecoveryChartPresentState,
  type RecoveryChartState,
} from "../workbook/recovery-states.js";
import { formatRecoveryChartAddress } from "../workbook/recovery/chart-state.js";
import { getErrorMessage } from "../utils/errors.js";
import { requireCapability } from "../office/capabilities.js";
import type { ChartListItemDetails, ChartsDetails } from "./tool-details.js";
import { recoveryCheckpointUnavailable } from "./recovery-metadata.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import type {
  MutationFinalizeDependencies,
  MutationRecoveryStep,
} from "./mutation/types.js";

const FRIENDLY_CHART_TYPES = [
  "column",
  "column_stacked",
  "column_stacked_100",
  "bar",
  "bar_stacked",
  "bar_stacked_100",
  "line",
  "line_markers",
  "area",
  "area_stacked",
  "pie",
  "doughnut",
  "scatter",
  "scatter_lines",
  "scatter_smooth",
  "radar",
] as const;

type FriendlyChartType = (typeof FRIENDLY_CHART_TYPES)[number];
type SeriesByParam = "auto" | "columns" | "rows";
type LegendPositionParam = "none" | "right" | "left" | "top" | "bottom";

type SupportedExcelChartType =
  | "ColumnClustered"
  | "ColumnStacked"
  | "ColumnStacked100"
  | "BarClustered"
  | "BarStacked"
  | "BarStacked100"
  | "Line"
  | "LineMarkers"
  | "Area"
  | "AreaStacked"
  | "Pie"
  | "Doughnut"
  | "XYScatter"
  | "XYScatterLines"
  | "XYScatterSmooth"
  | "Radar";

type SupportedExcelSeriesBy = "Auto" | "Columns" | "Rows";
type SupportedExcelLegendPosition = "Right" | "Left" | "Top" | "Bottom";

const CHART_TYPE_MAP: Record<FriendlyChartType, SupportedExcelChartType> = {
  column: "ColumnClustered",
  column_stacked: "ColumnStacked",
  column_stacked_100: "ColumnStacked100",
  bar: "BarClustered",
  bar_stacked: "BarStacked",
  bar_stacked_100: "BarStacked100",
  line: "Line",
  line_markers: "LineMarkers",
  area: "Area",
  area_stacked: "AreaStacked",
  pie: "Pie",
  doughnut: "Doughnut",
  scatter: "XYScatter",
  scatter_lines: "XYScatterLines",
  scatter_smooth: "XYScatterSmooth",
  radar: "Radar",
};

const SERIES_BY_MAP: Record<SeriesByParam, SupportedExcelSeriesBy> = {
  auto: "Auto",
  columns: "Columns",
  rows: "Rows",
};

const LEGEND_POSITION_MAP: Record<
  Exclude<LegendPositionParam, "none">,
  SupportedExcelLegendPosition
> = {
  right: "Right",
  left: "Left",
  top: "Top",
  bottom: "Bottom",
};

const DEFAULT_IMAGE_WIDTH = 600;
const MAX_IMAGE_WIDTH = 1_200;
const FALLBACK_IMAGE_ASPECT_RATIO = 0.6;

const CHART_DELETE_NO_BACKUP_REASON =
  "Chart deletes are not backed up in v1 because Office.js does not expose an arbitrary chart's source range for faithful recreation.";
const CHART_DELETE_NO_BACKUP_NOTE =
  "ℹ️ Backup not created. Excel does not expose the source range for an arbitrary existing chart, so `workbook_history` cannot faithfully recreate deleted charts in v1.";
const CHART_SOURCE_RANGE_RESTORE_NOTE =
  "ℹ️ Chart property backup only. Excel does not expose the previous source range, so `workbook_history` cannot revert this update's data-source change.";

const schema = EXCEL_CHARTS_PARAMETERS;

type Params = Static<typeof schema>;
type ChartAction = Params["action"];

interface ChartTarget {
  sheet: Excel.Worksheet;
  chart: Excel.Chart;
}

interface ResolvedRange {
  sheet: Excel.Worksheet;
  range: Excel.Range;
}

interface PositionAnchors {
  start: string;
  end?: string;
}

interface ChartsExecutionResult {
  result: AgentToolResult<ChartsDetails>;
  outputAddress?: string;
  changedCount?: number;
  auditSummary?: string;
  recoveryState?: RecoveryChartState;
  sourceRangeChanged?: boolean;
}

interface ChartsToolDependencies {
  executeAction: (params: Params) => Promise<ChartsExecutionResult>;
  captureChartPresent: (
    name: string,
    sheetName?: string,
  ) => Promise<RecoveryChartPresentState>;
  appendRecoverySnapshot: (
    args: AppendChartRecoverySnapshotArgs,
  ) => Promise<WorkbookRecoverySnapshot | null>;
  appendAuditEntry: (
    entry: AppendWorkbookChangeAuditEntryArgs,
  ) => Promise<void>;
  dispatchSnapshotCreated: (snapshot: WorkbookRecoverySnapshot) => void;
}

export function toExcelChartType(chartType: string): SupportedExcelChartType {
  if (isFriendlyChartType(chartType)) {
    return CHART_TYPE_MAP[chartType];
  }

  throw new Error(
    `Invalid chart_type "${chartType}". Use one of: ${FRIENDLY_CHART_TYPES.join(", ")}.`,
  );
}

function isFriendlyChartType(value: string): value is FriendlyChartType {
  return FRIENDLY_CHART_TYPES.some((candidate) => candidate === value);
}

function toExcelSeriesBy(
  seriesBy: SeriesByParam | undefined,
): SupportedExcelSeriesBy | undefined {
  return seriesBy ? SERIES_BY_MAP[seriesBy] : undefined;
}

function normalizeImageWidth(width: number | undefined): number {
  if (typeof width !== "number" || !Number.isFinite(width))
    return DEFAULT_IMAGE_WIDTH;
  const rounded = Math.round(width);
  if (rounded <= 0) return DEFAULT_IMAGE_WIDTH;
  return Math.min(rounded, MAX_IMAGE_WIDTH);
}

function normalizeImageHeight(
  width: number,
  chartWidth: number,
  chartHeight: number,
): number {
  if (chartWidth > 0 && chartHeight > 0) {
    return Math.max(1, Math.round(width * (chartHeight / chartWidth)));
  }

  return Math.max(1, Math.round(width * FALLBACK_IMAGE_ASPECT_RATIO));
}

function chartDetailsAddress(sheetName: string, chartName: string): string {
  return formatRecoveryChartAddress(sheetName, chartName);
}

function isMutatingChartsAction(action: ChartAction): boolean {
  return action === "create" || action === "update" || action === "delete";
}

function requireName(params: Params): string {
  const name = params.name?.trim();
  if (!name) {
    throw new Error(`name is required for ${params.action}`);
  }
  return name;
}

function requireSourceRange(params: Params): string {
  const sourceRange = params.source_range?.trim();
  if (!sourceRange) {
    throw new Error(`source_range is required for ${params.action}`);
  }
  return sourceRange;
}

function requireChartType(params: Params): FriendlyChartType {
  const chartType = params.chart_type;
  if (!chartType) {
    throw new Error("chart_type is required for create");
  }

  if (!isFriendlyChartType(chartType)) {
    throw new Error(`Invalid chart_type "${String(chartType)}".`);
  }

  return chartType;
}

function validateParams(params: Params): void {
  if (params.action === "create") {
    requireSourceRange(params);
    requireChartType(params);
  }

  if (
    params.action === "update" ||
    params.action === "delete" ||
    params.action === "get_image"
  ) {
    requireName(params);
  }

  if (params.chart_type) {
    toExcelChartType(params.chart_type);
  }
}

function resolveRange(
  context: Excel.RequestContext,
  ref: string,
  defaultSheet: Excel.Worksheet | undefined,
): ResolvedRange {
  const parsed = parseRangeRef(ref);
  const sheet = parsed.sheet
    ? context.workbook.worksheets.getItem(parsed.sheet)
    : (defaultSheet ?? context.workbook.worksheets.getActiveWorksheet());

  return {
    sheet,
    range: sheet.getRange(parsed.address),
  };
}

function resolveCreateChartSheet(
  context: Excel.RequestContext,
  params: Params,
): Excel.Worksheet {
  if (params.sheet) {
    return context.workbook.worksheets.getItem(params.sheet);
  }

  const sourceRange = requireSourceRange(params);
  const parsedSource = parseRangeRef(sourceRange);
  if (parsedSource.sheet) {
    return context.workbook.worksheets.getItem(parsedSource.sheet);
  }

  return context.workbook.worksheets.getActiveWorksheet();
}

function parsePositionAnchors(position: string): PositionAnchors {
  const parsed = parseRangeRef(position);
  const parts = parsed.address
    .split(":")
    .map((part) => part.trim())
    .filter((part) => part.length > 0);

  if (parts.length === 0 || parts.length > 2) {
    throw new Error(
      `Invalid chart position "${position}". Use a cell or two-cell range like D2:J18.`,
    );
  }

  const qualify = (address: string): string =>
    parsed.sheet ? qualifiedAddress(parsed.sheet, address) : address;
  const start = parts[0];
  if (!start) {
    throw new Error(`Invalid chart position "${position}".`);
  }

  const second = parts[1];
  return {
    start: qualify(start),
    ...(second ? { end: qualify(second) } : {}),
  };
}

function setTitle(
  title: Excel.ChartTitle | Excel.ChartAxisTitle,
  text: string,
): void {
  title.text = text;
  title.visible = text.trim().length > 0;
}

function applyChartProperties(chart: Excel.Chart, params: Params): void {
  if (params.chart_type) {
    chart.chartType = toExcelChartType(params.chart_type);
  }

  if (params.title !== undefined) {
    setTitle(chart.title, params.title);
  }

  if (params.legend_position !== undefined) {
    if (params.legend_position === "none") {
      chart.legend.visible = false;
    } else {
      chart.legend.position = LEGEND_POSITION_MAP[params.legend_position];
      chart.legend.visible = true;
    }
  }

  if (params.x_axis_title !== undefined) {
    setTitle(chart.axes.categoryAxis.title, params.x_axis_title);
  }

  if (params.y_axis_title !== undefined) {
    setTitle(chart.axes.valueAxis.title, params.y_axis_title);
  }

  if (params.position !== undefined) {
    const anchors = parsePositionAnchors(params.position);
    chart.setPosition(anchors.start, anchors.end);
  }
}

async function findChartsByName(
  context: Excel.RequestContext,
  name: string,
  sheetName: string | undefined,
): Promise<ChartTarget[]> {
  if (sheetName) {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load("name");
    sheet.charts.load("items/name");
    await context.sync();

    return sheet.charts.items
      .filter((chart) => chart.name === name)
      .map((chart) => ({ sheet, chart }));
  }

  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  for (const sheet of sheets.items) {
    sheet.charts.load("items/name");
  }
  await context.sync();

  const matches: ChartTarget[] = [];
  for (const sheet of sheets.items) {
    for (const chart of sheet.charts.items) {
      if (chart.name === name) {
        matches.push({ sheet, chart });
      }
    }
  }

  return matches;
}

async function findChartByName(
  context: Excel.RequestContext,
  name: string,
  sheetName: string | undefined,
): Promise<ChartTarget> {
  const matches = await findChartsByName(context, name, sheetName);
  if (matches.length === 0) {
    const location = sheetName ? ` on sheet "${sheetName}"` : "";
    throw new Error(`Chart "${name}" was not found${location}.`);
  }

  if (matches.length > 1) {
    throw new Error(
      `Multiple charts named "${name}" were found. Provide sheet to disambiguate.`,
    );
  }

  const match = matches[0];
  if (!match) {
    throw new Error(`Chart "${name}" was not found.`);
  }

  return match;
}

function chartPositionDetails(
  chart: Excel.Chart,
): ChartListItemDetails["position"] {
  return {
    top: typeof chart.top === "number" ? chart.top : 0,
    left: typeof chart.left === "number" ? chart.left : 0,
    width: typeof chart.width === "number" ? chart.width : 0,
    height: typeof chart.height === "number" ? chart.height : 0,
  };
}

function chartListItem(
  sheet: Excel.Worksheet,
  chart: Excel.Chart,
): ChartListItemDetails {
  return {
    name: chart.name,
    chartType: String(chart.chartType),
    title: chart.title.visible ? chart.title.text : "",
    worksheet: sheet.name,
    position: chartPositionDetails(chart),
  };
}

function formatRect(position: ChartListItemDetails["position"]): string {
  return `top ${position.top}, left ${position.left}, width ${position.width}, height ${position.height}`;
}

function formatChartList(
  items: ChartListItemDetails[],
  sheetName: string | undefined,
): string {
  if (items.length === 0) {
    return sheetName
      ? `No charts found on "${sheetName}".`
      : "No charts found in this workbook.";
  }

  const lines: string[] = [
    `Charts (${items.length})`,
    "",
    "| Worksheet | Name | Type | Title | Position |",
    "| --- | --- | --- | --- | --- |",
  ];

  for (const item of items) {
    const title = item.title || "(none)";
    lines.push(
      `| ${item.worksheet} | ${item.name} | ${item.chartType} | ${title} | ${formatRect(item.position)} |`,
    );
  }

  return lines.join("\n");
}

async function executeList(params: Params): Promise<ChartsExecutionResult> {
  return excelRun(async (context) => {
    const sheets: Excel.Worksheet[] = [];

    if (params.sheet) {
      const sheet = context.workbook.worksheets.getItem(params.sheet);
      sheet.load("name");
      sheets.push(sheet);
      await context.sync();
    } else {
      const worksheetCollection = context.workbook.worksheets;
      worksheetCollection.load("items/name");
      await context.sync();
      sheets.push(...worksheetCollection.items);
    }

    for (const sheet of sheets) {
      sheet.charts.load(
        "items/name,items/chartType,items/top,items/left,items/width,items/height",
      );
    }
    await context.sync();

    for (const sheet of sheets) {
      for (const chart of sheet.charts.items) {
        chart.title.load("text,visible");
      }
    }
    await context.sync();

    const charts = sheets.flatMap((sheet) =>
      sheet.charts.items.map((chart) => chartListItem(sheet, chart)),
    );

    return {
      result: {
        content: [
          { type: "text", text: formatChartList(charts, params.sheet) },
        ],
        details: {
          kind: "charts",
          action: "list",
          count: charts.length,
          charts,
        },
      },
    };
  });
}

async function executeCreate(params: Params): Promise<ChartsExecutionResult> {
  return excelRun(async (context) => {
    const chartSheet = resolveCreateChartSheet(context, params);
    const source = resolveRange(
      context,
      requireSourceRange(params),
      chartSheet,
    );
    const chartType = toExcelChartType(requireChartType(params));
    const seriesBy = toExcelSeriesBy(params.series_by);

    chartSheet.load("name");
    source.sheet.load("name");
    source.range.load("address");

    const chart = chartSheet.charts.add(chartType, source.range, seriesBy);
    applyChartProperties(chart, params);

    if (params.name !== undefined && params.name.trim().length > 0) {
      chart.name = params.name.trim();
    }

    chart.load("name,chartType,top,left,width,height");
    await context.sync();

    // Best-effort: capture the stable chart id for rename-proof recovery.
    // Hosts without Chart.id support must still be able to create charts.
    let chartId: string | undefined;
    try {
      chart.load("id");
      await context.sync();
      chartId =
        typeof chart.id === "string" && chart.id.length > 0
          ? chart.id
          : undefined;
    } catch {
      chartId = undefined;
    }

    const sourceAddress = qualifiedAddress(
      source.sheet.name,
      source.range.address,
    );
    const address = chartDetailsAddress(chartSheet.name, chart.name);
    const recoveryState: RecoveryChartAbsentState = {
      kind: "chart_absent",
      sheetName: chartSheet.name,
      name: chart.name,
      ...(chartId !== undefined ? { chartId } : {}),
    };

    return {
      result: {
        content: [
          {
            type: "text",
            text: `Chart '${chart.name}' created — ${sourceAddress}`,
          },
        ],
        details: {
          kind: "charts",
          action: "create",
          name: chart.name,
          address,
          sourceRange: sourceAddress,
        },
      },
      outputAddress: address,
      changedCount: 1,
      auditSummary: `created chart ${chart.name} from ${sourceAddress}`,
      recoveryState,
    };
  });
}

async function executeUpdate(params: Params): Promise<ChartsExecutionResult> {
  return excelRun(async (context) => {
    const target = await findChartByName(
      context,
      requireName(params),
      params.sheet,
    );
    const chart = target.chart;
    let sourceForAddress: ResolvedRange | null = null;

    if (params.source_range) {
      const source = resolveRange(context, params.source_range, target.sheet);
      source.sheet.load("name");
      source.range.load("address");
      chart.setData(source.range, toExcelSeriesBy(params.series_by));
      sourceForAddress = source;
    }

    applyChartProperties(chart, params);

    if (params.new_name !== undefined && params.new_name.trim().length > 0) {
      chart.name = params.new_name.trim();
    }

    target.sheet.load("name");
    chart.load("name,chartType,top,left,width,height");
    await context.sync();

    const address = chartDetailsAddress(target.sheet.name, chart.name);
    const sourceAddress = sourceForAddress
      ? qualifiedAddress(
          sourceForAddress.sheet.name,
          sourceForAddress.range.address,
        )
      : undefined;
    const sourceSuffix = sourceAddress ? ` — source ${sourceAddress}` : "";

    return {
      result: {
        content: [
          {
            type: "text",
            text: `Chart '${chart.name}' updated${sourceSuffix}.`,
          },
        ],
        details: {
          kind: "charts",
          action: "update",
          name: chart.name,
          address,
          ...(sourceAddress !== undefined
            ? { sourceRange: sourceAddress }
            : {}),
        },
      },
      outputAddress: address,
      changedCount: 1,
      auditSummary: `updated chart ${chart.name}${sourceAddress ? ` source ${sourceAddress}` : ""}`,
      sourceRangeChanged: Boolean(sourceAddress),
    };
  });
}

async function executeDelete(params: Params): Promise<ChartsExecutionResult> {
  return excelRun(async (context) => {
    const target = await findChartByName(
      context,
      requireName(params),
      params.sheet,
    );
    target.sheet.load("name");
    target.chart.load("name");
    await context.sync();

    const chartName = target.chart.name;
    const address = chartDetailsAddress(target.sheet.name, chartName);
    target.chart.delete();
    await context.sync();

    return {
      result: {
        content: [{ type: "text", text: `Chart '${chartName}' deleted.` }],
        details: {
          kind: "charts",
          action: "delete",
          name: chartName,
          address,
          recovery: recoveryCheckpointUnavailable(
            CHART_DELETE_NO_BACKUP_REASON,
          ),
        },
      },
      outputAddress: address,
      changedCount: 1,
      auditSummary: `deleted chart ${chartName}`,
    };
  });
}

async function executeGetImage(params: Params): Promise<ChartsExecutionResult> {
  // `Chart.getImage` is ExcelApi 1.7 — fail loud on an older host.
  const gap = requireCapability("excel.range.image");
  if (gap !== null) throw new Error(gap);

  return excelRun(async (context) => {
    const target = await findChartByName(
      context,
      requireName(params),
      params.sheet,
    );
    target.sheet.load("name");
    target.chart.load("name,width,height");
    await context.sync();

    const width = normalizeImageWidth(params.width);
    const chartWidth =
      typeof target.chart.width === "number" ? target.chart.width : 0;
    const chartHeight =
      typeof target.chart.height === "number" ? target.chart.height : 0;
    const height = normalizeImageHeight(width, chartWidth, chartHeight);
    const image = target.chart.getImage(width, height, "Fit");
    await context.sync();

    const address = chartDetailsAddress(target.sheet.name, target.chart.name);

    return {
      result: {
        content: [
          {
            type: "text",
            text: `Captured chart '${target.chart.name}' (${width}×${height} px).`,
          },
          { type: "image", data: image.value, mimeType: "image/png" },
        ],
        details: {
          kind: "charts",
          action: "get_image",
          name: target.chart.name,
          address,
          image: {
            base64: image.value,
            mimeType: "image/png",
            width,
            height,
          },
        },
      },
    };
  });
}

export async function executeChartsAction(
  params: Params,
): Promise<ChartsExecutionResult> {
  switch (params.action) {
    case "list":
      return executeList(params);
    case "create":
      return executeCreate(params);
    case "update":
      return executeUpdate(params);
    case "delete":
      return executeDelete(params);
    case "get_image":
      return executeGetImage(params);
  }
}

const defaultDependencies: ChartsToolDependencies = {
  executeAction: executeChartsAction,
  captureChartPresent: (name, sheetName) =>
    captureChartPresentState(name, sheetName),
  appendRecoverySnapshot: (args) => getWorkbookRecoveryLog().appendChart(args),
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
  dispatchSnapshotCreated: (snapshot) => {
    dispatchWorkbookSnapshotCreated({
      snapshotId: snapshot.id,
      toolName: snapshot.toolName,
      address: snapshot.address,
      changedCount: snapshot.changedCount,
    });
  },
};

function buildRecoveryStep(
  dependencies: ChartsToolDependencies,
  toolCallId: string,
  output: AgentToolResult<ChartsDetails>,
  result: ChartsExecutionResult,
  captureError: string | null,
): MutationRecoveryStep<ChartsDetails> | undefined {
  const action = output.details.action;
  if (action === "delete") {
    return {
      result: output,
      appendRecoverySnapshot: () => Promise.resolve(null),
      appendResultNote: appendMutationResultNote,
      unavailableReason: CHART_DELETE_NO_BACKUP_REASON,
      unavailableNote: CHART_DELETE_NO_BACKUP_NOTE,
    };
  }

  if (!result.recoveryState) {
    if (action === "update") {
      return {
        result: output,
        appendRecoverySnapshot: () => Promise.resolve(null),
        appendResultNote: appendMutationResultNote,
        unavailableReason: captureError ?? "Chart backup capture was skipped.",
        unavailableNote: backupUnavailableNote(captureError),
      };
    }

    return undefined;
  }

  const chartState = result.recoveryState;

  return {
    result: output,
    appendRecoverySnapshot: () =>
      dependencies.appendRecoverySnapshot({
        toolName: "charts",
        toolCallId,
        address: result.outputAddress ?? output.details.address ?? "chart",
        changedCount: result.changedCount ?? 1,
        chartState,
      }),
    appendResultNote: appendMutationResultNote,
    unavailableReason: captureError ?? "Chart backup capture was skipped.",
    unavailableNote: backupUnavailableNote(captureError),
    dispatchSnapshotCreated: (snapshot) =>
      dependencies.dispatchSnapshotCreated(snapshot),
  };
}

/**
 * User-visible note for a missing chart backup.
 *
 * A capture *failure* must name itself in the result text — the structured
 * `details.recovery.reason` is not enough for the model to act on, and a
 * generic "backup not created" note hides a real error.
 */
function backupUnavailableNote(captureError: string | null): string {
  return captureError === null
    ? "ℹ️ Backup not created for this chart mutation."
    : `⚠️ ${captureError}. The chart change was applied without a restore point.`;
}

export function createChartsTool(
  dependencies: Partial<ChartsToolDependencies> = {},
): AgentTool<typeof schema, ChartsDetails> {
  const resolvedDependencies: ChartsToolDependencies = {
    executeAction:
      dependencies.executeAction ?? defaultDependencies.executeAction,
    captureChartPresent:
      dependencies.captureChartPresent ??
      defaultDependencies.captureChartPresent,
    appendRecoverySnapshot:
      dependencies.appendRecoverySnapshot ??
      defaultDependencies.appendRecoverySnapshot,
    appendAuditEntry:
      dependencies.appendAuditEntry ?? defaultDependencies.appendAuditEntry,
    dispatchSnapshotCreated:
      dependencies.dispatchSnapshotCreated ??
      defaultDependencies.dispatchSnapshotCreated,
  };

  return {
    name: "charts",
    label: "Charts",
    description:
      "Manage workbook charts: list charts, create charts from ranges, update chart properties/source data, " +
      "delete charts, and capture chart images for visual verification.",
    parameters: schema,
    execute: async (
      toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ChartsDetails>> => {
      const isMutation = isMutatingChartsAction(params.action);
      const mutationFinalizeDependencies: MutationFinalizeDependencies = {
        appendAuditEntry: (entry) =>
          resolvedDependencies.appendAuditEntry(entry),
      };

      try {
        validateParams(params);

        let beforeChartState: RecoveryChartPresentState | null = null;
        let captureError: string | null = null;
        if (params.action === "update") {
          try {
            beforeChartState = await resolvedDependencies.captureChartPresent(
              requireName(params),
              params.sheet,
            );
          } catch (error) {
            captureError = `Chart backup capture failed: ${getErrorMessage(error)}`;
          }
        }

        const result = await resolvedDependencies.executeAction(params);
        if (params.action === "update" && beforeChartState) {
          result.recoveryState = beforeChartState;
        }

        if (!isMutation) {
          return result.result;
        }

        const outputAddress =
          result.outputAddress ?? result.result.details.address;
        const recovery = buildRecoveryStep(
          resolvedDependencies,
          toolCallId,
          result.result,
          result,
          captureError,
        );
        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "charts",
            toolCallId,
            blocked: false,
            ...(outputAddress !== undefined ? { outputAddress } : {}),
            changedCount: result.changedCount ?? 1,
            changes: [],
            summary: result.auditSummary ?? `${params.action} chart`,
          },
          ...(recovery !== undefined ? { recovery } : {}),
        });

        if (result.sourceRangeChanged) {
          appendMutationResultNote(
            result.result,
            CHART_SOURCE_RANGE_RESTORE_NOTE,
          );
        }

        return result.result;
      } catch (error) {
        const message = getErrorMessage(error);

        if (isMutation) {
          const outputAddress =
            params.name ?? params.source_range ?? params.sheet;
          await finalizeMutationOperation(mutationFinalizeDependencies, {
            auditEntry: {
              toolName: "charts",
              toolCallId,
              blocked: true,
              ...(outputAddress !== undefined ? { outputAddress } : {}),
              changedCount: 0,
              changes: [],
              summary: `error: ${message}`,
            },
          });
        }

        return {
          content: [
            { type: "text", text: `Error (${params.action}): ${message}` },
          ],
          details: {
            kind: "charts",
            action: params.action,
            ...(params.name !== undefined ? { name: params.name } : {}),
            ...(params.source_range !== undefined
              ? { sourceRange: params.source_range }
              : {}),
          },
        };
      }
    },
  };
}
