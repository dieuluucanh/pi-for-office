/**
 * Harness for mocking the global `Excel` (and the subset of `Office` the
 * capability probe reads) so the live Excel tools can be exercised under Node.
 *
 * @see ../../src/storage/local/app-storage.ts — `installFakeAppStorage` installs
 * the same `settings` surface `getResolvedConventions` reads.
 *
 * The fake records every assignment on the range format/font/fill/border
 * surface, which is what the format_cells and conditional_format tests assert.
 * A1 addresses are parsed so `rowCount`/`columnCount` are realistic; capture
 * paths that read a property the fake leaves `undefined` degrade to the same
 * "checkpoint unsupported" branch the real tools already handle.
 */

import {
  setAppStorage,
  type AppStorage,
} from "../../src/storage/local/app-storage.ts";

export interface FakeBorder {
  style?: string;
  weight?: string;
  color?: string;
  readonly loaded: string[];
  load(props?: string): void;
}

export interface FakeFont {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strikethrough?: boolean;
  color?: string;
  size?: number;
  name?: string;
  readonly loaded: string[];
  load(props?: string): void;
}

export interface FakeFill {
  color?: string;
  readonly loaded: string[];
  load(props?: string): void;
}

export interface FakeRangeFormat {
  readonly font: FakeFont;
  readonly fill: FakeFill;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  wrapText?: boolean;
  indentLevel?: number;
  columnWidth?: number;
  rowHeight?: number;
  readonly loaded: string[];
  readonly borders: FakeBorders;
  readonly autofitColumnsCalls: number;
  readonly autofitRowsCalls: number;
  load(props?: string): void;
  autofitColumns(): void;
  autofitRows(): void;
}

export interface FakeBorders {
  readonly items: Map<string, FakeBorder>;
  getItem(edge: string): FakeBorder;
}

export interface FakeConditionalFormatDataBar {
  barDirection?: string;
  axisColor?: string;
  showDataBarOnly?: boolean;
  lowerBoundRule?: Record<string, DynamicValue>;
  upperBoundRule?: Record<string, DynamicValue>;
  readonly positiveFormat: { fillColor?: string; borderColor?: string };
  readonly negativeFormat: { fillColor?: string; borderColor?: string };
}

export interface FakeColorScaleCriterion {
  type?: string;
  formula?: string;
  color?: string;
}

export interface FakeConditionalFormat {
  /** `String(Excel.ConditionalFormatType.<family>)` captured at add() time. */
  readonly type: string;
  stopIfTrue?: boolean;
  readonly custom: {
    rule: { formula?: string };
    format: FakeConditionalRangeFormat;
  };
  readonly cellValue: {
    rule: { formula1?: string; formula2?: string; operator?: string };
    format: FakeConditionalRangeFormat;
  };
  readonly dataBar: FakeConditionalFormatDataBar;
  readonly colorScale: {
    criteria: {
      minimum?: FakeColorScaleCriterion;
      midpoint?: FakeColorScaleCriterion;
      maximum?: FakeColorScaleCriterion;
    };
  };
  readonly iconSet: {
    style?: string;
    reverseIconOrder?: boolean;
    showIconOnly?: boolean;
    criteria?: FakeIconCriterion[];
  };
  readonly topBottom: {
    rule: { type?: string; rank?: number };
    format: FakeConditionalRangeFormat;
  };
  readonly preset: {
    rule: { criterion?: string };
    format: FakeConditionalRangeFormat;
  };
  readonly textComparison: {
    rule: { operator?: string; text?: string };
    format: FakeConditionalRangeFormat;
  };
}

export interface FakeIconCriterion {
  type?: string;
  formula?: string;
  operator?: string;
}

export interface FakeConditionalRangeFormat {
  readonly fill: { color?: string };
  readonly font: {
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: string;
  };
}

export interface FakeConditionalFormats {
  readonly added: FakeConditionalFormat[];
  /**
   * Pre-existing rules the capture path iterates. Kept empty on purpose:
   * capture/apply fidelity is covered by the recovery-log suites, while these
   * tests assert the add path and the gate.
   */
  readonly items: FakeConditionalFormat[];
  readonly clearAllCalls: number;
  count: number;
  load(props?: string): void;
  add(type: { toString(): string }): FakeConditionalFormat;
  getCount(): { value: number };
  clearAll(): void;
}

export interface FakeRange {
  readonly address: string;
  rowCount: number;
  columnCount: number;
  /** Values matrix, for the overview's header/preview reads. */
  values: DynamicValue[][];
  readonly format: FakeRangeFormat;
  numberFormat: string[][] | string;
  readonly conditionalFormats: FakeConditionalFormats;
  readonly mergedAreas: string[];
  readonly calls: string[];
  load(props?: string): void;
  merge(): void;
  unmerge(): void;
  getEntireRow(): FakeRange;
  getEntireColumn(): FakeRange;
  getColumn(index: number): FakeRange;
  getRow(index: number): FakeRange;
  getUsedRangeOrNullObject(): FakeRange & { isNullObject: boolean };
  getMergedAreasOrNullObject(): {
    isNullObject: boolean;
    areas: { items: FakeRange[] };
    load(props?: string): void;
  };
}

export interface FakeRangeAreas {
  readonly areas: { items: FakeRange[]; load(props?: string): void };
  readonly address: string;
  readonly format: FakeRangeFormat;
  readonly calls: string[];
  load(props?: string): void;
}

export interface FakeWorksheet {
  readonly name: string;
  readonly calls: string[];
  /** Mirrors Office.js `*OrNullObject` results. */
  isNullObject: boolean;
  /** 0-based sheet position, as Office.js reports it. */
  position: number;
  showGridlines: boolean;
  showHeadings: boolean;
  /** Every range this sheet handed out via `getRange`, in call order. */
  readonly createdRanges: FakeRange[];
  readonly createdAreas: FakeRangeAreas[];
  readonly shapes: { items: DynamicValue[]; load(props?: string): void };
  readonly tables: { items: DynamicValue[]; load(props?: string): void };
  readonly charts: {
    items: DynamicValue[];
    count: number;
    load(props?: string): void;
  };
  readonly comments: { items: DynamicValue[]; load(props?: string): void };
  readonly pivotTables: {
    getCount(): { value: number };
    load(props?: string): void;
  };
  visibility: string;
  load(props?: string): void;
  getRange(address: string): FakeRange;
  getRanges(addresses: string): FakeRangeAreas;
  getUsedRangeOrNullObject(): FakeRange & { isNullObject: boolean };
}

export interface FakeWorksheetCollection {
  readonly items: FakeWorksheet[];
  load(props?: string): void;
  getItem(name: string): FakeWorksheet;
  getItemOrNullObject(name: string): FakeWorksheet;
  getItemAt(index: number): FakeWorksheet;
  getActiveWorksheet(): FakeWorksheet;
  add(name?: string): FakeWorksheet;
}

export interface FakeWorkbook {
  readonly worksheets: FakeWorksheetCollection;
  /** Workbook-level named ranges (empty by default). */
  readonly names: { items: DynamicValue[]; load(props?: string): void };
  name: string;
  load(props?: string): void;
}

export interface FakeContext {
  readonly workbook: FakeWorkbook;
  syncCount: number;
  sync(): Promise<void>;
}

export interface FakeExcelGlobal {
  readonly context: FakeContext;
  run<T>(callback: (context: FakeContext) => Promise<T>): Promise<T>;
  readonly ConditionalFormatType: Record<string, { toString(): string }>;
  readonly ConditionalFormatColorCriterionType?: Record<string, string>;
  readonly ConditionalFormatIconSet?: Record<string, string>;
  readonly ConditionalFormatIconRuleType?: Record<string, string>;
  readonly ConditionalFormatTopBottomCriterionType?: Record<string, string>;
  readonly ConditionalFormatPresetCriterion?: Record<string, string>;
}

/** Parse "A1:B3" (or "A1") into row/column counts. */
export function parseA1Counts(address: string): {
  rowCount: number;
  columnCount: number;
} {
  const clean = address.includes("!")
    ? address.slice(address.indexOf("!") + 1)
    : address;
  const [start, end] = clean.split(":");
  const startCell = /^\$?([A-Z]+)\$?(\d+)$/u.exec(start?.trim() ?? "");
  const endCell = /^\$?([A-Z]+)\$?(\d+)$/u.exec((end ?? start)?.trim() ?? "");
  if (!startCell || !endCell) {
    return { rowCount: 1, columnCount: 1 };
  }
  const columnIndex = (letters: string): number =>
    letters
      .toUpperCase()
      .split("")
      .reduce((total, letter) => total * 26 + (letter.charCodeAt(0) - 64), 0);
  return {
    rowCount: Math.max(1, Number(endCell[2]) - Number(startCell[2]) + 1),
    columnCount: Math.max(
      1,
      columnIndex(endCell[1] ?? "A") - columnIndex(startCell[1] ?? "A") + 1,
    ),
  };
}

function makeBorder(): FakeBorder {
  const loaded: string[] = [];
  return {
    loaded,
    load(props?: string) {
      loaded.push(props ?? "");
    },
  };
}

function makeFont(): FakeFont {
  const loaded: string[] = [];
  return {
    loaded,
    load(props?: string) {
      loaded.push(props ?? "");
    },
  };
}

function makeFill(): FakeFill {
  const loaded: string[] = [];
  return {
    loaded,
    load(props?: string) {
      loaded.push(props ?? "");
    },
  };
}

function makeConditionalRangeFormat(): FakeConditionalRangeFormat {
  return { fill: {}, font: {} };
}

function makeConditionalFormat(type: {
  toString(): string;
}): FakeConditionalFormat {
  return {
    type: String(type),
    custom: {
      rule: {},
      format: makeConditionalRangeFormat(),
    },
    cellValue: {
      rule: {},
      format: makeConditionalRangeFormat(),
    },
    dataBar: {
      positiveFormat: {},
      negativeFormat: {},
    },
    colorScale: { criteria: {} },
    iconSet: {},
    topBottom: {
      rule: {},
      format: makeConditionalRangeFormat(),
    },
    preset: {
      rule: {},
      format: makeConditionalRangeFormat(),
    },
    textComparison: {
      rule: {},
      format: makeConditionalRangeFormat(),
    },
  };
}

function makeConditionalFormats(): FakeConditionalFormats {
  const added: FakeConditionalFormat[] = [];
  const state = {
    added,
    items: [] as FakeConditionalFormat[],
    clearAllCalls: 0,
    count: 0,
    load(): void {},
    add(type: { toString(): string }): FakeConditionalFormat {
      const cf = makeConditionalFormat(type);
      added.push(cf);
      state.count += 1;
      return cf;
    },
    getCount(): { value: number } {
      return { value: state.count };
    },
    clearAll(): void {
      state.clearAllCalls += 1;
      state.count = 0;
    },
  };
  return state;
}

export function makeRangeFormat(): FakeRangeFormat {
  const loaded: string[] = [];
  const borders: FakeBorders = {
    items: new Map<string, FakeBorder>(),
    getItem(edge: string): FakeBorder {
      let border = borders.items.get(edge);
      if (!border) {
        border = makeBorder();
        borders.items.set(edge, border);
      }
      return border;
    },
  };
  const format = {
    font: makeFont(),
    fill: makeFill(),
    loaded,
    borders,
    autofitColumnsCalls: 0,
    autofitRowsCalls: 0,
    load(props?: string) {
      loaded.push(props ?? "");
    },
    autofitColumns() {
      format.autofitColumnsCalls += 1;
    },
    autofitRows() {
      format.autofitRowsCalls += 1;
    },
  };
  return format;
}

export function makeRange(address: string): FakeRange {
  const { rowCount, columnCount } = parseA1Counts(address);
  const calls: string[] = [];
  const mergedAreas: string[] = [];
  const format = makeRangeFormat();
  const conditionalFormats = makeConditionalFormats();
  const range: FakeRange = {
    address,
    rowCount,
    columnCount,
    values: [],
    format,
    numberFormat: Array.from({ length: rowCount }, () =>
      Array.from({ length: columnCount }, () => "General"),
    ),
    conditionalFormats,
    mergedAreas,
    calls,
    load(props?: string) {
      calls.push(props ?? "");
    },
    merge() {
      if (!mergedAreas.includes(address)) mergedAreas.push(address);
    },
    unmerge() {
      mergedAreas.length = 0;
    },
    getEntireRow(): FakeRange {
      const row = makeRange(address);
      row.rowCount = 1;
      return row;
    },
    getEntireColumn(): FakeRange {
      const col = makeRange(address);
      col.columnCount = 1;
      return col;
    },
    getColumn(): FakeRange {
      return makeRange(address);
    },
    getRow(): FakeRange {
      return makeRange(address);
    },
    getUsedRangeOrNullObject() {
      return { ...range, isNullObject: false };
    },
    getMergedAreasOrNullObject() {
      const items: FakeRange[] = mergedAreas.map((merged) => makeRange(merged));
      return {
        isNullObject: items.length === 0,
        areas: {
          items,
          load() {},
        },
        load() {},
      };
    },
  };
  return range;
}

export function makeWorksheet(name: string): FakeWorksheet {
  const calls: string[] = [];
  const createdRanges: FakeRange[] = [];
  const createdAreas: FakeRangeAreas[] = [];
  const sheet: FakeWorksheet = {
    name,
    calls,
    isNullObject: false,
    position: 0,
    showGridlines: true,
    showHeadings: true,
    createdRanges,
    createdAreas,
    shapes: { items: [], load: () => {} },
    tables: { items: [], load: () => {} },
    charts: { items: [], count: 0, load: () => {} },
    comments: { items: [], load: () => {} },
    pivotTables: { getCount: () => ({ value: 0 }), load: () => {} },
    visibility: "Visible",
    load(props?: string) {
      calls.push(props ?? "");
    },
    getRange(address: string): FakeRange {
      const range = makeRange(address);
      createdRanges.push(range);
      return range;
    },
    getRanges(addresses: string): FakeRangeAreas {
      const items = addresses
        .split(",")
        .map((address) => makeRange(address.trim()));
      const format = makeRangeFormat();
      const areaCalls: string[] = [];
      const areas: FakeRangeAreas = {
        areas: {
          items,
          load() {},
        },
        address: addresses,
        format,
        calls: areaCalls,
        load(props?: string) {
          areaCalls.push(props ?? "");
        },
      };
      createdAreas.push(areas);
      return areas;
    },
    getUsedRangeOrNullObject() {
      const range = makeRange("A1");
      return { ...range, isNullObject: false };
    },
  };
  return sheet;
}

export function makeContext(sheetNames: string[] = ["Sheet1"]): FakeContext {
  const items = sheetNames.map((name) => makeWorksheet(name));
  items.forEach((sheet, index) => {
    sheet.position = index;
  });
  const worksheets: FakeWorksheetCollection = {
    items,
    load() {},
    getItem(name: string): FakeWorksheet {
      const sheet = items.find((candidate) => candidate.name === name);
      if (!sheet) throw new Error(`Worksheet ${name} not found.`);
      return sheet;
    },
    getItemOrNullObject(name: string): FakeWorksheet {
      const sheet = items.find((candidate) => candidate.name === name);
      if (sheet) return sheet;
      const placeholder = makeWorksheet(name);
      placeholder.isNullObject = true;
      return placeholder;
    },
    getItemAt(index: number): FakeWorksheet {
      const sheet = items[index];
      if (!sheet) throw new Error(`No worksheet at index ${index}.`);
      return sheet;
    },
    getActiveWorksheet(): FakeWorksheet {
      const first = items[0];
      if (!first) throw new Error("No active worksheet.");
      return first;
    },
    add(name?: string): FakeWorksheet {
      const sheet = makeWorksheet(name ?? `Sheet${items.length + 1}`);
      sheet.position = items.length;
      items.push(sheet);
      return sheet;
    },
  };
  const workbook: FakeWorkbook = {
    worksheets,
    names: { items: [], load: () => {} },
    name: "Workbook1",
    load() {},
  };
  const context: FakeContext = {
    workbook,
    syncCount: 0,
    sync(): Promise<void> {
      context.syncCount += 1;
      return Promise.resolve();
    },
  };
  return context;
}

function makeConditionalFormatType(): Record<string, { toString(): string }> {
  const names = [
    "custom",
    "cellValue",
    "dataBar",
    "colorScale",
    "iconSet",
    "topBottom",
    "presetCriteria",
    "containsText",
  ];
  const map: Record<string, { toString(): string }> = {};
  for (const name of names) {
    map[name] = { toString: () => name };
  }
  return map;
}

/** Install a fake global `Excel` and return the context it serves. */
export function makeExcelGlobal(context: FakeContext): FakeExcelGlobal {
  return {
    context,
    run<T>(callback: (ctx: FakeContext) => Promise<T>): Promise<T> {
      return callback(context);
    },
    ConditionalFormatType: makeConditionalFormatType(),
  };
}

export interface FakeOfficeCapabilities {
  /** Set names the host claims to support, e.g. { ExcelApi: "1.9" }. */
  supported: Record<string, string>;
  platform?: string;
}

/** A `Office.context.requirements` double driven by a declared set map. */
function makeOfficeGlobal(capabilities: FakeOfficeCapabilities): DynamicValue {
  return {
    context: {
      platform: capabilities.platform ?? "PC",
      requirements: {
        isSetSupported(set: string, minVersion: string): boolean {
          const supportedVersion = capabilities.supported[set];
          if (supportedVersion === undefined) return false;
          return compareVersions(supportedVersion, minVersion) >= 0;
        },
      },
    },
  };
}

/** Compare dotted requirement-set versions ("1.10" > "1.9"). */
export function compareVersions(a: string, b: string): number {
  const parse = (value: string): number[] =>
    value.split(".").map((part) => Number.parseInt(part, 10));
  const left = parse(a);
  const right = parse(b);
  for (let index = 0; index < Math.max(left.length, right.length); index += 1) {
    const diff = (left[index] ?? 0) - (right[index] ?? 0);
    if (diff !== 0) return diff > 0 ? 1 : -1;
  }
  return 0;
}

export interface FakeExcelEnvOptions {
  sheets?: string[];
  /** Declared Office requirement sets; omit to run with no `Office` global. */
  capabilities?: FakeOfficeCapabilities;
}

/**
 * Install a fake `Excel` (and optionally `Office`) global, run `fn`, then
 * restore the previous globals. The context handed to `fn` is the same one the
 * fake `Excel.run` serves, so tests can assert on recorded assignments.
 */
export async function withFakeExcel<T>(
  fn: (context: FakeContext, excel: FakeExcelGlobal) => T | Promise<T>,
  options: FakeExcelEnvOptions = {},
): Promise<T> {
  const context = makeContext(options.sheets ?? ["Sheet1"]);
  const excel = makeExcelGlobal(context);

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  const hadOffice = Reflect.has(globalThis, "Office");
  const previousOffice = Reflect.get(globalThis, "Office");

  Reflect.set(globalThis, "Excel", excel);
  if (options.capabilities) {
    Reflect.set(globalThis, "Office", makeOfficeGlobal(options.capabilities));
  }

  try {
    return await fn(context, excel);
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
    if (hadOffice) Reflect.set(globalThis, "Office", previousOffice);
    else Reflect.deleteProperty(globalThis, "Office");
  }
}

/** Install a minimal `AppStorage` double so the conventions store resolves. */
export function installFakeAppStorage(): void {
  const values = new Map<string, DynamicValue>();
  const settings = {
    get: (key: string): Promise<DynamicValue> =>
      Promise.resolve(values.get(key) ?? null),
    set: (key: string, value: DynamicValue): Promise<void> => {
      values.set(key, value);
      return Promise.resolve();
    },
  };
  const appStorage: DynamicValue = { settings };
  setAppStorage(appStorage as AppStorage);
}
