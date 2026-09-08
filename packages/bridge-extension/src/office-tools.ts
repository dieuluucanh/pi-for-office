/**
 * Office tool catalog — the Office.js operations the Pi agent can drive
 * through the bridge.
 *
 * Each descriptor is host-scoped (excel / word / powerpoint). The task-pane
 * add-in executes the actual Office.js call; this extension only describes the
 * tool to the LLM and routes the call over the bridge.
 *
 * Keep `op` values in sync with the add-in's `bridge/tool-registry.ts` (same
 * repository, `packages/add-in`). The pane validates args again at runtime, so
 * this file is the *contract*, not the enforcement point.
 */

import { Type, type TSchema } from "typebox";
import { StringEnum } from "@earendil-works/pi-ai";
import type { OfficeHostApp } from "./protocol.js";

export interface OfficeToolDescriptor {
  /** Host app that can execute this op. */
  host: OfficeHostApp;
  /** Payload op id, namespaced by host: "excel.read_range". */
  op: string;
  /** Pi-registered tool name, e.g. "office_excel_read_range". */
  name: string;
  label: string;
  description: string;
  promptGuidelines?: string[];
  parameters: TSchema;
}

/** Office apps we currently expose tools for (map host → tool prefix). */
export const HOST_APP_LABEL: Record<OfficeHostApp, string> = {
  excel: "Excel",
  word: "Word",
  powerpoint: "PowerPoint",
};

function officeToolName(host: OfficeHostApp, op: string): string {
  return `office_${host}_${op}`;
}

/* ── Excel ───────────────────────────────────────────────────────────── */

const EXCEL_READ_RANGE_SCHEMA = Type.Object({
  range: Type.String({
    description:
      'Cell range in A1 notation, e.g. "A1:D10" or "Sheet2!A1:B5". ' +
      "Uses the active sheet when no sheet is specified.",
  }),
  mode: Type.Optional(
    StringEnum(["compact", "csv", "detailed"], {
      description:
        '"compact" (default): markdown table. "csv": raw values. "detailed": with formulas/formats.',
    }),
  ),
});

/* ── Word ────────────────────────────────────────────────────────────── */

const WORD_READ_SCOPE = StringEnum(["all", "selection"], {
  description: '"all": whole document. "selection": currently selected text only.',
});

/* ── PowerPoint ──────────────────────────────────────────────────────── */

const PPT_SLIDE_INDEX = Type.Integer({
  minimum: 1,
  description: "1-based slide index.",
});

/* ── Catalog ─────────────────────────────────────────────────────────── */

export const OFFICE_TOOL_DESCRIPTORS: OfficeToolDescriptor[] = [
  /* Excel */
  {
    host: "excel",
    op: "get_overview",
    name: officeToolName("excel", "get_overview"),
    label: "Excel Workbook Overview",
    description:
      "Read a compact overview of the attached Excel workbook: sheet names, used ranges, " +
      "table names, and named ranges. Call this first before any range operation.",
    promptGuidelines: [
      "Call office_excel_get_overview before office_excel_read_range to learn the workbook structure.",
    ],
    parameters: Type.Object({}),
  },
  {
    host: "excel",
    op: "read_range",
    name: officeToolName("excel", "read_range"),
    label: "Excel Read Range",
    description:
      "Read cell values (and optionally formulas/formatting) from a range in the attached Excel workbook.",
    parameters: EXCEL_READ_RANGE_SCHEMA,
  },
  {
    host: "excel",
    op: "write_cells",
    name: officeToolName("excel", "write_cells"),
    label: "Excel Write Cells",
    description:
      "Write a 2D array of values into the attached Excel workbook, starting at a top-left cell. " +
      "values[row][col]; the array is written down and to the right from start_cell.",
    promptGuidelines: [
      "Prefer office_excel_write_cells in a single batched call instead of many small edits.",
      "Always verify with office_excel_read_range after office_excel_write_cells when the change is user-visible.",
    ],
    parameters: Type.Object({
      start_cell: Type.String({
        description: 'Top-left cell to write from, e.g. "A1" or "Sheet2!B3".',
      }),
      values: Type.Array(Type.Array(Type.Any()), {
        description: "2D array of cell values (rows × cols).",
      }),
    }),
  },
  {
    host: "excel",
    op: "fill_formula",
    name: officeToolName("excel", "fill_formula"),
    label: "Excel Fill Formula",
    description:
      "Write a formula into a single contiguous range of the attached Excel workbook. " +
      "Relative references adjust as the formula fills.",
    parameters: Type.Object({
      range: Type.String({ description: 'Target range, e.g. "B2:B20" or "Sheet1!C3:F20".' }),
      formula: Type.String({
        description: 'Formula starting with "=", e.g. "=SUM(B2:B10)".',
      }),
    }),
  },

  /* Word */
  {
    host: "word",
    op: "get_overview",
    name: officeToolName("word", "get_overview"),
    label: "Word Document Overview",
    description:
      "Read a compact overview of the attached Word document: heading outline, paragraph count, " +
      "table count, and word count. Call this first before editing.",
    promptGuidelines: [
      "Call office_word_get_overview before office_word_insert_text or office_word_replace_text.",
    ],
    parameters: Type.Object({}),
  },
  {
    host: "word",
    op: "read_document",
    name: officeToolName("word", "read_document"),
    label: "Word Read Document",
    description:
      "Read text from the attached Word document: the whole body or the current selection.",
    parameters: Type.Object({
      scope: Type.Optional(WORD_READ_SCOPE),
      maxChars: Type.Optional(
        Type.Integer({
          minimum: 100,
          maximum: 200000,
          description: "Cap on characters returned (default 20000).",
        }),
      ),
    }),
  },
  {
    host: "word",
    op: "insert_text",
    name: officeToolName("word", "insert_text"),
    label: "Word Insert Text",
    description:
      "Insert text at the start or end of the attached Word document, or replace the current selection.",
    parameters: Type.Object({
      text: Type.String({ description: "Text to insert." }),
      location: Type.Optional(
        StringEnum(["start", "end", "replace_selection"], {
          description: '"end" (default) appends to the document. "replace_selection" overwrites the selection.',
        }),
      ),
    }),
  },
  {
    host: "word",
    op: "replace_text",
    name: officeToolName("word", "replace_text"),
    label: "Word Replace Text",
    description:
      "Find and replace literal text in the attached Word document. Returns how many occurrences were replaced.",
    parameters: Type.Object({
      find: Type.String({ description: "Literal text to find." }),
      replace: Type.String({ description: "Replacement text." }),
      matchCase: Type.Optional(Type.Boolean({ description: "Case-sensitive match (default false)." })),
    }),
  },

  /* PowerPoint */
  {
    host: "powerpoint",
    op: "get_overview",
    name: officeToolName("powerpoint", "get_overview"),
    label: "PowerPoint Overview",
    description:
      "Read a compact overview of the attached presentation: slide count, each slide's title and " +
      "shape count. Call this first before any slide operation.",
    promptGuidelines: [
      "Call office_powerpoint_get_overview before office_powerpoint_read_slide or office_powerpoint_add_slide.",
    ],
    parameters: Type.Object({}),
  },
  {
    host: "powerpoint",
    op: "read_slide",
    name: officeToolName("powerpoint", "read_slide"),
    label: "PowerPoint Read Slide",
    description:
      "Read all text content of one slide in the attached presentation (shapes, text frames, notes).",
    parameters: Type.Object({
      slideIndex: PPT_SLIDE_INDEX,
    }),
  },
  {
    host: "powerpoint",
    op: "add_slide",
    name: officeToolName("powerpoint", "add_slide"),
    label: "PowerPoint Add Slide",
    description:
      "Append a new slide to the attached presentation and navigate to it. Uses the default layout.",
    parameters: Type.Object({}),
  },
  {
    host: "powerpoint",
    op: "add_text_box",
    name: officeToolName("powerpoint", "add_text_box"),
    label: "PowerPoint Add Text Box",
    description:
      "Add a text box with the given text to a slide. Coordinates/geometry are in points.",
    parameters: Type.Object({
      slideIndex: PPT_SLIDE_INDEX,
      text: Type.String({ description: "Text box content." }),
      x: Type.Optional(Type.Number({ description: "Left edge in points (default centered)." })),
      y: Type.Optional(Type.Number({ description: "Top edge in points (default centered)." })),
      width: Type.Optional(Type.Number({ description: "Width in points (default 400)." })),
      height: Type.Optional(Type.Number({ description: "Height in points (default 60)." })),
    }),
  },
];

/** Index by op id for fast lookup. */
export const OFFICE_TOOL_BY_OP: ReadonlyMap<string, OfficeToolDescriptor> = new Map(
  OFFICE_TOOL_DESCRIPTORS.map((d) => [`${d.host}.${d.op}`, d]),
);

/** All tool names registered by this extension. */
export const OFFICE_TOOL_NAMES: readonly string[] = OFFICE_TOOL_DESCRIPTORS.map((d) => d.name);

/** The office host this tool name drives, or null when unknown. */
export function hostForToolName(name: string): OfficeHostApp | null {
  for (const d of OFFICE_TOOL_DESCRIPTORS) {
    if (d.name === name) return d.host;
  }
  return null;
}
