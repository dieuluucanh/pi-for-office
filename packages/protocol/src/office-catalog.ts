/**
 * Office op catalog — the single source of truth for every operation the Pi
 * agent can drive through the office bridge.
 *
 * Both sides derive from this file:
 *
 *  - The Pi extension (`@dieulc/pi-office-bridge`) registers one
 *    `office_<host>_<op>` tool per entry.
 *  - The task-pane add-in builds its bridge op registry (`ALL_BRIDGE_OPS`)
 *    from the same op ids and delegates execution to the shared local tool
 *    implementations. Word/PowerPoint local tools import the parameter
 *    schemas from here, so bridge and local mode can never drift.
 *
 * The catalog is **server-authoritative**: a pane may only *narrow* the set of
 * ops it advertises in `hello` — it never supplies tool names, schemas, or
 * descriptions. This keeps the loopback WebSocket surface safe against a rogue
 * local page trying to inject tool definitions into the agent.
 */

import { Type, type TLiteral, type TSchema, type TUnion } from "typebox";
import type { OfficeHostApp } from "./protocol.js";

export interface OfficeCatalogEntry {
  /** Host app that can execute this op. */
  host: OfficeHostApp;
  /** Payload op id, namespaced by host: "excel.read_range". */
  op: string;
  /** Pi-registered tool name, e.g. "office_excel_read_range". */
  name: string;
  /** Short UI/command label. */
  label: string;
  /** Long description shown to the LLM. */
  description: string;
  /** One-line description listed under "Available tools". */
  promptSnippet?: string;
  /** Tool-specific guideline bullets (pi appends them flat; each must name the tool). */
  promptGuidelines?: string[];
  /** TypeBox schema for the tool's arguments. */
  parameters: TSchema;
  /** Catalog version this entry was introduced in. */
  since: number;
  /** True when the op was exposed by the legacy 0.2.x bridge (v1 op set). */
  legacyV1?: boolean;
}

/** Increment when entries or schemas change incompatibly. */
export const CATALOG_VERSION = 3;

/** Office apps we expose tools for (map host → tool prefix). */
export const HOST_APP_LABEL: Record<OfficeHostApp, string> = {
  excel: "Excel",
  word: "Word",
  powerpoint: "PowerPoint",
};

/** String enum helper (protocol must not depend on pi-ai's StringEnum). */
export function catalogEnum<T extends readonly string[]>(
  values: T,
  options?: { description?: string },
): TUnion<[TLiteral<T[number]>]> {
  // SAFETY: the runtime array holds one TLiteral per input value; the tuple
  // cast only affects the static type so `Static` preserves the literal union.
  const schemas = values.map((v) => Type.Literal(v)) as unknown as [
    TLiteral<T[number]>,
  ];
  return Type.Union(schemas, options);
}

export function officeToolName(host: OfficeHostApp, op: string): string {
  return `office_${host}_${op}`;
}

/* ── Shared parameter fragments ─────────────────────────────────────── */

const WORD_ALIGNMENT = catalogEnum(
  ["Left", "Centered", "Right", "Justified"] as const,
  {
    description:
      'Paragraph alignment. Word values: "Left", "Centered" (note the capital C), "Right", "Justified".',
  },
);

const PPT_ALIGNMENT = catalogEnum(
  ["Left", "Center", "Right", "Justify"] as const,
  {
    description:
      'Horizontal alignment. PowerPoint values: "Left", "Center", "Right", "Justify".',
  },
);

const WORD_FONT_PROPS = {
  bold: Type.Optional(
    Type.Boolean({ description: "Set bold (true) or unbold (false)." }),
  ),
  italic: Type.Optional(
    Type.Boolean({ description: "Set italic (true) or unitalicize (false)." }),
  ),
  underline: Type.Optional(
    Type.Boolean({ description: "Underline the text (single underline)." }),
  ),
  size: Type.Optional(
    Type.Number({ description: "Font size in points (e.g. 16)." }),
  ),
  name: Type.Optional(
    Type.String({ description: 'Font name (e.g. "Times New Roman").' }),
  ),
  color: Type.Optional(
    Type.String({ description: 'Font color as #RRGGBB (e.g. "#000000").' }),
  ),
  highlight: Type.Optional(
    Type.String({
      description:
        'Text highlight (marker) color as #RRGGBB, or "none" to clear. This is what a ' +
        '"fill"/"background"/"highlight" request on text maps to. Available on Word ' +
        "desktop and Word for the web (WordApi 1.1); on Word for Windows desktop a value " +
        "outside the 15 named highlight colors snaps to the nearest one.",
    }),
  ),
};

const WORD_PARAGRAPH_PROPS = {
  alignment: Type.Optional(WORD_ALIGNMENT),
  style: Type.Optional(
    Type.String({
      description:
        'Paragraph style name, e.g. "Heading 1", "Title", "Normal". Style names are locale-sensitive; a fallback to direct formatting is applied when the style cannot be verified.',
    }),
  ),
  spaceBefore: Type.Optional(
    Type.Number({ description: "Space before the paragraph, in points." }),
  ),
  spaceAfter: Type.Optional(
    Type.Number({ description: "Space after the paragraph, in points." }),
  ),
  lineSpacing: Type.Optional(
    Type.Number({ description: "Line spacing, in points." }),
  ),
  firstLineIndent: Type.Optional(
    Type.Number({ description: "First-line indent, in points." }),
  ),
  leftIndent: Type.Optional(
    Type.Number({ description: "Left indent, in points." }),
  ),
  shading: Type.Optional(
    Type.String({
      description:
        'Background fill ("shading") as #RRGGBB. Desktop-only: needs WordApiDesktop 1.3, so it ' +
        "reports an error on Word for the web. Its span depends on the tool: word_format_range " +
        "fills only the matched text, while word_insert_text / word_insert_blocks fill the whole " +
        "paragraph. Use `highlight` instead for a character-level marker that works on every platform.",
    }),
  ),
};

/**
 * Short human labels for the Word font schema properties, used to generate the
 * prompt's formatting list.
 *
 * Typed as `Record<keyof typeof WORD_FONT_PROPS, string>`, so the compiler
 * enforces coverage: adding a property to the schema fails typecheck until a
 * label exists here, and the prompt then picks it up automatically. That is what
 * stops the prompt from advertising a property the tools do not accept — or
 * silently omitting one they do.
 */
export const WORD_FONT_PROP_LABELS: Readonly<
  Record<keyof typeof WORD_FONT_PROPS, string>
> = {
  bold: "bold",
  italic: "italic",
  underline: "underline",
  size: "font size (points)",
  name: "font name",
  color: "color (#RRGGBB)",
  highlight: 'text highlight/background fill (#RRGGBB, or "none")',
};

/** Labels for the Word paragraph schema properties (see WORD_FONT_PROP_LABELS). */
export const WORD_PARAGRAPH_PROP_LABELS: Readonly<
  Record<keyof typeof WORD_PARAGRAPH_PROPS, string>
> = {
  alignment: 'alignment ("Left"/"Centered"/"Right"/"Justified")',
  style: "style",
  spaceBefore: "spacing",
  spaceAfter: "spacing",
  lineSpacing: "line spacing",
  firstLineIndent: "indents",
  leftIndent: "indents",
  shading: 'background fill ("shading", desktop-only)',
};

/**
 * Render the Word formatting properties as a comma-separated list, derived from
 * the schemas above. Duplicate labels (e.g. two spacing properties) collapse.
 */
export function wordFormattingPropertyList(): string {
  const labels = [
    ...Object.values(WORD_FONT_PROP_LABELS),
    ...Object.values(WORD_PARAGRAPH_PROP_LABELS),
  ];
  return [...new Set(labels)].join(", ");
}

/**
 * PowerPoint underline styles (`PowerPoint.ShapeFontUnderlineStyle`,
 * PowerPointApi 1.4) — 17 values.
 */
export const POWERPOINT_UNDERLINE_STYLES = [
  "None",
  "Single",
  "Double",
  "Heavy",
  "Dotted",
  "DottedHeavy",
  "Dash",
  "DashHeavy",
  "DashLong",
  "DashLongHeavy",
  "DotDash",
  "DotDashHeavy",
  "DotDotDash",
  "DotDotDashHeavy",
  "Wavy",
  "WavyHeavy",
  "WavyDouble",
] as const;

/**
 * PowerPoint font effects (**PowerPointApi 1.8**) shared by both text tools.
 * Newer than the 1.4 font properties, so it gets its own gate.
 */
const PPT_FONT_EFFECT_PROPS = {
  strikethrough: Type.Optional(
    Type.Boolean({
      description: "Strikethrough the text. Needs PowerPointApi 1.8.",
    }),
  ),
  doubleStrikethrough: Type.Optional(
    Type.Boolean({
      description: "Double strikethrough. Needs PowerPointApi 1.8.",
    }),
  ),
  smallCaps: Type.Optional(
    Type.Boolean({ description: "Small caps. Needs PowerPointApi 1.8." }),
  ),
  allCaps: Type.Optional(
    Type.Boolean({ description: "All caps. Needs PowerPointApi 1.8." }),
  ),
  subscript: Type.Optional(
    Type.Boolean({ description: "Subscript. Needs PowerPointApi 1.8." }),
  ),
  superscript: Type.Optional(
    Type.Boolean({ description: "Superscript. Needs PowerPointApi 1.8." }),
  ),
};

const PPT_SLIDE_INDEX = Type.Integer({
  minimum: 1,
  description: "1-based slide index.",
});

const PPT_TEXT_FORMAT_PROPS = {
  bold: Type.Optional(
    Type.Boolean({ description: "Set bold (true) or unbold (false)." }),
  ),
  italic: Type.Optional(
    Type.Boolean({ description: "Set italic (true) or unitalicize (false)." }),
  ),
  underline: Type.Optional(
    Type.Union(
      [
        Type.Boolean(),
        ...POWERPOINT_UNDERLINE_STYLES.map((style) => Type.Literal(style)),
      ],
      {
        description:
          'Underline. true = "Single", false = "None"; or name a style directly: ' +
          '"Double", "Heavy", "Dotted", "Wavy", "DashLong", … Needs PowerPointApi 1.4.',
      },
    ),
  ),
  fontSize: Type.Optional(
    Type.Number({ description: "Font size in points (e.g. 24)." }),
  ),
  fontName: Type.Optional(
    Type.String({ description: 'Font name (e.g. "Arial").' }),
  ),
  fontColor: Type.Optional(
    Type.String({ description: 'Font color as #RRGGBB (e.g. "#333333").' }),
  ),
  alignment: Type.Optional(PPT_ALIGNMENT),
  ...PPT_FONT_EFFECT_PROPS,
};

/**
 * Shape fill (**PowerPointApi 1.4**) shared by the text-box and slide tools.
 */
const PPT_SHAPE_FILL_PROPS = {
  fillColor: Type.Optional(
    Type.String({
      description:
        'Solid shape fill color as #RRGGBB (e.g. "#FFFF00"). Needs PowerPointApi 1.4.',
    }),
  ),
  fillTransparency: Type.Optional(
    Type.Number({
      minimum: 0,
      maximum: 1,
      description:
        "Fill transparency, 0 (opaque) to 1 (clear). Needs PowerPointApi 1.4.",
    }),
  ),
  fillClear: Type.Optional(
    Type.Boolean({
      description: "Remove the shape fill entirely. Needs PowerPointApi 1.4.",
    }),
  ),
};

/** PowerPoint shape line styles (`PowerPoint.ShapeLineStyle`, PowerPointApi 1.4). */
export const POWERPOINT_LINE_STYLES = [
  "Single",
  "ThickBetweenThin",
  "ThickThin",
  "ThinThick",
  "ThinThin",
] as const;

/**
 * PowerPoint shape line dash styles (`PowerPoint.ShapeLineDashStyle`,
 * PowerPointApi 1.4).
 */
export const POWERPOINT_LINE_DASH_STYLES = [
  "Dash",
  "DashDot",
  "DashDotDot",
  "LongDash",
  "LongDashDot",
  "RoundDot",
  "Solid",
  "SquareDot",
  "LongDashDotDot",
  "SystemDash",
  "SystemDot",
  "SystemDashDot",
] as const;

/**
 * PowerPoint bullet styles (`PowerPoint.BulletStyle`, PowerPointApi 1.10).
 */
export const POWERPOINT_BULLET_STYLES = [
  "AlphabetLowercasePeriod",
  "AlphabetUppercasePeriod",
  "ArabicNumeralParenthesisRight",
  "ArabicNumeralPeriod",
  "RomanLowercaseParenthesesBoth",
  "RomanLowercaseParenthesisRight",
  "RomanLowercasePeriod",
  "RomanUppercasePeriod",
  "AlphabetLowercaseParenthesesBoth",
  "AlphabetLowercaseParenthesisRight",
  "AlphabetUppercaseParenthesesBoth",
  "AlphabetUppercaseParenthesisRight",
  "ArabicNumeralParenthesesBoth",
  "ArabicNumeralPlain",
  "RomanUppercaseParenthesesBoth",
  "RomanUppercaseParenthesisRight",
  "SimplifiedChinesePlain",
  "SimplifiedChinesePeriod",
  "CircleNumberDoubleBytePlain",
  "CircleNumberWideDoubleByteWhitePlain",
  "CircleNumberWideDoubleByteBlackPlain",
  "TraditionalChinesePlain",
  "TraditionalChinesePeriod",
  "ArabicAlphabetDash",
  "ArabicAbjadDash",
  "HebrewAlphabetDash",
  "KanjiKoreanPlain",
  "KanjiKoreanPeriod",
  "ArabicDoubleBytePlain",
  "ArabicDoubleBytePeriod",
  "ThaiAlphabetPeriod",
  "ThaiAlphabetParenthesisRight",
  "ThaiAlphabetParenthesesBoth",
  "ThaiNumeralPeriod",
  "ThaiNumeralParenthesisRight",
  "ThaiNumeralParenthesesBoth",
  "HindiAlphabetPeriod",
  "HindiNumeralPeriod",
  "KanjiSimplifiedChineseDoubleBytePeriod",
  "HindiNumeralParenthesisRight",
  "HindiAlphabet1Period",
] as const;

/** PowerPoint bullet types (`PowerPoint.BulletType`, PowerPointApi 1.10). */
export const POWERPOINT_BULLET_TYPES = [
  "None",
  "Numbered",
  "Unnumbered",
] as const;

/**
 * Paragraph/bullet properties shared by both text tools.
 *
 * The split matters: `bulletVisible` is PowerPointApi 1.4, while bullet style
 * and paragraph `indentLevel` are 1.10 — so a 1.4 host still gets bullet
 * visibility.
 */
const PPT_PARAGRAPH_PROPS = {
  bulletVisible: Type.Optional(
    Type.Boolean({
      description:
        "Show (true) or hide (false) bullets. Needs PowerPointApi 1.4.",
    }),
  ),
  bulletStyle: Type.Optional(
    catalogEnum(POWERPOINT_BULLET_STYLES, {
      description:
        'Numbered bullet style (e.g. "ArabicNumeralPeriod"). Needs PowerPointApi 1.10.',
    }),
  ),
  bulletType: Type.Optional(
    catalogEnum(POWERPOINT_BULLET_TYPES, {
      description:
        'Bullet type: "None", "Numbered" or "Unnumbered". Needs PowerPointApi 1.10.',
    }),
  ),
  indentLevel: Type.Optional(
    Type.Number({
      description:
        "Paragraph indent level (0 = no indent). Needs PowerPointApi 1.10.",
    }),
  ),
};

/**
 * Shape line formatting (**PowerPointApi 1.4**) shared by both text tools.
 */
const PPT_SHAPE_LINE_PROPS = {
  lineColor: Type.Optional(
    Type.String({
      description:
        'Line color as #RRGGBB (e.g. "#000000"). Needs PowerPointApi 1.4.',
    }),
  ),
  lineWeight: Type.Optional(
    Type.Number({
      description: "Line weight in points. Needs PowerPointApi 1.4.",
    }),
  ),
  lineStyle: Type.Optional(
    catalogEnum(POWERPOINT_LINE_STYLES, {
      description: 'Line style. Needs PowerPointApi 1.4. Default "Single".',
    }),
  ),
  lineDashStyle: Type.Optional(
    catalogEnum(POWERPOINT_LINE_DASH_STYLES, {
      description:
        'Line dash style (e.g. "Dash", "Solid"). Needs PowerPointApi 1.4.',
    }),
  ),
  lineVisible: Type.Optional(
    Type.Boolean({
      description:
        "Show (true) or hide (false) the shape outline. Needs PowerPointApi 1.4.",
    }),
  ),
  lineTransparency: Type.Optional(
    Type.Number({
      minimum: 0,
      maximum: 1,
      description:
        "Line transparency, 0 (opaque) to 1 (clear). Needs PowerPointApi 1.4.",
    }),
  ),
};

/**
 * Slide background fill (**PowerPointApi 1.10**) — newer than shape fill, so
 * it is gated separately and only offered by `powerpoint_format_slide`.
 */
const PPT_SLIDE_BACKGROUND_PROPS = {
  slideBackgroundColor: Type.Optional(
    Type.String({
      description:
        "Solid slide background color as #RRGGBB. Needs PowerPointApi 1.10.",
    }),
  ),
  slideBackgroundTransparency: Type.Optional(
    Type.Number({
      minimum: 0,
      maximum: 1,
      description:
        "Background transparency, 0 (opaque) to 1 (clear). Needs PowerPointApi 1.10.",
    }),
  ),
  slideBackgroundClear: Type.Optional(
    Type.Boolean({
      description:
        "Reset the slide background to the layout/master default. Needs PowerPointApi 1.10.",
    }),
  ),
};

/* ── Word schemas (imported by the add-in's local Word tools) ────────── */

export const WORD_GET_OVERVIEW_PARAMETERS = Type.Object({});

export const WORD_GET_CAPABILITIES_PARAMETERS = Type.Object({});
export const EXCEL_GET_CAPABILITIES_PARAMETERS = Type.Object({});
export const POWERPOINT_GET_CAPABILITIES_PARAMETERS = Type.Object({});

export const WORD_READ_DOCUMENT_PARAMETERS = Type.Object({
  scope: Type.Optional(
    catalogEnum(["all", "selection"] as const, {
      description:
        '"all": whole document. "selection": currently selected text only.',
    }),
  ),
  maxChars: Type.Optional(
    Type.Integer({
      minimum: 100,
      maximum: 200000,
      description: "Cap on characters returned (default 20000).",
    }),
  ),
});

export const WORD_INSERT_TEXT_PARAMETERS = Type.Object({
  text: Type.String({
    description: "Text to insert. Newlines create separate paragraphs.",
  }),
  location: Type.Optional(
    catalogEnum(["start", "end", "replace_selection"] as const, {
      description:
        '"end" (default) appends to the document. "replace_selection" overwrites the selection.',
    }),
  ),
  format: Type.Optional(
    catalogEnum(["text", "markdown"] as const, {
      description:
        '"text" (default): plain text with newline→paragraph. "markdown": parse a markdown subset.',
    }),
  ),
  ...WORD_FONT_PROPS,
  ...WORD_PARAGRAPH_PROPS,
});

export const WORD_REPLACE_TEXT_PARAMETERS = Type.Object({
  find: Type.String({ description: "Literal text to find." }),
  replace: Type.String({ description: "Replacement text." }),
  matchCase: Type.Optional(
    Type.Boolean({ description: "Case-sensitive match (default false)." }),
  ),
});

export const WORD_FORMAT_RANGE_PARAMETERS = Type.Object({
  text: Type.String({
    description:
      "Literal text to find and format. Every occurrence is formatted.",
  }),
  matchCase: Type.Optional(
    Type.Boolean({ description: "Case-sensitive search (default false)." }),
  ),
  ...WORD_FONT_PROPS,
  ...WORD_PARAGRAPH_PROPS,
});

export const WORD_INSERT_BLOCKS_PARAMETERS = Type.Object({
  location: Type.Optional(
    catalogEnum(["start", "end", "replace_selection"] as const, {
      description:
        '"end" (default) appends to the document. "replace_selection" overwrites the selection.',
    }),
  ),
  blocks: Type.Array(
    Type.Object({
      text: Type.String({ description: "Text content of the block." }),
      type: Type.Optional(
        catalogEnum(
          ["paragraph", "heading", "listItem", "pageBreak"] as const,
          {
            description:
              'Block type: "paragraph" (default), "heading" (use with level), "listItem", or "pageBreak" (text is ignored).',
          },
        ),
      ),
      level: Type.Optional(
        Type.Integer({
          minimum: 1,
          maximum: 6,
          description: "Heading level 1-6 (for type: heading).",
        }),
      ),
      listType: Type.Optional(
        catalogEnum(["bullet", "number"] as const, {
          description:
            'List marker style for type: listItem ("bullet" or "number").',
        }),
      ),
      ...WORD_FONT_PROPS,
      ...WORD_PARAGRAPH_PROPS,
    }),
    { description: "Ordered list of blocks to insert." },
  ),
});

export const WORD_INSERT_TABLE_PARAMETERS = Type.Object({
  cells: Type.Array(Type.Array(Type.String()), {
    description: "2D array of cell values (rows × cols).",
  }),
  headerRow: Type.Optional(
    Type.Boolean({
      description: "Bold the first row as a header. Default: false.",
    }),
  ),
  style: Type.Optional(
    Type.String({
      description:
        'Table style name, e.g. "Grid Table 4 - Accent 1". Locale-sensitive; degraded gracefully.',
    }),
  ),
  alignment: Type.Optional(WORD_ALIGNMENT),
  location: Type.Optional(
    catalogEnum(["start", "end"] as const, {
      description:
        '"end" (default) appends to the document. "start" inserts at the beginning.',
    }),
  ),
});

export const WORD_INSERT_PAGE_BREAK_PARAMETERS = Type.Object({
  location: Type.Optional(
    catalogEnum(["start", "end"] as const, {
      description:
        '"end" (default) appends to the document. "start" inserts at the beginning.',
    }),
  ),
});

export const WORD_INSERT_IMAGE_PARAMETERS = Type.Object({
  base64: Type.String({
    description:
      "Base64-encoded image, optionally prefixed with data:<mime>;base64,",
  }),
  width: Type.Optional(
    Type.Number({ description: "Width in points. Default: keep source size." }),
  ),
  height: Type.Optional(
    Type.Number({
      description: "Height in points. Default: keep source size.",
    }),
  ),
  alignment: Type.Optional(WORD_ALIGNMENT),
  location: Type.Optional(
    catalogEnum(["start", "end"] as const, {
      description:
        '"end" (default) appends to the document. "start" inserts at the beginning.',
    }),
  ),
});

export const WORD_INSERT_HYPERLINK_PARAMETERS = Type.Object({
  text: Type.String({ description: "Display text for the hyperlink." }),
  url: Type.String({ description: "Target URL (e.g. https://example.com)." }),
  screenTip: Type.Optional(
    Type.String({ description: "Optional hover tooltip." }),
  ),
  location: Type.Optional(
    catalogEnum(["start", "end"] as const, {
      description:
        '"end" (default) appends to the document. "start" inserts at the beginning.',
    }),
  ),
});

/* ── PowerPoint schemas ──────────────────────────────────────────────── */

export const POWERPOINT_GET_OVERVIEW_PARAMETERS = Type.Object({});

export const POWERPOINT_READ_SLIDE_PARAMETERS = Type.Object({
  slideIndex: PPT_SLIDE_INDEX,
});

export const POWERPOINT_ADD_SLIDE_PARAMETERS = Type.Object({});

export const POWERPOINT_ADD_TEXT_BOX_PARAMETERS = Type.Object({
  slideIndex: PPT_SLIDE_INDEX,
  text: Type.String({ description: "Text box content." }),
  x: Type.Optional(
    Type.Number({ description: "Left edge in points (default centered)." }),
  ),
  y: Type.Optional(
    Type.Number({ description: "Top edge in points (default centered)." }),
  ),
  width: Type.Optional(
    Type.Number({ description: "Width in points (default 400)." }),
  ),
  height: Type.Optional(
    Type.Number({ description: "Height in points (default 60)." }),
  ),
  ...PPT_TEXT_FORMAT_PROPS,
  ...PPT_PARAGRAPH_PROPS,
  ...PPT_SHAPE_FILL_PROPS,
  ...PPT_SHAPE_LINE_PROPS,
});

export const POWERPOINT_FORMAT_SLIDE_PARAMETERS = Type.Object({
  slideIndex: PPT_SLIDE_INDEX,
  ...PPT_TEXT_FORMAT_PROPS,
  ...PPT_PARAGRAPH_PROPS,
  ...PPT_SHAPE_FILL_PROPS,
  ...PPT_SHAPE_LINE_PROPS,
  ...PPT_SLIDE_BACKGROUND_PROPS,
});

/* ── Excel schemas ──────────────────────────────────────────────────── */

export const EXCEL_READ_RANGE_PARAMETERS = Type.Object({
  range: Type.String({
    description:
      'Cell range in A1 notation, e.g. "A1:D10", "Sheet2!A1:B5". ' +
      "If no sheet is specified, uses the active sheet.",
  }),
  mode: Type.Optional(
    catalogEnum(["compact", "csv", "detailed"] as const, {
      description:
        '"compact" (default): markdown table of values. "csv": values only as CSV. "detailed": includes formulas and number formats.',
    }),
  ),
});

export const EXCEL_WRITE_CELLS_PARAMETERS = Type.Object({
  start_cell: Type.String({
    description:
      'Top-left cell to write from, e.g. "A1", "Sheet2!B3". If no sheet is specified, uses the active sheet.',
  }),
  values: Type.Array(Type.Array(Type.Any()), {
    description:
      '2D array of values. Each inner array is a row. Strings starting with "=" are formulas. ' +
      'Example: [["Name", "Total"], ["Alice", "=SUM(B2:B10)"]]',
  }),
  allow_overwrite: Type.Optional(
    Type.Boolean({
      description:
        "Set to true to overwrite existing data. Default: false. " +
        "If false and the target range contains values or formulas, the write is blocked " +
        "and the existing data is returned so you can ask the user.",
    }),
  ),
});

export const EXCEL_FILL_FORMULA_PARAMETERS = Type.Object({
  range: Type.String({
    description:
      'Target range to fill, e.g. "B2:B20" or "Sheet1!C3:F20". Single contiguous range only.',
  }),
  formula: Type.String({
    description:
      'Formula to fill, starting with "=", e.g. "=SUM(B2:B10)". Relative refs will adjust.',
  }),
  allow_overwrite: Type.Optional(
    Type.Boolean({
      description:
        "Set to true to overwrite existing data. Default: false. " +
        "If false and the target range contains data, the fill is blocked.",
    }),
  ),
});

export const EXCEL_SEARCH_WORKBOOK_PARAMETERS = Type.Object({
  query: Type.String({
    description:
      'Search term. For formula search, use references like "Sheet1!" to find cross-sheet links.',
  }),
  search_formulas: Type.Optional(
    Type.Boolean({
      description:
        "If true, search in formula text instead of values. " +
        'Useful for finding cross-sheet references (e.g. query "Inputs!" to find all cells referencing Inputs sheet).',
    }),
  ),
  use_regex: Type.Optional(
    Type.Boolean({
      description:
        "If true, treat the query as a regular expression (case-insensitive).",
    }),
  ),
  offset: Type.Optional(
    Type.Number({
      description: "Skip the first N matches (pagination). Default: 0.",
    }),
  ),
  sheet: Type.Optional(
    Type.String({
      description:
        "Restrict search to this sheet. If omitted, searches all sheets.",
    }),
  ),
  max_results: Type.Optional(
    Type.Number({
      description: "Maximum number of results to return. Default: 20.",
    }),
  ),
  context_rows: Type.Optional(
    Type.Number({
      description:
        "Number of rows above and below each match to include as context. " +
        "Default: 0 (no context). Use 2-5 when searching for labels to see surrounding structure.",
    }),
  ),
});

export const EXCEL_MODIFY_STRUCTURE_PARAMETERS = Type.Object({
  action: catalogEnum(
    [
      "insert_rows",
      "delete_rows",
      "insert_columns",
      "delete_columns",
      "add_sheet",
      "delete_sheet",
      "rename_sheet",
      "duplicate_sheet",
      "hide_sheet",
      "unhide_sheet",
    ] as const,
    { description: "The structural modification to perform." },
  ),
  sheet: Type.Optional(
    Type.String({
      description:
        "Target sheet name. Required for sheet operations and row/column operations on a specific sheet. " +
        "If omitted for row/column ops, uses the active sheet.",
    }),
  ),
  position: Type.Optional(
    Type.Number({
      description:
        "For insert_rows/delete_rows: the 1-indexed row number. " +
        "For insert_columns/delete_columns: the 1-indexed column number. " +
        "For add_sheet: the 0-indexed position to insert the new sheet.",
    }),
  ),
  count: Type.Optional(
    Type.Number({
      description: "Number of rows or columns to insert/delete. Default: 1.",
    }),
  ),
  new_name: Type.Optional(
    Type.String({
      description:
        "New name for rename_sheet or add_sheet. Also used for duplicate_sheet target name.",
    }),
  ),
});

/**
 * Excel underline styles (`Excel.RangeUnderlineStyle`, **ExcelApi 1.1** — no gate
 * needed; the previous boolean was a bridge limitation, not a host one).
 */
export const EXCEL_UNDERLINE_STYLES = [
  "None",
  "Single",
  "Double",
  "SingleAccountant",
  "DoubleAccountant",
] as const;

/**
 * Excel border line styles (`Excel.BorderLineStyle`, **ExcelApi 1.1**).
 *
 * Exactly these eight — the earlier audit listed a `Hair` style that does not
 * exist in the enum.
 */
export const EXCEL_BORDER_LINE_STYLES = [
  "None",
  "Continuous",
  "Dash",
  "DashDot",
  "DashDotDot",
  "Dot",
  "Double",
  "SlantDashDot",
] as const;

/** Excel horizontal alignments (`Excel.HorizontalAlignment`, ExcelApi 1.1). */
export const EXCEL_HORIZONTAL_ALIGNMENTS = [
  "General",
  "Left",
  "Center",
  "Right",
  "Fill",
  "Justify",
  "CenterAcrossSelection",
  "Distributed",
] as const;

/**
 * Excel conditional-format rule families this tool exposes.
 *
 * Every member maps to a real `Excel.ConditionalFormatType` (all **ExcelApi
 * 1.6**) — `formula` is the friendly name for `custom`. There is deliberately
 * no above-average or duplicate/unique family: those are `presetCriteria`
 * criteria, not separate types.
 */
export const EXCEL_CONDITIONAL_FORMAT_TYPES = [
  "formula",
  "cell_value",
  "data_bar",
  "color_scale",
  "icon_set",
  "top_bottom",
  "preset_criteria",
  "contains_text",
] as const;

/** Excel data-bar directions (`Excel.ConditionalDataBarDirection`, ExcelApi 1.6). */
export const EXCEL_DATA_BAR_DIRECTIONS = [
  "Context",
  "LeftToRight",
  "RightToLeft",
] as const;

/** Excel top/bottom criteria (`Excel.ConditionalTopBottomCriterionType`, ExcelApi 1.6). */
export const EXCEL_TOP_BOTTOM_CRITERIA = [
  "TopItems",
  "TopPercent",
  "BottomItems",
  "BottomPercent",
] as const;

/** Excel text-comparison operators (`Excel.ConditionalTextOperator`, ExcelApi 1.6). */
export const EXCEL_TEXT_OPERATORS = [
  "Contains",
  "NotContains",
  "BeginsWith",
  "EndsWith",
] as const;

/** Excel preset criteria (`Excel.ConditionalFormatPresetCriterion`, ExcelApi 1.6). */
export const EXCEL_PRESET_CRITERIA = [
  "Blanks",
  "NonBlanks",
  "Errors",
  "NonErrors",
  "Yesterday",
  "Today",
  "Tomorrow",
  "LastSevenDays",
  "LastWeek",
  "ThisWeek",
  "NextWeek",
  "LastMonth",
  "ThisMonth",
  "NextMonth",
  "AboveAverage",
  "BelowAverage",
  "EqualOrAboveAverage",
  "EqualOrBelowAverage",
  "OneStdDevAboveAverage",
  "OneStdDevBelowAverage",
  "TwoStdDevAboveAverage",
  "TwoStdDevBelowAverage",
  "ThreeStdDevAboveAverage",
  "ThreeStdDevBelowAverage",
  "UniqueValues",
  "DuplicateValues",
] as const;

/** Excel icon sets (`Excel.IconSet`, ExcelApi 1.6). */
export const EXCEL_ICON_SET_STYLES = [
  "ThreeArrows",
  "ThreeArrowsGray",
  "ThreeFlags",
  "ThreeTrafficLights1",
  "ThreeTrafficLights2",
  "ThreeSigns",
  "ThreeSymbols",
  "ThreeSymbols2",
  "FourArrows",
  "FourArrowsGray",
  "FourRedToBlack",
  "FourRating",
  "FourTrafficLights",
  "FiveArrows",
  "FiveArrowsGray",
  "FiveRating",
  "FiveQuarters",
  "ThreeStars",
  "ThreeTriangles",
  "FiveBoxes",
] as const;

/**
 * Excel color-scale / data-bar / icon criterion basis
 * (`Excel.ConditionalFormatColorCriterionType`, ExcelApi 1.6).
 */
export const EXCEL_COLOR_SCALE_CRITERION_TYPES = [
  "LowestValue",
  "HighestValue",
  "Number",
  "Percent",
  "Formula",
  "Percentile",
] as const;

export const EXCEL_FORMAT_CELLS_PARAMETERS = Type.Object({
  range: Type.String({
    description:
      'Range to format, e.g. "A1:D1", "Sheet2!B3:B20". Supports comma/semicolon-separated ranges on the same sheet (e.g. "A1:B2, D1:D2").',
  }),
  style: Type.Optional(
    Type.Union([Type.String(), Type.Array(Type.String())], {
      description:
        "Named style(s) to apply. Compose as array (left-to-right). " +
        'Format: "number", "integer", "currency", "percent", "ratio", "text". ' +
        'Structural: "header", "total-row", "subtotal", "input", "blank-section". ' +
        'Example: ["currency", "total-row"] = currency format + bold + top border.',
    }),
  ),
  bold: Type.Optional(Type.Boolean({ description: "Set bold." })),
  italic: Type.Optional(Type.Boolean({ description: "Set italic." })),
  underline: Type.Optional(
    Type.Union(
      [Type.Boolean(), ...EXCEL_UNDERLINE_STYLES.map((s) => Type.Literal(s))],
      {
        description:
          'Underline. Pass true = "Single", false = "None"; or name a style directly: ' +
          '"None", "Single", "Double", "SingleAccountant", "DoubleAccountant".',
      },
    ),
  ),
  strikethrough: Type.Optional(
    Type.Boolean({
      description: "Strikethrough the text. Needs ExcelApi 1.9.",
    }),
  ),
  indent_level: Type.Optional(
    Type.Number({
      description:
        "Indent level, 0-250. Needs ExcelApi 1.9; reports an error on older hosts.",
    }),
  ),
  font_color: Type.Optional(
    Type.String({ description: 'Font color as hex, e.g. "#0000FF" for blue.' }),
  ),
  font_size: Type.Optional(
    Type.Number({ description: "Font size in points." }),
  ),
  font_name: Type.Optional(
    Type.String({ description: 'Font name, e.g. "Arial", "Calibri".' }),
  ),
  fill_color: Type.Optional(
    Type.String({
      description: 'Background fill color as hex, e.g. "#FFFF00" for yellow.',
    }),
  ),
  number_format: Type.Optional(
    Type.String({
      description:
        'Preset name ("number", "integer", "currency", "percent", "ratio", "text") ' +
        "or raw Excel format string. Overrides style's number format.",
    }),
  ),
  number_format_dp: Type.Optional(
    Type.Number({
      description: "Override decimal places for a number format preset.",
    }),
  ),
  currency_symbol: Type.Optional(
    Type.String({
      description:
        'Override currency symbol, e.g. "£", "€". Only with currency preset.',
    }),
  ),
  horizontal_alignment: Type.Optional(
    catalogEnum(EXCEL_HORIZONTAL_ALIGNMENTS, {
      description:
        '"Left", "Center", "Right", "General", "Fill", "Justify", "CenterAcrossSelection", or "Distributed".',
    }),
  ),
  vertical_alignment: Type.Optional(
    Type.String({ description: '"Top", "Center", "Bottom".' }),
  ),
  wrap_text: Type.Optional(
    Type.Boolean({ description: "Enable text wrapping." }),
  ),
  column_width: Type.Optional(
    Type.Number({
      description:
        "Set column width in Excel character-width units (assumes Arial 10).",
    }),
  ),
  row_height: Type.Optional(
    Type.Number({ description: "Set row height in points." }),
  ),
  auto_fit: Type.Optional(
    Type.Boolean({
      description: "Auto-fit column widths to content. Default: false.",
    }),
  ),
  borders: Type.Optional(
    catalogEnum(["thin", "medium", "thick", "none"] as const, {
      description:
        "Border weight for ALL edges (shorthand). Individual edge params override this.",
    }),
  ),
  border_top: Type.Optional(
    catalogEnum(["thin", "medium", "thick", "none"] as const, {
      description: "Top border weight.",
    }),
  ),
  border_bottom: Type.Optional(
    catalogEnum(["thin", "medium", "thick", "none"] as const, {
      description: "Bottom border weight.",
    }),
  ),
  border_left: Type.Optional(
    catalogEnum(["thin", "medium", "thick", "none"] as const, {
      description: "Left border weight.",
    }),
  ),
  border_right: Type.Optional(
    catalogEnum(["thin", "medium", "thick", "none"] as const, {
      description: "Right border weight.",
    }),
  ),
  border_color: Type.Optional(
    Type.String({
      description:
        'Hex color for borders (e.g. "#000000"). Applies to all borders set in this call. Default: automatic (black).',
    }),
  ),
  border_style: Type.Optional(
    catalogEnum(EXCEL_BORDER_LINE_STYLES, {
      description:
        'Border line style applied to every border set in this call. Default "Continuous". ' +
        '"Dash", "DashDot", "DashDotDot", "Dot", "Double" and "SlantDashDot" give the ' +
        "full Excel line-style set (ExcelApi 1.1).",
    }),
  ),
  merge: Type.Optional(
    Type.Boolean({ description: "Merge the range into a single cell." }),
  ),
});

export const EXCEL_CONDITIONAL_FORMAT_PARAMETERS = Type.Object({
  action: Type.Union([Type.Literal("add"), Type.Literal("clear")], {
    description:
      '"add" to create a rule, "clear" to remove all rules in the range.',
  }),
  range: Type.String({
    description: 'Target range, e.g. "A1:D10" or "Sheet2!B2:B50".',
  }),
  type: Type.Optional(
    catalogEnum(EXCEL_CONDITIONAL_FORMAT_TYPES, {
      description:
        'Rule family for "add" (all ExcelApi 1.6). "formula" = custom formula, ' +
        '"cell_value" = comparison, "data_bar", "color_scale", "icon_set", ' +
        '"top_bottom", "preset_criteria" (blanks/average/dates/unique), ' +
        '"contains_text".',
    }),
  ),
  formula: Type.Optional(
    Type.String({
      description: 'Custom formula for type="formula", e.g. "=A1>0".',
    }),
  ),
  operator: Type.Optional(
    catalogEnum(
      [
        "Between",
        "NotBetween",
        "EqualTo",
        "NotEqualTo",
        "GreaterThan",
        "LessThan",
        "GreaterThanOrEqual",
        "LessThanOrEqual",
      ] as const,
      { description: "Cell value operator (required for cell_value rules)." },
    ),
  ),
  value: Type.Optional(
    Type.Union([Type.String(), Type.Number()], {
      description:
        'Cell value comparison target (required for cell_value rules). Use numbers or formulas like "=$B$2".',
    }),
  ),
  value2: Type.Optional(
    Type.Union([Type.String(), Type.Number()], {
      description: "Second value for Between/NotBetween operators (optional).",
    }),
  ),
  text_operator: Type.Optional(
    catalogEnum(EXCEL_TEXT_OPERATORS, {
      description:
        'Text comparison for type="contains_text". Default "Contains".',
    }),
  ),
  text: Type.Optional(
    Type.String({
      description: 'Text to match for type="contains_text".',
    }),
  ),
  top_bottom_type: Type.Optional(
    catalogEnum(EXCEL_TOP_BOTTOM_CRITERIA, {
      description: 'Rank basis for type="top_bottom". Default "TopItems".',
    }),
  ),
  top_bottom_rank: Type.Optional(
    Type.Number({
      description:
        'Rank for type="top_bottom": 1-1000 for item ranks, 1-100 for percent ranks. Default 10.',
    }),
  ),
  preset_criterion: Type.Optional(
    catalogEnum(EXCEL_PRESET_CRITERIA, {
      description:
        'Preset criterion for type="preset_criteria", e.g. "AboveAverage", "Blanks", "Today", "DuplicateValues".',
    }),
  ),
  data_bar_fill_color: Type.Optional(
    Type.String({
      description: 'Positive data-bar fill hex, e.g. "#638EC6".',
    }),
  ),
  data_bar_negative_fill_color: Type.Optional(
    Type.String({ description: "Negative data-bar fill hex." }),
  ),
  data_bar_direction: Type.Optional(
    catalogEnum(EXCEL_DATA_BAR_DIRECTIONS, {
      description: 'Data-bar direction. Default "Context".',
    }),
  ),
  data_bar_show_only: Type.Optional(
    Type.Boolean({
      description: "Show the data bar only, hiding the cell value.",
    }),
  ),
  data_bar_axis_color: Type.Optional(
    Type.String({ description: "Data-bar axis color hex." }),
  ),
  color_scale_min_color: Type.Optional(
    Type.String({
      description: 'Minimum color for type="color_scale". Default "#F8696B".',
    }),
  ),
  color_scale_mid_color: Type.Optional(
    Type.String({
      description:
        'Midpoint color for type="color_scale". Omit for a two-color scale.',
    }),
  ),
  color_scale_max_color: Type.Optional(
    Type.String({
      description: 'Maximum color for type="color_scale". Default "#63BE7B".',
    }),
  ),
  icon_set_style: Type.Optional(
    catalogEnum(EXCEL_ICON_SET_STYLES, {
      description:
        'Icon set for type="icon_set". Default "ThreeTrafficLights1".',
    }),
  ),
  icon_set_reverse_order: Type.Optional(
    Type.Boolean({ description: "Reverse the icon order." }),
  ),
  icon_set_show_icon_only: Type.Optional(
    Type.Boolean({
      description: "Show only the icon, hiding the cell value.",
    }),
  ),
  fill_color: Type.Optional(
    Type.String({ description: 'Fill color hex, e.g. "#FFFDE0".' }),
  ),
  font_color: Type.Optional(
    Type.String({ description: 'Font color hex, e.g. "#000000".' }),
  ),
  bold: Type.Optional(Type.Boolean({ description: "Bold text." })),
  italic: Type.Optional(Type.Boolean({ description: "Italic text." })),
  underline: Type.Optional(Type.Boolean({ description: "Underline text." })),
  stop_if_true: Type.Optional(
    Type.Boolean({ description: "Stop evaluating later rules if true." }),
  ),
});

export const EXCEL_CHARTS_PARAMETERS = Type.Object({
  action: catalogEnum(
    ["list", "create", "update", "delete", "get_image"] as const,
    {
      description:
        "Chart operation: list, create, update, delete, or get_image.",
    },
  ),
  sheet: Type.Optional(
    Type.String({
      description:
        "Worksheet name. For list, limits output to one sheet. For create/update, defaults to the active sheet or the source range sheet.",
    }),
  ),
  name: Type.Optional(
    Type.String({
      description:
        "Chart name. Required for update, delete, and get_image. Optional assigned name for create.",
    }),
  ),
  new_name: Type.Optional(
    Type.String({ description: "New chart name for update." }),
  ),
  source_range: Type.Optional(
    Type.String({
      description:
        "Source data range for create/update, e.g. `Sheet1!A1:B12` or `A1:B12` relative to sheet.",
    }),
  ),
  chart_type: Type.Optional(
    catalogEnum(
      [
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
      ] as const,
      {
        description:
          "Chart type: column, column_stacked, column_stacked_100, bar, bar_stacked, " +
          "bar_stacked_100, line, line_markers, area, area_stacked, pie, doughnut, " +
          "scatter, scatter_lines, scatter_smooth, or radar.",
      },
    ),
  ),
  series_by: Type.Optional(
    catalogEnum(["auto", "columns", "rows"] as const, {
      description:
        "How source rows/columns become series: auto, columns, or rows.",
    }),
  ),
  title: Type.Optional(
    Type.String({ description: "Chart title. Empty string hides the title." }),
  ),
  legend_position: Type.Optional(
    catalogEnum(["none", "right", "left", "top", "bottom"] as const, {
      description: "Legend position, or none to hide the legend.",
    }),
  ),
  x_axis_title: Type.Optional(
    Type.String({
      description: "Category/X axis title. Empty string hides it.",
    }),
  ),
  y_axis_title: Type.Optional(
    Type.String({ description: "Value/Y axis title. Empty string hides it." }),
  ),
  position: Type.Optional(
    Type.String({
      description: "Anchor range for chart placement, e.g. `D2:J18`.",
    }),
  ),
  width: Type.Optional(
    Type.Number({
      description:
        "Image width in pixels for get_image. Defaults to 600; capped at 1200.",
    }),
  ),
});

export const EXCEL_TRACE_DEPENDENCIES_PARAMETERS = Type.Object({
  cell: Type.String({
    description:
      'Cell to trace, e.g. "D10", "Sheet2!F5". Must be a single cell, not a range.',
  }),
  mode: Type.Optional(
    catalogEnum(["precedents", "dependents"] as const, {
      description:
        "Trace direction: precedents (upstream) or dependents (downstream). Default: precedents.",
    }),
  ),
  depth: Type.Optional(
    Type.Number({
      description:
        "How many levels of dependencies to trace. Default: 2. Max: 5.",
    }),
  ),
});

export const EXCEL_EXPLAIN_FORMULA_PARAMETERS = Type.Object({
  cell: Type.String({
    description: 'Single formula cell to explain, e.g. "D10" or "Sheet2!F5".',
  }),
  max_references: Type.Optional(
    Type.Number({
      description:
        "Max number of direct references to preview. Default: 8. Max: 20.",
    }),
  ),
});

export const EXCEL_VIEW_SETTINGS_PARAMETERS = Type.Object({
  action: catalogEnum(
    [
      "get",
      "show_gridlines",
      "hide_gridlines",
      "show_headings",
      "hide_headings",
      "freeze_rows",
      "freeze_columns",
      "freeze_at",
      "unfreeze",
      "set_tab_color",
      "hide_sheet",
      "show_sheet",
      "very_hide_sheet",
      "set_standard_width",
      "activate",
    ] as const,
    { description: "The view setting to read or change." },
  ),
  sheet: Type.Optional(
    Type.String({
      description:
        "Target sheet name. Defaults to the active sheet for most actions. " +
        "Required for hide/show/very_hide and activate.",
    }),
  ),
  count: Type.Optional(
    Type.Number({
      description:
        "Number of rows or columns to freeze. Required for freeze_rows/freeze_columns.",
    }),
  ),
  range: Type.Optional(
    Type.String({
      description:
        'Cell range for freeze_at (e.g. "B3"). Everything above and to the left of this cell will be frozen.',
    }),
  ),
  color: Type.Optional(
    Type.String({
      description:
        'Tab color in #RRGGBB format (e.g. "#FF6600"). Use "" to clear.',
    }),
  ),
  width: Type.Optional(
    Type.Number({
      description:
        "Standard (default) column width for the worksheet, in Excel character-width units. " +
        "Required for set_standard_width.",
    }),
  ),
});

export const EXCEL_COMMENTS_PARAMETERS = Type.Object({
  action: catalogEnum(
    ["read", "add", "update", "reply", "delete", "resolve", "reopen"] as const,
    {
      description:
        "Comment operation: read (list comments in range), add (new comment on cell), " +
        "update (edit existing comment text), reply (add threaded reply), " +
        "delete (remove comment + replies), resolve/reopen (toggle thread status).",
    },
  ),
  range: Type.String({
    description:
      'Target cell or range in A1 notation, e.g. "A1", "B2:D10", "Sheet2!A1". ' +
      "Range supported for read; other actions require a single cell.",
  }),
  content: Type.Optional(
    Type.String({
      description: "Comment text. Required for add, update, and reply actions.",
    }),
  ),
});

export const EXCEL_WORKBOOK_HISTORY_PARAMETERS = Type.Object({
  action: Type.Optional(
    catalogEnum(["list", "restore", "delete", "clear"] as const, {
      description:
        "Operation to run. list (default): show recent backups; " +
        "restore: revert one backup; delete: remove one backup; clear: remove all backups for current workbook.",
    }),
  ),
  snapshot_id: Type.Optional(
    Type.String({
      description:
        "Backup id for restore/delete. If omitted, the latest backup is used.",
    }),
  ),
  limit: Type.Optional(
    Type.Integer({
      minimum: 1,
      maximum: 50,
      description: "Max backups to list (list action only). Default: 10.",
    }),
  ),
});

export const EXCEL_GET_OVERVIEW_PARAMETERS = Type.Object({
  sheet: Type.Optional(
    Type.String({
      description:
        "If provided, return detailed info for this specific sheet " +
        "(dimensions, headers, tables, named ranges, objects, and a data preview). " +
        "If omitted, return the workbook-level overview.",
    }),
  ),
});

/**
 * Capability-calibration rule appended to the Word formatting descriptions.
 *
 * These descriptions previously ended with "Formatting is fully supported —
 * never tell the user it is not". That wording stopped the model *refusing* to
 * format, but it also licensed it to invent capabilities — e.g. answering that
 * Word has no text highlight when `font.highlightColor` is WordApi 1.1. Naming
 * the exposed properties plus this rule prevents both failures.
 */
const WORD_FORMAT_CALIBRATION =
  "Only the formatting properties listed above are exposed — if the user asks " +
  "for one that is not listed, say it is not available rather than claiming support.";

/* ── Catalog entries ────────────────────────────────────────────────── */

export const OFFICE_CATALOG: readonly OfficeCatalogEntry[] = [
  /* ── Excel ─────────────────────────────────────────────────────── */
  {
    host: "excel",
    op: "get_overview",
    name: officeToolName("excel", "get_overview"),
    label: "Excel Workbook Overview",
    description:
      "Read a compact overview of the attached Excel workbook: sheet names, used ranges, " +
      "table names, and named ranges. Call this first before any range operation.",
    promptSnippet: "Outline the attached Excel workbook",
    promptGuidelines: [
      "Call office_excel_get_overview before office_excel_read_range to learn the workbook structure.",
    ],
    parameters: EXCEL_GET_OVERVIEW_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "excel",
    op: "read_range",
    name: officeToolName("excel", "read_range"),
    label: "Excel Read Range",
    description:
      "Read cell values (and optionally formulas/formatting) from a range in the attached Excel workbook.",
    promptSnippet: "Read cells from the attached Excel workbook",
    parameters: EXCEL_READ_RANGE_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "excel",
    op: "write_cells",
    name: officeToolName("excel", "write_cells"),
    label: "Excel Write Cells",
    description:
      "Write a 2D array of values into the attached Excel workbook, starting at a top-left cell. " +
      "values[row][col]; the array is written down and to the right from start_cell.",
    promptSnippet: "Write values/formulas into the attached Excel workbook",
    promptGuidelines: [
      "Prefer office_excel_write_cells in a single batched call instead of many small edits.",
      "Always verify with office_excel_read_range after office_excel_write_cells when the change is user-visible.",
    ],
    parameters: EXCEL_WRITE_CELLS_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "excel",
    op: "fill_formula",
    name: officeToolName("excel", "fill_formula"),
    label: "Excel Fill Formula",
    description:
      "Write a formula into a single contiguous range of the attached Excel workbook. " +
      "Relative references adjust as the formula fills.",
    promptSnippet: "Fill a formula across an Excel range",
    parameters: EXCEL_FILL_FORMULA_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "excel",
    op: "search_workbook",
    name: officeToolName("excel", "search_workbook"),
    label: "Excel Search Workbook",
    description:
      "Search for text, values, or formulas across the attached Excel workbook. " +
      "Returns matching cells with sheet name, address, value, and formula. " +
      "Use to find data, locate cells by label, or trace cross-sheet references.",
    promptSnippet:
      "Search for text/values/formulas in the attached Excel workbook",
    parameters: EXCEL_SEARCH_WORKBOOK_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "modify_structure",
    name: officeToolName("excel", "modify_structure"),
    label: "Excel Modify Structure",
    description:
      "Modify the workbook structure of the attached Excel workbook: insert/delete rows and columns, " +
      "add/delete/rename/duplicate/hide/unhide sheets. Be careful with deletions — there is no undo.",
    promptSnippet:
      "Insert/delete rows, columns, or sheets in the attached Excel workbook",
    parameters: EXCEL_MODIFY_STRUCTURE_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "format_cells",
    name: officeToolName("excel", "format_cells"),
    label: "Excel Format Cells",
    description:
      "Apply formatting to a range of cells in the attached Excel workbook (supports comma-separated ranges on one sheet). " +
      'Use named styles for common patterns: style: "currency" or style: ["currency", "total-row"]. ' +
      "Individual params (bold, fill_color, etc.) override style properties. " +
      "Does NOT modify cell values — use office_excel_write_cells for that.",
    promptSnippet:
      "Format cells (font, fill, borders, alignment, number format) in the attached Excel workbook",
    parameters: EXCEL_FORMAT_CELLS_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "conditional_format",
    name: officeToolName("excel", "conditional_format"),
    label: "Excel Conditional Format",
    description:
      "Add or clear conditional formatting rules in the attached Excel workbook. " +
      "Supports custom formula and cell value rules.",
    promptSnippet:
      "Add or clear conditional formatting rules in the attached Excel workbook",
    parameters: EXCEL_CONDITIONAL_FORMAT_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "charts",
    name: officeToolName("excel", "charts"),
    label: "Excel Charts",
    description:
      "List, create, update, delete, and capture images of charts in the attached Excel workbook.",
    promptSnippet:
      "List/create/update/delete charts in the attached Excel workbook",
    parameters: EXCEL_CHARTS_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "trace_dependencies",
    name: officeToolName("excel", "trace_dependencies"),
    label: "Excel Trace Dependencies",
    description:
      "Return formula lineage for a cell in the attached Excel workbook: precedents (upstream) or dependents (downstream).",
    promptSnippet: "Trace formula precedents/dependents of an Excel cell",
    parameters: EXCEL_TRACE_DEPENDENCIES_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "explain_formula",
    name: officeToolName("excel", "explain_formula"),
    label: "Excel Explain Formula",
    description:
      "Explain a formula cell in the attached Excel workbook in plain language, including direct input references and current values.",
    promptSnippet: "Explain what an Excel formula cell does",
    parameters: EXCEL_EXPLAIN_FORMULA_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "view_settings",
    name: officeToolName("excel", "view_settings"),
    label: "Excel View Settings",
    description:
      "Read or change worksheet view/navigation settings in the attached Excel workbook: gridlines, row/column headings, " +
      "freeze panes, tab color, sheet visibility, sheet activation, and standard width.",
    promptSnippet:
      "Change Excel view settings (gridlines, freeze panes, tab color, visibility)",
    parameters: EXCEL_VIEW_SETTINGS_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "comments",
    name: officeToolName("excel", "comments"),
    label: "Excel Comments",
    description:
      "Read, add, update, reply, delete, resolve, and reopen cell comments in the attached Excel workbook.",
    promptSnippet: "Read or manage Excel cell comments and reply threads",
    parameters: EXCEL_COMMENTS_PARAMETERS,
    since: 1,
  },
  {
    host: "excel",
    op: "workbook_history",
    name: officeToolName("excel", "workbook_history"),
    label: "Excel Workbook History",
    description:
      "List, restore, and manage automatic workbook backups created before edits in the attached Excel workbook.",
    promptSnippet:
      "List/restore automatic backups of the attached Excel workbook",
    parameters: EXCEL_WORKBOOK_HISTORY_PARAMETERS,
    since: 1,
  },

  {
    host: "excel",
    op: "get_capabilities",
    name: officeToolName("excel", "get_capabilities"),
    label: "Excel Capabilities",
    description:
      "Report which Excel capabilities the current host actually supports: the platform, and " +
      "whether each formatting/structure capability is available or gated by a newer ExcelApi " +
      "version. Read-only — it never changes the workbook. Call this before telling the user a " +
      "capability is unavailable, so the answer reflects the real host rather than an assumption.",
    promptSnippet: "Report which Excel capabilities the current host supports",
    promptGuidelines: [
      "Call office_excel_get_capabilities before claiming an Excel feature is unavailable.",
    ],
    parameters: EXCEL_GET_CAPABILITIES_PARAMETERS,
    since: 3,
  },

  /* ── Word ──────────────────────────────────────────────────────── */
  {
    host: "word",
    op: "get_overview",
    name: officeToolName("word", "get_overview"),
    label: "Word Document Overview",
    description:
      "Read a compact overview of the attached Word document: heading outline (with text), paragraph count, " +
      "table count, and word count. Call this first before editing.",
    promptSnippet: "Outline the attached Word document",
    promptGuidelines: [
      "Call office_word_get_overview before office_word_insert_text or office_word_replace_text.",
    ],
    parameters: WORD_GET_OVERVIEW_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "word",
    op: "get_capabilities",
    name: officeToolName("word", "get_capabilities"),
    label: "Word Capabilities",
    description:
      "Report which Word capabilities the current host actually supports: the platform, and " +
      "whether each formatting/insertion capability is available or desktop-only. Read-only — " +
      "it never changes the document. Call this before telling the user that a capability is " +
      "unavailable, so the answer reflects the real host rather than an assumption.",
    promptSnippet: "Report which Word capabilities the current host supports",
    promptGuidelines: [
      "Call office_word_get_capabilities before claiming a Word formatting feature is unavailable.",
    ],
    parameters: WORD_GET_CAPABILITIES_PARAMETERS,
    since: 2,
  },
  {
    host: "word",
    op: "read_document",
    name: officeToolName("word", "read_document"),
    label: "Word Read Document",
    description:
      "Read text from the attached Word document: the whole body or the current selection. " +
      "Paragraphs are preserved (one line per paragraph) with heading/list markers.",
    promptSnippet: "Read the attached Word document (or selection)",
    parameters: WORD_READ_DOCUMENT_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "word",
    op: "insert_text",
    name: officeToolName("word", "insert_text"),
    label: "Word Insert Text",
    description:
      "Insert text into the attached Word document — at the start, at the end (default), " +
      "or replacing the current selection. Multi-line text becomes separate paragraphs. " +
      "Optionally format the inserted text in the same call: font (bold, italic, underline, size, name, color), " +
      'paragraph alignment ("Left"/"Centered"/"Right"/"Justified"), style names (e.g. "Heading 1", "Title"), ' +
      'spacing, and indents. Use format: "markdown" to insert a markdown document (headings, bold/italic, bullets). ' +
      WORD_FORMAT_CALIBRATION,
    promptSnippet: "Insert (and format) text into the attached Word document",
    promptGuidelines: [
      'For a formatted document use office_word_insert_blocks (precise per-block control) or office_word_insert_text with format: "markdown".',
      "Never emit HTML for Word documents; use office_word_insert_text / office_word_insert_blocks instead.",
    ],
    parameters: WORD_INSERT_TEXT_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "word",
    op: "replace_text",
    name: officeToolName("word", "replace_text"),
    label: "Word Replace Text",
    description:
      "Find and replace literal text in the attached Word document. Returns how many occurrences were replaced.",
    promptSnippet:
      "Find and replace literal text in the attached Word document",
    parameters: WORD_REPLACE_TEXT_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "word",
    op: "format_range",
    name: officeToolName("word", "format_range"),
    label: "Word Format Range",
    description:
      "Find text by literal content in the attached Word document and apply formatting to every match: " +
      "bold, italic, underline, font size (points), font name, color (#RRGGBB), paragraph alignment, " +
      'paragraph style (e.g. "Heading 1"), spacing, and indents. ' +
      "Use this to bold/size/center existing content — e.g. format a document title. " +
      WORD_FORMAT_CALIBRATION,
    promptSnippet: "Format existing Word text (bold, size, alignment, style)",
    parameters: WORD_FORMAT_RANGE_PARAMETERS,
    since: 1,
  },
  {
    host: "word",
    op: "insert_blocks",
    name: officeToolName("word", "insert_blocks"),
    label: "Word Insert Blocks",
    description:
      "Assemble a structured formatted document in the attached Word document from blocks: " +
      "paragraphs, headings (level 1-6), bullet/numbered list items, and page breaks. " +
      "Each block can carry its own font (bold/italic/underline/size/name/color), " +
      "paragraph alignment, style name, spacing and indents. Use this for a complete formatted document " +
      "in one call — e.g. a title, then body paragraphs, then a right-aligned signature block.",
    promptSnippet:
      "Insert a fully formatted document (title, headings, lists, alignment) into Word",
    promptGuidelines: [
      "Prefer office_word_insert_blocks over many small inserts when composing a formatted document.",
      'Align office_word_insert_blocks signatures/dates right (alignment: "Right") and keep heading levels ≤ 6.',
      "Never emit HTML for Word documents; use office_word_insert_blocks / office_word_insert_text instead.",
    ],
    parameters: WORD_INSERT_BLOCKS_PARAMETERS,
    since: 1,
  },
  {
    host: "word",
    op: "insert_table",
    name: officeToolName("word", "insert_table"),
    label: "Word Insert Table",
    description:
      "Insert a table into the attached Word document with values[row][col]. " +
      "Optionally bold the header row, apply a table style/alignment.",
    promptSnippet: "Insert a table into the attached Word document",
    parameters: WORD_INSERT_TABLE_PARAMETERS,
    since: 1,
  },
  {
    host: "word",
    op: "insert_page_break",
    name: officeToolName("word", "insert_page_break"),
    label: "Word Insert Page Break",
    description:
      "Insert a page break at the start or end of the attached Word document.",
    promptSnippet: "Insert a page break in the attached Word document",
    parameters: WORD_INSERT_PAGE_BREAK_PARAMETERS,
    since: 1,
  },
  {
    host: "word",
    op: "insert_image",
    name: officeToolName("word", "insert_image"),
    label: "Word Insert Image",
    description:
      "Insert an inline image into the attached Word document from a base64 string (or a data URL). " +
      "Max decoded size ~5 MB. To embed a local file, encode it first: on macOS/Linux run `base64 -w0 <file>` " +
      "in bash and pass the output; on Windows run `certutil -encode <file> tmp.b64` and read the file.",
    promptSnippet: "Insert an image (base64) into the attached Word document",
    parameters: WORD_INSERT_IMAGE_PARAMETERS,
    since: 1,
  },
  {
    host: "word",
    op: "insert_hyperlink",
    name: officeToolName("word", "insert_hyperlink"),
    label: "Word Insert Hyperlink",
    description:
      "Insert a clickable hyperlink with display text into the attached Word document.",
    promptSnippet: "Insert a hyperlink into the attached Word document",
    parameters: WORD_INSERT_HYPERLINK_PARAMETERS,
    since: 1,
  },

  /* ── PowerPoint ─────────────────────────────────────────────────── */
  {
    host: "powerpoint",
    op: "get_overview",
    name: officeToolName("powerpoint", "get_overview"),
    label: "PowerPoint Overview",
    description:
      "Read a compact overview of the attached presentation: slide count, each slide's title and " +
      "shape count. Call this first before any slide operation.",
    promptSnippet: "Outline the attached PowerPoint presentation",
    promptGuidelines: [
      "Call office_powerpoint_get_overview before office_powerpoint_read_slide or office_powerpoint_add_slide.",
    ],
    parameters: POWERPOINT_GET_OVERVIEW_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "powerpoint",
    op: "read_slide",
    name: officeToolName("powerpoint", "read_slide"),
    label: "PowerPoint Read Slide",
    description:
      "Read all text content of one slide in the attached presentation (shapes, text frames, notes).",
    promptSnippet: "Read the text of a PowerPoint slide",
    parameters: POWERPOINT_READ_SLIDE_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "powerpoint",
    op: "add_slide",
    name: officeToolName("powerpoint", "add_slide"),
    label: "PowerPoint Add Slide",
    description:
      "Append a new slide to the attached presentation and navigate to it. Uses the default layout.",
    promptSnippet: "Append a slide to the attached PowerPoint presentation",
    parameters: POWERPOINT_ADD_SLIDE_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "powerpoint",
    op: "add_text_box",
    name: officeToolName("powerpoint", "add_text_box"),
    label: "PowerPoint Add Text Box",
    description:
      "Add a text box with the given text to a slide. Coordinates/geometry are in points. " +
      "Supports formatting in the same call (bold, italic, underline, fontSize/fontName/fontColor, " +
      'alignment "Left"/"Center"/"Right"/"Justify") — always pass formatting for titles and headings.',
    promptSnippet: "Add a (formatted) text box to a PowerPoint slide",
    parameters: POWERPOINT_ADD_TEXT_BOX_PARAMETERS,
    since: 1,
    legacyV1: true,
  },
  {
    host: "powerpoint",
    op: "format_slide",
    name: officeToolName("powerpoint", "format_slide"),
    label: "PowerPoint Format Slide",
    description:
      "Apply formatting (bold, italic, underline, fontSize in points, fontName, fontColor #RRGGBB, " +
      'alignment "Left"/"Center"/"Right"/"Justify") to every text box on a slide of the attached ' +
      "PowerPoint presentation. Use this to restyle existing content, e.g. make all text on a slide bold and centered.",
    promptSnippet: "Restyle all text boxes on a PowerPoint slide",
    parameters: POWERPOINT_FORMAT_SLIDE_PARAMETERS,
    since: 1,
  },
  {
    host: "powerpoint",
    op: "get_capabilities",
    name: officeToolName("powerpoint", "get_capabilities"),
    label: "PowerPoint Capabilities",
    description:
      "Report which PowerPoint capabilities the current host actually supports: the platform, and " +
      "whether each formatting capability is available or gated by a newer PowerPointApi version. " +
      "Read-only — it never changes the presentation. Call this before telling the user a capability " +
      "is unavailable, so the answer reflects the real host rather than an assumption.",
    promptSnippet:
      "Report which PowerPoint capabilities the current host supports",
    promptGuidelines: [
      "Call office_powerpoint_get_capabilities before claiming a PowerPoint feature is unavailable.",
    ],
    parameters: POWERPOINT_GET_CAPABILITIES_PARAMETERS,
    since: 3,
  },
];

/** Index by "<host>.<op>" for fast lookup. */
export const OFFICE_CATALOG_BY_OP: ReadonlyMap<string, OfficeCatalogEntry> =
  new Map(OFFICE_CATALOG.map((entry) => [`${entry.host}.${entry.op}`, entry]));

/**
 * Canonical tool descriptions keyed by "<host>.<op>" — the single source of
 * truth shared by the bridge catalog entries above and the add-in's local
 * (browser-only) tool factories.
 *
 * Local factories MUST NOT re-declare these strings. Duplicated, hand-maintained
 * descriptions are what allowed a tool description to drift from — and
 * overstate — what the implementation actually did.
 */
export const OFFICE_TOOL_DESCRIPTIONS: Readonly<Record<string, string>> =
  Object.fromEntries(
    OFFICE_CATALOG.map((entry) => [
      `${entry.host}.${entry.op}`,
      entry.description,
    ]),
  );

/**
 * Resolve a `<host>.<op>` id to its canonical tool description. Throws for
 * unknown ids so local tools import the single source of truth and drift is
 * impossible (mirrors `catalogSchemaFor`).
 */
export function officeToolDescription(opId: string): string {
  const description: string | undefined = OFFICE_TOOL_DESCRIPTIONS[opId];
  if (description === undefined) {
    throw new Error(`office catalog: no description for op "${opId}"`);
  }
  return description;
}

/** All Pi tool names in the catalog ("office_<host>_<op>"). */
export const OFFICE_TOOL_NAMES: readonly string[] = OFFICE_CATALOG.map(
  (entry) => entry.name,
);

/** Entries for one host. */
export function catalogForHost(
  host: OfficeHostApp,
): readonly OfficeCatalogEntry[] {
  return OFFICE_CATALOG.filter((entry) => entry.host === host);
}

/** The office host this tool name drives, or null when unknown. */
export function hostForToolName(name: string): OfficeHostApp | null {
  const entry = OFFICE_CATALOG.find((candidate) => candidate.name === name);
  return entry?.host ?? null;
}

/**
 * The op set exposed by the legacy 0.2.x bridge. Panes that do not advertise
 * `ops` are treated as legacy and only these ops are activated for them.
 */
export const LEGACY_V1_OPS: readonly string[] = OFFICE_CATALOG.flatMap(
  (entry) => (entry.legacyV1 === true ? [`${entry.host}.${entry.op}`] : []),
);

/** Op ids for one host — i.e. `<host>.<op>` strings. */
export function catalogOpIdsForHost(host: OfficeHostApp): readonly string[] {
  return catalogForHost(host).map((entry) => `${entry.host}.${entry.op}`);
}

/**
 * Resolve a `<host>.<op>` id to its shared parameter schema. Throws for
 * unknown ids so local tools import the single source of truth and drift is
 * impossible.
 */
export function catalogSchemaFor(opId: string): TSchema {
  const entry = OFFICE_CATALOG_BY_OP.get(opId);
  if (!entry) {
    throw new Error(`office catalog: no schema for op "${opId}"`);
  }
  return entry.parameters;
}
