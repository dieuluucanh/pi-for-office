/**
 * Office.js wrappers — thin abstraction over PowerPoint.run().
 *
 * Mirrors `excel/helpers.ts` (excelRun) so local PowerPoint tools follow
 * the same execution pattern as the Excel core tools.
 *
 * Requirement sets (verified against the installed `@types/office-js`):
 * - shape access `PowerPointApi 1.3`
 * - text frame / font / paragraph format `1.4`
 * - shape fill and line `1.4`
 * - bullet *visibility* `1.4`; bullet *style* and paragraph `indentLevel` `1.10`
 * - font effects (strikethrough, small caps, …) `1.8`
 * - slide background fill `1.10`
 *
 * Callers gate with `requireCapability(...)`; these helpers only apply values.
 */

import {
 POWERPOINT_BULLET_STYLES,
 POWERPOINT_BULLET_TYPES,
 POWERPOINT_LINE_DASH_STYLES,
 POWERPOINT_LINE_STYLES,
 POWERPOINT_UNDERLINE_STYLES,
} from "@dieulc/pi-office-protocol/office-catalog";
import {
 capabilityUnsupported,
 requireCapability,
} from "../office/capabilities.js";

export type PowerPointLineStyle = (typeof POWERPOINT_LINE_STYLES)[number];
export type PowerPointLineDashStyle =
 (typeof POWERPOINT_LINE_DASH_STYLES)[number];
export type PowerPointBulletStyle = (typeof POWERPOINT_BULLET_STYLES)[number];
export type PowerPointBulletType = (typeof POWERPOINT_BULLET_TYPES)[number];
export type PowerPointUnderlineStyle =
 (typeof POWERPOINT_UNDERLINE_STYLES)[number];

/**
 * Run an Office.js PowerPoint operation with a typed context.
 * Wraps PowerPoint.run() (proxy objects → load() → context.sync()).
 */
export async function powerPointRun<T>(
 fn: (context: PowerPoint.RequestContext) => Promise<T>,
): Promise<T> {
 return PowerPoint.run(fn);
}

// ============================================================================
// Text formatting
// ============================================================================

/**
 * Underline styles (`PowerPoint.ShapeFontUnderlineStyle`, PowerPointApi 1.4).
 * Re-exported from the catalog so the schema and the implementation share one
 * definition.
 */
export { POWERPOINT_UNDERLINE_STYLES };

/**
 * Formatting options shared by the PowerPoint tools (all optional).
 *
 * Grouping (and therefore the gate each property needs) is documented per
 * field so a caller never applies a 1.10 property on a 1.4 host.
 */
export interface PowerPointTextFormat {
 bold?: boolean;
 italic?: boolean;
 /** `true`/`false` remain the legacy alias for "Single"/"None" (1.4). */
 underline?: boolean | PowerPointUnderlineStyle;
 /** Font size in points. */
 fontSize?: number;
 /** Font name, e.g. "Arial". */
 fontName?: string;
 /** Font color as #RRGGBB. */
 fontColor?: string;
 /** Horizontal alignment. PPT values: "Left" | "Center" | "Right" | "Justify". */
 alignment?: "Left" | "Center" | "Right" | "Justify";
 /** Bullet visibility (PowerPointApi 1.4). */
 bulletVisible?: boolean;
 /** Numbered bullet style (PowerPointApi 1.10). */
 bulletStyle?: PowerPointBulletStyle;
 /** Bullet type (PowerPointApi 1.10). */
 bulletType?: PowerPointBulletType;
 /** Paragraph indent level (PowerPointApi 1.10). */
 indentLevel?: number;
 /** Strikethrough (PowerPointApi 1.8). */
 strikethrough?: boolean;
 /** Double strikethrough (PowerPointApi 1.8). */
 doubleStrikethrough?: boolean;
 /** Small caps (PowerPointApi 1.8). */
 smallCaps?: boolean;
 /** All caps (PowerPointApi 1.8). */
 allCaps?: boolean;
 /** Subscript (PowerPointApi 1.8). */
 subscript?: boolean;
 /** Superscript (PowerPointApi 1.8). */
 superscript?: boolean;
}

/** Resolve the `underline` alias to an underline style name, or undefined. */
export function resolvePowerPointUnderline(
 value: boolean | PowerPointUnderlineStyle | undefined,
): PowerPointUnderlineStyle | undefined {
 if (value === undefined) return undefined;
 if (value === true) return "Single";
 if (value === false) return "None";
 return POWERPOINT_UNDERLINE_STYLES.find((style) => style === value);
}

/**
 * Every text-level property `setPowerPointTextFormat` understands.
 *
 * This list is the single source of truth for "is there any text formatting
 * requested?" — the tool schemas advertise exactly these keys, and the
 * parity guard checks that each one has an assignment path below.
 */
export const POWERPOINT_TEXT_FORMAT_KEYS = [
 "bold",
 "italic",
 "underline",
 "fontSize",
 "fontName",
 "fontColor",
 "alignment",
 "bulletVisible",
 "bulletStyle",
 "bulletType",
 "indentLevel",
 "strikethrough",
 "doubleStrikethrough",
 "smallCaps",
 "allCaps",
 "subscript",
 "superscript",
] as const satisfies ReadonlyArray<keyof PowerPointTextFormat>;

/** Core font/alignment properties — PowerPointApi 1.4. */
const POWERPOINT_FONT_KEYS = [
 "bold",
 "italic",
 "fontSize",
 "fontName",
 "fontColor",
 "alignment",
] as const satisfies ReadonlyArray<keyof PowerPointTextFormat>;

/** Underline (any form, including the legacy boolean) — PowerPointApi 1.4. */
const POWERPOINT_UNDERLINE_KEYS = [
 "underline",
] as const satisfies ReadonlyArray<keyof PowerPointTextFormat>;

/** Bullet visibility — PowerPointApi 1.4. */
const POWERPOINT_BULLET_VISIBILITY_KEYS = [
 "bulletVisible",
] as const satisfies ReadonlyArray<keyof PowerPointTextFormat>;

/** Bullet style/type — PowerPointApi 1.10. */
const POWERPOINT_BULLET_STYLE_KEYS = [
 "bulletStyle",
 "bulletType",
] as const satisfies ReadonlyArray<keyof PowerPointTextFormat>;

/** Paragraph indent — PowerPointApi 1.10. */
const POWERPOINT_INDENT_KEYS = ["indentLevel"] as const satisfies ReadonlyArray<
 keyof PowerPointTextFormat
>;

/** Font effects — PowerPointApi 1.8. */
const POWERPOINT_FONT_EFFECT_KEYS = [
 "strikethrough",
 "doubleStrikethrough",
 "smallCaps",
 "allCaps",
 "subscript",
 "superscript",
] as const satisfies ReadonlyArray<keyof PowerPointTextFormat>;

function anyKeySet(
 format: PowerPointTextFormat,
 keys: ReadonlyArray<keyof PowerPointTextFormat>,
): boolean {
 return keys.some((key) => format[key] !== undefined);
}

/** True when at least one text property was requested. */
export function hasPowerPointTextFormat(format: PowerPointTextFormat): boolean {
 return anyKeySet(format, POWERPOINT_TEXT_FORMAT_KEYS);
}

/** True when bullet visibility (PowerPointApi 1.4) was requested. */
export function hasPowerPointBulletVisibility(
 format: PowerPointTextFormat,
): boolean {
 return anyKeySet(format, POWERPOINT_BULLET_VISIBILITY_KEYS);
}

/** True when bullet style/type (PowerPointApi 1.10) was requested. */
export function hasPowerPointBulletStyle(
 format: PowerPointTextFormat,
): boolean {
 return anyKeySet(format, POWERPOINT_BULLET_STYLE_KEYS);
}

/** True when paragraph indent level (PowerPointApi 1.10) was requested. */
export function hasPowerPointIndentLevel(
 format: PowerPointTextFormat,
): boolean {
 return anyKeySet(format, POWERPOINT_INDENT_KEYS);
}

/** True when a font effect (PowerPointApi 1.8) was requested. */
export function hasPowerPointFontEffects(
 format: PowerPointTextFormat,
): boolean {
 return anyKeySet(format, POWERPOINT_FONT_EFFECT_KEYS);
}

/**
 * Fail-loud gate for the text-format request: the first requirement set the
 * host lacks, or null when every requested group is available.
 *
 * Ordering matters — higher-set groups are checked first, so a 1.3 host asked
 * for `indentLevel` is told about PowerPointApi 1.10 rather than the generic
 * 1.4 text-format requirement.
 */
export function powerPointTextFormatGap(
 format: PowerPointTextFormat,
): string | null {
 const groups: Array<[ReadonlyArray<keyof PowerPointTextFormat>, string]> = [
  [POWERPOINT_BULLET_STYLE_KEYS, "powerpoint.bullet.style"],
  [POWERPOINT_INDENT_KEYS, "powerpoint.paragraph.indent"],
  [POWERPOINT_FONT_EFFECT_KEYS, "powerpoint.font.effects"],
  [POWERPOINT_UNDERLINE_KEYS, "powerpoint.font.underlineStyles"],
  [POWERPOINT_BULLET_VISIBILITY_KEYS, "powerpoint.bullet.visible"],
  [POWERPOINT_FONT_KEYS, "powerpoint.textFormat"],
 ];

 for (const [keys, capabilityId] of groups) {
  if (!anyKeySet(format, keys)) continue;
  const gap = requireCapability(capabilityId);
  if (gap !== null) return gap;
 }

 return null;
}

/**
 * Fail-loud guard for PowerPoint shape access (`PowerPointApi 1.3`).
 * Returns null when the host may attempt the call.
 */
export function powerPointShapesGateMessage(): string | null {
 return requireCapability("powerpoint.shapes");
}

/**
 * True only when the host is known to lack shape access (`PowerPointApi 1.3`).
 */
export function powerPointShapesUnsupported(): boolean {
 return capabilityUnsupported("powerpoint.shapes");
}

/**
 * Apply text formatting to a PowerPoint text range proxy (e.g.
 * `shape.textFrame.textRange`). Assignments are queued and committed by the
 * caller's context.sync() — safe to call inside one PowerPoint.run batch.
 */
export function setPowerPointTextFormat(
 textRange: PowerPoint.TextRange,
 format: PowerPointTextFormat,
): void {
 if (format.bold !== undefined) textRange.font.bold = format.bold;
 if (format.italic !== undefined) textRange.font.italic = format.italic;
 if (format.underline !== undefined) {
  textRange.font.underline =
   resolvePowerPointUnderline(format.underline) ?? "None";
 }
 if (format.fontSize !== undefined) textRange.font.size = format.fontSize;
 if (format.fontName !== undefined) textRange.font.name = format.fontName;
 if (format.fontColor !== undefined) textRange.font.color = format.fontColor;
 if (format.alignment !== undefined) {
  textRange.paragraphFormat.horizontalAlignment = format.alignment;
 }
 if (format.bulletStyle !== undefined) {
  textRange.paragraphFormat.bulletFormat.style = format.bulletStyle;
 }
 if (format.bulletType !== undefined) {
  textRange.paragraphFormat.bulletFormat.type = format.bulletType;
 }
 if (format.indentLevel !== undefined) {
  textRange.paragraphFormat.indentLevel = format.indentLevel;
 }
 // Visibility last: hiding bullets must not be undone by the style writes.
 if (format.bulletVisible !== undefined) {
  textRange.paragraphFormat.bulletFormat.visible = format.bulletVisible;
 }
 if (format.strikethrough !== undefined) {
  textRange.font.strikethrough = format.strikethrough;
 }
 if (format.doubleStrikethrough !== undefined) {
  textRange.font.doubleStrikethrough = format.doubleStrikethrough;
 }
 if (format.smallCaps !== undefined) {
  textRange.font.smallCaps = format.smallCaps;
 }
 if (format.allCaps !== undefined) {
  textRange.font.allCaps = format.allCaps;
 }
 if (format.subscript !== undefined) {
  textRange.font.subscript = format.subscript;
 }
 if (format.superscript !== undefined) {
  textRange.font.superscript = format.superscript;
 }
}

// ============================================================================
// Shape fill (PowerPointApi 1.4)
// ============================================================================

/** Shape fill options. All PowerPointApi 1.4. */
export interface PowerPointShapeFill {
 /** Solid fill color as #RRGGBB. */
 fillColor?: string;
 /** Transparency, 0 (opaque) to 1 (clear). */
 fillTransparency?: number;
 /** Remove the fill entirely. */
 fillClear?: boolean;
}

/** True when at least one fill property was requested. */
export function hasPowerPointShapeFill(fill: PowerPointShapeFill): boolean {
 return (
  fill.fillColor !== undefined ||
  fill.fillTransparency !== undefined ||
  fill.fillClear === true
 );
}

/**
 * Apply shape fill. `fillClear` runs first so a call that sets both clears the
 * previous fill and then applies the solid color (documented precedence).
 */
export function setPowerPointShapeFill(
 shape: PowerPoint.Shape,
 fill: PowerPointShapeFill,
): void {
 const target = shape.fill;
 if (fill.fillClear === true) target.clear();
 if (fill.fillColor !== undefined) target.setSolidColor(fill.fillColor);
 if (fill.fillTransparency !== undefined)
  target.transparency = fill.fillTransparency;
}

// ============================================================================
// Shape line (PowerPointApi 1.4)
// ============================================================================

/** Shape line options. All PowerPointApi 1.4. */
export interface PowerPointShapeLine {
 /** Line color as #RRGGBB. */
 lineColor?: string;
 /** Line weight in points. */
 lineWeight?: number;
 /** `PowerPoint.ShapeLineStyle`. */
 lineStyle?: PowerPointLineStyle;
 /** `PowerPoint.ShapeLineDashStyle`. */
 lineDashStyle?: PowerPointLineDashStyle;
 /** Show/hide the outline. */
 lineVisible?: boolean;
 /** Transparency, 0 (opaque) to 1 (clear). */
 lineTransparency?: number;
}

/** True when at least one line property was requested. */
export function hasPowerPointShapeLine(line: PowerPointShapeLine): boolean {
 return (
  line.lineColor !== undefined ||
  line.lineWeight !== undefined ||
  line.lineStyle !== undefined ||
  line.lineDashStyle !== undefined ||
  line.lineVisible !== undefined ||
  line.lineTransparency !== undefined
 );
}

/** Apply shape line formatting to a shape proxy. */
export function setPowerPointShapeLine(
 shape: PowerPoint.Shape,
 line: PowerPointShapeLine,
): void {
 const target = shape.lineFormat;
 if (line.lineColor !== undefined) target.color = line.lineColor;
 if (line.lineWeight !== undefined) target.weight = line.lineWeight;
 if (line.lineStyle !== undefined) target.style = line.lineStyle;
 if (line.lineDashStyle !== undefined) target.dashStyle = line.lineDashStyle;
 if (line.lineTransparency !== undefined)
  target.transparency = line.lineTransparency;
 // Visibility last: hiding the line must not be undone by the writes above.
 if (line.lineVisible !== undefined) target.visible = line.lineVisible;
}

// ============================================================================
// Slide background fill (PowerPointApi 1.10)
// ============================================================================

/** Slide background fill options. All PowerPointApi 1.10. */
export interface PowerPointSlideBackground {
 /** Solid background color as #RRGGBB. */
 slideBackgroundColor?: string;
 /** Background transparency, 0 (opaque) to 1 (clear). */
 slideBackgroundTransparency?: number;
 /** Reset the background to the layout/master default. */
 slideBackgroundClear?: boolean;
}

/** True when at least one background property was requested. */
export function hasPowerPointSlideBackground(
 background: PowerPointSlideBackground,
): boolean {
 return (
  background.slideBackgroundColor !== undefined ||
  background.slideBackgroundTransparency !== undefined ||
  background.slideBackgroundClear === true
 );
}

/**
 * Apply slide background fill. `slideBackgroundClear` resets to the
 * layout/master default; setting a color/transparency switches the background
 * to a solid fill.
 */
export function setPowerPointSlideBackground(
 slide: PowerPoint.Slide,
 background: PowerPointSlideBackground,
): void {
 if (background.slideBackgroundClear === true) {
  slide.background.reset();
  return;
 }

 if (
  background.slideBackgroundColor === undefined &&
  background.slideBackgroundTransparency === undefined
 ) {
  return;
 }

 const options: PowerPoint.SlideBackgroundSolidFillOptions = {};
 if (background.slideBackgroundColor !== undefined) {
  options.color = background.slideBackgroundColor;
 }
 if (background.slideBackgroundTransparency !== undefined) {
  options.transparency = background.slideBackgroundTransparency;
 }

 slide.background.fill.setSolidFill(options);
}
