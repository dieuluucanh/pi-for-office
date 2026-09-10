/**
 * Office.js wrappers — thin abstraction over PowerPoint.run().
 *
 * Mirrors `excel/helpers.ts` (excelRun) so local PowerPoint tools follow
 * the same execution pattern as the Excel core tools.
 */

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
// Formatting
// ============================================================================

/** Formatting options shared by the PowerPoint tools (all optional). */
export interface PowerPointTextFormat {
 bold?: boolean;
 italic?: boolean;
 underline?: boolean;
 /** Font size in points. */
 fontSize?: number;
 /** Font name, e.g. "Arial". */
 fontName?: string;
 /** Font color as #RRGGBB. */
 fontColor?: string;
 /** Horizontal alignment. PPT values: "Left" | "Center" | "Right" | "Justify". */
 alignment?: "Left" | "Center" | "Right" | "Justify";
}

/**
 * Apply formatting to a PowerPoint text range proxy (e.g.
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
  textRange.font.underline = format.underline ? "Single" : "None";
 }
 if (format.fontSize !== undefined) textRange.font.size = format.fontSize;
 if (format.fontName !== undefined) textRange.font.name = format.fontName;
 if (format.fontColor !== undefined) textRange.font.color = format.fontColor;
 if (format.alignment !== undefined) {
  textRange.paragraphFormat.horizontalAlignment = format.alignment;
 }
}
