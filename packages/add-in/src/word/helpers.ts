/**
 * Office.js wrappers — thin abstraction over Word.run().
 *
 * Mirrors `excel/helpers.ts` (excelRun) so local Word document tools follow
 * the same execution pattern as the Excel core tools.
 */

/**
 * Run an Office.js Word operation with a typed context.
 * Wraps Word.run() (proxy objects → load() → context.sync()).
 */
export async function wordRun<T>(
 fn: (context: Word.RequestContext) => Promise<T>,
): Promise<T> {
 return Word.run(fn);
}

/** Approximate word count (split on whitespace), mirroring bridge word-ops. */
export function wordCount(text: string): number {
 const trimmed = text.trim();
 if (trimmed.length === 0) return 0;
 return trimmed.split(/\s+/).length;
}

// ============================================================================
// Formatting
// ============================================================================

/** Formatting options shared by the Word tools (all optional — set what's given). */
export interface WordTextFormat {
 bold?: boolean;
 italic?: boolean;
 underline?: boolean;
 /** Font size in points. */
 size?: number;
 /** Font name, e.g. "Arial". */
 name?: string;
 /** Font color as #RRGGBB. */
 color?: string;
 /** Paragraph alignment. Word enum values — note "Centered", not "center". */
 alignment?: "Left" | "Centered" | "Right" | "Justified";
}

/**
 * Apply font properties to a Word range proxy. Assignments are queued and
 * committed by the caller's context.sync() — safe to call for many ranges
 * inside one Word.run batch without extra syncs.
 */
export function setWordFontProps(
 range: Word.Range,
 format: WordTextFormat,
): void {
 if (format.bold !== undefined) range.font.bold = format.bold;
 if (format.italic !== undefined) range.font.italic = format.italic;
 if (format.underline !== undefined) {
  range.font.underline = format.underline ? "Single" : "None";
 }
 if (format.size !== undefined) range.font.size = format.size;
 if (format.name !== undefined) range.font.name = format.name;
 if (format.color !== undefined) range.font.color = format.color;
}

/**
 * Apply paragraph alignment to every paragraph covered by a Word range.
 * Alignment lives on Word.Paragraph (Word.Range has no paragraphFormat), so
 * the paragraph proxies must be loaded and synced before iterating.
 */
export async function applyWordAlignment(
 range: Word.Range,
 alignment: NonNullable<WordTextFormat["alignment"]>,
): Promise<void> {
 range.paragraphs.load("items");
 await range.paragraphs.context.sync();
 for (const paragraph of range.paragraphs.items) {
  paragraph.alignment = alignment;
 }
}
