/**
 * Office.js wrappers — thin abstraction over Word.run().
 *
 * Mirrors `excel/helpers.ts` (excelRun) so local Word document tools follow
 * the same execution pattern as the Excel core tools. Also hosts the
 * paragraph-block insertion engine shared by word_insert_text,
 * word_insert_blocks and the markdown path.
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
 * Paragraph-level formatting (lives on Word.Paragraph, not Word.Range).
 */
export interface WordParagraphFormat {
 /** Paragraph style name, e.g. "Heading 1", "Title". Locale-sensitive. */
 style?: string;
 alignment?: WordTextFormat["alignment"];
 /** Space before the paragraph, in points. */
 spaceBefore?: number;
 /** Space after the paragraph, in points. */
 spaceAfter?: number;
 /** Line spacing, in points. */
 lineSpacing?: number;
 /** First-line indent, in points. */
 firstLineIndent?: number;
 /** Left indent, in points. */
 leftIndent?: number;
}

/** Union of everything a block can carry (font + paragraph). */
export type WordBlockFormat = WordTextFormat & WordParagraphFormat;

/** One insertable block of a formatted document. */
export interface WordInsertBlock {
 /** Text content (ignored for pageBreak blocks). */
 text: string;
 type: "paragraph" | "heading" | "listItem" | "pageBreak";
 /** Heading level 1-6 for type "heading". */
 level?: number;
 /** List marker style for type "listItem". */
 listType?: "bullet" | "number";
 format?: WordBlockFormat;
 /** Pre-parsed inline runs (markdown). When empty, the whole text is one run. */
 runs?: WordInlineRun[];
}

/** An inline run with its own font formatting (markdown bold/italic/code). */
export interface WordInlineRun {
 text: string;
 bold?: boolean;
 italic?: boolean;
 code?: boolean;
}

/** Where blocks are inserted. */
export type WordInsertLocation = "start" | "end" | "replace_selection";

/**
 * Any Office proxy that carries a `.font` (Word.Range, Word.Paragraph) — used
 * so formatting helpers accept both without full proxy types.
 */
export interface FontCarrier {
 font: Word.Font;
}

/** Font props from an inline run, applied to a font carrier. */
function applyRunFont(target: FontCarrier, run: WordInlineRun): void {
 if (run.bold !== undefined) target.font.bold = run.bold;
 if (run.italic !== undefined) target.font.italic = run.italic;
 if (run.code === true) target.font.name = "Consolas";
}

/**
 * Apply font properties to a font carrier. Assignments are queued and
 * committed by the caller's context.sync().
 */
export function setWordFontProps(
 target: FontCarrier,
 format: WordTextFormat,
): void {
 if (format.bold !== undefined) target.font.bold = format.bold;
 if (format.italic !== undefined) target.font.italic = format.italic;
 if (format.underline !== undefined) {
  target.font.underline = format.underline ? "Single" : "None";
 }
 if (format.size !== undefined) target.font.size = format.size;
 if (format.name !== undefined) target.font.name = format.name;
 if (format.color !== undefined) target.font.color = format.color;
}

/** Apply paragraph-level formatting to a Word.Paragraph proxy. */
export function setWordParagraphProps(
 paragraph: Word.Paragraph,
 format: WordParagraphFormat | undefined,
): void {
 if (format === undefined) return;
 if (format.style !== undefined) paragraph.style = format.style;
 if (format.alignment !== undefined) paragraph.alignment = format.alignment;
 if (format.spaceBefore !== undefined)
  paragraph.spaceBefore = format.spaceBefore;
 if (format.spaceAfter !== undefined) paragraph.spaceAfter = format.spaceAfter;
 if (format.lineSpacing !== undefined)
  paragraph.lineSpacing = format.lineSpacing;
 if (format.firstLineIndent !== undefined)
  paragraph.firstLineIndent = format.firstLineIndent;
 if (format.leftIndent !== undefined) paragraph.leftIndent = format.leftIndent;
}

/** Resolve a block's runs (defaults to a single run of its text). */
function resolveRuns(block: WordInsertBlock): readonly WordInlineRun[] {
 const text = block.text ?? "";
 return block.runs && block.runs.length > 0 ? block.runs : [{ text }];
}

/** Concrete heading defaults (definite props — no `undefined` leaks). */
function defaultHeadingDefaults(level: number): {
 bold: boolean;
 size: number;
 alignment: "Left";
 spaceAfter: number;
} {
 if (level <= 1)
  return { bold: true, size: 18, alignment: "Left", spaceAfter: 12 };
 if (level === 2)
  return { bold: true, size: 15, alignment: "Left", spaceAfter: 8 };
 return { bold: true, size: 13, alignment: "Left", spaceAfter: 6 };
}

/** Apply heading/list/paragraph formatting to a freshly created paragraph. */
export function applyBlockFormatting(
 paragraph: Word.Paragraph,
 block: WordInsertBlock,
): void {
 if (block.type === "heading") {
  const level = Math.min(Math.max(block.level ?? 1, 1), 6);
  // Try the built-in style first (localized Word may ignore the English name);
  // explicit default formatting guarantees a heading look regardless.
  if (block.format?.style !== undefined) {
   setWordParagraphProps(paragraph, block.format);
  } else {
   paragraph.style = `Heading ${level}`;
  }
  const defaults = defaultHeadingDefaults(level);
  setWordFontProps(paragraph, {
   ...(block.format?.bold !== undefined
    ? { bold: block.format.bold }
    : { bold: defaults.bold }),
   ...(block.format?.size !== undefined
    ? { size: block.format.size }
    : { size: defaults.size }),
  });
  setWordParagraphProps(paragraph, {
   ...(block.format ?? {}),
   ...(block.format?.alignment === undefined
    ? { alignment: defaults.alignment }
    : {}),
   ...(block.format?.spaceAfter === undefined
    ? { spaceAfter: defaults.spaceAfter }
    : {}),
  });
  return;
 }

 if (block.type === "listItem") {
  // No list-creation API exists; the built-in List Bullet / List Number styles
  // are locale-sensitive, so always add the hanging indent so the marker has
  // room, and let the style land when Word resolves it.
  paragraph.style = block.listType === "number" ? "List Number" : "List Bullet";
  setWordParagraphProps(paragraph, {
   ...(block.format ?? {}),
   ...(block.format?.leftIndent === undefined ? { leftIndent: 24 } : {}),
  });
  setWordFontProps(paragraph, block.format ?? {});
  return;
 }

 setWordParagraphProps(paragraph, block.format);
 setWordFontProps(paragraph, block.format ?? {});
}

/**
 * Insert a block as a paragraph at the given body location, filling its runs
 * in order inside the single paragraph.
 */
function createBodyParagraph(
 body: Word.Body,
 block: WordInsertBlock,
 insertLocation: Word.InsertLocation.start | Word.InsertLocation.end,
): Word.Paragraph {
 const runs = resolveRuns(block);
 const firstText = runs[0]?.text ?? "";
 const paragraph = body.insertParagraph(firstText, insertLocation);
 if (runs[0]) applyRunFont(paragraph, runs[0]);

 let range: Word.Range = paragraph.getRange();
 for (const run of runs.slice(1)) {
  range = range.insertText(run.text, Word.InsertLocation.end);
  applyRunFont(range, run);
 }

 applyBlockFormatting(paragraph, block);
 return paragraph;
}

/**
 * Insert an ordered list of blocks at the given location, returning the
 * inserted paragraphs. This is the single engine behind word_insert_text
 * (plain + markdown) and word_insert_blocks.
 *
 * - end/start: each block becomes a paragraph appended via body.insertParagraph
 *   (start inserts in reverse so the first block ends up first).
 * - replace_selection: the first block replaces the selection; subsequent
 *   blocks are inserted "after" the previously inserted paragraph.
 */
export function insertWordBlocks(
 context: Word.RequestContext,
 blocks: readonly WordInsertBlock[],
 location: WordInsertLocation,
): Word.Paragraph[] {
 const inserted: Word.Paragraph[] = [];
 if (blocks.length === 0) return inserted;

 if (location === "replace_selection") {
  const selection = context.document.getSelection();
  let anchor: Word.Range = selection;
  let isFirst = true;
  for (const block of blocks) {
   if (block.type === "pageBreak") {
    anchor.insertBreak("Page", Word.InsertLocation.after);
    continue;
   }
   const runs = resolveRuns(block);
   const firstText = runs[0]?.text ?? "";
   let paragraph: Word.Paragraph;
   if (isFirst) {
    anchor = anchor.insertText(firstText, Word.InsertLocation.replace);
    if (runs[0]) applyRunFont(anchor, runs[0]);
    for (const run of runs.slice(1)) {
     anchor = anchor.insertText(run.text, Word.InsertLocation.end);
     applyRunFont(anchor, run);
    }
    paragraph = anchor.paragraphs.getFirst();
   } else {
    paragraph = anchor.insertParagraph(firstText, Word.InsertLocation.after);
    if (runs[0]) applyRunFont(paragraph, runs[0]);
    let range: Word.Range = paragraph.getRange();
    for (const run of runs.slice(1)) {
     range = range.insertText(run.text, Word.InsertLocation.end);
     applyRunFont(range, run);
    }
   }
   applyBlockFormatting(paragraph, block);
   inserted.push(paragraph);
   anchor = paragraph.getRange();
   isFirst = false;
  }
  return inserted;
 }

 const body = context.document.body;
 const insertLocation =
  location === "start" ? Word.InsertLocation.start : Word.InsertLocation.end;
 // "start" must insert in reverse so the first block ends up on top.
 const ordered = location === "start" ? [...blocks].reverse() : blocks;

 for (const block of ordered) {
  if (block.type === "pageBreak") {
   body.insertBreak("Page", insertLocation);
   continue;
  }
  inserted.push(createBodyParagraph(body, block, insertLocation));
 }
 return inserted;
}

/** Return paragraph count + a compact heading outline for the document. */
export async function readDocumentStructure(
 context: Word.RequestContext,
 maxHeadings = 200,
): Promise<{ paragraphCount: number; tableCount: number; headings: string[] }> {
 const body = context.document.body;
 const paragraphs = body.paragraphs;
 paragraphs.load("items/style");
 const tables = body.tables;
 tables.load("items");
 await context.sync();

 const headings: string[] = [];
 for (const para of paragraphs.items) {
  const style = para.style;
  if (typeof style === "string" && style.toLowerCase().startsWith("heading")) {
   const level = style.replace(/^Heading\s*/i, "");
   headings.push(`H${level || "?"}`);
   if (headings.length >= maxHeadings) break;
  }
 }
 return {
  paragraphCount: paragraphs.items.length,
  tableCount: tables.items.length,
  headings,
 };
}
