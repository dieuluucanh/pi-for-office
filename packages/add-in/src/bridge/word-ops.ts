/**
 * Word bridge executors — the first Word.js operations in pi-for-office.
 *
 * Runs inside the Word task pane via the `Word.run` batch model:
 * proxy objects → load() → context.sync() → read/mutate → sync().
 */

import { guardExecutor, type OfficeOpExecutor } from "./ops.js";

const MAX_READ_CHARS = 200_000;

/** Count words in text (approximation split on whitespace). */
function wordCount(text: string): number {
  const trimmed = text.trim();
  if (trimmed.length === 0) return 0;
  return trimmed.split(/\s+/).length;
}

async function readBodyOverview() {
  return Word.run(async (context) => {
    const body = context.document.body;

    // Paragraph proxies: style only (no text) to build the heading outline.
    const paragraphs = body.paragraphs;
    paragraphs.load("items/style");
    const tables = body.tables;
    tables.load("items");

    await context.sync();

    const headings: string[] = [];
    let paragraphCount = 0;
    for (const para of paragraphs.items) {
      paragraphCount += 1;
      const style = para.style;
      if (typeof style === "string" && style.toLowerCase().startsWith("heading")) {
        // Keep only the style name — extracting heading text for the whole doc
        // would require loading every paragraph's text (see read_document).
        const level = style.replace(/^Heading\s*/i, "");
        headings.push(`H${level || "?"}`);
      }
      if (headings.length >= 200 || paragraphCount >= 200_000) break;
    }

    return {
      paragraphCount: paragraphs.items.length,
      tableCount: tables.items.length,
      headingCountByLevel: headings.reduce<Record<string, number>>((acc, h) => {
        acc[h] = (acc[h] ?? 0) + 1;
        return acc;
      }, {}),
    };
  });
}

export const WORD_OPS: ReadonlyMap<string, OfficeOpExecutor> = new Map<
  string,
  OfficeOpExecutor
>([
  [
    "word.get_overview",
    guardExecutor(async () => {
      const overview = await readBodyOverview();

      const lines: string[] = [`**Document overview**`];
      lines.push(`- Paragraphs: ${overview.paragraphCount}`);
      lines.push(`- Tables: ${overview.tableCount}`);
      const headingEntries = Object.entries(overview.headingCountByLevel);
      if (headingEntries.length > 0) {
        lines.push(
          `- Headings: ${headingEntries.map(([level, n]) => `${level}×${n}`).join(", ")}`,
        );
      } else {
        lines.push("- Headings: none detected");
      }
      lines.push("Use word.read_document to read text; word.insert_text / word.replace_text to edit.");

      return { text: lines.join("\n"), details: overview };
    }),
  ],
  [
    "word.read_document",
    guardExecutor(async (args) => {
      const scope = args.scope === "selection" ? "selection" : "all";
      const maxChars =
        typeof args.maxChars === "number" ? Math.min(Math.max(args.maxChars, 100), MAX_READ_CHARS) : 20_000;

      const text = await Word.run(async (context) => {
        if (scope === "selection") {
          const selection = context.document.getSelection();
          selection.load("text");
          await context.sync();
          return selection.text ?? "";
        }
        const body = context.document.body;
        body.load("text");
        await context.sync();
        return body.text ?? "";
      });

      const truncated = text.length > maxChars;
      const shown = truncated ? text.slice(0, maxChars) : text;

      const lines: string[] = [];
      if (scope === "selection") {
        lines.push(`**Selection (${shown.length} chars shown)**`);
      } else {
        lines.push(`**Document text** (${text.length} chars total${truncated ? `, showing first ${maxChars}` : ""})`);
      }
      if (shown.trim().length === 0) {
        lines.push("_Empty._");
      } else {
        lines.push("");
        lines.push(shown);
        if (truncated) {
          lines.push("");
          lines.push(`_…truncated (${text.length - maxChars} chars remaining). Pass maxChars to read more._`);
        }
      }
      return { text: lines.join("\n"), details: { scope, totalChars: text.length, wordCount: wordCount(text) } };
    }),
  ],
  [
    "word.insert_text",
    guardExecutor(async (args) => {
      const text = typeof args.text === "string" ? args.text : "";
      const location = args.location === "start" || args.location === "replace_selection" ? args.location : "end";
      if (text.length === 0) {
        return { text: "Nothing to insert (empty text).", isError: true };
      }

      await Word.run(async (context) => {
        const body = context.document.body;
        if (location === "replace_selection") {
          const selection = context.document.getSelection();
          selection.insertText(text, Word.InsertLocation.replace);
        } else {
          body.insertText(text, Word.InsertLocation.start); // start/end both supported on body
        }
        await context.sync();
      });

      return {
        text:
          location === "replace_selection"
            ? `Replaced the current selection with ${text.length} chars.`
            : `Inserted ${text.length} chars at the ${location} of the document.`,
        details: { location, insertedChars: text.length },
      };
    }),
  ],
  [
    "word.replace_text",
    guardExecutor(async (args) => {
      const find = typeof args.find === "string" ? args.find : "";
      const replace = typeof args.replace === "string" ? args.replace : "";
      const matchCase = args.matchCase === true;
      if (find.length === 0) {
        return { text: "Missing 'find' text.", isError: true };
      }

      const replaced = await Word.run(async (context) => {
        const results = context.document.body.search(find, { matchCase });
        results.load("length");
        await context.sync();
        const count = results.items.length;
        for (const item of results.items) {
          item.insertText(replace, Word.InsertLocation.replace);
        }
        await context.sync();
        return count;
      });

      return {
        text: replaced === 0
          ? `No occurrences of "${find}" found.`
          : `Replaced ${replaced} occurrence${replaced === 1 ? "" : "s"} of "${find}".`,
        details: { find, replaced },
      };
    }),
  ],
]);
