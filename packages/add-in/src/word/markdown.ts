/**
 * Markdown subset → Word blocks.
 *
 * Converts markdown-like text into the same `WordInsertBlock` structure the
 * Word insertion engine consumes. Supported:
 *
 *  - `#`…`######` headings
 *  - `-` / `*` / `+` bullet list items
 *  - `1.` numbered list items
 *  - `>` blockquote (indented paragraph)
 *  - `---` horizontal rule → page break
 *  - inline `**bold**`, `*italic*`, `` `code` `` runs
 *
 * Everything else becomes a normal paragraph. Blank lines separate blocks.
 */

import type { WordInsertBlock, WordInlineRun } from "./helpers.js";

/** Split a line into inline runs (bold / italic / code segments). */
export function parseInlineRuns(line: string): WordInlineRun[] {
  const runs: WordInlineRun[] = [];
  const rest = line;
  const pattern = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g;
  let cursor = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(rest)) !== null) {
    const before = rest.slice(cursor, match.index);
    if (before.length > 0) runs.push({ text: before });
    const token = match[0];
    if (token.startsWith("**")) {
      runs.push({ text: token.slice(2, -2), bold: true });
    } else if (token.startsWith("`")) {
      runs.push({ text: token.slice(1, -1), code: true });
    } else {
      runs.push({ text: token.slice(1, -1), italic: true });
    }
    cursor = match.index + token.length;
  }
  const tail = rest.slice(cursor);
  if (tail.length > 0) runs.push({ text: tail });
  return runs.length > 0 ? runs : [{ text: line }];
}

/** True when the line looks like a block structure start. */
function isHeading(line: string): number {
  const m = /^(#{1,6})\s+(.*)$/.exec(line);
  return m ? (m[1]?.length ?? 0) : 0;
}

function isBullet(line: string): string | null {
  const m = /^\s*[-*+]\s+(.*)$/.exec(line);
  return m ? (m[1] ?? null) : null;
}

function isNumbered(line: string): string | null {
  const m = /^\s*(?:\d+\.)\s+(.*)$/.exec(line);
  return m ? (m[1] ?? null) : null;
}

function isBlockquote(line: string): string | null {
  const m = /^\s*>\s?(.*)$/.exec(line);
  return m ? (m[1] ?? null) : null;
}

function isHorizontalRule(line: string): boolean {
  return /^\s*---+\s*$/.test(line);
}

/**
 * Convert markdown text to ordered Word blocks. Unsupported constructs are
 * degraded gracefully (never dropped silently when avoidable).
 */
export function parseMarkdownToBlocks(markdown: string): WordInsertBlock[] {
  const blocks: WordInsertBlock[] = [];
  const lines = markdown.split(/\r?\n/u);

  let paragraphBuffer: string[] = [];
  const flushParagraph = () => {
    if (paragraphBuffer.length === 0) return;
    const text = paragraphBuffer.join(" ");
    blocks.push({ text, type: "paragraph", runs: parseInlineRuns(text) });
    paragraphBuffer = [];
  };

  for (const rawLine of lines) {
    const line = rawLine.trimEnd();
    if (line.trim().length === 0) {
      flushParagraph();
      continue;
    }

    const level = isHeading(line);
    if (level > 0) {
      flushParagraph();
      const text = line.replace(/^#{1,6}\s+/, "");
      blocks.push({
        text,
        type: "heading",
        level,
        runs: parseInlineRuns(text),
      });
      continue;
    }

    const bullet = isBullet(line);
    if (bullet !== null) {
      flushParagraph();
      const text = bullet.trim();
      blocks.push({
        text,
        type: "listItem",
        listType: "bullet",
        runs: parseInlineRuns(text),
      });
      continue;
    }

    const numbered = isNumbered(line);
    if (numbered !== null) {
      flushParagraph();
      const text = numbered.trim();
      blocks.push({
        text,
        type: "listItem",
        listType: "number",
        runs: parseInlineRuns(text),
      });
      continue;
    }

    if (isHorizontalRule(line)) {
      flushParagraph();
      blocks.push({ text: "", type: "pageBreak" });
      continue;
    }

    const quote = isBlockquote(line);
    if (quote !== null) {
      flushParagraph();
      const text = quote.trim();
      blocks.push({
        text,
        type: "paragraph",
        runs: parseInlineRuns(text),
        format: { leftIndent: 36, italic: true },
      });
      continue;
    }

    paragraphBuffer.push(line.trim());
  }
  flushParagraph();
  return blocks;
}
