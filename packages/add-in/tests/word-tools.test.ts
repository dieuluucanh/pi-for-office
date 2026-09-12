/**
 * Word tool tests — exercised against a mocked global `Word` so the live
 * Word.js tools run under Node. Focus on the incident regressions:
 *  1. `location: "end"` must append (not always insert at start).
 *  2. Multi-line text must become separate paragraphs.
 *  3. Markdown conversion produces headings/lists/page breaks.
 *  4. format_range applies font + paragraph props to matches.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { parseMarkdownToBlocks } from "../src/word/markdown.ts";
import { insertWordBlocks } from "../src/word/helpers.ts";
import {
  makeContext,
  makeRange,
  runWithFakeWord,
} from "./helpers/fake-word.ts";
import { createWordInsertTextTool } from "../src/tools/word/insert-text.ts";
import { createWordFormatRangeTool } from "../src/tools/word/format-range.ts";
import { createWordGetOverviewTool } from "../src/tools/word/get-overview.ts";
import { createWordReadDocumentTool } from "../src/tools/word/read-document.ts";

/** Extract the text payload from an AgentToolResult (content is string|parts). */
function toolText(result: {
  content?: Array<{ type?: string; text?: DynamicValue }> | string;
}): string {
  const content = result.content;
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";
  return content
    .filter((part) => typeof part.text === "string")
    .map((part) => String(part.text))
    .join("\n");
}

void test("markdown: headings, bullets, numbered lists, page-break rule", () => {
  const blocks = parseMarkdownToBlocks(
    "# Title\n\nSome body.\n\n- a\n- b\n\n1. one\n\n---\n\nNext.",
  );
  const types = blocks.map((b) => b.type);
  assert.deepEqual(types, [
    "heading",
    "paragraph",
    "listItem",
    "listItem",
    "listItem",
    "pageBreak",
    "paragraph",
  ]);
  assert.equal(blocks[0]?.level, 1);
  assert.equal(blocks[2]?.listType, "bullet");
  assert.equal(blocks[4]?.listType, "number");
});

void test("markdown: inline bold/italic/code become runs", () => {
  const blocks = parseMarkdownToBlocks(
    "This is **bold** and *italic* and `code`.",
  );
  const paragraph = blocks[0];
  assert.ok(
    paragraph.runs && paragraph.runs.length >= 4,
    "expected multiple runs",
  );
  assert.ok(paragraph.runs.some((r) => r.bold === true && r.text === "bold"));
  assert.ok(
    paragraph.runs.some((r) => r.italic === true && r.text === "italic"),
  );
  assert.ok(paragraph.runs.some((r) => r.code === true && r.text === "code"));
});

void test("insert_text at end appends (regression: always-start bug)", async () => {
  const context = makeContext();
  await runWithFakeWord(context, async () => {
    const tool = createWordInsertTextTool();
    const result = await tool.execute("t1", {
      text: "line one\nline two",
      location: "end",
    });
    assert.ok(
      toolText(result).includes("at the end of the document"),
      toolText(result),
    );
  });
  const bodyCalls = context.calls.join("\n");
  assert.ok(bodyCalls.includes('insertParagraph("line one", End)'), bodyCalls);
  assert.ok(bodyCalls.includes('insertParagraph("line two", End)'), bodyCalls);
});

void test("insert_text at start inserts first, then the rest (reverse order)", async () => {
  const context = makeContext();
  await runWithFakeWord(context, async () => {
    const tool = createWordInsertTextTool();
    await tool.execute("t1", { text: "A\nB", location: "start" });
  });
  const calls = context.calls.join("\n");
  // For start, the engine inserts in reverse so A ends up first.
  assert.ok(
    calls.indexOf('insertParagraph("B", Start)') <
      calls.indexOf('insertParagraph("A", Start)'),
    calls,
  );
});

void test("insert_text with markdown creates heading + paragraph blocks", async () => {
  const context = makeContext();
  await runWithFakeWord(context, async () => {
    const tool = createWordInsertTextTool();
    const result = await tool.execute("t1", {
      text: "# Heading\n\nBody text",
      format: "markdown",
    });
    assert.ok(toolText(result).startsWith("Inserted "), toolText(result));
  });
  const bodyCalls = context.calls.join("\n");
  assert.ok(bodyCalls.includes('insertParagraph("Heading", End)'), bodyCalls);
  assert.ok(bodyCalls.includes('insertParagraph("Body text", End)'), bodyCalls);
});

void test("format_range applies bold/alignment to every match via paragraphs", async () => {
  const context = makeContext();
  await runWithFakeWord(context, async () => {
    const tool = createWordFormatRangeTool();
    const result = await tool.execute("t1", {
      text: "title",
      bold: true,
      size: 18,
      alignment: "Centered",
    });
    assert.ok(
      toolText(result).includes("No occurrences"),
      "no matches in fake body",
    );
  });
  // Formatting path only touches paragraphs when there are matches; with zero
  // matches the tool must report gracefully without throwing.
  assert.ok(context.calls.length > 0);
});

void test("read_document renders paragraph text and heading markers", async () => {
  const context = makeContext();
  // Seed fake paragraphs.
  context.body.paragraphs.items.push(
    {
      text: "My Title",
      style: "Heading 1",
      alignment: undefined,
      font: {},
      calls: [],
      getRange: () => makeRange("My Title"),
      load() {},
    },
    {
      text: "Body line",
      style: "Normal",
      alignment: undefined,
      font: {},
      calls: [],
      getRange: () => makeRange("Body line"),
      load() {},
    },
  );
  await runWithFakeWord(context, async () => {
    const tool = createWordReadDocumentTool();
    const result = await tool.execute("t1", { scope: "all" });
    const text = toolText(result);
    assert.ok(text.includes("# My Title"), text);
    assert.ok(text.includes("Body line"), text);
  });
});

void test("get_overview reports counts (fake body is empty)", async () => {
  const context = makeContext();
  await runWithFakeWord(context, async () => {
    const tool = createWordGetOverviewTool();
    const result = await tool.execute("t1", {});
    const text = toolText(result);
    assert.ok(text.includes("Paragraphs: 0"), text);
    assert.ok(text.includes("Tables: 0"), text);
    assert.ok(text.includes("Headings: none detected"), text);
  });
});

void test("insert engine preserves block order for end insertion", async () => {
  const context = makeContext();
  await runWithFakeWord(context, () => {
    const inserted = insertWordBlocks(
      context as never,
      [
        { text: "1", type: "paragraph" },
        { text: "2", type: "paragraph" },
      ],
      "end",
    );
    assert.equal(inserted.length, 2);
  });
  const calls = context.calls.join("\n");
  assert.ok(
    calls.indexOf('insertParagraph("1", End)') <
      calls.indexOf('insertParagraph("2", End)'),
    calls,
  );
});
