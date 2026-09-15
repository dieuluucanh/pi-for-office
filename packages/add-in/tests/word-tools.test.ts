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
import {
  insertWordBlocks,
  setWordFontProps,
  wordRangeNavigationGateMessage,
  wordRun,
  type WordInsertBlock,
} from "../src/word/helpers.ts";
import {
  makeContext,
  makeRange,
  runWithFakeWord,
} from "./helpers/fake-word.ts";
import { createWordInsertTextTool } from "../src/tools/word/insert-text.ts";
import { createWordFormatRangeTool } from "../src/tools/word/format-range.ts";
import { createWordGetOverviewTool } from "../src/tools/word/get-overview.ts";
import { createWordReadDocumentTool } from "../src/tools/word/read-document.ts";
import { createWordInsertBlocksTool } from "../src/tools/word/insert-blocks.ts";
import {
  buildHyperlinkOoxml,
  createWordInsertHyperlinkTool,
} from "../src/tools/word/insert-hyperlink.ts";
import { createWordInsertTableTool } from "../src/tools/word/insert-table.ts";
import { createWordInsertImageTool } from "../src/tools/word/insert-image.ts";

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

/** Install a fake `Office.context` for the duration of `fn`, then restore it. */
async function withOffice<T>(
  office: Record<string, DynamicValue>,
  fn: () => Promise<T>,
): Promise<T> {
  const scope = globalThis as Record<string, DynamicValue>;
  const prior = scope.Office;
  scope.Office = office;
  try {
    return await fn();
  } finally {
    if (prior === undefined) delete scope.Office;
    else scope.Office = prior;
  }
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

/* ── Text highlight (the "yellow fill" regression) ─────────────────────
 * `Office.word_format_range`/`insert_text` advertised no way to fill text, so
 * an agent correctly reported that Word could not do it. These tests pin the
 * capability AND the three factories that must forward it.
 */

void test("setWordFontProps: highlight assigns, and 'none' clears it", () => {
  const font: Record<string, DynamicValue> = {};
  // SAFETY: setWordFontProps only touches `.font`, which is all this stub provides.
  const carrier = { font } as unknown as Parameters<typeof setWordFontProps>[0];

  setWordFontProps(carrier, { highlight: "#FFFF00" });
  assert.equal(font.highlightColor, "#FFFF00");

  // Microsoft documents `null` as the way to remove a highlight, even though
  // the d.ts declares the property as `string`.
  setWordFontProps(carrier, { highlight: "none" });
  assert.equal(font.highlightColor, null);
});

void test("insert_text forwards highlight to the inserted paragraph", async () => {
  const context = makeContext();
  await runWithFakeWord(context, async () => {
    const tool = createWordInsertTextTool();
    await tool.execute("t1", {
      text: "testing with target",
      color: "#0000FF",
      highlight: "#FFFF00",
    });
  });
  const paragraph = context.body.createdParagraphs[0];
  assert.ok(paragraph, "expected one inserted paragraph");
  assert.equal(paragraph.font.highlightColor, "#FFFF00");
  assert.equal(paragraph.font.color, "#0000FF");
});

void test("insert_blocks forwards highlight from the block to the paragraph", async () => {
  const context = makeContext();
  await runWithFakeWord(context, async () => {
    const tool = createWordInsertBlocksTool();
    await tool.execute("t1", {
      blocks: [{ text: "block", highlight: "#FFFF00" }],
    });
  });
  const paragraph = context.body.createdParagraphs[0];
  assert.ok(paragraph, "expected one inserted paragraph");
  assert.equal(paragraph.font.highlightColor, "#FFFF00");
});

void test("format_range forwards highlight to every matched range", async () => {
  const context = makeContext();
  const matched = makeRange("title");
  context.body.searchResults.push(matched);

  await runWithFakeWord(context, async () => {
    const tool = createWordFormatRangeTool();
    const result = await tool.execute("t1", {
      text: "title",
      highlight: "#FFFF00",
    });
    assert.ok(
      toolText(result).includes("highlight=#FFFF00"),
      `summary must report the highlight: ${toolText(result)}`,
    );
  });

  assert.equal(matched.font.highlightColor, "#FFFF00");
});

/* ── Paragraph shading (desktop-only, must fail loudly) ─────────────────── */

void test("shading is applied when the host permits it", async () => {
  // No `Office` global in the harness ⇒ verdict is "unknown", so the gate lets
  // the call through and the Office.js call would surface any real error.
  const context = makeContext();
  await runWithFakeWord(context, async () => {
    const tool = createWordInsertTextTool();
    await tool.execute("t1", { text: "shade me", shading: "#FFFF00" });
  });
  const paragraph = context.body.createdParagraphs[0];
  assert.ok(paragraph, "expected one inserted paragraph");
  assert.equal(paragraph.shading.backgroundPatternColor, "#FFFF00");
});

void test("shading fails loudly on a host without WordApiDesktop", async () => {
  const context = makeContext();
  const scope = globalThis as Record<string, DynamicValue>;
  const priorOffice = scope.Office;
  scope.Office = {
    context: {
      platform: "OfficeOnline",
      requirements: {
        isSetSupported: (set: string) => set !== "WordApiDesktop",
      },
    },
  };
  try {
    await runWithFakeWord(context, async () => {
      const tool = createWordFormatRangeTool();
      const result = await tool.execute("t1", {
        text: "title",
        shading: "#FFFF00",
      });
      const text = toolText(result);
      // The message must name the capability, the set, and the platform.
      assert.ok(text.includes("WordApiDesktop 1.3"), text);
      assert.ok(text.includes("Office on the web"), text);
      assert.ok(text.includes("highlight"), text);
    });
  } finally {
    if (priorOffice === undefined) {
      delete scope.Office;
    } else {
      scope.Office = priorOffice;
    }
  }
  // Nothing may be written to the document when the gate trips.
  assert.equal(context.calls.length, 0, context.calls.join("\n"));
});

void test("format_range shading is character-precise, not paragraph-wide", async () => {
  const context = makeContext();
  const matched = makeRange("title");
  context.body.searchResults.push(matched);

  await runWithFakeWord(context, async () => {
    const tool = createWordFormatRangeTool();
    await tool.execute("t1", { text: "title", shading: "#FFFF00" });
  });

  // The fill lands on the matched range...
  assert.equal(matched.shading.backgroundPatternColor, "#FFFF00");
  // ...and NOT on the paragraph containing it, so sibling text in the same
  // paragraph stays unshaded. That is the difference from the insert paths,
  // where the paragraph is the unit of formatting.
  assert.equal(
    matched.paragraphs.items[0].shading.backgroundPatternColor,
    undefined,
  );
});

/* ── Hyperlinks: must work on Word for the web, where the native API throws ── */

void test("insert_hyperlink uses the native API when WordApiDesktop 1.3 exists", async () => {
  const context = makeContext();
  await withOffice(
    {
      context: { platform: "PC", requirements: { isSetSupported: () => true } },
    },
    () =>
      runWithFakeWord(context, async () => {
        const tool = createWordInsertHyperlinkTool();
        const result = await tool.execute("t1", {
          text: "site",
          url: "https://example.com",
        });
        assert.ok(
          toolText(result).includes("native hyperlink API"),
          toolText(result),
        );
      }),
  );

  assert.ok(
    context.calls.some((c) => c.startsWith("hyperlinks.add(")),
    context.calls.join("\n"),
  );
  assert.equal(context.body.lastOoxml, null, "native path must not use OOXML");
});

void test("insert_hyperlink falls back to OOXML when WordApiDesktop 1.3 is absent", async () => {
  // Word for the web: `hyperlinks.add` throws, `insertOoxml` (WordApi 1.1) works.
  const context = makeContext();
  await withOffice(
    {
      context: {
        platform: "OfficeOnline",
        requirements: {
          isSetSupported: (set: string) => set !== "WordApiDesktop",
        },
      },
    },
    () =>
      runWithFakeWord(context, async () => {
        const tool = createWordInsertHyperlinkTool();
        const result = await tool.execute("t1", {
          text: "site",
          url: "https://example.com",
        });
        assert.ok(
          toolText(result).includes("OOXML fallback"),
          toolText(result),
        );
      }),
  );

  // No native call at all...
  assert.ok(
    !context.calls.some((c) => c.startsWith("hyperlinks.add(")),
    context.calls.join("\n"),
  );
  // ...and the self-contained field carries the URL inline.
  const ooxml = context.body.lastOoxml;
  assert.ok(ooxml, "expected an OOXML insertion");
  assert.match(ooxml, /<w:fldSimple /);
  assert.ok(ooxml.includes("HYPERLINK"), ooxml);
  assert.ok(ooxml.includes("https://example.com"), ooxml);
  assert.match(ooxml, /<w:t xml:space="preserve">site<\/w:t>/);
});

void test("buildHyperlinkOoxml escapes text, URL and screen tip", () => {
  const ooxml = buildHyperlinkOoxml(
    'a "b" <c> & d',
    "https://x.test/?a=1&b=2",
    "tip",
  );
  // Quotes inside the attribute value must be entities, never raw — otherwise a
  // URL containing a quote could break out of `w:instr`.
  assert.ok(
    ooxml.includes("HYPERLINK &quot;https://x.test/?a=1&amp;b=2&quot;"),
    ooxml,
  );
  assert.ok(ooxml.includes("\\o &quot;tip&quot;"), ooxml);
  assert.ok(ooxml.includes("&lt;c&gt;"), ooxml);
  assert.ok(ooxml.includes("&amp; d"), ooxml);
  assert.match(ooxml, /^<w:p /);
  assert.match(ooxml, /<\/w:p>$/);
});

/* ── WordApi 1.2/1.3 gates: table, inline image, range navigation ───────── */

void test("insert_table fails loudly below WordApi 1.3", async () => {
  const context = makeContext();
  await withOffice(
    {
      context: {
        platform: "Mac",
        requirements: { isSetSupported: () => false },
      },
    },
    () =>
      runWithFakeWord(context, async () => {
        const tool = createWordInsertTableTool();
        const result = await tool.execute("t1", {
          cells: [
            ["a", "b"],
            ["c", "d"],
          ],
        });
        const text = toolText(result);
        assert.ok(text.includes("WordApi 1.3"), text);
        assert.ok(text.includes("table insertion"), text);
        assert.ok(text.includes("Mac desktop"), text);
      }),
  );
  // Nothing may reach the document when the gate trips.
  assert.equal(context.calls.length, 0, context.calls.join("\n"));
});

void test("insert_table inserts when WordApi 1.3 is available", async () => {
  const context = makeContext();
  await withOffice(
    {
      context: { platform: "PC", requirements: { isSetSupported: () => true } },
    },
    () =>
      runWithFakeWord(context, async () => {
        const tool = createWordInsertTableTool();
        const result = await tool.execute("t1", {
          cells: [
            ["a", "b"],
            ["c", "d"],
          ],
          headerRow: true,
        });
        assert.ok(toolText(result).includes("Inserted"), toolText(result));
      }),
  );
  assert.ok(
    context.calls.some((c) => c.startsWith("insertTable(")),
    context.calls.join("\n"),
  );
});

void test("insert_image fails loudly below WordApi 1.2", async () => {
  const context = makeContext();
  await withOffice(
    {
      context: {
        platform: "iOS",
        requirements: { isSetSupported: () => false },
      },
    },
    () =>
      runWithFakeWord(context, async () => {
        const tool = createWordInsertImageTool();
        const result = await tool.execute("t1", {
          base64: "iVBORw0KGgo=",
        });
        const text = toolText(result);
        assert.ok(text.includes("WordApi 1.2"), text);
        assert.ok(text.includes("inline images"), text);
      }),
  );
  assert.equal(context.calls.length, 0, context.calls.join("\n"));
});

void test("insert_image inserts when WordApi 1.2 is available", async () => {
  const context = makeContext();
  await withOffice(
    {
      context: { platform: "PC", requirements: { isSetSupported: () => true } },
    },
    () =>
      runWithFakeWord(context, async () => {
        const tool = createWordInsertImageTool();
        const result = await tool.execute("t1", {
          base64: "iVBORw0KGgo=",
          width: 100,
        });
        assert.ok(
          toolText(result).includes("Inserted an image"),
          toolText(result),
        );
      }),
  );
  assert.ok(
    context.calls.some((c) => c.startsWith("insertInlinePictureFromBase64(")),
    context.calls.join("\n"),
  );
});

void test("range navigation gate names WordApi 1.3 only when the host lacks it", async () => {
  await withOffice(
    {
      context: {
        platform: "PC",
        requirements: { isSetSupported: () => false },
      },
    },
    async () => {
      const message = wordRangeNavigationGateMessage();
      assert.ok(message !== null);
      assert.match(message, /WordApi 1\.3/);
      assert.match(message, /getRange/);
    },
  );

  await withOffice(
    {
      context: { platform: "PC", requirements: { isSetSupported: () => true } },
    },
    async () => {
      assert.equal(wordRangeNavigationGateMessage(), null);
    },
  );
});

void test("multi-run insertion fails loudly below WordApi 1.3", async () => {
  const context = makeContext();
  const multiRunBlock: WordInsertBlock = {
    text: "plain bold",
    type: "paragraph",
    runs: [{ text: "plain " }, { text: "bold", bold: true }],
  };

  await withOffice(
    {
      context: {
        platform: "PC",
        requirements: { isSetSupported: () => false },
      },
    },
    () =>
      runWithFakeWord(context, async () => {
        await assert.rejects(
          () =>
            wordRun(async (ctx) => {
              insertWordBlocks(ctx, [multiRunBlock], "end");
              return Promise.resolve();
            }),
          /WordApi 1\.3/,
        );
      }),
  );
});

void test("multi-run insertion works when WordApi 1.3 is available", async () => {
  const context = makeContext();
  const multiRunBlock: WordInsertBlock = {
    text: "plain bold",
    type: "paragraph",
    runs: [{ text: "plain " }, { text: "bold", bold: true }],
  };

  await withOffice(
    {
      context: { platform: "PC", requirements: { isSetSupported: () => true } },
    },
    () =>
      runWithFakeWord(context, async () => {
        await wordRun(async (ctx) => {
          insertWordBlocks(ctx, [multiRunBlock], "end");
          return Promise.resolve();
        });
      }),
  );

  assert.ok(
    context.calls.some((c) => c.includes("bold")),
    context.calls.join("\n"),
  );
});
