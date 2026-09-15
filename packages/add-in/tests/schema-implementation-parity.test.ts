/**
 * Schema ↔ implementation parity guard.
 *
 * The incident this guards against: `word_format_range` advertised a
 * `highlight` property in its schema, but `setWordFontProps` never assigned
 * `font.highlightColor` — so the agent promised a fill it could not apply.
 *
 * The guard reads the **implementation source** (the same file the runtime
 * calls) and the **advertised schema** (the live TypeBox object), never a
 * hand-copied property list. Delete an assignment and the corresponding test
 * fails, which is the mutation check the plan requires.
 */

import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";

import {
  EXCEL_CONDITIONAL_FORMAT_PARAMETERS,
  EXCEL_FORMAT_CELLS_PARAMETERS,
  POWERPOINT_ADD_TEXT_BOX_PARAMETERS,
  POWERPOINT_FORMAT_SLIDE_PARAMETERS,
  WORD_FONT_PROP_LABELS,
  WORD_PARAGRAPH_PROP_LABELS,
} from "@dieulc/pi-office-protocol/office-catalog";

/** Read a source file relative to this test file. */
async function implementationSource(relativePath: string): Promise<string> {
  return readFile(new URL(relativePath, import.meta.url), "utf8");
}

/**
 * Strip comments so documentation prose cannot satisfy the guard.
 * (A JSDoc line mentioning `WordTextFormat.highlight` once satisfied the
 * `format.highlight` check, which is exactly the false pass to avoid.)
 */
function withoutComments(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\//gu, "").replace(/\/\/[^\n]*/gu, "");
}

/**
 * True when `source` reads/assigns `member` for real.
 *
 * Accepts `member !==`, `member ===`, or a bare `member =` that is not part of
 * a comparison — so a comment or a string literal cannot fake it.
 */
function readsOrAssigns(source: string, member: string): boolean {
  if (source.includes(`${member} !==`) || source.includes(`${member} ===`)) {
    return true;
  }
  let index = source.indexOf(`${member} =`);
  while (index !== -1) {
    if (source[index + member.length + 2] !== "=") return true;
    index = source.indexOf(`${member} =`, index + 1);
  }
  return false;
}

/** Property names a TypeBox object schema advertises. */
function schemaKeys(schema: unknown): string[] {
  const properties = (schema as { properties?: Record<string, unknown> })
    .properties;
  return Object.keys(properties ?? {});
}

/** PowerPoint params that are not formatting (layout / targeting). */
const POWERPOINT_CONTROL_KEYS = new Set([
  "slideIndex",
  "text",
  "x",
  "y",
  "width",
  "height",
]);

void test("every Word formatting property is assigned by a Word helper", async () => {
  const helper = withoutComments(
    await implementationSource("../src/word/helpers.ts"),
  );
  // Derived from the catalog's label maps — the same keys WORD_FONT_PROPS /
  // WORD_PARAGRAPH_PROPS declare, so a new schema property is included
  // automatically.
  const formattingKeys = [
    ...Object.keys(WORD_FONT_PROP_LABELS),
    ...Object.keys(WORD_PARAGRAPH_PROP_LABELS),
  ];
  assert.ok(
    formattingKeys.includes("highlight"),
    "expected the highlight prop",
  );

  const missing = formattingKeys.filter(
    (key) => !readsOrAssigns(helper, `format.${key}`),
  );
  assert.deepEqual(
    missing,
    [],
    `Word properties with no assignment path in word/helpers.ts: ${missing.join(", ")}`,
  );
});

void test("every PowerPoint formatting property is assigned by a PPT helper", async () => {
  const helper = withoutComments(
    await implementationSource("../src/powerpoint/helpers.ts"),
  );
  const keys = [
    ...schemaKeys(POWERPOINT_ADD_TEXT_BOX_PARAMETERS),
    ...schemaKeys(POWERPOINT_FORMAT_SLIDE_PARAMETERS),
  ].filter((key) => !POWERPOINT_CONTROL_KEYS.has(key));

  const missing = [...new Set(keys)].filter(
    (key) => !readsOrAssigns(helper, `.${key}`),
  );
  assert.deepEqual(
    missing,
    [],
    `PowerPoint properties with no assignment path in powerpoint/helpers.ts: ${missing.join(", ")}`,
  );
});

void test("every format_cells property is read by the format_cells implementation", async () => {
  // Border params are normalised in their own module; both files are part of
  // the implementation surface.
  const sources = (
    await Promise.all([
      implementationSource("../src/tools/format-cells.ts"),
      implementationSource("../src/tools/format-cells-borders.ts"),
    ])
  ).map((source) => withoutComments(source));
  const missing = schemaKeys(EXCEL_FORMAT_CELLS_PARAMETERS).filter(
    (key) => !sources.some((source) => source.includes(`params.${key}`)),
  );
  assert.deepEqual(
    missing,
    [],
    `format_cells properties never read by format-cells.ts: ${missing.join(", ")}`,
  );
});

void test("every conditional_format property is read by the conditional_format implementation", async () => {
  const implementation = withoutComments(
    await implementationSource("../src/tools/conditional-format.ts"),
  );
  const missing = schemaKeys(EXCEL_CONDITIONAL_FORMAT_PARAMETERS).filter(
    (key) => !implementation.includes(`params.${key}`),
  );
  assert.deepEqual(
    missing,
    [],
    `conditional_format properties never read by conditional-format.ts: ${missing.join(", ")}`,
  );
});

void test("the Word guard is not vacuous (highlight has a real assignment)", async () => {
  // Pins the exact regression from the incident: a matcher that silently
  // matched nothing would pass every test above.
  const helper = withoutComments(
    await implementationSource("../src/word/helpers.ts"),
  );
  assert.ok(readsOrAssigns(helper, "format.highlight"));
  assert.ok(helper.includes("font.highlightColor ="));
});
