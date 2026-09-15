/**
 * Capability ↔ schema consistency guard.
 *
 * A capability declared in `*_CAPABILITIES` must be *reachable*: either a
 * caller gates a tool parameter with it (naming the id outside
 * `capabilities.ts`), or it is an explicit probe-only entry. A declared
 * capability nothing consults is a claim the capability report makes that the
 * tool surface does not back.
 */

import assert from "node:assert/strict";
import { readdir, readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { join } from "node:path";
import { test } from "node:test";

import {
  EXCEL_CAPABILITIES,
  POWERPOINT_CAPABILITIES,
  PROBE_ONLY_CAPABILITY_IDS,
  WORD_CAPABILITIES,
  type CapabilitySpec,
} from "../src/office/capabilities.ts";

const ALL_CAPABILITIES: readonly CapabilitySpec[] = [
  ...WORD_CAPABILITIES,
  ...EXCEL_CAPABILITIES,
  ...POWERPOINT_CAPABILITIES,
];

async function collectTypeScriptFiles(dir: string): Promise<string[]> {
  const entries = await readdir(dir, { withFileTypes: true });
  const files: string[] = [];
  for (const entry of entries) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await collectTypeScriptFiles(full)));
    } else if (entry.isFile() && entry.name.endsWith(".ts")) {
      files.push(full);
    }
  }
  return files;
}

/** Strip comments so prose cannot satisfy the reachability check. */
function withoutComments(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\//gu, "").replace(/\/\/[^\n]*/gu, "");
}

/** Every `src/**.ts` file except the registry that declares the ids. */
async function callerSources(): Promise<string> {
  const root = fileURLToPath(new URL("../src", import.meta.url));
  const files = await collectTypeScriptFiles(root);
  const chunks: string[] = [];
  for (const file of files) {
    // Normalise separators: on Windows `fileURLToPath` yields backslashes, so
    // a plain `endsWith("office/capabilities.ts")` would NOT exclude the
    // declaring registry and the guard would pass vacuously.
    const normalised = file.replace(/\\/gu, "/");
    if (normalised.endsWith("office/capabilities.ts")) continue;
    chunks.push(withoutComments(await readFile(file, "utf8")));
  }
  return chunks.join("\n");
}

void test("every declared capability is gated by a caller or marked probe-only", async () => {
  const sources = await callerSources();

  const unreachable = ALL_CAPABILITIES.filter(
    (spec) =>
      !PROBE_ONLY_CAPABILITY_IDS.has(spec.id) &&
      !sources.includes(`"${spec.id}"`),
  ).map((spec) => spec.id);

  assert.deepEqual(
    unreachable,
    [],
    `capabilities nothing consumes (add a gate or mark probe-only): ${unreachable.join(", ")}`,
  );
});

void test("probe-only ids are real capability ids", () => {
  const known = new Set(ALL_CAPABILITIES.map((spec) => spec.id));
  const unknown = [...PROBE_ONLY_CAPABILITY_IDS].filter((id) => !known.has(id));
  assert.deepEqual(
    unknown,
    [],
    `probe-only ids with no spec: ${unknown.join(", ")}`,
  );
});

void test("capability ids are unique and host-prefixed", () => {
  const seen = new Set<string>();
  for (const spec of ALL_CAPABILITIES) {
    assert.ok(!seen.has(spec.id), `duplicate capability id: ${spec.id}`);
    seen.add(spec.id);
    assert.match(spec.id, /^(word|excel|powerpoint)\./u);
  }
});

void test("the reachability guard is not vacuous (a gated id is found, a typo is not)", async () => {
  const sources = await callerSources();
  // `word.table.insert` is gated in word/helpers.ts.
  assert.ok(sources.includes(`"word.table.insert"`));
  assert.ok(!PROBE_ONLY_CAPABILITY_IDS.has("word.table.insert"));
  // A typo'd id must not resolve.
  assert.ok(!sources.includes(`"word.table.inserted"`));
});
