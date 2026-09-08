import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";

void test("oauth storage does not call localStorage APIs", async () => {
  const source = await readFile(new URL("../src/auth/oauth-storage.ts", import.meta.url), "utf8");
  const localStorageApiPattern = /\blocalStorage\s*\.\s*(?:getItem|setItem|removeItem)\b/u;

  assert.equal(localStorageApiPattern.test(source), false);
});
