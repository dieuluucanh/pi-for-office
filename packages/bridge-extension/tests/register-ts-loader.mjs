/**
 * Register the add-in's test-time ESM resolver for plain-`node` runners in
 * this package.
 *
 * The interop test imports add-in TypeScript sources directly (e.g.
 * `pane-client.ts`), which use explicit ".js" import specifiers for the
 * bundler. Under plain `node` those specifiers must retry as ".ts" — exactly
 * the job of `test-ts-import-loader.mjs` — so we register it before the test
 * module graph resolves.
 *
 * Run: node --import ./tests/register-ts-loader.mjs tests/pane-interop.mjs
 */
import { register } from "node:module";

const loaderUrl = new URL(
  "../../add-in/scripts/test-ts-import-loader.mjs",
  import.meta.url,
);
register(loaderUrl.href);