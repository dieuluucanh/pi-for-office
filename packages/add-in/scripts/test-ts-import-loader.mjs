/**
 * Test-time ESM resolver for Node's --experimental-strip-types runner.
 *
 * Source files use explicit ".js" import specifiers (for browser/bundler output),
 * while tests execute TypeScript sources directly. This loader retries relative
 * ".js" specifiers as ".ts" so node --test can resolve source modules.
 *
 * It also supports Vite-style `?raw` text imports for local markdown files used
 * by bundled docs/skills catalogs, and stubs Vite asset imports (images) that
 * the bundler would otherwise turn into hashed URLs — Node has no loader for a
 * `.png` and would abort the whole test file with ERR_UNKNOWN_FILE_EXTENSION.
 */

import { readFile } from "node:fs/promises";

/** Extensions Vite resolves to an asset URL instead of executing. */
const ASSET_EXTENSIONS = [
  ".png",
  ".jpg",
  ".jpeg",
  ".gif",
  ".svg",
  ".webp",
  ".bmp",
  ".ico",
  ".avif",
];

function assetStub(url) {
  const withoutQuery = url.split("?")[0].split("#")[0];
  const extension = ASSET_EXTENSIONS.find((candidate) =>
    withoutQuery.toLowerCase().endsWith(candidate),
  );
  if (extension === undefined) return null;
  // Mirrors what the bundler hands the module graph: a URL string.
  return `export default ${JSON.stringify(withoutQuery)};\n`;
}

function hasCode(error) {
  return typeof error === "object" && error !== null && "code" in error;
}

function toRawFileUrl(rawUrl) {
  const url = new URL(rawUrl);
  url.search = "";
  url.hash = "";
  return url;
}

export async function resolve(specifier, context, defaultResolve) {
  try {
    return await defaultResolve(specifier, context, defaultResolve);
  } catch (error) {
    const isModuleNotFound =
      hasCode(error) && error.code === "ERR_MODULE_NOT_FOUND";
    const isRelative =
      specifier.startsWith("./") || specifier.startsWith("../");

    if (!isModuleNotFound || !isRelative) {
      throw error;
    }

    const [pathPart, queryPart] = specifier.split("?", 2);
    if (!pathPart.endsWith(".js")) {
      throw error;
    }

    const suffix = queryPart ? `?${queryPart}` : "";
    const tsSpecifier = `${pathPart.slice(0, -3)}.ts${suffix}`;
    return defaultResolve(tsSpecifier, context, defaultResolve);
  }
}

export async function load(url, context, defaultLoad) {
  if (url.endsWith("?raw")) {
    const fileUrl = toRawFileUrl(url);
    const sourceText = await readFile(fileUrl, "utf8");
    const source = `export default ${JSON.stringify(sourceText)};`;
    return {
      format: "module",
      shortCircuit: true,
      source,
    };
  }

  const assetSource = assetStub(url);
  if (assetSource !== null) {
    return {
      format: "module",
      shortCircuit: true,
      source: assetSource,
    };
  }

  return defaultLoad(url, context, defaultLoad);
}
