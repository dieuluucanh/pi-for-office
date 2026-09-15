#!/usr/bin/env node
/**
 * Package-contents guard — verifies that every relative module reachable from a
 * publishable package's entry points is actually inside the tarball that
 * `npm publish` would upload.
 *
 * Why this exists: the Pi bridge extension ships raw TypeScript and is loaded
 * from `pi.extensions`. Its `files` allowlist once omitted two runtime modules,
 * producing a published package that failed on load with
 * `Cannot find module './pane-replies.js'`. Local typecheck/tests resolve files
 * from disk, so only this artifact-level check catches that class of bug.
 *
 * Usage:
 *   node scripts/check-package-contents.mjs <packageDir> [<packageDir> ...]
 *
 * Exits non-zero and lists each missing `importer → specifier` on failure.
 */

import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { execPath } from "node:process";
import { fileURLToPath } from "node:url";

/** On Windows npm is a .cmd shim; invoke npm's JS entry via node instead. */
function npmSpawn() {
  if (process.platform !== "win32") return { cmd: "npm", args: [] };
  const candidate = path.join(
    path.dirname(execPath),
    "node_modules",
    "npm",
    "bin",
    "npm-cli.js",
  );
  if (existsSync(candidate)) return { cmd: execPath, args: [candidate] };
  return { cmd: "npm.cmd", args: [] };
}

/** Normalize a package-relative path to forward slashes, no leading ./ */
function normalize(p) {
  return String(p ?? "")
    .replace(/\\/g, "/")
    .replace(/^\.\//, "")
    .replace(/^\/+/, "");
}

/** Run `npm pack --dry-run --json` and return the packed relative paths. */
export function packDryRun(packageDir) {
  const { cmd, args: pre } = npmSpawn();
  const res = spawnSync(cmd, [...pre, "pack", "--dry-run", "--json"], {
    cwd: packageDir,
    encoding: "utf8",
    ...(process.platform === "win32" && cmd === "npm.cmd"
      ? { shell: true }
      : {}),
  });
  if (res.status !== 0) {
    throw new Error(
      `npm pack --dry-run failed in ${packageDir}: ` +
        (res.stderr || res.stdout || "(no output)"),
    );
  }
  // A `prepack` script may print progress lines before the JSON payload.
  const lines = String(res.stdout ?? "").split("\n");
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    if (!t.startsWith("[") && !t.startsWith("{")) continue;
    try {
      const parsed = JSON.parse(lines.slice(i).join("\n"));
      const entry = Array.isArray(parsed) ? parsed[0] : parsed;
      const files = Array.isArray(entry?.files) ? entry.files : [];
      return files.map((f) => normalize(typeof f === "string" ? f : f.path));
    } catch {
      // Not the JSON start — keep scanning.
    }
  }
  throw new Error(`could not parse npm pack --json output in ${packageDir}`);
}

const SPECIFIER_RE =
  /(?:import|export)\s+(?:[^'"]*?\s+from\s+)?["']([^"']+)["']|import\s*\(\s*["']([^"']+)["']\s*\)|require\s*\(\s*["']([^"']+)["']\s*\)/g;

/** Relative import/require specifiers in a source file. */
function relativeSpecifiers(source) {
  const out = [];
  for (const match of source.matchAll(SPECIFIER_RE)) {
    const spec = match[1] ?? match[2] ?? match[3];
    if (spec && (spec.startsWith("./") || spec.startsWith("../"))) out.push(spec);
  }
  return out;
}

/** Runtime entry files declared by a package.json. */
function entryFiles(pkg) {
  const entries = [];
  const push = (value) => {
    if (typeof value === "string") entries.push(normalize(value));
  };
  if (Array.isArray(pkg.pi?.extensions)) pkg.pi.extensions.forEach(push);
  push(pkg.main);
  push(pkg.module);
  if (typeof pkg.bin === "string") push(pkg.bin);
  else if (pkg.bin && typeof pkg.bin === "object") {
    Object.values(pkg.bin).forEach(push);
  }
  const walkExports = (value) => {
    if (typeof value === "string") push(value);
    else if (Array.isArray(value)) value.forEach(walkExports);
    else if (value && typeof value === "object") {
      Object.values(value).forEach(walkExports);
    }
  };
  if (pkg.exports) walkExports(pkg.exports);
  return [...new Set(entries)];
}

/** Resolve one relative specifier to a packed or on-disk file. */
function resolveSpecifier(importerRel, spec, packedSet, packageDir) {
  const base = path.posix.normalize(
    path.posix.join(path.posix.dirname(importerRel), spec),
  );
  const candidates = [base];
  if (base.endsWith(".js")) {
    // Pi's TS loader rewrites .js specifiers to .ts sources.
    candidates.push(base.replace(/\.js$/, ".ts"));
    candidates.push(base.replace(/\.js$/, ".tsx"));
  }
  if (!/\.[a-z0-9]+$/i.test(base)) {
    for (const ext of [".ts", ".tsx", ".js", ".mjs", ".cjs"]) {
      candidates.push(`${base}${ext}`);
    }
    candidates.push(
      `${base}/index.ts`,
      `${base}/index.js`,
      `${base}/index.mjs`,
    );
  }
  for (const candidate of candidates) {
    if (packedSet.has(candidate)) return { rel: candidate, packed: true };
  }
  for (const candidate of candidates) {
    const abs = path.join(packageDir, ...candidate.split("/"));
    if (existsSync(abs) && statSync(abs).isFile()) {
      return { rel: candidate, packed: false };
    }
  }
  return null;
}

/**
 * Verify every relative module reachable from the package entry points is
 * present in `packedPaths`. Returns `{ ok, missing, checked }`.
 */
export function verifyPackageContents(packageDir, packedPaths) {
  const pkg = JSON.parse(
    readFileSync(path.join(packageDir, "package.json"), "utf8"),
  );
  const packedSet = new Set(packedPaths.map(normalize));
  const visited = new Set();
  const queue = entryFiles(pkg);
  const missing = [];
  let checked = 0;

  while (queue.length > 0) {
    const rel = queue.shift();
    if (!rel || visited.has(rel)) continue;
    visited.add(rel);

    if (!packedSet.has(rel)) {
      // Only report a missing entry when it exists locally (a platform-specific
      // optional entry may legitimately be absent from this platform's pack).
      const abs = path.join(packageDir, ...rel.split("/"));
      if (existsSync(abs) && statSync(abs).isFile()) {
        missing.push({ from: "(entry)", specifier: rel });
      }
      continue;
    }

    let source = "";
    try {
      source = readFileSync(
        path.join(packageDir, ...rel.split("/")),
        "utf8",
      );
    } catch {
      continue;
    }
    checked += 1;

    for (const spec of relativeSpecifiers(source)) {
      const resolved = resolveSpecifier(rel, spec, packedSet, packageDir);
      if (resolved === null || !resolved.packed) {
        missing.push({ from: rel, specifier: spec });
        continue;
      }
      queue.push(resolved.rel);
    }
  }

  return { ok: missing.length === 0, missing, checked };
}

/** CLI: check one or more package directories (relative to cwd). */
export function runCli(argv) {
  if (argv.length === 0) {
    console.error(
      "usage: node scripts/check-package-contents.mjs <packageDir> [...]",
    );
    process.exit(2);
  }
  let failed = false;
  for (const dir of argv) {
    const packageDir = path.resolve(process.cwd(), dir);
    try {
      const packed = packDryRun(packageDir);
      const result = verifyPackageContents(packageDir, packed);
      if (result.ok) {
        console.log(
          `[package-contents] ok: ${dir} ` +
            `(${result.checked} module(s) reachable, ${packed.length} file(s) packed)`,
        );
      } else {
        failed = true;
        console.error(
          `[package-contents] FAIL: ${dir} — ${result.missing.length} module(s) missing from the tarball:`,
        );
        for (const m of result.missing) {
          console.error(`   ${m.from} → ${m.specifier}`);
        }
      }
    } catch (error) {
      failed = true;
      console.error(
        `[package-contents] error in ${dir}: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }
  process.exit(failed ? 1 : 0);
}

const invokedPath = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (invokedPath === fileURLToPath(import.meta.url)) {
  runCli(process.argv.slice(2));
}
