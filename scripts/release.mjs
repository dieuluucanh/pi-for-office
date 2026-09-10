#!/usr/bin/env node
/**
 * Release helper — one tracked flow for publishing the npm packages and
 * verifying the GitHub Pages add-in deployment.
 *
 *   node scripts/release.mjs status          # read-only: versions + npm auth
 *   node scripts/release.mjs check           # typecheck/tests + pack --dry-run
 *   node scripts/release.mjs publish [opts]  # bump + publish in dependency order
 *
 * Publish order (dependency-aware):
 *   1. @dieulc/pi-office-protocol   (packages/protocol)
 *   2. @dieulc/pi-office-bridge     (packages/bridge-extension)
 *   3. pi-for-office-proxy          (packages/add-in/pkg/proxy)
 *   4. pi-for-office-python-bridge  (packages/add-in/pkg/python-bridge)
 *   5. pi-for-office-tmux-bridge    (packages/add-in/pkg/tmux-bridge)
 *
 * The add-in itself (packages/add-in) is NOT published to npm — it deploys
 * via GitHub Pages on push to main (see docs/releasing.md).
 *
 * Options for `publish`:
 *   --only <pkgName>          publish just one package
 *   --bump <patch|minor|major> bump version before publishing (all or --only)
 *   --tag <dist-tag>          publish under a custom dist-tag
 *   --otp <code>              npm one-time password (2FA)
 *   --allow-dirty             skip the git-clean check
 *   --skip-checks             skip the `check` phase
 *   --yes                     skip the interactive confirmation
 */

import { spawnSync } from "node:child_process";
import { readFileSync, existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execPath } from "node:process";
import { createInterface } from "node:readline";

const rootDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

// On Windows npm is a .cmd shim; spawning it through a shell is deprecated
// in modern Node, so we invoke npm's JS entry via the current node instead.
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
  // Fallback (e.g. nvm/volta shims): tolerate the shell deprecation.
  return { cmd: "npm.cmd", args: [] };
}

function run(cmd, args, opts = {}) {
  const res = spawnSync(cmd, args, {
    cwd: opts.cwd ?? rootDir,
    encoding: "utf8",
    ...(process.platform === "win32" && cmd === "npm.cmd"
      ? { shell: true }
      : {}),
    env: { ...process.env, ...(opts.env ?? {}) },
    stdio: ["pipe", "pipe", "pipe"],
  });
  return res;
}

const PACKAGES = [
  {
    name: "@dieulc/pi-office-protocol",
    dir: "packages/protocol",
    scoped: true,
  },
  {
    name: "@dieulc/pi-office-bridge",
    dir: "packages/bridge-extension",
    scoped: true,
  },
  {
    name: "pi-for-office-proxy",
    dir: "packages/add-in/pkg/proxy",
    scoped: false,
  },
  {
    name: "pi-for-office-python-bridge",
    dir: "packages/add-in/pkg/python-bridge",
    scoped: false,
  },
  {
    name: "pi-for-office-tmux-bridge",
    dir: "packages/add-in/pkg/tmux-bridge",
    scoped: false,
  },
];

function npm(args, opts = {}) {
  const { cmd, args: pre } = npmSpawn();
  return run(cmd, [...pre, ...args], opts);
}

function pkgJsonPackage(dir) {
  try {
    const raw = readFileSync(path.join(rootDir, dir, "package.json"), "utf8");
    return JSON.parse(raw);
  } catch (error) {
    console.error(
      `[release] cannot read package.json at ${dir}:`,
      error instanceof Error ? error.message : error,
    );
    process.exit(1);
  }
}

function stripAnsi(s) {
  // eslint-disable-next-line no-control-regex
  return String(s ?? "").replace(/\u001B\[[0-9;]*m/g, "");
}

function errText(res) {
  const msg = (res?.stderr ?? "") || (res?.stdout ?? "");
  return stripAnsi(msg).trim() || "(no output)";
}

function confirm(prompt) {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => {
    rl.question(`${prompt} [y/N] `, (answer) => {
      rl.close();
      resolve(answer.trim().toLowerCase() === "y");
    });
  });
}

// ── status ────────────────────────────────────────────────────────────

async function status() {
  console.log("\n── npm identity ──\n");
  const whoami = npm(["whoami"]);
  if (whoami.status === 0) {
    console.log(`logged in as: ${stripAnsi(whoami.stdout ?? "").trim()}`);
  } else {
    console.log(`⚠  not logged in: ${errText(whoami)}`);
  }

  console.log("\n── packages ──\n");
  const rows = [];
  for (const pkg of PACKAGES) {
    const local = pkgJsonPackage(pkg.dir).version;
    const view = npm(["view", pkg.name, "version", "--json"]);
    let latest = "(never published)";
    if (view.status === 0) {
      try {
        const parsed = JSON.parse(view.stdout);
        latest = Array.isArray(parsed) ? parsed.join(", ") : String(parsed);
      } catch {
        latest = stripAnsi(view.stdout);
      }
    }
    const publishedHere =
      npm(["view", `${pkg.name}@${local}`, "version"]).status === 0;
    rows.push({
      pkg: pkg.name,
      local,
      latest: latest.length > 40 ? `${latest.slice(0, 40)}…` : latest,
      publishedHere: publishedHere ? "yes" : "no",
    });
  }
  const w = (s, n) => String(s).padEnd(n);
  console.log(
    `${w("package", 32)}${w("local", 8)}${w("npm latest", 44)}published@local`,
  );
  for (const r of rows) {
    console.log(
      `${w(r.pkg, 32)}${w(r.local, 8)}${w(r.latest, 44)}${r.publishedHere}`,
    );
  }

  console.log("\n── add-in deployment ──\n");
  console.log("add-in (packages/add-in) is NOT npm-published.");
  console.log(
    "It deploys via GitHub Pages (push to main → https://dieuluucanh.github.io/pi-for-office/).",
  );
  console.log("See docs/releasing.md.");
}

// ── check ─────────────────────────────────────────────────────────────

async function check({ only, verb = "check" }) {
  const targets = only ? PACKAGES.filter((p) => p.name === only) : PACKAGES;
  if (targets.length === 0) {
    console.error(`Unknown package: ${only}`);
    process.exit(1);
  }

  console.log(`\n── ${verb}: root typecheck ──\n`);
  const tc = npm(["run", "typecheck"]);
  if (tc.status !== 0) {
    console.error(errText(tc));
    console.error(`[release] root typecheck failed — aborting`);
    process.exit(1);
  }
  console.log("[release] root typecheck ok");

  // bridge-extension tests (unit-level for the wire protocol + /health)
  const bridge = PACKAGES.find((p) => p.name === "@dieulc/pi-office-bridge");
  if (!only || only === bridge.name) {
    console.log(`\n── ${verb}: @dieulc/pi-office-bridge tests ──\n`);
    const test = npm(["test"], { cwd: path.join(rootDir, bridge.dir) });
    if (test.status !== 0) {
      console.error(errText(test));
      console.error(`[release] bridge-extension tests failed — aborting`);
      process.exit(1);
    }
    console.log("[release] bridge-extension tests ok");
  }

  console.log(`\n── ${verb}: pack --dry-run for publishables ──\n`);
  for (const pkg of targets) {
    const pack = npm(["pack", "--dry-run", "--json"], {
      cwd: path.join(rootDir, pkg.dir),
    });
    if (pack.status !== 0) {
      console.error(errText(pack));
      console.error(
        `[release] pack --dry-run failed for ${pkg.name} — aborting`,
      );
      process.exit(1);
    }
    try {
      // prepack sync scripts print progress lines before the JSON array;
      // scan for the first line that actually parses as JSON.
      const lines = (pack.stdout ?? "").split("\n");
      let parsed = null;
      for (let i = 0; i < lines.length; i++) {
        const t = lines[i].trim();
        if (t.startsWith("[") || t.startsWith("{")) {
          try {
            parsed = JSON.parse(lines.slice(i).join("\n"));
            break;
          } catch {
            // Not the JSON start — keep scanning.
          }
        }
      }
      if (!parsed) throw new Error("no JSON in pack output");
      const entry = Array.isArray(parsed) ? parsed[0] : parsed;
      const files = Array.isArray(entry.files) ? entry.files : [];
      const names = files
        .map((f) => (typeof f === "string" ? f : f.path))
        .slice(0, 30);
      console.log(
        `${pkg.name}@${pkgJsonPackage(pkg.dir).version}: ${names.length} file(s)`,
      );
      for (const n of names) console.log(`   ${n}`);
    } catch {
      const tail = stripAnsi(pack.stdout ?? "")
        .split("\n")
        .slice(-3)
        .join(" | ");
      console.log(`${pkg.name}: (no --json output) ${tail}`);
    }
  }
}

// ── publish ───────────────────────────────────────────────────────────

async function publish({ only, bump, tag, otp, allowDirty, skipChecks, yes }) {
  // 1. Auth
  const whoami = npm(["whoami"]);
  if (whoami.status !== 0) {
    console.error("[release] not logged in to npm — run `npm login` first");
    process.exit(1);
  }
  console.log(
    `[release] publishing as ${stripAnsi(whoami.stdout ?? "").trim()}`,
  );

  // 2. Git clean
  if (!allowDirty) {
    const dirty = run("git", ["status", "--porcelain"]);
    if (dirty.status === 0 && dirty.stdout && dirty.stdout.trim().length > 0) {
      console.error(
        "[release] working tree is dirty — commit first or pass --allow-dirty",
      );
      process.exit(1);
    }
  }

  // 3. Pre-publish checks unless opted out
  if (!skipChecks) {
    await check({ only, verb: "preflight check" });
  }

  const targets = only ? PACKAGES.filter((p) => p.name === only) : PACKAGES;
  if (targets.length === 0) {
    console.error(`Unknown package: ${only}`);
    process.exit(1);
  }

  // 4. Version bump
  if (bump) {
    if (!["patch", "minor", "major"].includes(bump)) {
      console.error("--bump must be one of: patch | minor | major");
      process.exit(1);
    }
    for (const pkg of targets) {
      const v = npm(["version", bump, "--no-git-tag-version"], {
        cwd: path.join(rootDir, pkg.dir),
      });
      if (v.status !== 0) {
        console.error(
          `[release] version bump failed for ${pkg.name}:`,
          errText(v),
        );
        process.exit(1);
      }
      console.log(
        `[release] ${pkg.name} → ${stripAnsi(v.stdout ?? "").trim()}`,
      );
    }
  }

  // 5. Plan + confirm
  console.log("\n── publish plan ──\n");
  for (const pkg of targets) {
    const version = pkgJsonPackage(pkg.dir).version;
    console.log(
      `  npm publish ${pkg.name}@${version}${tag ? ` --tag ${tag}` : ""}${scopedAccess(pkg) ? " --access public" : ""}`,
    );
  }
  if (!yes) {
    const ok = await confirm("\nPublish these packages to npm?");
    if (!ok) {
      console.log("[release] aborted by user");
      process.exit(0);
    }
  }

  // 6. Publish in order
  for (const pkg of targets) {
    const version = pkgJsonPackage(pkg.dir).version;

    // Skip already-published versions
    const already = npm(["view", `${pkg.name}@${version}`, "version"]);
    if (already.status === 0) {
      console.log(
        `[release] ${pkg.name}@${version} already published — skipping`,
      );
      continue;
    }

    const args = ["publish"];
    if (tag) args.push("--tag", tag);
    if (otp) args.push("--otp", otp);
    const pub = npm(args, { cwd: path.join(rootDir, pkg.dir) });
    if (pub.status !== 0) {
      console.error(`[release] publish failed for ${pkg.name}:`, errText(pub));
      process.exit(1);
    }
    console.log(`[release] published ${pkg.name}@${version}`);
  }

  // 7. Verify
  console.log("\n── verification ──\n");
  for (const pkg of targets) {
    const version = pkgJsonPackage(pkg.dir).version;
    const v = npm(["view", `${pkg.name}@${version}`, "version"]);
    console.log(
      `  ${pkg.name}@${version} → ${v.status === 0 ? "visible on npm ✓" : "NOT VISIBLE ✗"}`,
    );
  }

  console.log(
    "\n[release] done. See docs/releasing.md for post-publish verification.",
  );
}

function scopedAccess(pkg) {
  return (
    pkg.scoped && pkgJsonPackage(pkg.dir).publishConfig?.access === "public"
  );
}

// ── CLI ───────────────────────────────────────────────────────────────

const args = process.argv.slice(2);
const command = args[0];

function flagValue(name) {
  const idx = args.indexOf(name);
  return idx >= 0 ? args[idx + 1] : undefined;
}
function hasFlag(name) {
  return args.includes(name);
}

switch (command) {
  case "status":
    await status();
    break;
  case "check":
    await check({ only: flagValue("--only") });
    break;
  case "publish":
    await publish({
      only: flagValue("--only"),
      bump: flagValue("--bump"),
      tag: flagValue("--tag"),
      otp: flagValue("--otp"),
      allowDirty: hasFlag("--allow-dirty"),
      skipChecks: hasFlag("--skip-checks"),
      yes: hasFlag("--yes"),
    });
    break;
  case "help":
  case "--help":
  case "-h":
    console.log(
      `usage: node scripts/release.mjs <status|check|publish> [options]`,
    );
    console.log(`\nstatus   : read-only version + auth report`);
    console.log(
      `check    : typecheck + bridge tests + pack --dry-run for publishables`,
    );
    console.log(`publish  : bump (optional) + publish in dependency order`);
    console.log(`\npublish options:`);
    console.log(`  --only <name>           publish one package`);
    console.log(`  --bump <patch|minor|major>`);
    console.log(`  --tag <dist-tag>`);
    console.log(`  --otp <code>`);
    console.log(`  --allow-dirty | --skip-checks | --yes`);
    break;
  default:
    console.error(
      `unknown command: ${command ?? "(none)"} — use status | check | publish`,
    );
    process.exit(1);
}
