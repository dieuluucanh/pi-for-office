/**
 * Architecture page fact guard.
 *
 * The public architecture page (`public/architecture/index.html`, served at
 * https://dieuluucanh.github.io/pi-for-office/architecture/) states concrete
 * facts about this codebase: tool counts, capability counts, bridge ports,
 * limits, compaction ratios, test-suite counts, and so on.
 *
 * Every such fact is bound to the code with a marker:
 *
 *     <span data-fact="tools.core">17</span>
 *
 * This test derives each expected value from the sources (constants, catalogs,
 * directory listings, package scripts) and fails when the page drifts. It also
 * enforces the page's link/anchor contract and forbids retired terminology.
 *
 * Run: `npm run check:architecture-facts`
 * Maintainer guide: `docs/architecture-page.md`
 */

import assert from "node:assert/strict";
import { readdir, readFile } from "node:fs/promises";
import { test } from "node:test";

const PAGE_URL = new URL("../public/architecture/index.html", import.meta.url);

/** `<span data-fact="key">value</span>` — value is the whole inner text. */
const FACT_PATTERN = /<span[^>]*\bdata-fact="([^"]+)"[^>]*>([^<]*)<\/span>/gu;

async function readSource(relativePath: string): Promise<string> {
  return readFile(new URL(`../${relativePath}`, import.meta.url), "utf8");
}

async function readPage(): Promise<string> {
  return readFile(PAGE_URL, "utf8");
}

/** Capture group 1 of a pattern; fails loudly when the source shape changed. */
function capture(source: string, pattern: RegExp, label: string): string {
  const match = pattern.exec(source);
  const value = match?.[1];
  if (value === undefined) {
    throw new Error(
      `Fact guard: could not derive ${label}. Pattern ${String(pattern)} did not match its source; ` +
        "update the guard (and docs/architecture-page.md) before changing the source shape.",
    );
  }
  return value;
}

function toInt(value: string): number {
  return Number.parseInt(value.replaceAll("_", ""), 10);
}

function captureInt(source: string, pattern: RegExp, label: string): number {
  return toInt(capture(source, pattern, label));
}

/** Count matches of a global pattern. */
function countMatches(source: string, pattern: RegExp, label: string): number {
  const matches = [...source.matchAll(pattern)];
  if (matches.length === 0) {
    throw new Error(
      `Fact guard: counted 0 matches for ${label} (pattern ${String(pattern)}).`,
    );
  }
  return matches.length;
}

function requireSnippets(
  source: string,
  snippets: readonly string[],
  label: string,
): void {
  const missing = snippets.filter((snippet) => !source.includes(snippet));
  assert.equal(
    missing.length,
    0,
    `${label} is missing expected entries: ${missing.join(", ")}`,
  );
}

/** Thousands separators, matching the page's display format. */
function formatInt(value: number): string {
  return String(value).replace(/\B(?=(\d{3})+(?!\d))/gu, ",");
}

function formatMegabytes(bytes: number): string {
  return `${bytes / 1_000_000} MB`;
}

/**
 * Non-core built-in tool factories in `src/tools/index.ts`. Membership is the
 * contract: adding/removing a factory must update this list and the page.
 */
const BUILTIN_EXTRA_TOOL_FACTORIES: readonly string[] = [
  "createTmuxTool",
  "createPythonRunTool",
  "createLibreOfficeConvertTool",
  "createPythonTransformRangeTool",
  "createFilesTool",
  "createExecuteOfficeJsTool",
  "createExecuteWpsJsTool",
  "createExtensionsManagerTool",
];

/** Keyboard-shortcut matchers and the label the page must document for each. */
const SHORTCUT_CONTRACT: readonly (readonly [
  matcher: string,
  pageLabel: string,
])[] = [
  ["isCreateTabShortcut", "⌘/Ctrl+T"],
  ["isCloseActiveTabShortcut", "⌘/Ctrl+W"],
  ["isUndoCloseTabShortcut", "⌘/Ctrl+Z"],
  ["isReopenLastClosedShortcut", "⌘/Ctrl+Shift+T"],
  ["isFocusInputShortcut", "F2"],
  ["getAdjacentTabDirectionFromShortcut", "Ctrl+PgUp"],
  ["isRestoreQueuedMessagesShortcut", "Alt+↑"],
  ["shouldBlurEditorFromEscape", "Esc"],
  ["shouldAbortFromEscape", "Esc"],
];

/** Retired terminology and stale facts that must never come back. */
const FORBIDDEN_PATTERNS: readonly (readonly [RegExp, string])[] = [
  [/pi-web-ui/i, "pi-web-ui was removed in the first-party UI migration"],
  [/nicobailon/i, "the visual-explainer credit was removed"],
  [/whimsical/i, "the whimsical.ts credit was removed"],
  [/Armin Ronacher/u, "the whimsical.ts credit was removed"],
  [
    /v0\.(?:9|10)\.\d+-pre/u,
    "the page must not assert a version (dropped by decision)",
  ],
  [
    /18[-\s]capabilit/iu,
    "the extension capability count is derived (currently 20)",
  ],
  [/11[-\s]overlay/iu, "the overlay count is derived (currently 12)"],
  [/\b100 test files\b/iu, "the test-file count is derived (currently 147)"],
  [/\b3 suites\b/iu, "the test-suite count is derived (currently 9)"],
  [
    /\b5 snapshot kinds\b/iu,
    "the recovery snapshot-kind count is derived (currently 6)",
  ],
  [/\b24 (?:built-in )?tools\b/iu, "the tool counts are derived"],
  [
    /256\s*KB limit/iu,
    "the extension HTTP response cap is 1 MB (see data-fact extensions.httpBody)",
  ],
  [/1,221 lines/u, "do not print source line counts (they drift continuously)"],
  [
    /href="\//u,
    "internal links must be relative so they work under the GitHub Pages subpath",
  ],
];

async function countTestFiles(): Promise<number> {
  const entries = await readdir(new URL("../tests/", import.meta.url));
  // The fact guard is a repo check, not part of the app's test suites.
  return entries.filter(
    (name) =>
      /\.test\.(?:ts|mjs)$/u.test(name) &&
      name !== "architecture-page-facts.test.ts",
  ).length;
}

async function countCheckScripts(): Promise<number> {
  const entries = await readdir(new URL("../scripts/", import.meta.url));
  return entries.filter((name) => /^check-.*\.mjs$/u.test(name)).length;
}

async function countBuiltinCommandNames(): Promise<number> {
  const builtinsDir = new URL("../src/commands/builtins/", import.meta.url);
  const entries = await readdir(builtinsDir, { withFileTypes: true });
  let literalNames = 0;

  for (const entry of entries) {
    if (!entry.isFile() || !entry.name.endsWith(".ts")) continue;
    const source = await readFile(new URL(entry.name, builtinsDir), "utf8");
    literalNames += [...source.matchAll(/\bname: "/gu)].length;
  }

  // `/tools` is registered through TOOLS_COMMAND_NAME rather than a literal.
  const naming = await readSource("src/integrations/naming.ts");
  const constantName = capture(
    naming,
    /TOOLS_COMMAND_NAME = "([^"]+)"/u,
    "TOOLS_COMMAND_NAME",
  );

  return literalNames + (constantName.length > 0 ? 1 : 0);
}

async function countTestSuiteScripts(): Promise<number> {
  const packageJson = await readSource("package.json");
  const testScript = capture(
    packageJson,
    /\n\s*"test": "([^"]+)"/u,
    "the aggregate test script",
  );
  return [...testScript.matchAll(/test:[a-z-]+/gu)].length;
}

async function deriveExpectedFacts(): Promise<Record<string, string>> {
  const [
    toolNames,
    toolsIndex,
    hostSelection,
    integrationsCatalog,
    permissions,
    storageStore,
    runtimeHelpers,
    overlayIds,
    languageIndex,
    oauthRegistry,
    streamProxy,
    compactionDefaults,
    windowBudgets,
    recoveryLog,
    recoveryConstants,
    auditLog,
    saveBoundaryMonitor,
    appStorage,
    serviceDefaults,
  ] = await Promise.all([
    readSource("src/tools/names.ts"),
    readSource("src/tools/index.ts"),
    readSource("src/tools/host-selection.ts"),
    readSource("src/integrations/catalog.ts"),
    readSource("src/extensions/permissions.ts"),
    readSource("src/extensions/storage-store.ts"),
    readSource("src/extensions/runtime-manager-helpers.ts"),
    readSource("src/ui/overlay-ids.ts"),
    readSource("src/language/index.ts"),
    readSource("src/auth/oauth-provider-registry.ts"),
    readSource("src/auth/stream-proxy.ts"),
    readSource("src/compaction/defaults.ts"),
    readSource("src/context/window-budgets.ts"),
    readSource("src/workbook/recovery-log.ts"),
    readSource("src/workbook/recovery/constants.ts"),
    readSource("src/audit/workbook-change-audit.ts"),
    readSource("src/workbook/save-boundary-monitor.ts"),
    readSource("src/storage/init-app-storage.ts"),
    readSource("src/config/local-service-defaults.ts"),
  ]);

  const coreToolBlock = capture(
    toolNames,
    /CORE_TOOL_NAMES = \[([\s\S]*?)\] as const/u,
    "CORE_TOOL_NAMES",
  );
  requireSnippets(
    toolsIndex,
    BUILTIN_EXTRA_TOOL_FACTORIES,
    "src/tools/index.ts",
  );

  const sharedCoreBlock = capture(
    toolsIndex,
    /HOST_AGNOSTIC_CORE_TOOL_NAMES = new Set<string>\(\[([\s\S]*?)\]\)/u,
    "HOST_AGNOSTIC_CORE_TOOL_NAMES",
  );

  const wpsOverrideBlock = capture(
    hostSelection,
    /WPS_CORE_TOOL_EXECUTE_OVERRIDES[^{]*\{([\s\S]*?)\n\};/u,
    "WPS_CORE_TOOL_EXECUTE_OVERRIDES",
  );

  const integrationToolNames = [
    ...integrationsCatalog.matchAll(/toolNames: \[([^\]]*)\]/gu),
  ]
    .flatMap((match) => (match[1] ?? "").split(","))
    .filter((entry) => entry.trim().length > 0);

  const capabilityBlock = capture(
    permissions,
    /getCapabilityDescriptors\(\): readonly CapabilityDescriptor\[\] \{\s*return \[([\s\S]*?)\n {2}\];/u,
    "getCapabilityDescriptors()",
  );

  const trustBlock = capture(
    permissions,
    /export type StoredExtensionTrust =([\s\S]*?);/u,
    "StoredExtensionTrust",
  );

  const supportedLanguages = capture(
    languageIndex,
    /SUPPORTED_LANGUAGES = \[([^\]]*)\]/u,
    "SUPPORTED_LANGUAGES",
  );

  const oauthProviderBlock = capture(
    oauthRegistry,
    /OAUTH_PROVIDERS: Map<string, BrowserOAuthProvider> = new Map\(\[([\s\S]*?)\]\);/u,
    "OAUTH_PROVIDERS",
  );

  const recoveryKindBlock = capture(
    recoveryLog,
    /export type WorkbookRecoverySnapshotKind =([\s\S]*?);/u,
    "WorkbookRecoverySnapshotKind",
  );

  const storageStoresBlock = capture(
    appStorage,
    /stores: \[([\s\S]*?)\],/u,
    "app storage stores",
  );

  const storageBytes = captureInt(
    storageStore,
    /MAX_EXTENSION_STORAGE_BYTES = ([\d_]+)/u,
    "MAX_EXTENSION_STORAGE_BYTES",
  );
  const httpBytes = captureInt(
    runtimeHelpers,
    /MAX_EXTENSION_HTTP_BODY_BYTES = ([\d_]+)/u,
    "MAX_EXTENSION_HTTP_BODY_BYTES",
  );
  const toolOutputFloorBytes =
    captureInt(
      windowBudgets,
      /MIN_TOOL_OUTPUT_MAX_BYTES = ([\d_]+)/u,
      "MIN_TOOL_OUTPUT_MAX_BYTES (KB factor)",
    ) *
    captureInt(
      windowBudgets,
      /MIN_TOOL_OUTPUT_MAX_BYTES = [\d_]+ \* ([\d_]+)/u,
      "MIN_TOOL_OUTPUT_MAX_BYTES",
    );

  const expected: Record<string, string> = {
    "tools.core": String((coreToolBlock.match(/"[a-z_]+"/gu) ?? []).length),
    "tools.builtinExtra": String(BUILTIN_EXTRA_TOOL_FACTORIES.length),
    "tools.word": String(
      countMatches(
        toolsIndex,
        /createWord[A-Za-z]+Tool\(/gu,
        "Word tool factories",
      ),
    ),
    "tools.powerpoint": String(
      countMatches(
        toolsIndex,
        /createPowerPoint[A-Za-z]+Tool\(/gu,
        "PowerPoint tool factories",
      ),
    ),
    "tools.integrations": String(integrationToolNames.length),

    "hosts.coreShared": String(
      (sharedCoreBlock.match(/"[a-z_]+"/gu) ?? []).length,
    ),
    "hosts.wpsOverrides": String(
      (wpsOverrideBlock.match(/^\s+[a-z_]+:/gmu) ?? []).length,
    ),

    "extensions.capabilities": String(
      (capabilityBlock.match(/\{ capability: "/gu) ?? []).length,
    ),
    "extensions.trustTypes": String(
      (trustBlock.match(/"[a-z-]+"/gu) ?? []).length,
    ),
    "extensions.storage": formatMegabytes(storageBytes),
    "extensions.httpBody": formatMegabytes(httpBytes),

    "ui.overlays": String(
      countMatches(overlayIds, /_OVERLAY_ID = "/gu, "overlay ids"),
    ),
    "ui.commands": String(await countBuiltinCommandNames()),
    "ui.locales": String((supportedLanguages.match(/"[^"]+"/gu) ?? []).length),

    "model.oauthProviders": String(
      countMatches(oauthProviderBlock, /^\s*\[/gmu, "OAuth providers"),
    ),
    "model.payloadSnapshots": String(
      captureInt(
        streamProxy,
        /MAX_PAYLOAD_SNAPSHOTS = ([\d_]+)/u,
        "MAX_PAYLOAD_SNAPSHOTS",
      ),
    ),
    "model.sessionContexts": String(
      captureInt(
        streamProxy,
        /MAX_SESSION_CONTEXTS = ([\d_]+)/u,
        "MAX_SESSION_CONTEXTS",
      ),
    ),

    "compaction.largePercent": String(
      Math.round(
        Number.parseFloat(
          capture(
            compactionDefaults,
            /LARGE_CONTEXT_HARD_RATIO = ([\d.]+)/u,
            "LARGE_CONTEXT_HARD_RATIO",
          ),
        ) * 100,
      ),
    ),
    "compaction.xlPercent": String(
      Math.round(
        Number.parseFloat(
          capture(
            compactionDefaults,
            /XL_CONTEXT_HARD_RATIO = ([\d.]+)/u,
            "XL_CONTEXT_HARD_RATIO",
          ),
        ) * 100,
      ),
    ),

    "budget.toolOutputKb": `${toolOutputFloorBytes / 1024} KB`,
    "budget.toolOutputLines": String(
      captureInt(
        windowBudgets,
        /MIN_TOOL_OUTPUT_MAX_LINES = ([\d_]+)/u,
        "MIN_TOOL_OUTPUT_MAX_LINES",
      ),
    ),

    "recovery.kinds": String(
      (recoveryKindBlock.match(/"[a-z_]+"/gu) ?? []).length,
    ),
    "recovery.maxEntries": String(
      captureInt(
        recoveryConstants,
        /MAX_RECOVERY_ENTRIES = ([\d_]+)/u,
        "MAX_RECOVERY_ENTRIES",
      ),
    ),
    "recovery.maxCells": formatInt(
      captureInt(
        recoveryConstants,
        /MAX_RECOVERY_CELLS = ([\d_]+)/u,
        "MAX_RECOVERY_CELLS",
      ),
    ),
    "audit.maxEntries": String(
      captureInt(
        auditLog,
        /MAX_AUDIT_ENTRIES = ([\d_]+)/u,
        "MAX_AUDIT_ENTRIES",
      ),
    ),
    "save.pollSeconds": String(
      captureInt(
        saveBoundaryMonitor,
        /DEFAULT_POLL_INTERVAL_MS = ([\d_]+)/u,
        "DEFAULT_POLL_INTERVAL_MS",
      ) / 1000,
    ),
    "storage.stores": String(
      (storageStoresBlock.match(/get(?:Metadata)?Config\(\)/gu) ?? []).length,
    ),

    "tests.files": String(await countTestFiles()),
    "tests.suites": String(await countTestSuiteScripts()),
    "checks.scripts": String(await countCheckScripts()),

    "bridge.pi.prodPort": capture(
      serviceDefaults,
      /BRIDGE_PROD_PORT = ([\d_]+)/u,
      "BRIDGE_PROD_PORT",
    ),
    "bridge.pi.devPort": capture(
      serviceDefaults,
      /BRIDGE_DEV_PORT = ([\d_]+)/u,
      "BRIDGE_DEV_PORT",
    ),
    "bridge.python.prodPort": capture(
      serviceDefaults,
      /PYTHON_PROD_PORT = ([\d_]+)/u,
      "PYTHON_PROD_PORT",
    ),
    "bridge.python.devPort": capture(
      serviceDefaults,
      /PYTHON_DEV_PORT = ([\d_]+)/u,
      "PYTHON_DEV_PORT",
    ),
    "bridge.tmux.prodPort": capture(
      serviceDefaults,
      /TMUX_PROD_PORT = ([\d_]+)/u,
      "TMUX_PROD_PORT",
    ),
    "bridge.tmux.devPort": capture(
      serviceDefaults,
      /TMUX_DEV_PORT = ([\d_]+)/u,
      "TMUX_DEV_PORT",
    ),
    "bridge.proxy.prodPort": capture(
      serviceDefaults,
      /PROXY_PROD_PORT = ([\d_]+)/u,
      "PROXY_PROD_PORT",
    ),
    "bridge.proxy.devPort": capture(
      serviceDefaults,
      /PROXY_DEV_PORT = ([\d_]+)/u,
      "PROXY_DEV_PORT",
    ),
  };

  return Object.fromEntries(
    Object.entries(expected).map(([key, value]) => [
      key,
      value.replaceAll("_", ""),
    ]),
  );
}

function extractPageFacts(page: string): Map<string, string[]> {
  const facts = new Map<string, string[]>();

  for (const match of page.matchAll(FACT_PATTERN)) {
    const key = match[1];
    const value = (match[2] ?? "").trim();
    if (key === undefined) continue;
    const values = facts.get(key) ?? [];
    values.push(value);
    facts.set(key, values);
  }

  return facts;
}

void test("architecture page facts match the codebase", async () => {
  const [page, expected] = await Promise.all([
    readPage(),
    deriveExpectedFacts(),
  ]);
  const actual = extractPageFacts(page);
  const problems: string[] = [];

  for (const [key, expectedValue] of Object.entries(expected)) {
    const actualValues = actual.get(key);
    if (actualValues === undefined) {
      problems.push(
        `missing marker: data-fact="${key}" (code says "${expectedValue}")`,
      );
      continue;
    }
    for (const actualValue of actualValues) {
      if (actualValue !== expectedValue) {
        problems.push(
          `stale value: data-fact="${key}" is "${actualValue}", code says "${expectedValue}"`,
        );
      }
    }
  }

  for (const key of actual.keys()) {
    if (expected[key] === undefined) {
      problems.push(
        `unknown marker: data-fact="${key}" has no derivation in the guard`,
      );
    }
  }

  assert.equal(
    problems.length,
    0,
    `Architecture page facts drifted (${problems.length} problem(s)):\n- ${problems.join("\n- ")}`,
  );
});

void test("architecture page uses no retired terminology", async () => {
  const page = await readPage();
  const offenders: string[] = [];

  for (const [pattern, reason] of FORBIDDEN_PATTERNS) {
    if (pattern.test(page)) {
      offenders.push(`${String(pattern)} — ${reason}`);
    }
  }

  assert.equal(
    offenders.length,
    0,
    `Architecture page contains retired terminology:\n- ${offenders.join("\n- ")}`,
  );
});

void test("architecture page internal links and TOC anchors resolve", async () => {
  const page = await readPage();

  const ids = new Set(
    [...page.matchAll(/\bid="([^"]+)"/gu)].map((match) => match[1]),
  );
  const anchorTargets = [...page.matchAll(/href="#([^"]+)"/gu)].map(
    (match) => match[1],
  );
  const missingAnchors = anchorTargets.filter(
    (target) => target !== undefined && !ids.has(target),
  );

  assert.equal(
    missingAnchors.length,
    0,
    `Architecture page has anchors with no target: ${missingAnchors.join(", ")}`,
  );

  requireSnippets(
    page,
    [
      'href="../"',
      'href="../#install"',
      'href="../#connect"',
      'href="./"',
      'class="nav-mark"',
    ],
    "Architecture page nav",
  );
});

void test("architecture page documents the implemented keyboard shortcuts", async () => {
  const [page, matchers] = await Promise.all([
    readPage(),
    readSource("src/taskpane/keyboard-shortcuts/matchers.ts"),
  ]);

  const missingMatchers = SHORTCUT_CONTRACT.map(([matcher]) => matcher).filter(
    (matcher) => !matchers.includes(matcher),
  );
  const missingLabels = SHORTCUT_CONTRACT.map(([, label]) => label).filter(
    (label) => !page.includes(label),
  );

  assert.equal(
    missingMatchers.length,
    0,
    `Shortcut matchers not found: ${missingMatchers.join(", ")}`,
  );
  assert.equal(
    missingLabels.length,
    0,
    `Shortcut labels missing from the page: ${missingLabels.join(", ")}`,
  );
});
