/**
 * Guard: a transient surface cannot ship without a way out.
 *
 * Every module that renders a notice/toast/banner must either use the shared
 * `createDismissControl()` (the × used by toasts, banners and popover headers)
 * or render an explicit `__dismiss` control. The docs matrix must name every
 * guarded module, so adding a surface forces a documentation decision too.
 */

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

interface GuardedNotice {
  path: string;
  /** Why this module is on the list (kept for the failure message). */
  surface: string;
}

const GUARDED_NOTICES: readonly GuardedNotice[] = [
  { path: "src/ui/toast.ts", surface: "toast stack" },
  { path: "src/ui/error-banner.ts", surface: "runtime error banner" },
  { path: "src/ui/proxy-banner.ts", surface: "proxy helper banner" },
  { path: "src/taskpane/status-popovers.ts", surface: "status-bar popovers" },
  { path: "src/ui/disclosure-bar.ts", surface: "onboarding disclosure bar" },
  { path: "src/ui/bridge-setup-card.ts", surface: "bridge setup card" },
  { path: "src/ui/web-search-setup-card.ts", surface: "web-search setup card" },
  { path: "src/ui/overlay-dialog.ts", surface: "overlay dialogs" },
  { path: "src/ui/settings-shell.ts", surface: "settings shell" },
];

const DISMISS_CONTROL = /createDismissControl\(/u;
const EXPLICIT_DISMISS = /__(?:dismiss|close)\b/u;

void test("every notice module renders a close/dismiss control", () => {
  const failures: string[] = [];

  for (const notice of GUARDED_NOTICES) {
    const source = readFileSync(notice.path, "utf8");
    if (!DISMISS_CONTROL.test(source) && !EXPLICIT_DISMISS.test(source)) {
      failures.push(`${notice.path} (${notice.surface})`);
    }
  }

  assert.deepEqual(failures, [], `notice modules without a close control: ${failures.join(", ")}`);
});

void test("the dismiss matrix names every guarded module", () => {
  const doc = readFileSync("docs/ui-layout.md", "utf8");
  const missing = GUARDED_NOTICES
    .filter((notice) => !doc.includes(notice.path))
    .map((notice) => notice.path);

  assert.deepEqual(missing, [], `docs/ui-layout.md must document: ${missing.join(", ")}`);
});

void test("the dismiss matrix records the non-dismissable exceptions", () => {
  const doc = readFileSync("docs/ui-layout.md", "utf8");

  assert.match(doc, /fatal bootstrap error/u);
  assert.match(doc, /thinking-level flash bar/u);
});
