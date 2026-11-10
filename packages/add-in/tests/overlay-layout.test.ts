/**
 * CSS invariants for the responsive overlay/toast system (docs/ui-layout.md).
 *
 * These pin the *layout model*, not pixel values: the toast must never go back
 * to the centred shrink-to-fit box that produced one glyph per line, and cards
 * must stay `min(100%, cap)` with `max-height: 100%` inside a padded overlay.
 */

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

function read(relativePath: string): string {
  const css = readFileSync(`src/ui/theme/${relativePath}`, "utf8");
  // Comments explain the old (broken) rules by name; strip them so the
  // assertions below only see live declarations.
  return css.replace(/\/\*[\s\S]*?\*\//gu, "");
}

/** Text of the first rule block for an exact selector. */
function block(css: string, selector: string): string {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
  const match = new RegExp(`${escaped}\\s*\\{([^}]*)\\}`, "u").exec(css);
  assert.ok(match, `missing rule for ${selector}`);
  return match[1] ?? "";
}

void test("the toast host spans the pane instead of anchoring at 50%", () => {
  const css = read("components/toasts.css");
  const host = block(css, ".pi-toast-host");

  assert.match(host, /left:\s*8px;/u);
  assert.match(host, /right:\s*8px;/u);
  assert.match(host, /pointer-events:\s*none;/u);
  assert.match(host, /z-index:\s*var\(--pi-z-toast\);/u);
});

void test("the toast itself is a flex item, not a centred abspos box", () => {
  const css = read("components/toasts.css");
  const toast = block(css, ".pi-toast");

  assert.doesNotMatch(toast, /left:\s*50%/u, "left: 50% caps the width at half the pane");
  assert.doesNotMatch(toast, /translateX/u);
  assert.doesNotMatch(toast, /position:\s*fixed;/u);
  assert.match(toast, /max-width:\s*min\(var\(--pi-toast-max\), 100%\);/u);
  assert.match(toast, /pointer-events:\s*auto;/u);
});

void test("the toast message keeps a readable measure and the action wraps", () => {
  const css = read("components/toasts.css");
  const message = block(css, ".pi-toast__message");
  const action = block(css, ".pi-toast__action");

  assert.match(message, /flex:\s*1 1 12ch;/u);
  assert.match(message, /min-width:\s*0;/u);
  assert.doesNotMatch(action, /white-space:\s*nowrap;/u);
});

void test("every toast variant exposes the shared close control", () => {
  const css = read("components/toasts.css");

  assert.ok(block(css, ".pi-toast__close").length > 0);
});

void test("overlay cards are fluid with a cap and fill the pane height", () => {
  const css = read("overlays/primitives.css");

  const base = block(css, ".pi-overlay-card");
  assert.match(base, /width:\s*min\(100%, var\(--pi-dialog-w-m\)\);/u);
  assert.match(base, /max-height:\s*100%;/u);
  assert.doesNotMatch(base, /\bvh\b/u, "no viewport-height ceiling on cards");

  assert.match(
    block(css, ".pi-overlay-card--s"),
    /width:\s*min\(100%, var\(--pi-dialog-w-s\)\);/u,
  );
  assert.match(
    block(css, ".pi-overlay-card--l"),
    /width:\s*min\(100%, var\(--pi-dialog-w-l\)\);/u,
  );
});

void test("the overlay provides the gutter the cards rely on", () => {
  const css = read("components/welcome.css");
  const overlay = block(css, ".pi-welcome-overlay");
  const card = block(css, ".pi-welcome-card");

  assert.match(overlay, /padding:\s*var\(--pi-dialog-gutter\);/u);
  assert.match(overlay, /z-index:\s*var\(--pi-z-overlay\);/u);
  assert.match(card, /max-height:\s*100%;/u);
  assert.match(card, /width:\s*min\(100%, var\(--pi-dialog-w-prompt\)\);/u);
});

void test("no dead short-viewport override remains for the welcome card", () => {
  const css = read("components/welcome.css");

  // components.css loads before overlays.css, so such an override silently
  // loses the cascade tie — the exact bug this test exists for.
  assert.doesNotMatch(css, /@media \(max-height:/u);
});

void test("the settings shell fills the pane and owns the single scroller", () => {
  const css = read("overlays/settings-shell.css");

  const shell = block(css, ".pi-set-shell");
  assert.match(shell, /max-height:\s*100%;/u);
  assert.doesNotMatch(shell, /82vh/u);

  const body = block(css, ".pi-set-shell__body");
  assert.match(body, /overflow-y:\s*auto;/u);
  assert.match(body, /min-height:\s*0;/u);
});

void test("the resume dialog scrolls once, in the body", () => {
  const css = read("overlays/provider-resume-shortcuts.css");

  const dialog = block(css, ".pi-resume-dialog");
  assert.match(dialog, /max-height:\s*100%;/u);
  assert.doesNotMatch(dialog, /80vh/u);

  assert.match(block(css, ".pi-resume-body"), /overflow-y:\s*auto;/u);
  assert.match(block(css, ".pi-resume-groups"), /overflow:\s*visible;/u);
  assert.match(block(css, ".pi-resume-list"), /overflow:\s*visible;/u);
});

void test("the shortcuts list no longer caps its own height inside the shell", () => {
  const css = read("overlays/provider-resume-shortcuts.css");

  assert.doesNotMatch(block(css, ".pi-shortcuts-list"), /max-height:/u);
});

void test("standalone dialog cards are fluid too", () => {
  const css = read("dialogs.css");

  assert.match(block(css, ".pi-model-selector-card"), /width:\s*min\(100%, var\(--pi-dialog-w-prompt\)\);/u);
  assert.match(block(css, ".pi-api-key-card"), /width:\s*min\(100%, var\(--pi-dialog-w-prompt\)\);/u);
  assert.match(block(css, ".pi-api-key-card"), /max-height:\s*100%;/u);
});
