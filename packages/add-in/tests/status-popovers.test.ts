import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { installFakeDom } from "./fake-dom.test.ts";
import {
  createStatusPopoverHeader,
  resolveStatusPopoverLayout,
} from "../src/taskpane/status-popovers.ts";

void test("thinking popover stays inside a 290x400 short taskpane", () => {
  const layout = resolveStatusPopoverLayout({
    anchor: { top: 350, bottom: 380, right: 280 },
    viewportWidth: 290,
    viewportHeight: 400,
    popoverWidth: 274,
    popoverHeight: 488,
  });

  assert.deepEqual(layout, {
    left: 8,
    top: 8,
    maxHeight: 384,
  });
  assert.ok(layout.top + layout.maxHeight <= 400 - 8);
});

void test("thinking popover remains above its anchor when it fits", () => {
  const layout = resolveStatusPopoverLayout({
    anchor: { top: 800, bottom: 830, right: 780 },
    viewportWidth: 800,
    viewportHeight: 900,
    popoverWidth: 290,
    popoverHeight: 300,
  });

  assert.deepEqual(layout, {
    left: 490,
    top: 492,
    maxHeight: 884,
  });
});

void test("status popover CSS constrains the shell and scrolls its choices", () => {
  const css = readFileSync("src/ui/theme/components/status-bar.css", "utf8");

  assert.match(
    css,
    /\.pi-status-popover\s*\{[^}]*max-height:\s*calc\(100vh - 16px\);[^}]*overflow:\s*hidden;/su,
  );
  assert.match(
    css,
    /\.pi-status-popover__list,[^}]*min-height:\s*0;[^}]*overflow-y:\s*auto;/su,
  );
  assert.match(
    css,
    /\.pi-status-popover__header\s*\{[^}]*display:\s*flex;[^}]*flex-shrink:\s*0;/su,
  );
  assert.match(css, /\.pi-status-popover__close\s*\{[^}]*flex-shrink:\s*0;/su);
});

void test("status popover header exposes an explicit close control", () => {
  const dom = installFakeDom();
  try {
    const { header, title } = createStatusPopoverHeader("Thinking level");

    assert.equal(title.textContent, "Thinking level");
    assert.ok(header.className.includes("pi-status-popover__header"));

    const close = header.querySelector<HTMLElement>(".pi-status-popover__close");
    assert.ok(close, "close control present");
    assert.equal(close.textContent, "×");
    assert.ok((close.getAttribute("aria-label") ?? "").length > 0);

    // Closing with no mounted popover is a no-op, not a throw.
    close.dispatchEvent(new Event("click"));
  } finally {
    dom.restore();
  }
});
