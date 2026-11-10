import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { installFakeDom } from "./fake-dom.test.ts";
import {
  clearErrorBanner,
  mountErrorBanner,
} from "../src/ui/error-banner.ts";

function bannerRoot(): HTMLElement {
  const root = document.createElement("div");
  root.id = "error-root";
  document.body.appendChild(root);
  return root;
}

void test("error banner shows the message with a close control", () => {
  const dom = installFakeDom();
  try {
    const root = bannerRoot();
    mountErrorBanner(root, { message: "LLM error: rate limited" });

    const banner = root.querySelector(".pi-error");
    assert.ok(banner, "banner mounted");
    assert.equal(banner.getAttribute("role"), "alert");
    assert.equal(root.querySelector(".pi-error__message")?.textContent, "LLM error: rate limited");

    const close = root.querySelector(".pi-error__close");
    assert.ok(close, "close control present");
    assert.equal(close.textContent, "×");
    assert.ok((close.getAttribute("aria-label") ?? "").length > 0);
  } finally {
    dom.restore();
  }
});

void test("closing the banner removes it and reports the dismissal", () => {
  const dom = installFakeDom();
  try {
    const root = bannerRoot();
    let dismissed = 0;
    mountErrorBanner(root, {
      message: "Proxy unreachable",
      onDismiss: () => {
        dismissed += 1;
      },
    });

    root.querySelector(".pi-error__close")?.dispatchEvent(new Event("click"));

    assert.equal(dismissed, 1);
    assert.equal(root.querySelector(".pi-error"), null);
  } finally {
    dom.restore();
  }
});

void test("a new error re-shows the banner after a dismissal", () => {
  const dom = installFakeDom();
  try {
    const root = bannerRoot();
    mountErrorBanner(root, { message: "First failure" });
    root.querySelector(".pi-error__close")?.dispatchEvent(new Event("click"));

    mountErrorBanner(root, { message: "Second failure" });

    assert.equal(root.querySelectorAll(".pi-error").length, 1);
    assert.equal(root.querySelector(".pi-error__message")?.textContent, "Second failure");
  } finally {
    dom.restore();
  }
});

void test("action buttons still call their handlers", () => {
  const dom = installFakeDom();
  try {
    const root = bannerRoot();
    let retried = 0;
    let compacted = 0;

    mountErrorBanner(root, {
      message: "Context window exceeded",
      actions: [
        {
          label: "Retry",
          onClick: () => {
            retried += 1;
          },
        },
        {
          label: "Compact + retry",
          variant: "cancel",
          onClick: () => {
            compacted += 1;
          },
        },
      ],
    });

    const buttons = root.querySelectorAll(".pi-error__action");
    assert.equal(buttons.length, 2);
    buttons[0]?.dispatchEvent(new Event("click"));
    buttons[1]?.dispatchEvent(new Event("click"));

    assert.equal(retried, 1);
    assert.equal(compacted, 1);
    assert.ok(buttons[1]?.className.includes("pi-error__action--cancel"));
  } finally {
    dom.restore();
  }
});

void test("clearErrorBanner empties the root", () => {
  const dom = installFakeDom();
  try {
    const root = bannerRoot();
    mountErrorBanner(root, { message: "No API key for openai" });
    clearErrorBanner(root);

    assert.equal(root.querySelector(".pi-error"), null);
    assert.equal(root.querySelectorAll(".pi-error").length, 0);
  } finally {
    dom.restore();
  }
});

void test("error banner CSS keeps the close reachable and the message wrapping", () => {
  const css = readFileSync("src/ui/theme/components/empty-state.css", "utf8");

  assert.match(css, /\.pi-error__row\s*\{[^}]*display:\s*flex;[^}]*align-items:\s*flex-start;/su);
  assert.match(css, /\.pi-error__message\s*\{[^}]*min-width:\s*0;[^}]*overflow-wrap:\s*anywhere;/su);
  assert.match(css, /\.pi-error__close\s*\{[^}]*flex-shrink:\s*0;/su);
});
