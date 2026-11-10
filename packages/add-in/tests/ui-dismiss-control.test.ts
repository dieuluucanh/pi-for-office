import assert from "node:assert/strict";
import { test } from "node:test";

import { installFakeDom } from "./fake-dom.test.ts";
import { createDismissControl } from "../src/ui/dismiss-control.ts";
import { createOverlayCloseButton } from "../src/ui/overlay-dialog.ts";

void test("dismiss control is a labelled, type=button × with the shared class", () => {
  const dom = installFakeDom();
  try {
    const button = createDismissControl({ onDismiss: () => {} });

    assert.equal(button.type, "button");
    assert.equal(button.textContent, "×");
    assert.ok(button.className.split(" ").includes("pi-overlay-close"));
    assert.equal(button.getAttribute("aria-label"), button.title);
    assert.ok((button.getAttribute("aria-label") ?? "").length > 0);
  } finally {
    dom.restore();
  }
});

void test("dismiss control runs the callback once and swallows the click", () => {
  const dom = installFakeDom();
  try {
    let dismissals = 0;
    let parentClicks = 0;

    const host = document.createElement("div");
    document.body.appendChild(host);
    host.addEventListener("click", () => {
      parentClicks += 1;
    });

    const button = createDismissControl({
      onDismiss: () => {
        dismissals += 1;
      },
    });
    host.appendChild(button);

    button.dispatchEvent(new Event("click", { bubbles: true }));

    assert.equal(dismissals, 1);
    assert.equal(parentClicks, 0, "dismiss must not bubble into the parent row/backdrop");
  } finally {
    dom.restore();
  }
});

void test("dismiss control accepts a caller modifier class and label", () => {
  const dom = installFakeDom();
  try {
    const button = createDismissControl({
      onDismiss: () => {},
      label: "Dismiss notification",
      className: "pi-toast__close",
    });

    const classes = button.className.split(" ");
    assert.ok(classes.includes("pi-overlay-close"));
    assert.ok(classes.includes("pi-toast__close"));
    assert.equal(button.getAttribute("aria-label"), "Dismiss notification");
  } finally {
    dom.restore();
  }
});

void test("overlay close button is the same control", () => {
  const dom = installFakeDom();
  try {
    let closed = 0;
    const button = createOverlayCloseButton({
      onClose: () => {
        closed += 1;
      },
      label: "Close resume session",
    });

    assert.equal(button.className, "pi-overlay-close");
    assert.equal(button.getAttribute("aria-label"), "Close resume session");
    button.dispatchEvent(new Event("click"));
    assert.equal(closed, 1);
  } finally {
    dom.restore();
  }
});
