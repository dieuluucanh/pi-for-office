import assert from "node:assert/strict";
import { test } from "node:test";

import { installFakeDom } from "./fake-dom.test.ts";
import { createDisclosureBar } from "../src/ui/disclosure-bar.ts";

void test("disclosure bar exposes a close control next to the text", () => {
  const dom = installFakeDom();
  try {
    const bar = createDisclosureBar({ providerCount: 1 });
    assert.ok(bar, "bar renders for a configured provider");
    document.body.appendChild(bar);

    const head = bar.querySelector(".pi-disclosure-bar__head");
    assert.ok(head, "head row groups text + close");

    const close = bar.querySelector<HTMLElement>(".pi-disclosure-bar__close");
    assert.ok(close, "close control present");
    assert.equal(close.textContent, "×");
    assert.ok((close.getAttribute("aria-label") ?? "").length > 0);
  } finally {
    dom.restore();
  }
});

void test("the close control dismisses the bar like Got it", () => {
  const dom = installFakeDom();
  try {
    const bar = createDisclosureBar({ providerCount: 1 });
    assert.ok(bar);
    document.body.appendChild(bar);

    bar.querySelector<HTMLElement>(".pi-disclosure-bar__close")
      ?.dispatchEvent(new Event("click"));

    assert.equal(document.querySelector(".pi-disclosure-bar"), null);
  } finally {
    dom.restore();
  }
});

void test("the bar is not rendered without a configured provider", () => {
  const dom = installFakeDom();
  try {
    assert.equal(createDisclosureBar({ providerCount: 0 }), null);
  } finally {
    dom.restore();
  }
});
