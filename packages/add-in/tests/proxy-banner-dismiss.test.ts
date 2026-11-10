import assert from "node:assert/strict";
import { test } from "node:test";

import { installFakeDom } from "./fake-dom.test.ts";
import { createProxyBanner } from "../src/ui/proxy-banner.ts";

function mountBanner(): {
  root: HTMLElement;
  update: (state: "detected" | "not-detected" | "unknown" | "disabled") => void;
} {
  const banner = createProxyBanner();
  document.body.appendChild(banner.root);
  return banner;
}

void test("proxy banner has a close control that hides it", () => {
  const dom = installFakeDom();
  try {
    const banner = mountBanner();
    banner.update("not-detected");
    assert.equal(banner.root.hidden, false);

    const close = banner.root.querySelector(".pi-proxy-banner__close");
    assert.ok(close, "close control present");
    assert.equal(close.textContent, "×");

    close.dispatchEvent(new Event("click"));
    assert.equal(banner.root.hidden, true);
  } finally {
    dom.restore();
  }
});

void test("a dismissed banner stays hidden while the state is unchanged", () => {
  const dom = installFakeDom();
  try {
    const banner = mountBanner();
    banner.update("not-detected");
    banner.root.querySelector(".pi-proxy-banner__close")?.dispatchEvent(new Event("click"));

    banner.update("not-detected");
    assert.equal(banner.root.hidden, true);
  } finally {
    dom.restore();
  }
});

void test("a later outage re-notifies after the proxy was detected", () => {
  const dom = installFakeDom();
  try {
    const banner = mountBanner();
    banner.update("not-detected");
    banner.root.querySelector(".pi-proxy-banner__close")?.dispatchEvent(new Event("click"));
    assert.equal(banner.root.hidden, true);

    banner.update("detected");
    assert.equal(banner.root.hidden, true);

    banner.update("not-detected");
    assert.equal(banner.root.hidden, false, "new outage shows the hint again");
  } finally {
    dom.restore();
  }
});

void test("proxy disabled or unknown never shows the banner", () => {
  const dom = installFakeDom();
  try {
    const banner = mountBanner();

    banner.update("disabled");
    assert.equal(banner.root.hidden, true);

    banner.update("unknown");
    assert.equal(banner.root.hidden, true);
  } finally {
    dom.restore();
  }
});

void test("the how-to-fix disclosure still works", () => {
  const dom = installFakeDom();
  try {
    const banner = mountBanner();
    banner.update("not-detected");

    const action = banner.root.querySelector<HTMLElement>(".pi-proxy-banner__action");
    const details = banner.root.querySelector<HTMLElement>(".pi-proxy-banner__details");
    assert.ok(action && details);
    assert.equal(details.hidden, true);

    action.dispatchEvent(new Event("click"));
    assert.equal(details.hidden, false);
    assert.ok(banner.root.classList.contains("is-open"));
  } finally {
    dom.restore();
  }
});
