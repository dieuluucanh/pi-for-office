import assert from "node:assert/strict";
import { test } from "node:test";

import { installFakeDom } from "./fake-dom.test.ts";
import { createExtensionAPI } from "../src/commands/extension-api.ts";
import type { ExtensionCapability } from "../src/extensions/permissions.ts";
import { clearToasts, showToast, isActionToastVisible } from "../src/ui/toast.ts";

function createApi(allowed: readonly ExtensionCapability[]) {
  const gate = (capability: ExtensionCapability): boolean => allowed.includes(capability);

  return createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    isCapabilityEnabled: gate,
  });
}

function toastMessages(): string[] {
  return Array.from(document.querySelectorAll(".pi-toast__message"))
    .map((node) => node.textContent ?? "");
}

void test("extension toast uses the shared toast host (one implementation)", () => {
  const dom = installFakeDom();
  clearToasts();

  try {
    const api = createApi(["ui.toast"]);
    api.toast("Hello from extension");

    const hosts = document.querySelectorAll(".pi-toast-host");
    assert.equal(hosts.length, 1, "a single host serves app and extension toasts");
    assert.deepEqual(toastMessages(), ["Hello from extension"]);

    const close = document.querySelector(".pi-toast__close");
    assert.ok(close, "extension toasts are dismissable too");
  } finally {
    clearToasts();
    dom.restore();
  }
});

void test("an app toast after an extension toast updates the same host", () => {
  const dom = installFakeDom();
  clearToasts();

  try {
    const api = createApi(["ui.toast"]);
    api.toast("First");
    showToast("Second");

    // Regression: the old defaultToast() wrote textContent into #pi-toast,
    // detaching the content nodes toast.ts keeps references to.
    assert.deepEqual(toastMessages(), ["First", "Second"]);
    assert.equal(document.querySelectorAll(".pi-toast-host").length, 1);
  } finally {
    clearToasts();
    dom.restore();
  }
});

void test("extension toast is denied without the capability", () => {
  const dom = installFakeDom();
  clearToasts();

  try {
    const api = createApi([]);

    assert.throws(() => {
      api.toast("nope");
    }, /ui\.toast/u);
    assert.deepEqual(toastMessages(), []);
  } finally {
    clearToasts();
    dom.restore();
  }
});

void test("the shared toast stack keeps the undo affordance exclusive", () => {
  const dom = installFakeDom();
  clearToasts();

  try {
    showToast("Saved");
    assert.equal(isActionToastVisible(), false);
  } finally {
    clearToasts();
    dom.restore();
  }
});
