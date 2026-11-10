import assert from "node:assert/strict";
import { test } from "node:test";

import { installFakeDom } from "./fake-dom.test.ts";
import {
  clearToasts,
  dismissToast,
  isActionToastVisible,
  showActionToast,
  showToast,
} from "../src/ui/toast.ts";
import { MAX_VISIBLE_TOASTS } from "../src/ui/toast-model.ts";

const delay = (ms: number): Promise<void> => new Promise((resolve) => {
  setTimeout(resolve, ms);
});

function toastNodes(): HTMLElement[] {
  return Array.from(document.querySelectorAll<HTMLElement>(".pi-toast"));
}

function visibleToasts(): HTMLElement[] {
  return toastNodes().filter((node) => node.classList.contains("visible"));
}

function closeButtons(scope: HTMLElement): HTMLElement[] {
  return Array.from(scope.querySelectorAll<HTMLElement>(".pi-toast__close"));
}

function withToastDom(run: () => void | Promise<void>): () => Promise<void> {
  return async () => {
    const dom = installFakeDom();
    clearToasts();
    try {
      await run();
    } finally {
      clearToasts();
      dom.restore();
    }
  };
}

void test("toasts mount in one host and each carries a close control", withToastDom(() => {
  showToast("Session moved to this workbook");

  const host = document.getElementById("pi-toast-host");
  assert.ok(host instanceof HTMLElement, "toast host is mounted");
  assert.ok(host.classList.contains("pi-toast-host"));

  const [toast] = toastNodes();
  assert.ok(toast, "one toast rendered");
  assert.ok(toast.classList.contains("visible"));
  assert.equal(toast.getAttribute("role"), "status");

  const close = closeButtons(toast);
  assert.equal(close.length, 1);
  assert.equal(close[0]?.textContent, "×");
  assert.ok((close[0]?.getAttribute("aria-label") ?? "").length > 0);

  const message = toast.querySelector(".pi-toast__message");
  assert.equal(message?.textContent, "Session moved to this workbook");
}));

void test("stacking is capped and eviction keeps the newest toasts", withToastDom(() => {
  for (let index = 1; index <= 4; index += 1) {
    showToast(`Notice ${index}`);
  }

  const messages = visibleToasts().map((toast) => (
    toast.querySelector(".pi-toast__message")?.textContent
  ));

  assert.equal(messages.length, MAX_VISIBLE_TOASTS);
  assert.deepEqual(messages, ["Notice 2", "Notice 3", "Notice 4"]);
}));

void test("the close control removes only its own toast", withToastDom(() => {
  showToast("First notice");
  showToast("Second notice");

  const first = visibleToasts()[0];
  assert.ok(first);
  closeButtons(first)[0]?.dispatchEvent(new Event("click"));

  const remaining = visibleToasts().map((toast) => (
    toast.querySelector(".pi-toast__message")?.textContent
  ));
  assert.deepEqual(remaining, ["Second notice"]);
}));

void test("a toast hides itself after its duration", withToastDom(async () => {
  showToast("Brief", 20);
  assert.equal(visibleToasts().length, 1);

  await delay(220);
  assert.equal(toastNodes().length, 0);
}));

void test("hover pauses the timer and leaving resumes it", withToastDom(async () => {
  showToast("Read me", 40);
  const toast = visibleToasts()[0];
  assert.ok(toast);

  toast.dispatchEvent(new Event("mouseenter"));
  await delay(90);
  assert.equal(visibleToasts().length, 1, "still visible while hovered");

  toast.dispatchEvent(new Event("mouseleave"));
  await delay(240);
  assert.equal(toastNodes().length, 0);
}));

void test("focus inside the toast pauses its timer", withToastDom(async () => {
  showToast("Read me", 40);
  const toast = visibleToasts()[0];
  assert.ok(toast);

  toast.dispatchEvent(new Event("focusin", { bubbles: true }));
  await delay(90);
  assert.equal(visibleToasts().length, 1, "still visible while focused");

  toast.dispatchEvent(new Event("focusout", { bubbles: true }));
  await delay(240);
  assert.equal(toastNodes().length, 0);
}));

void test("repeating a message resets the existing toast instead of stacking", withToastDom(() => {
  showToast("Session moved");
  showToast("Session moved");

  assert.equal(visibleToasts().length, 1);
}));

void test("error toasts are announced assertively", withToastDom(() => {
  showToast("Could not save the session");

  const [toast] = toastNodes();
  assert.ok(toast?.classList.contains("pi-toast--error"));
  assert.equal(toast.getAttribute("role"), "alert");
  assert.equal(toast.getAttribute("aria-live"), "assertive");
}));

void test("only the newest action toast keeps its button", withToastDom(() => {
  showActionToast({
    message: "Closed Chat 1",
    actionLabel: "Undo",
    onAction: () => {},
    duration: 5000,
  });

  assert.equal(isActionToastVisible(), true);
  assert.equal(document.querySelectorAll(".pi-toast__action").length, 1);

  showActionToast({
    message: "Closed Chat 2",
    actionLabel: "Undo",
    onAction: () => {},
    duration: 5000,
  });

  const actionButtons = document.querySelectorAll(".pi-toast__action");
  assert.equal(actionButtons.length, 1, "one live undo affordance at a time");
  assert.equal(actionButtons[0]?.textContent, "Undo");

  const newest = visibleToasts().find((toast) => (
    toast.querySelector(".pi-toast__message")?.textContent === "Closed Chat 2"
  ));
  assert.ok(newest?.classList.contains("pi-toast--action"));
  assert.equal(isActionToastVisible(), true);
}));

void test("the action button runs its callback and dismisses that toast", withToastDom(() => {
  let undone = 0;
  showActionToast({
    message: "Closed Chat 3",
    actionLabel: "Undo",
    onAction: () => {
      undone += 1;
    },
    duration: 5000,
  });

  document.querySelector(".pi-toast__action")?.dispatchEvent(new Event("click"));

  assert.equal(undone, 1);
  assert.equal(visibleToasts().length, 0);
  assert.equal(isActionToastVisible(), false);
}));

void test("dismissToast and clearToasts remove toasts immediately", withToastDom(() => {
  const firstId = showToast("One", 60_000);
  showToast("Two", 60_000);
  assert.equal(visibleToasts().length, 2);

  dismissToast(firstId);
  assert.equal(visibleToasts().length, 1);

  clearToasts();
  assert.equal(toastNodes().length, 0);
  assert.equal(isActionToastVisible(), false);
}));
