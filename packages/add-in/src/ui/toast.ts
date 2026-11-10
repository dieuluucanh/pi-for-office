/**
 * Shared toast helper used across taskpane and commands.
 *
 * Layout model (docs/ui-layout.md): a fixed `.pi-toast-host` spanning the pane
 * with a small gutter holds normal flex-item toasts. There is deliberately no
 * absolutely-positioned, centred, shrink-to-fit toast: with `left: 50%` the
 * available width collapses to half the pane, and a `nowrap` action button then
 * squeezed the message to one character per line in a ~350px taskpane.
 *
 * Stacking/eviction rules live in `ui/toast-model.ts`. Every toast has its own
 * close control, its own timer, and pauses on hover/keyboard focus.
 */

import { t } from "../language/index.js";
import { createDismissControl } from "./dismiss-control.js";
import {
  pickUndoTarget,
  pushToast,
  removeToast,
  type PushToastResult,
  type ToastRecord,
  type ToastVariant,
} from "./toast-model.js";

export type { ToastVariant };

export interface ToastOptions {
  duration?: number;
  variant?: ToastVariant;
}

export interface ActionToastOptions {
  message: string;
  actionLabel: string;
  onAction: () => void;
  duration?: number;
}

interface ResolvedToastOptions {
  duration: number;
  variant: ToastVariant;
}

interface ToastAction {
  label: string;
  onAction: () => void;
}

interface ToastNode {
  record: ToastRecord;
  root: HTMLDivElement;
  close: HTMLButtonElement;
  action: HTMLButtonElement | null;
  actionLabel: string | null;
  timer: ReturnType<typeof setTimeout> | null;
  leaveTimer: ReturnType<typeof setTimeout> | null;
  remainingMs: number;
  startedAt: number;
  paused: boolean;
  leaving: boolean;
}

const ERROR_TOAST_PATTERN = /\b(fail(?:ed|ure)?|error|invalid|denied|blocked|could\s*not|couldn't|can\s*not|can't|unable|timed\s*out)\b/iu;
const DEFAULT_INFO_DURATION_MS = 2000;
const DEFAULT_ERROR_DURATION_MS = 6000;
const ACTION_TOAST_DEFAULT_MS = 7000;
/** Fade-out budget; keep in sync with `--duration-normal` in tokens.css. */
const LEAVE_DURATION_MS = 160;
const HOST_ID = "pi-toast-host";

let hostElement: HTMLDivElement | null = null;
let records: ToastRecord[] = [];
let sequence = 0;
const nodes = new Map<string, ToastNode>();

function nextSequence(): number {
  sequence += 1;
  return sequence;
}

function ensureHost(): HTMLDivElement {
  // A `/reload` (or a fresh test document) leaves the previous host behind; a
  // host from another document must never be reused.
  if (hostElement?.isConnected && hostElement.ownerDocument === document) {
    return hostElement;
  }

  const existing = document.getElementById(HOST_ID);
  if (existing instanceof HTMLElement) {
    hostElement = existing as HTMLDivElement;
    return hostElement;
  }

  const host = document.createElement("div");
  host.id = HOST_ID;
  host.className = "pi-toast-host";
  document.body.appendChild(host);
  hostElement = host;
  return host;
}

function inferToastVariant(message: string): ToastVariant {
  return ERROR_TOAST_PATTERN.test(message) ? "error" : "info";
}

function normalizeToastOptions(
  message: string,
  durationOrOptions: number | ToastOptions | undefined,
): ResolvedToastOptions {
  if (typeof durationOrOptions === "number") {
    return {
      duration: durationOrOptions,
      variant: inferToastVariant(message),
    };
  }

  const variant = durationOrOptions?.variant ?? inferToastVariant(message);
  const duration = durationOrOptions?.duration
    ?? (variant === "error" ? DEFAULT_ERROR_DURATION_MS : DEFAULT_INFO_DURATION_MS);

  return { duration, variant };
}

function clearToastTimer(node: ToastNode): void {
  if (node.timer !== null) {
    clearTimeout(node.timer);
    node.timer = null;
  }
}

function scheduleToastTimer(node: ToastNode): void {
  node.timer = setTimeout(() => {
    node.timer = null;
    hideToast(node.record.id);
  }, node.remainingMs);
}

function startToastTimer(node: ToastNode): void {
  clearToastTimer(node);
  node.remainingMs = node.record.duration;
  node.startedAt = Date.now();
  node.paused = false;

  if (node.record.duration <= 0) {
    hideToast(node.record.id);
    return;
  }

  scheduleToastTimer(node);
}

function pauseToastTimer(node: ToastNode): void {
  if (node.paused || node.leaving || node.timer === null) {
    return;
  }

  clearToastTimer(node);
  node.remainingMs = Math.max(0, node.remainingMs - (Date.now() - node.startedAt));
  node.paused = true;
}

function resumeToastTimer(node: ToastNode): void {
  if (!node.paused || node.leaving) {
    return;
  }

  node.paused = false;
  node.startedAt = Date.now();
  scheduleToastTimer(node);
}

function finalizeToast(node: ToastNode): void {
  clearToastTimer(node);

  if (node.leaveTimer !== null) {
    clearTimeout(node.leaveTimer);
    node.leaveTimer = null;
  }

  nodes.delete(node.record.id);
  node.root.remove();
}

/** Remove one toast now (user pressed ×) or fade it out (timer elapsed). */
function hideToast(id: string, immediate = false): void {
  const node = nodes.get(id);
  if (!node || node.leaving) {
    return;
  }

  node.leaving = true;
  clearToastTimer(node);
  records = removeToast(records, id);
  node.root.classList.remove("visible");

  if (immediate) {
    finalizeToast(node);
    return;
  }

  node.leaveTimer = setTimeout(() => {
    node.leaveTimer = null;
    finalizeToast(node);
  }, LEAVE_DURATION_MS);
}

function dropToastAction(node: ToastNode): void {
  if (!node.action) {
    return;
  }

  node.action.remove();
  node.action = null;
  node.actionLabel = null;
  node.root.classList.remove("pi-toast--action");
}

function isFocusWithinToast(node: ToastNode, event: Event): boolean {
  const related = (event as FocusEvent).relatedTarget;
  if (related === null || related === undefined) {
    return false;
  }

  if (typeof Node === "undefined" || !(related instanceof Node)) {
    return false;
  }

  return node.root.contains(related);
}

function buildToastElement(
  record: ToastRecord,
  action: ToastAction | undefined,
): ToastNode {
  const root = document.createElement("div");
  root.className = "pi-toast";
  root.setAttribute("aria-atomic", "true");

  const content = document.createElement("div");
  content.className = "pi-toast__content";

  const message = document.createElement("span");
  message.className = "pi-toast__message";
  message.textContent = record.message;

  content.appendChild(message);

  let actionButton: HTMLButtonElement | null = null;
  if (action) {
    actionButton = document.createElement("button");
    actionButton.type = "button";
    actionButton.className = "pi-toast__action";
    actionButton.textContent = action.label;
    actionButton.addEventListener("click", () => {
      action.onAction();
      hideToast(record.id, true);
    });
    content.appendChild(actionButton);
    root.classList.add("pi-toast--action");
  }

  const node: ToastNode = {
    record,
    root,
    close: createDismissControl({
      onDismiss: () => {
        hideToast(record.id);
      },
      label: t("toast.close"),
      className: "pi-toast__close",
    }),
    action: actionButton,
    actionLabel: action?.label ?? null,
    timer: null,
    leaveTimer: null,
    remainingMs: record.duration,
    startedAt: 0,
    paused: false,
    leaving: false,
  };

  root.append(content, node.close);

  if (record.variant === "error") {
    root.classList.add("pi-toast--error");
    root.setAttribute("role", "alert");
    root.setAttribute("aria-live", "assertive");
  } else {
    root.setAttribute("role", "status");
    root.setAttribute("aria-live", "polite");
  }

  root.addEventListener("mouseenter", () => {
    pauseToastTimer(node);
  });
  root.addEventListener("mouseleave", () => {
    resumeToastTimer(node);
  });
  root.addEventListener("focusin", () => {
    pauseToastTimer(node);
  });
  root.addEventListener("focusout", (event) => {
    if (isFocusWithinToast(node, event)) {
      return;
    }

    resumeToastTimer(node);
  });

  return node;
}

function mountToast(record: ToastRecord, action?: ToastAction): ToastNode {
  const node = buildToastElement(record, action);
  ensureHost().appendChild(node.root);
  nodes.set(record.id, node);
  node.root.classList.add("visible");
  startToastTimer(node);
  return node;
}

function applyPushResult(result: PushToastResult, action?: ToastAction): void {
  if (result.evictedId !== null) {
    hideToast(result.evictedId, true);
  }

  for (const id of result.demotedIds) {
    const demoted = nodes.get(id);
    if (demoted) {
      dropToastAction(demoted);
    }
  }

  if (result.reused) {
    const existing = nodes.get(result.record.id);
    if (existing) {
      existing.record = result.record;
      startToastTimer(existing);
      return;
    }
  }

  mountToast(result.record, action);
}

export function showToast(message: string, duration?: number): string;
export function showToast(message: string, options?: ToastOptions): string;
export function showToast(
  message: string,
  durationOrOptions?: number | ToastOptions,
): string {
  const { duration, variant } = normalizeToastOptions(message, durationOrOptions);
  const result = pushToast(
    records,
    { message, variant, duration, sticky: false },
    nextSequence(),
  );

  records = result.list;
  applyPushResult(result);
  return result.record.id;
}

export function showActionToast(opts: ActionToastOptions): string {
  const result = pushToast(
    records,
    {
      message: opts.message,
      variant: "info",
      duration: opts.duration ?? ACTION_TOAST_DEFAULT_MS,
      sticky: true,
    },
    nextSequence(),
  );

  records = result.list;
  applyPushResult(result, { label: opts.actionLabel, onAction: opts.onAction });
  return result.record.id;
}

export function dismissToast(id: string): void {
  hideToast(id, true);
}

export function clearToasts(): void {
  for (const node of [...nodes.values()]) {
    finalizeToast(node);
  }

  records = [];
}

/** True while a toast with an action button (undo/retry) is on screen. */
export function isActionToastVisible(): boolean {
  const target = pickUndoTarget(records);
  if (!target) {
    return false;
  }

  const node = nodes.get(target.id);
  return Boolean(node && !node.leaving && node.root.isConnected);
}
