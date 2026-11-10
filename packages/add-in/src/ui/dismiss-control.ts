/**
 * Shared dismiss (×) control for every transient surface.
 *
 * One implementation for toasts, the runtime error banner, the proxy banner,
 * status popovers and overlay headers, so the affordance, the accessible name
 * and the visual language never drift. See docs/ui-layout.md (dismiss matrix).
 *
 * Deliberately DOM-only and style-free: sizing lives in CSS via the caller's
 * `className` modifier (e.g. `pi-toast__close`).
 */

import { t } from "../language/index.js";

export interface DismissControlOptions {
  /** Called on click. */
  onDismiss: () => void;
  /** Accessible name; defaults to `dialog.close` ("Close dialog"). */
  label?: string;
  /** Extra class names appended after `.pi-overlay-close`. */
  className?: string;
}

export const DISMISS_CONTROL_CLASS = "pi-overlay-close";

export function createDismissControl(options: DismissControlOptions): HTMLButtonElement {
  const label = options.label ?? t("dialog.close");

  const button = document.createElement("button");
  button.type = "button";
  button.className = options.className
    ? `${DISMISS_CONTROL_CLASS} ${options.className}`
    : DISMISS_CONTROL_CLASS;
  button.textContent = "×";
  button.setAttribute("aria-label", label);
  button.title = label;

  button.addEventListener("click", (event) => {
    // Never let a dismiss bubble into a parent click handler (session rows,
    // overlay backdrops, popover anchors).
    event.preventDefault();
    event.stopPropagation();
    options.onDismiss();
  });

  return button;
}
