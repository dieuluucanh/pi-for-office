/**
 * Runtime error banner (rendered into `#error-root`).
 *
 * Dismissable: an LLM/CORS/no-key error is informational — the user may want to
 * read the transcript or fix the key first. The *fatal* bootstrap error keeps
 * using `ui/loading.ts` `renderError` and deliberately has no close control
 * (a blank pane with no explanation is worse).
 *
 * Dismissal is local state: the next `mountErrorBanner` shows the banner again,
 * and `clearErrorBanner` (successful turn / new message) empties the root.
 */

import { t } from "../language/index.js";
import { createDismissControl } from "./dismiss-control.js";
import type { ErrorBannerAction } from "./loading.js";

export type { ErrorBannerAction };

export interface ErrorBannerOptions {
  message: string;
  actions?: readonly ErrorBannerAction[];
  onDismiss?: () => void;
}

export function mountErrorBanner(root: HTMLElement, opts: ErrorBannerOptions): void {
  const banner = document.createElement("div");
  banner.className = "pi-error";
  banner.setAttribute("role", "alert");

  const row = document.createElement("div");
  row.className = "pi-error__row";

  const message = document.createElement("p");
  message.className = "pi-error__message";
  message.textContent = opts.message;

  const close = createDismissControl({
    onDismiss: () => {
      banner.remove();
      opts.onDismiss?.();
    },
    label: t("error-banner.dismiss"),
    className: "pi-error__close",
  });

  row.append(message, close);
  banner.appendChild(row);

  if (opts.actions && opts.actions.length > 0) {
    const actions = document.createElement("div");
    actions.className = "pi-error__actions";

    for (const action of opts.actions) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `pi-error__action pi-error__action--${action.variant ?? "ok"}`;
      button.textContent = action.label;
      button.addEventListener("click", () => {
        void action.onClick();
      });
      actions.appendChild(button);
    }

    banner.appendChild(actions);
  }

  root.replaceChildren(banner);
}

export function clearErrorBanner(root: HTMLElement): void {
  root.replaceChildren();
}
