/**
 * Pi for Office — Loading and error state components.
 *
 * Extracted for easy swapping / versioning.
 */

import { html, type TemplateResult } from "lit";
import { t } from "../language/index.js";

/**
 * Render the loading spinner.
 */
export function renderLoading(): TemplateResult {
 return html`
    <div class="pi-loading">
      <div class="pi-loading__spinner">
        <div class="pi-loading__ring"></div>
        <div class="pi-loading__ring pi-loading__ring--inner"></div>
      </div>
      <span class="pi-loading__text">${t("loading.initializing")}</span>
    </div>
  `;
}

export interface ErrorBannerAction {
 label: string;
 onClick: () => void | Promise<void>;
 variant?: "ok" | "cancel";
}

/**
 * Show an error message, optionally with action buttons (retry / compact /
 * new session). Rendered into #error via showErrorBanner.
 */
export function renderError(
 message: string,
 actions?: readonly ErrorBannerAction[],
): TemplateResult {
 const actionButtons =
  actions && actions.length > 0
   ? html`<div class="pi-error__actions">
          ${actions.map(
           (action) => html`<button
              type="button"
              class="pi-error__action pi-error__action--${action.variant ?? "ok"}"
              @click=${() => void action.onClick()}
            >${action.label}</button>`,
          )}
        </div>`
   : null;

 return html`<div class="pi-error">${message}${actionButtons}</div>`;
}
