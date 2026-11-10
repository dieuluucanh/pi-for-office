import {
  closeOverlayById,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { CONFIRM_DIALOG_OVERLAY_ID } from "./overlay-ids.js";
import { t } from "../language/index.js";

function getConfirmationUiUnavailableError(): string {
  return t("confirm.unavailable");
}

export type ConfirmButtonTone = "primary" | "danger";

export interface ConfirmDialogOptions {
  title: string;
  message: string;
  confirmLabel?: string;
  cancelLabel?: string;
  /**
   * When set, renders a persistent "Always allow" action between Cancel and
   * the confirm button. Choosing it resolves the decision with
   * `remember: true` so the caller can persist the grant.
   */
  alwaysLabel?: string;
  confirmButtonTone?: ConfirmButtonTone;
  overlayId?: string;
  restoreFocusOnClose?: boolean;
  cardClassName?: string;
}

export interface ConfirmationDialogDecision {
  /** True when the user confirmed (either once or persistently). */
  confirmed: boolean;
  /** True only when the user chose the persistent "Always allow" action. */
  remember: boolean;
}

function canRenderConfirmationDialog(): boolean {
  if (typeof document === "undefined") {
    return false;
  }

  return document.body instanceof HTMLElement;
}

function getConfirmButtonClassName(tone: ConfirmButtonTone | undefined): string {
  return tone === "danger" ? "pi-overlay-btn--danger" : "pi-overlay-btn--primary";
}

/**
 * Boolean-compatible confirmation prompt. Use
 * `requestConfirmationDialogDecision` when the "Always allow" action is
 * needed.
 */
export function requestConfirmationDialog(options: ConfirmDialogOptions): Promise<boolean> {
  return requestConfirmationDialogDecision(options).then(
    (decision) => decision.confirmed,
  );
}

/**
 * Confirmation prompt with an optional persistent "Always allow" action.
 *
 * Resolves `{ confirmed, remember }`:
 * - confirm button → `{ confirmed: true, remember: false }`
 * - always button  → `{ confirmed: true, remember: true }`
 * - cancel/escape  → `{ confirmed: false, remember: false }`
 */
export function requestConfirmationDialogDecision(
  options: ConfirmDialogOptions,
): Promise<ConfirmationDialogDecision> {
  if (!canRenderConfirmationDialog()) {
    return Promise.reject(new Error(getConfirmationUiUnavailableError()));
  }

  const overlayId = options.overlayId ?? CONFIRM_DIALOG_OVERLAY_ID;
  closeOverlayById(overlayId);

  return new Promise((resolve) => {
    const dialog = createOverlayDialog({
      overlayId,
      cardClassName: options.cardClassName ?? "pi-welcome-card pi-overlay-card pi-overlay-card--s",
      ...(options.restoreFocusOnClose !== undefined ? { restoreFocusOnClose: options.restoreFocusOnClose } : {}),
    });

    let settled = false;

    const settle = (decision: ConfirmationDialogDecision): void => {
      if (settled) {
        return;
      }

      settled = true;
      resolve(decision);
    };

    const cancel = (): void => {
      settle({ confirmed: false, remember: false });
      dialog.close();
    };

    const confirm = (): void => {
      settle({ confirmed: true, remember: false });
      dialog.close();
    };

    const alwaysAllow = (): void => {
      settle({ confirmed: true, remember: true });
      dialog.close();
    };

    const { header } = createOverlayHeader({
      onClose: cancel,
      closeLabel: options.cancelLabel ?? t("confirm.cancel"),
      title: options.title,
    });

    const body = document.createElement("div");
    body.className = "pi-overlay-body";

    const message = document.createElement("p");
    message.className = "pi-overlay-subtitle pi-confirm-dialog__message";
    message.textContent = options.message;

    const actions = document.createElement("div");
    actions.className = "pi-overlay-actions";

    const cancelButton = createOverlayButton({
      text: options.cancelLabel ?? t("confirm.cancel"),
    });

    const confirmButton = createOverlayButton({
      text: options.confirmLabel ?? t("confirm.confirm"),
      className: getConfirmButtonClassName(options.confirmButtonTone),
    });

    cancelButton.addEventListener("click", cancel);
    confirmButton.addEventListener("click", confirm);

    dialog.addCleanup(() => {
      cancelButton.removeEventListener("click", cancel);
      confirmButton.removeEventListener("click", confirm);
      settle({ confirmed: false, remember: false });
    });

    actions.append(cancelButton);

    if (options.alwaysLabel !== undefined) {
      const alwaysButton = createOverlayButton({
        text: options.alwaysLabel,
      });

      alwaysButton.addEventListener("click", alwaysAllow);
      dialog.addCleanup(() => {
        alwaysButton.removeEventListener("click", alwaysAllow);
      });

      actions.append(alwaysButton);
    }

    actions.append(confirmButton);
    body.appendChild(message);
    dialog.card.append(header, body, actions);
    dialog.mount();
  });
}
