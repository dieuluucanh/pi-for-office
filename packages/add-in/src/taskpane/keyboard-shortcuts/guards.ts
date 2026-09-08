/**
 * Guard helpers for keyboard shortcut side effects.
 */

import { getTextEntryElement } from "../../utils/text-entry.js";

export function textEntryHasContent(target: EventTarget | null | undefined): boolean {
  const element = getTextEntryElement(target);
  if (!element) return false;

  if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) {
    return element.value.length > 0;
  }

  if (element instanceof HTMLSelectElement) {
    return element.value.length > 0;
  }

  return (element.textContent ?? "").length > 0;
}

export function shouldHandleUndoCloseTabShortcut(opts: {
  canUndoCloseTab: boolean;
  isTextEntry: boolean;
  textEntryHasContent: boolean;
  actionToastVisible: boolean;
}): boolean {
  if (!opts.canUndoCloseTab) return false;
  if (opts.actionToastVisible) return true;
  if (!opts.isTextEntry) return true;
  if (!opts.textEntryHasContent) return true;
  return false;
}
