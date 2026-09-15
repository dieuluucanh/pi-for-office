/**
 * Pane focus/visibility reporting for the bridge.
 *
 * The server resolves a missing or `"active"` document target to the focused
 * document, so the pane reports window focus and visibility with `pane_state`
 * frames. Signals are best-effort and layered:
 *
 *  - `window` focus/blur and `document.visibilitychange` (all hosts),
 *  - `Excel.Workbook.onActivated` for Excel, which is more reliable than iframe
 *    focus because it fires when the workbook window is activated.
 *
 * Everything degrades to a no-op outside a browser (tests/SSR) and never
 * throws out of a host API failure.
 */

import type { OfficeHostApp } from "@dieulc/pi-office-protocol";

/** Focus/visibility snapshot reported to the server. */
export interface PaneFocusState {
  focused: boolean;
  visible: boolean;
  /** Timestamp of the last change (ms since epoch). */
  activeAt: number;
}

export interface PaneFocusReporter {
  /** Latest state (without waiting for the debounce). */
  snapshot(): PaneFocusState;
  /** Start listening; returns a cleanup function. */
  start(): () => void;
}

const DEFAULT_DEBOUNCE_MS = 150;

function readDomState(): { focused: boolean; visible: boolean } {
  if (typeof document === "undefined") {
    return { focused: false, visible: true };
  }
  return {
    focused: document.hasFocus(),
    visible: document.visibilityState !== "hidden",
  };
}

/** Minimal Excel batch-context shape used by the activation binding. */
interface ExcelActivationContext {
  workbook: { onActivated: { add: (handler: () => void) => unknown } };
  sync: () => Promise<unknown>;
}

/**
 * Subscribe to `Excel.Workbook.onActivated` when the Excel JS API is present.
 * Returns a cleanup function; silently no-ops on other hosts or older builds.
 */
export function bindExcelWorkbookActivation(
  markActive: () => void,
): () => void {
  const excel = (globalThis as { Excel?: unknown }).Excel as
    | {
        run?: (
          batch: (context: ExcelActivationContext) => Promise<unknown>,
        ) => Promise<unknown>;
      }
    | undefined;
  if (excel === undefined || typeof excel.run !== "function") {
    return () => {};
  }
  try {
    void excel
      .run(async (context) => {
        context.workbook.onActivated.add(() => markActive());
        await context.sync();
      })
      .catch(() => {
        // Older Excel builds / mocked runtimes: focus events still work.
      });
  } catch {
    // Defensive: a broken host API must never break the bridge connection.
  }
  return () => {};
}

/**
 * Create a debounced focus/visibility reporter for one pane.
 *
 * `onChange` fires at most once per debounce window with the latest state; the
 * reporter also exposes an immediate `snapshot()` for the handshake.
 */
export function createPaneFocusReporter(options: {
  host: OfficeHostApp;
  onChange: (state: PaneFocusState) => void;
  debounceMs?: number;
}): PaneFocusReporter {
  const debounceMs = options.debounceMs ?? DEFAULT_DEBOUNCE_MS;
  let state: PaneFocusState = { ...readDomState(), activeAt: Date.now() };
  let timer: ReturnType<typeof setTimeout> | null = null;
  let cleanups: Array<() => void> = [];

  const update = (patch: Partial<PaneFocusState>): void => {
    state = { ...state, ...patch, activeAt: Date.now() };
    if (timer !== null) clearTimeout(timer);
    timer = setTimeout(() => {
      timer = null;
      options.onChange(state);
    }, debounceMs);
  };

  return {
    snapshot: () => state,
    start: () => {
      if (typeof window === "undefined" || typeof document === "undefined") {
        return () => {};
      }

      const onFocus = () => update({ focused: true });
      const onBlur = () => update({ focused: false });
      const onVisibility = () =>
        update({
          visible: document.visibilityState !== "hidden",
          focused: document.hasFocus(),
        });

      window.addEventListener("focus", onFocus);
      window.addEventListener("blur", onBlur);
      document.addEventListener("visibilitychange", onVisibility);
      cleanups.push(() => window.removeEventListener("focus", onFocus));
      cleanups.push(() => window.removeEventListener("blur", onBlur));
      cleanups.push(() =>
        document.removeEventListener("visibilitychange", onVisibility),
      );

      if (options.host === "excel") {
        cleanups.push(
          bindExcelWorkbookActivation(() => update({ focused: true })),
        );
      }

      return () => {
        if (timer !== null) {
          clearTimeout(timer);
          timer = null;
        }
        for (const cleanup of cleanups) cleanup();
        cleanups = [];
      };
    },
  };
}
