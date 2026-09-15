/**
 * Pane focus/visibility reporter tests (protocol v2 `pane_state` source).
 *
 * Uses the shared fake DOM plus a stubbed `window` EventTarget and
 * `document.hasFocus()`/`visibilityState`, so the debounce, cleanup and Excel
 * activation binding can be exercised without Office.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  bindExcelWorkbookActivation,
  createPaneFocusReporter,
  type PaneFocusState,
} from "../src/bridge/pane-focus.ts";
import { installFakeDom } from "./fake-dom.test.ts";

const flush = (): Promise<void> =>
  new Promise((resolve) => {
    setTimeout(resolve, 10);
  });

interface FocusEnv {
  setFocused(value: boolean): void;
  setVisible(value: boolean): void;
  windowTarget: EventTarget;
  restore(): void;
}

function installFocusEnv(): FocusEnv {
  const dom = installFakeDom();
  const previousWindow = Reflect.get(globalThis, "window");
  const windowTarget = new EventTarget();
  Reflect.set(globalThis, "window", windowTarget);

  const doc = dom.document as unknown as {
    hasFocus: () => boolean;
    visibilityState: string;
  };
  let focused = false;
  let visible = true;
  doc.hasFocus = () => focused;
  Object.defineProperty(doc, "visibilityState", {
    configurable: true,
    get: () => (visible ? "visible" : "hidden"),
  });

  return {
    setFocused: (value) => {
      focused = value;
    },
    setVisible: (value) => {
      visible = value;
    },
    windowTarget,
    restore: () => {
      Reflect.set(globalThis, "window", previousWindow);
      dom.restore();
    },
  };
}

void test("reports window focus/blur changes", async () => {
  const env = installFocusEnv();
  try {
    const events: PaneFocusState[] = [];
    const reporter = createPaneFocusReporter({
      host: "excel",
      debounceMs: 0,
      onChange: (state) => events.push(state),
    });
    const stop = reporter.start();

    env.setFocused(true);
    env.windowTarget.dispatchEvent(new Event("focus"));
    await flush();
    assert.equal(events.length, 1);
    assert.equal(events[events.length - 1]?.focused, true);
    assert.equal(events[events.length - 1]?.visible, true);

    env.setFocused(false);
    env.windowTarget.dispatchEvent(new Event("blur"));
    await flush();
    assert.equal(events.length, 2);
    assert.equal(events[events.length - 1]?.focused, false);

    stop();
  } finally {
    env.restore();
  }
});

void test("reports visibility changes and stops after cleanup", async () => {
  const env = installFocusEnv();
  try {
    const events: PaneFocusState[] = [];
    const reporter = createPaneFocusReporter({
      host: "word",
      debounceMs: 0,
      onChange: (state) => events.push(state),
    });
    const stop = reporter.start();

    env.setVisible(false);
    Reflect.get(globalThis, "document").dispatchEvent(
      new Event("visibilitychange"),
    );
    await flush();
    assert.equal(events.length, 1);
    assert.equal(events[events.length - 1]?.visible, false);

    stop();
    env.setVisible(true);
    Reflect.get(globalThis, "document").dispatchEvent(
      new Event("visibilitychange"),
    );
    await flush();
    assert.equal(events.length, 1, "no events after cleanup");
  } finally {
    env.restore();
  }
});

void test("selection changes use the latest snapshot without waiting", () => {
  const env = installFocusEnv();
  try {
    const reporter = createPaneFocusReporter({
      host: "excel",
      onChange: () => {},
    });
    env.setFocused(true);
    // snapshot() reflects the last known state (started unfocused).
    assert.equal(reporter.snapshot().focused, false);
    assert.equal(reporter.snapshot().visible, true);
    assert.ok(Number.isFinite(reporter.snapshot().activeAt));
  } finally {
    env.restore();
  }
});

void test("binds Excel.Workbook.onActivated when the Excel API is present", async () => {
  const previousExcel = Reflect.get(globalThis, "Excel");
  const handlers: Array<() => void> = [];
  Reflect.set(globalThis, "Excel", {
    run: (batch: (context: unknown) => Promise<unknown>) => {
      const context = {
        workbook: {
          onActivated: {
            add: (handler: () => void) => {
              handlers.push(handler);
            },
          },
        },
        sync: () => Promise.resolve(),
      };
      return batch(context);
    },
  });

  try {
    let active = 0;
    const stop = bindExcelWorkbookActivation(() => {
      active += 1;
    });
    await flush();
    assert.equal(handlers.length, 1, "onActivated handler registered");
    handlers[0]?.();
    assert.equal(active, 1);
    stop();
  } finally {
    Reflect.set(globalThis, "Excel", previousExcel);
    if (previousExcel === undefined) {
      Reflect.deleteProperty(globalThis, "Excel");
    }
  }
});

void test("Excel activation binding is a no-op without the Excel API", () => {
  assert.doesNotThrow(() => {
    const stop = bindExcelWorkbookActivation(() => {});
    stop();
  });
});
