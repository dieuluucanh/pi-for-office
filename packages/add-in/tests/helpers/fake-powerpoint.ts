/**
 * Harness for mocking the global `PowerPoint` (and the `Office` requirements
 * API the capability probe reads) so the live PowerPoint tools run under Node.
 *
 * Records every assignment the helpers make on the shape fill / line / font /
 * paragraph surface, which is what the PowerPoint tests assert.
 */

import {
  setAppStorage,
  type AppStorage,
} from "../../src/storage/local/app-storage.ts";

export interface FakeShapeFill {
  readonly cleared: number;
  readonly solidColors: string[];
  transparency?: number;
  clear(): void;
  setSolidColor(color: string): void;
}

export interface FakeShapeLineFormat {
  color?: string;
  weight?: number;
  style?: string;
  dashStyle?: string;
  visible?: boolean;
  transparency?: number;
}

export interface FakeShapeFont {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  size?: number;
  name?: string;
  color?: string;
  strikethrough?: boolean;
  doubleStrikethrough?: boolean;
  smallCaps?: boolean;
  allCaps?: boolean;
}

export interface FakeBulletFormat {
  visible?: boolean;
  style?: string;
  type?: string;
}

export interface FakeParagraphFormat {
  horizontalAlignment?: string;
  indentLevel?: number;
  readonly bulletFormat: FakeBulletFormat;
}

export interface FakeTextRange {
  text?: string;
  readonly font: FakeShapeFont;
  readonly paragraphFormat: FakeParagraphFormat;
  load(props?: string): void;
}

export interface FakeTextFrame {
  readonly hasText: boolean;
  readonly textRange: FakeTextRange;
}

export interface FakeShape {
  readonly id: string;
  readonly fill: FakeShapeFill;
  readonly lineFormat: FakeShapeLineFormat;
  readonly textFrame: FakeTextFrame;
  readonly loaded: string[];
  load(props?: string): void;
}

export interface FakeShapeCollection {
  readonly items: FakeShape[];
  readonly added: FakeShape[];
  /** The request context, as Office.js exposes it on proxies. */
  context: FakePowerPointContext | null;
  load(props?: string): void;
  addTextBox(text: string, options?: PowerPoint.ShapeAddOptions): FakeShape;
}

export interface FakeSlideBackgroundSolidFill {
  readonly options: Array<{ color?: string; transparency?: number }>;
}

export interface FakeSlideBackground {
  readonly resetCalls: number;
  readonly fill: {
    readonly solid: FakeSlideBackgroundSolidFill;
    readonly cleared: number;
    setSolidFill(options?: { color?: string; transparency?: number }): void;
    clear(): void;
  };
  reset(): void;
}

export interface FakeSlide {
  readonly id: string;
  readonly shapes: FakeShapeCollection;
  readonly background: FakeSlideBackground;
  readonly loaded: string[];
  load(props?: string): void;
}

export interface FakeSlideCollection {
  readonly items: FakeSlide[];
  readonly added: FakeSlide[];
  load(props?: string): void;
  getCount(): { value: number };
  add(): FakeSlide;
  getItemAt(index: number): FakeSlide;
}

export interface FakePresentation {
  readonly slides: FakeSlideCollection;
  readonly selectedSlideIds: string[];
  readonly selectionCalls: string[][];
  setSelectedSlides(ids: string[]): void;
}

export interface FakePowerPointContext {
  readonly presentation: FakePresentation;
  syncCount: number;
  sync(): Promise<void>;
}

export interface FakePowerPointGlobal {
  readonly context: FakePowerPointContext;
  run<T>(callback: (context: FakePowerPointContext) => Promise<T>): Promise<T>;
}

export interface FakeOfficeCapabilities {
  supported: Record<string, string>;
  platform?: string;
}

let shapeIdCounter = 0;
let slideIdCounter = 0;

export function makeShape(options: { hasText?: boolean } = {}): FakeShape {
  const fill: FakeShapeFill = {
    cleared: 0,
    solidColors: [],
    clear() {
      this.cleared += 1;
    },
    setSolidColor(color: string) {
      this.solidColors.push(color);
    },
  };
  const shape: FakeShape = {
    id: `shape-${(shapeIdCounter += 1)}`,
    fill,
    lineFormat: {},
    textFrame: {
      hasText: options.hasText ?? false,
      textRange: {
        font: {},
        paragraphFormat: { bulletFormat: {} },
        load() {},
      },
    },
    loaded: [],
    load(props?: string) {
      shape.loaded.push(props ?? "");
    },
  };
  return shape;
}

export function makeSlide(): FakeSlide {
  const shapes: FakeShape[] = [];
  const added: FakeShape[] = [];

  const shapeCollection: FakeShapeCollection = {
    items: shapes,
    added,
    context: null,
    load() {},
    addTextBox(_text: string): FakeShape {
      const shape = makeShape({ hasText: true });
      shapes.push(shape);
      added.push(shape);
      return shape;
    },
  };

  const background: FakeSlideBackground = {
    resetCalls: 0,
    fill: {
      solid: { options: [] },
      cleared: 0,
      setSolidFill(options?: { color?: string; transparency?: number }) {
        background.fill.solid.options.push(options ?? {});
      },
      clear() {
        background.fill.cleared += 1;
      },
    },
    reset() {
      background.resetCalls += 1;
    },
  };

  const slide: FakeSlide = {
    id: `slide-${(slideIdCounter += 1)}`,
    shapes: shapeCollection,
    background,
    loaded: [],
    load(props?: string) {
      slide.loaded.push(props ?? "");
    },
  };
  return slide;
}

export function makeContext(slideCount = 1): FakePowerPointContext {
  const items: FakeSlide[] = Array.from({ length: slideCount }, () =>
    makeSlide(),
  );
  const added: FakeSlide[] = [];

  const slides: FakeSlideCollection = {
    items,
    added,
    load() {},
    getCount(): { value: number } {
      return { value: items.length };
    },
    add(): FakeSlide {
      const slide = makeSlide();
      items.push(slide);
      added.push(slide);
      return slide;
    },
    getItemAt(index: number): FakeSlide {
      const slide = items[index];
      if (!slide) throw new Error(`No slide at index ${index}.`);
      return slide;
    },
  };

  const presentation: FakePresentation = {
    slides,
    selectedSlideIds: [],
    selectionCalls: [],
    setSelectedSlides(ids: string[]) {
      presentation.selectionCalls.push([...ids]);
      presentation.selectedSlideIds.length = 0;
      presentation.selectedSlideIds.push(...ids);
    },
  };

  const context: FakePowerPointContext = {
    presentation,
    syncCount: 0,
    sync(): Promise<void> {
      context.syncCount += 1;
      return Promise.resolve();
    },
  };
  // Proxies expose the request context in Office.js (e.g. `shapes.context`).
  for (const slide of items) {
    slide.shapes.context = context;
  }
  return context;
}

export function makePowerPointGlobal(
  context: FakePowerPointContext,
): FakePowerPointGlobal {
  return {
    context,
    run<T>(callback: (ctx: FakePowerPointContext) => Promise<T>): Promise<T> {
      return callback(context);
    },
  };
}

/** Compare dotted requirement-set versions ("1.10" > "1.9"). */
export function compareVersions(a: string, b: string): number {
  const parse = (value: string): number[] =>
    value.split(".").map((part) => Number.parseInt(part, 10));
  const left = parse(a);
  const right = parse(b);
  for (let index = 0; index < Math.max(left.length, right.length); index += 1) {
    const diff = (left[index] ?? 0) - (right[index] ?? 0);
    if (diff !== 0) return diff > 0 ? 1 : -1;
  }
  return 0;
}

function makeOfficeGlobal(capabilities: FakeOfficeCapabilities): DynamicValue {
  return {
    context: {
      platform: capabilities.platform ?? "PC",
      requirements: {
        isSetSupported(set: string, minVersion: string): boolean {
          const supportedVersion = capabilities.supported[set];
          if (supportedVersion === undefined) return false;
          return compareVersions(supportedVersion, minVersion) >= 0;
        },
      },
    },
  };
}

export interface FakePowerPointEnvOptions {
  slides?: number;
  /** Declared Office requirement sets; omit to run with no `Office` global. */
  capabilities?: FakeOfficeCapabilities;
}

/**
 * Install a fake `PowerPoint` (and optionally `Office`) global, run `fn`, then
 * restore the previous globals.
 */
export async function withFakePowerPoint<T>(
  fn: (context: FakePowerPointContext) => T | Promise<T>,
  options: FakePowerPointEnvOptions = {},
): Promise<T> {
  const context = makeContext(options.slides ?? 1);
  const powerPoint = makePowerPointGlobal(context);

  const hadPowerPoint = Reflect.has(globalThis, "PowerPoint");
  const previousPowerPoint = Reflect.get(globalThis, "PowerPoint");
  const hadOffice = Reflect.has(globalThis, "Office");
  const previousOffice = Reflect.get(globalThis, "Office");

  Reflect.set(globalThis, "PowerPoint", powerPoint);
  if (options.capabilities) {
    Reflect.set(globalThis, "Office", makeOfficeGlobal(options.capabilities));
  }

  try {
    return await fn(context);
  } finally {
    if (hadPowerPoint)
      Reflect.set(globalThis, "PowerPoint", previousPowerPoint);
    else Reflect.deleteProperty(globalThis, "PowerPoint");
    if (hadOffice) Reflect.set(globalThis, "Office", previousOffice);
    else Reflect.deleteProperty(globalThis, "Office");
  }
}

/** Install a minimal `AppStorage` double (PowerPoint tools read conventions). */
export function installFakeAppStorage(): void {
  const values = new Map<string, DynamicValue>();
  const settings = {
    get: (key: string): Promise<DynamicValue> =>
      Promise.resolve(values.get(key) ?? null),
    set: (key: string, value: DynamicValue): Promise<void> => {
      values.set(key, value);
      return Promise.resolve();
    },
  };
  const appStorage: DynamicValue = { settings };
  setAppStorage(appStorage as AppStorage);
}
