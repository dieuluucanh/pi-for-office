import assert from "node:assert/strict";

type ConstructorSnapshot = {
  Node?: DynamicValue;
  Element?: DynamicValue;
  HTMLElement?: DynamicValue;
  HTMLInputElement?: DynamicValue;
  HTMLTextAreaElement?: DynamicValue;
  HTMLSelectElement?: DynamicValue;
  CustomEvent?: DynamicValue;
  document?: DynamicValue;
};

type FakeListener = {
  callback: (event: Event) => void;
  capture: boolean;
  once: boolean;
};

type FakeListenerOptions = boolean | { capture?: boolean; once?: boolean };

/**
 * Minimal bubbling EventTarget.
 *
 * Node's own EventTarget does not bubble and leaves `cancelBubble` untouched,
 * so tests could not exercise delegated handlers (a toast close button inside
 * the toast host, overlay backdrop clicks, document-level Escape). This keeps a
 * real listener registry, runs the capture → target → bubble phases, and pins
 * `event.target` to the original dispatch target.
 */
class FakeNode extends EventTarget {
  parentElement: FakeElement | null = null;
  ownerDocument: FakeNode | null = null;
  private readonly listeners = new Map<string, FakeListener[]>();

  override addEventListener(
    type: string,
    callback: EventListenerOrEventListenerObject | null,
    options?: FakeListenerOptions,
  ): void {
    const listener = normalizeListener(callback);
    if (!listener) return;

    const capture = typeof options === "boolean" ? options : options?.capture === true;
    const once = typeof options === "object" && options !== null && options.once === true;
    const list = this.listeners.get(type) ?? [];

    if (list.some((entry) => entry.callback === listener && entry.capture === capture)) {
      return;
    }

    list.push({ callback: listener, capture, once });
    this.listeners.set(type, list);
  }

  override removeEventListener(
    type: string,
    callback: EventListenerOrEventListenerObject | null,
    options?: FakeListenerOptions,
  ): void {
    const listener = normalizeListener(callback);
    if (!listener) return;

    const capture = typeof options === "boolean" ? options : options?.capture === true;
    const list = this.listeners.get(type);
    if (!list) return;

    const index = list.findIndex(
      (entry) => entry.callback === listener && entry.capture === capture,
    );
    if (index >= 0) list.splice(index, 1);
  }

  override dispatchEvent(event: Event): boolean {
    const path: FakeNode[] = [this];

    let cursor: FakeNode | null = this.parentElement ?? this.ownerDocument;
    while (cursor) {
      path.push(cursor);
      cursor = cursor.parentElement ?? cursor.ownerDocument;
    }

    let stopped = false;
    let immediate = false;
    let current: FakeNode | null = null;

    const plainStopPropagation = event.stopPropagation.bind(event);
    const plainStopImmediate = event.stopImmediatePropagation.bind(event);

    Object.defineProperty(event, "stopPropagation", {
      configurable: true,
      value: () => {
        stopped = true;
        plainStopPropagation();
      },
    });
    Object.defineProperty(event, "stopImmediatePropagation", {
      configurable: true,
      value: () => {
        stopped = true;
        immediate = true;
        plainStopImmediate();
      },
    });
    Object.defineProperty(event, "target", {
      configurable: true,
      // Arrow keeps the lexical `this`, i.e. the node that dispatched.
      get: () => this,
    });
    Object.defineProperty(event, "currentTarget", {
      configurable: true,
      get: () => current,
    });

    const invoke = (node: FakeNode, phase: "capture" | "target" | "bubble"): void => {
      const list = this.listenersFor(node, event.type);
      if (list.length === 0) return;

      current = node;
      immediate = false;

      for (const listener of [...list]) {
        if (phase === "capture" && !listener.capture) continue;
        if (phase === "bubble" && listener.capture) continue;
        if (immediate) break;

        if (listener.once) {
          const index = list.indexOf(listener);
          if (index >= 0) list.splice(index, 1);
        }

        listener.callback.call(node, event);
      }
    };

    // Capture: root-most ancestor down to the parent of the target.
    for (let index = path.length - 1; index >= 1; index -= 1) {
      if (stopped) break;
      invoke(path[index], "capture");
    }

    // Target phase: capture and bubble listeners both run, in registration order.
    if (!stopped) invoke(this, "target");

    // Bubble: parent first, up to the root-most ancestor.
    if (!stopped && event.bubbles) {
      for (let index = 1; index < path.length; index += 1) {
        if (stopped) break;
        invoke(path[index], "bubble");
      }
    }

    current = null;
    return !event.defaultPrevented;
  }

  private listenersFor(node: FakeNode, type: string): FakeListener[] {
    return node.listeners.get(type) ?? [];
  }

  get isConnected(): boolean {
    return this.parentElement ? this.parentElement.isConnected : false;
  }

  remove(): void {
    if (!this.parentElement) {
      return;
    }

    const parent = this.parentElement;
    this.parentElement = null;
    parent.removeChild(this);
  }
}

function normalizeListener(
  callback: EventListenerOrEventListenerObject | null,
): ((event: Event) => void) | null {
  if (callback === null) return null;
  if (typeof callback === "function") return callback;

  return (event: Event) => {
    callback.handleEvent(event);
  };
}

/** Set-backed `element.classList`, mirrored into `className` like the real thing. */
class FakeClassList {
  private readonly classes: Set<string>;

  constructor(classes: Set<string>) {
    this.classes = classes;
  }

  get length(): number {
    return this.classes.size;
  }

  add(...tokens: string[]): void {
    for (const token of tokens) {
      if (token.length > 0) this.classes.add(token);
    }
  }

  remove(...tokens: string[]): void {
    for (const token of tokens) this.classes.delete(token);
  }

  contains(token: string): boolean {
    return this.classes.has(token);
  }

  toggle(token: string, force?: boolean): boolean {
    const isPresent = force ?? !this.classes.has(token);
    if (isPresent) {
      this.classes.add(token);
    } else {
      this.classes.delete(token);
    }

    return isPresent;
  }

  replace(oldToken: string, newToken: string): boolean {
    if (!this.classes.delete(oldToken)) return false;
    this.classes.add(newToken);
    return true;
  }

  item(index: number): string | null {
    return [...this.classes][index] ?? null;
  }

  toString(): string {
    return [...this.classes].join(" ");
  }

  [Symbol.iterator](): IterableIterator<string> {
    return this.classes.values();
  }
}

class FakeElement extends FakeNode {
  readonly tagName: string;
  readonly children: FakeElement[] = [];
  readonly dataset: Record<string, string> = {};
  readonly style: Record<string, string> = {};
  readonly classList: FakeClassList;

  id = "";
  hidden = false;
  textContent = "";

  private readonly classes = new Set<string>();
  private readonly attributes = new Map<string, string>();

  constructor(tagName: string) {
    super();
    this.tagName = tagName.toUpperCase();
    this.classList = new FakeClassList(this.classes);
  }

  get className(): string {
    return [...this.classes].join(" ");
  }

  set className(value: string) {
    this.classes.clear();
    for (const token of value.split(/\s+/u)) {
      if (token.length > 0) this.classes.add(token);
    }
  }

  override get isConnected(): boolean {
    if (this.parentElement) {
      return this.parentElement.isConnected;
    }

    return false;
  }

  appendChild(child: FakeElement): FakeElement {
    child.remove();
    child.parentElement = this;
    this.children.push(child);
    return child;
  }

  append(...nodes: Array<FakeElement | string>): void {
    for (const node of nodes) {
      if (typeof node === "string") {
        continue;
      }

      this.appendChild(node);
    }
  }

  replaceChildren(...nodes: FakeElement[]): void {
    for (const child of this.children) {
      child.parentElement = null;
    }

    this.children.length = 0;
    for (const node of nodes) {
      this.appendChild(node);
    }
  }

  removeChild(node: FakeNode): void {
    const index = this.children.findIndex((child) => child === node);
    if (index < 0) {
      return;
    }

    this.children[index].parentElement = null;
    this.children.splice(index, 1);
  }

  contains(target: Element): boolean {
    if (!(target instanceof FakeElement)) {
      return false;
    }

    if (target === this) {
      return true;
    }

    for (const child of this.children) {
      if (child.contains(target)) {
        return true;
      }
    }

    return false;
  }

  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
    if (name === "id") {
      this.id = value;
    }
    if (name === "class") {
      this.className = value;
    }
  }

  getAttribute(name: string): string | null {
    return this.attributes.get(name) ?? null;
  }

  querySelectorAll<T extends Element = Element>(selector: string): T[] {
    const matches: T[] = [];

    const visit = (node: FakeElement): void => {
      if (matchesSelector(node, selector)) {
        matches.push(node as DynamicValue as T);
      }

      for (const child of node.children) {
        visit(child);
      }
    };

    for (const child of this.children) {
      visit(child);
    }

    return matches;
  }

  querySelector<T extends Element = Element>(selector: string): T | null {
    return this.querySelectorAll<T>(selector)[0] ?? null;
  }

  closest<T extends Element = Element>(selector: string): T | null {
    const selectors = selector.split(",").map((part) => part.trim()).filter((part) => part.length > 0);

    for (const part of selectors) {
      if (matchesSelector(this, part)) {
        return this as DynamicValue as T;
      }
    }

    let parent = this.parentElement;
    while (parent) {
      for (const part of selectors) {
        if (matchesSelector(parent, part)) {
          return parent as DynamicValue as T;
        }
      }

      parent = parent.parentElement;
    }

    return null;
  }
}

class FakeHTMLElement extends FakeElement {
  constructor(tagName: string) {
    super(tagName);
  }

  focus(_options?: FocusOptions): void {
    if (this.ownerDocument instanceof FakeDocument) {
      this.ownerDocument.activeElement = this;
    }

    this.dispatchEvent(new Event("focus"));
    this.dispatchEvent(new Event("focusin", { bubbles: true }));
  }

  blur(): void {
    if (this.ownerDocument instanceof FakeDocument && this.ownerDocument.activeElement === this) {
      this.ownerDocument.activeElement = null;
    }

    this.dispatchEvent(new Event("blur"));
    this.dispatchEvent(new Event("focusout", { bubbles: true }));
  }
}

class FakeBodyElement extends FakeHTMLElement {
  constructor() {
    super("body");
  }

  override get isConnected(): boolean {
    return true;
  }
}
class FakeHTMLInputElement extends FakeHTMLElement {
  type = "text";

  constructor() {
    super("input");
  }
}

class FakeHTMLTextAreaElement extends FakeHTMLElement {
  value = "";

  constructor() {
    super("textarea");
  }
}

class FakeHTMLSelectElement extends FakeHTMLElement {
  constructor() {
    super("select");
  }
}

class FakeCustomEvent extends Event {
  readonly detail: DynamicValue;

  constructor(type: string, init?: CustomEventInit<DynamicValue>) {
    super(type, init);
    this.detail = init?.detail;
  }
}

class FakeDocument extends FakeNode {
  readonly body: FakeHTMLElement;
  readonly documentElement: FakeHTMLElement;
  activeElement: FakeElement | null = null;

  constructor() {
    super();
    this.documentElement = new FakeHTMLElement("html");
    this.documentElement.ownerDocument = this;
    this.body = new FakeBodyElement();
    this.body.ownerDocument = this;
  }

  createElement(tagName: string): Element {
    const normalized = tagName.trim().toLowerCase();
    let element: Element;

    if (normalized === "input") {
      element = new FakeHTMLInputElement() as DynamicValue as Element;
    } else if (normalized === "textarea") {
      element = new FakeHTMLTextAreaElement() as DynamicValue as Element;
    } else if (normalized === "select") {
      element = new FakeHTMLSelectElement() as DynamicValue as Element;
    } else {
      element = new FakeHTMLElement(normalized) as DynamicValue as Element;
    }

    (element as DynamicValue as FakeElement).ownerDocument = this;
    return element;
  }

  createElementNS(_namespace: string | null, qualifiedName: string): Element {
    return this.createElement(qualifiedName);
  }

  getElementById(id: string): HTMLElement | null {
    const visit = (node: FakeElement): FakeElement | null => {
      if (node.id === id) {
        return node;
      }

      for (const child of node.children) {
        const match = visit(child);
        if (match) return match;
      }

      return null;
    };

    const found = visit(this.body);
    return found as DynamicValue as HTMLElement | null;
  }

  querySelectorAll(selector: string): Element[] {
    const matches: Element[] = [];

    const visit = (node: FakeElement): void => {
      if (matchesSelector(node, selector)) {
        matches.push(node as DynamicValue as Element);
      }

      for (const child of node.children) {
        visit(child);
      }
    };

    visit(this.body);
    return matches;
  }

  querySelector(selector: string): Element | null {
    return this.querySelectorAll(selector)[0] ?? null;
  }
}

function matchesSelector(node: FakeElement, selector: string): boolean {
  const selectors = selector.split(",").map((part) => part.trim()).filter((part) => part.length > 0);
  if (selectors.length > 1) {
    return selectors.some((part) => matchesSelector(node, part));
  }

  if (selector === "textarea") {
    return node.tagName === "TEXTAREA";
  }

  if (selector === "input") {
    return node.tagName === "INPUT";
  }

  if (selector === "select") {
    return node.tagName === "SELECT";
  }

  if (selector === "[contenteditable='true']" || selector === "[contenteditable='plaintext-only']") {
    const value = node.getAttribute("contenteditable");
    return value === "true" || value === "plaintext-only";
  }

  if (selector === "[data-claims-escape='true']") {
    return node.dataset.claimsEscape === "true";
  }

  const attributePresence = /^\[([A-Za-z0-9_-]+)\]$/u.exec(selector);
  if (attributePresence) {
    return node.getAttribute(attributePresence[1] ?? "") !== null;
  }

  if (selector.startsWith(".")) {
    const cls = selector.slice(1);
    return node.className.split(/\s+/u).includes(cls);
  }

  if (selector.startsWith("#")) {
    return node.id === selector.slice(1);
  }

  // Bare tag name (e.g. "button"), case-insensitive.
  if (/^[A-Za-z][A-Za-z0-9-]*$/u.test(selector)) {
    return node.tagName === selector.toUpperCase();
  }

  return false;
}

export interface FakeDomHandle {
  document: Document;
  restore: () => void;
}

export function installFakeDom(): FakeDomHandle {
  const previous: ConstructorSnapshot = {
    Node: Reflect.get(globalThis, "Node"),
    Element: Reflect.get(globalThis, "Element"),
    HTMLElement: Reflect.get(globalThis, "HTMLElement"),
    HTMLInputElement: Reflect.get(globalThis, "HTMLInputElement"),
    HTMLTextAreaElement: Reflect.get(globalThis, "HTMLTextAreaElement"),
    HTMLSelectElement: Reflect.get(globalThis, "HTMLSelectElement"),
    CustomEvent: Reflect.get(globalThis, "CustomEvent"),
    document: Reflect.get(globalThis, "document"),
  };

  const fakeDocument = new FakeDocument();

  Reflect.set(globalThis, "Node", FakeNode);
  Reflect.set(globalThis, "Element", FakeElement);
  Reflect.set(globalThis, "HTMLElement", FakeHTMLElement);
  Reflect.set(globalThis, "HTMLInputElement", FakeHTMLInputElement);
  Reflect.set(globalThis, "HTMLTextAreaElement", FakeHTMLTextAreaElement);
  Reflect.set(globalThis, "HTMLSelectElement", FakeHTMLSelectElement);
  Reflect.set(globalThis, "CustomEvent", FakeCustomEvent);
  Reflect.set(globalThis, "document", fakeDocument);

  return {
    document: fakeDocument as DynamicValue as Document,
    restore: () => {
      restoreGlobal("Node", previous.Node);
      restoreGlobal("Element", previous.Element);
      restoreGlobal("HTMLElement", previous.HTMLElement);
      restoreGlobal("HTMLInputElement", previous.HTMLInputElement);
      restoreGlobal("HTMLTextAreaElement", previous.HTMLTextAreaElement);
      restoreGlobal("HTMLSelectElement", previous.HTMLSelectElement);
      restoreGlobal("CustomEvent", previous.CustomEvent);
      restoreGlobal("document", previous.document);
    },
  };
}

function restoreGlobal(key: string, value: DynamicValue): void {
  if (value === undefined) {
    const deleted = Reflect.deleteProperty(globalThis, key);
    assert.equal(deleted, true);
    return;
  }

  Reflect.set(globalThis, key, value);
}
