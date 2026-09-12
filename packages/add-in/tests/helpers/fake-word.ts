/**
 * Harness for mocking the global `Word` object so the live Word tools can be
 * exercised under Node. Provides a recording body/paragraph/range proxy
 * surface sufficient for the insertion engine, markdown path, overview read,
 * replace/search, and table/page-break/image/hyperlink tools.
 */

export interface FakeParagraph {
  text: string;
  style?: string;
  alignment?: string;
  font: Record<string, DynamicValue>;
  spaceBefore?: number;
  spaceAfter?: number;
  lineSpacing?: number;
  firstLineIndent?: number;
  leftIndent?: number;
  calls: string[];
  getRange(): FakeRange;
  load(_props?: string): void;
}

export interface FakeRange {
  text: string;
  font: Record<string, DynamicValue>;
  paragraphs: { getFirst(): FakeParagraph };
  hyperlinks: { add(range: FakeRange, opts?: Record<string, string>): void };
  calls: string[];
  insertText(text: string, loc: string): FakeRange;
  insertParagraph(text: string, loc: string): FakeParagraph;
  insertBreak(breakType: string, loc: string): void;
  load(_props?: string): void;
}

export interface FakeTable {
  style?: string;
  alignment?: string;
  rows: {
    getFirst(): { font: Record<string, DynamicValue> };
    load(_p?: string): void;
  };
  columns: { load(_p?: string): void };
  calls: string[];
  load(_props?: string): void;
}

function makeFont(): Record<string, DynamicValue> {
  const target: Record<string, DynamicValue> = {};
  return target;
}

export function makeRange(text = ""): FakeRange {
  const font = makeFont();
  const calls: string[] = [];
  const paragraph = createParagraph(text, calls);
  const range: FakeRange = {
    text,
    font,
    calls,
    paragraphs: { getFirst: () => paragraph },
    hyperlinks: {
      add: (_r, opts) => {
        calls.push(`hyperlinks.add(${JSON.stringify(opts ?? {})})`);
      },
    },
    insertText(newText, loc) {
      calls.push(`insertText(${JSON.stringify(newText)}, ${String(loc)})`);
      return makeRange(newText);
    },
    insertParagraph(newText, loc) {
      calls.push(`insertParagraph(${JSON.stringify(newText)}, ${String(loc)})`);
      return createParagraph(newText, calls);
    },
    insertBreak(breakType, loc) {
      calls.push(`insertBreak(${String(breakType)}, ${String(loc)})`);
    },
    load() {
      calls.push("range.load()");
    },
  };
  return range;
}

function createParagraph(text: string, sharedCalls: string[]): FakeParagraph {
  const paragraph: FakeParagraph = {
    text,
    style: undefined,
    alignment: undefined,
    font: makeFont(),
    calls: sharedCalls,
    getRange() {
      sharedCalls.push("paragraph.getRange()");
      return makeRange(text);
    },
    load() {
      sharedCalls.push("paragraph.load()");
    },
  };
  return paragraph;
}

/** Collection of proxies with a load() no-op (paragraphs/tables/rows). */
interface FakeCollection<T> {
  items: T[];
  load(_props?: string): void;
}

function makeCollection<T>(items: T[]): FakeCollection<T> {
  return {
    items,
    load() {
      // no-op: proxies are plain data in the fake.
    },
  };
}

export interface FakeSearchResults {
  items: FakeRange[];
  load(_props?: string): void;
}
export interface FakeBody {
  calls: string[];
  text: string;
  paragraphs: FakeCollection<FakeParagraph>;
  tables: FakeCollection<FakeTable>;
  load(_props?: string): void;
  search(find: string, _opts?: Record<string, DynamicValue>): FakeSearchResults;
  insertParagraph(text: string, loc: string): FakeParagraph;
  insertText(text: string, loc: string): FakeRange;
  insertTable(
    rows: number,
    cols: number,
    loc: string,
    values: string[][],
  ): FakeTable;
  insertBreak(breakType: string, loc: string): void;
  insertInlinePictureFromBase64(
    b64: string,
    loc: string,
  ): {
    width?: number;
    height?: number;
    calls: string[];
  };
}

export function makeBody(sharedCalls?: string[]): FakeBody {
  const calls: string[] = sharedCalls ?? [];
  const body: FakeBody = {
    calls,
    text: "",
    paragraphs: makeCollection([]),
    tables: makeCollection([]),
    load() {
      calls.push("body.load()");
    },
    search(find) {
      calls.push(`search(${JSON.stringify(find)})`);
      return { items: [], load() {} };
    },
    insertParagraph(text, loc) {
      calls.push(`insertParagraph(${JSON.stringify(text)}, ${String(loc)})`);
      return createParagraph(text, calls);
    },
    insertText(text, loc) {
      calls.push(`insertText(${JSON.stringify(text)}, ${String(loc)})`);
      return makeRange(text);
    },
    insertTable(rows, cols, loc, values) {
      calls.push(
        `insertTable(${rows},${cols},${String(loc)},${JSON.stringify(values)})`,
      );
      const table: FakeTable = {
        rows: { getFirst: () => ({ font: makeFont() }), load() {} },
        columns: { load() {} },
        calls,
        load() {},
      };
      return table;
    },
    insertBreak(breakType, loc) {
      calls.push(`insertBreak(${String(breakType)}, ${String(loc)})`);
    },
    insertInlinePictureFromBase64(b64, loc) {
      calls.push(
        `insertInlinePictureFromBase64(${b64.length} chars, ${String(loc)})`,
      );
      return { calls };
    },
  };
  return body;
}

export interface FakeContext {
  calls: string[];
  body: FakeBody;
  document: { body: FakeBody; getSelection(): FakeRange };
  sync(): Promise<void>;
}

export function makeContext(): FakeContext {
  const calls: string[] = [];
  const body = makeBody(calls);
  const context: FakeContext = {
    calls,
    body,
    document: {
      body,
      getSelection() {
        calls.push("getSelection()");
        return makeRange();
      },
    },
    sync(): Promise<void> {
      calls.push("sync()");
      return Promise.resolve();
    },
  };
  return context;
}

/** Call a tool factory's execute with the fake Word harness installed. */
export async function runWithFakeWord<T>(
  context: FakeContext,
  fn: () => T | Promise<T>,
): Promise<T> {
  const priorWord = (globalThis as Record<string, DynamicValue>).Word;
  (globalThis as Record<string, DynamicValue>).Word = {
    run: (callback: (ctx: FakeContext) => Promise<T>): Promise<T> =>
      callback(context),
    InsertLocation: {
      before: "Before",
      replace: "Replace",
      start: "Start",
      end: "End",
      after: "After",
    },
  };
  try {
    return await fn();
  } finally {
    if (priorWord === undefined) {
      delete (globalThis as Record<string, DynamicValue>).Word;
    } else {
      (globalThis as Record<string, DynamicValue>).Word = priorWord;
    }
  }
}
