/**
 * In-memory StorageBackend for tests.
 *
 * Mirrors the IndexedDB contract (including the failure modes the fix relies
 * on): index reads can be made to throw, and writes can be made to fail.
 */

import type {
  StorageBackend,
  StorageTransaction,
} from "../../src/storage/local/types.ts";

type Row = Record<string, DynamicValue> | string | number | boolean | null;

export class MemoryStorage implements StorageBackend {
  private readonly stores = new Map<string, Map<string, Row>>();
  private readonly indices = new Map<string, string>();

  /** Throw on `getAllFromIndex` (simulates a missing index). */
  failIndexReads = false;
  /** Reject every write (simulates quota/abort). */
  failWrites = false;

  constructor(config?: {
    stores?: Array<{
      name: string;
      keyPath?: string;
      indices?: Array<{ name: string; keyPath: string }>;
    }>;
  }) {
    for (const store of config?.stores ?? []) {
      this.ensureStore(store.name);
      for (const index of store.indices ?? []) {
        this.indices.set(`${store.name}:${index.name}`, index.keyPath);
      }
    }
  }

  private ensureStore(name: string): Map<string, Row> {
    let store = this.stores.get(name);
    if (!store) {
      store = new Map();
      this.stores.set(name, store);
    }
    return store;
  }

  private store(name: string): Map<string, Row> {
    const existing = this.stores.get(name);
    if (!existing) throw new Error(`Unknown store: ${name}`);
    return existing;
  }

  private field(value: Row, keyPath: string): string {
    if (typeof value !== "object" || value === null) return "";
    const raw = value[keyPath];
    if (typeof raw === "string") return raw;
    if (typeof raw === "number") return String(raw);
    return "";
  }

  get<T = DynamicValue>(storeName: string, key: string): Promise<T | null> {
    const value = this.store(storeName).get(key);
    if (value === undefined) return Promise.resolve(null);
    return Promise.resolve(structuredClone(value) as T);
  }

  set<T = DynamicValue>(
    storeName: string,
    key: string,
    value: T,
  ): Promise<void> {
    if (this.failWrites) return Promise.reject(new Error("write failed"));
    this.store(storeName).set(key, structuredClone(value) as Row);
    return Promise.resolve();
  }

  delete(storeName: string, key: string): Promise<void> {
    if (this.failWrites) return Promise.reject(new Error("write failed"));
    this.store(storeName).delete(key);
    return Promise.resolve();
  }

  keys(storeName: string, prefix?: string): Promise<string[]> {
    const keys = [...this.store(storeName).keys()];
    return Promise.resolve(
      prefix ? keys.filter((key) => key.startsWith(prefix)) : keys,
    );
  }

  getAll<T = DynamicValue>(storeName: string): Promise<T[]> {
    return Promise.resolve(
      [...this.store(storeName).values()].map(
        (value) => structuredClone(value) as T,
      ),
    );
  }

  getAllFromIndex<T = DynamicValue>(
    storeName: string,
    indexName: string,
    direction: "asc" | "desc" = "asc",
  ): Promise<T[]> {
    if (this.failIndexReads) {
      return Promise.reject(new Error(`Index ${indexName} is missing`));
    }

    const keyPath = this.indices.get(`${storeName}:${indexName}`) ?? indexName;

    const values = [...this.store(storeName).values()]
      .filter(
        (value) =>
          typeof value === "object" &&
          value !== null &&
          value[keyPath] !== undefined,
      )
      .sort((left, right) => {
        const a = this.field(left, keyPath);
        const b = this.field(right, keyPath);
        return direction === "desc" ? b.localeCompare(a) : a.localeCompare(b);
      });

    return Promise.resolve(values.map((value) => structuredClone(value) as T));
  }

  clear(storeName: string): Promise<void> {
    if (this.failWrites) return Promise.reject(new Error("write failed"));
    this.store(storeName).clear();
    return Promise.resolve();
  }

  has(storeName: string, key: string): Promise<boolean> {
    return Promise.resolve(this.store(storeName).has(key));
  }

  transaction<T>(
    storeNames: string[],
    mode: "readonly" | "readwrite",
    operation: (tx: StorageTransaction) => Promise<T>,
  ): Promise<T> {
    if (mode === "readwrite" && this.failWrites) {
      return Promise.reject(new Error("write failed"));
    }

    for (const storeName of storeNames) {
      this.store(storeName);
    }

    const tx: StorageTransaction = {
      get: <V = DynamicValue>(storeName: string, key: string) =>
        this.get<V>(storeName, key),
      set: <V = DynamicValue>(storeName: string, key: string, value: V) =>
        this.set<V>(storeName, key, value),
      delete: (storeName: string, key: string) => this.delete(storeName, key),
    };

    return operation(tx);
  }

  getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> {
    return Promise.resolve({ usage: 0, quota: 0, percent: 0 });
  }

  requestPersistence(): Promise<boolean> {
    return Promise.resolve(true);
  }

  /** Test helper: seed a raw row without going through validation. */
  seed(storeName: string, key: string, value: Row): void {
    this.ensureStore(storeName).set(key, structuredClone(value));
  }

  /** Test helper: read a raw row. */
  raw(storeName: string, key: string): Row | undefined {
    return this.store(storeName).get(key);
  }

  size(storeName: string): number {
    return this.store(storeName).size;
  }
}
