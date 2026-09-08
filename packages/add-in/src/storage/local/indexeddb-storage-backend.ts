/**
 * IndexedDB implementation of StorageBackend.
 * Provides multi-store key-value storage with transactions and quota management.
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 *
 * Schema compatibility invariant: this backend must open the existing
 * "pi-for-office" database (same store names, key paths, and indices) so
 * user data restores unchanged. When constructed with `migrateFrom`, it
 * copies legacy user data from the pre-rebrand database name ("pi-for-excel")
 * into the new database on the first successful open — then marks the
 * migration so it never runs again.
 */

import type { IndexedDBConfig, StorageTransaction } from "./types.js";
import type { StorageBackend } from "./types.js";

function toIdbError(error: DOMException | null): Error {
  return error ?? new Error("IndexedDB request failed");
}

/** Marker key written to the settings store after a legacy-db copy completes. */
const MIGRATION_MARKER_KEY = "__pi_migrated_legacy_db_v1";

/** All stores in this app use string (or numeric) keys; other key shapes are skipped. */
function normalizeIdbKey(key: IDBValidKey): string[] {
  if (typeof key === "string") return [key];
  if (typeof key === "number") return [key.toString()];
  return [];
}

export class IndexedDBStorageBackend implements StorageBackend {
  private dbPromise: Promise<IDBDatabase> | null = null;
  private readonly config: IndexedDBConfig;
  private readonly migrateFrom: string | undefined;
  private migrationAttempted = false;

  constructor(config: IndexedDBConfig, options?: { migrateFrom?: string }) {
    this.config = config;
    this.migrateFrom = options?.migrateFrom;
  }

  /**
   * Copy user data from a legacy database (pre-rebrand name) into the current
   * database. Same store names/key paths by schema invariant, so a straight
   * per-store key/value copy restores sessions, settings, keys, and catalogs.
   *
   * Never throws: a failed migration must not block the app from opening — the
   * new database simply starts empty (worst case the user reconnects).
   *
   * NOTE: must not call `this.get`/`this.set` (they re-enter getDB and deadlock).
   */
  private async migrateFromLegacyDatabase(db: IDBDatabase): Promise<void> {
    const legacyName = this.migrateFrom;
    if (!legacyName) return;

    try {
      // Only Chromium/Edge webviews expose indexedDB.databases(); on
      // WKWebView (macOS) we can't enumerate safely, so we skip migration.
      if (typeof indexedDB.databases !== "function") return;

      const databases = await indexedDB.databases();
      const legacyExists = databases.some((entry) => entry.name === legacyName);
      if (!legacyExists) return;

      const settingsStore = db.objectStoreNames.contains("settings")
        ? db.transaction("settings", "readonly").objectStore("settings")
        : null;
      if (settingsStore) {
        const marker = await this.promisifyRequest(
          settingsStore.get(MIGRATION_MARKER_KEY),
        );
        if (marker !== undefined) return;
      }

      const legacyDb = await new Promise<IDBDatabase>((resolve, reject) => {
        const req = indexedDB.open(legacyName);
        req.onerror = () => reject(toIdbError(req.error));
        req.onsuccess = () => resolve(req.result);
      });

      try {
        for (const storeConfig of this.config.stores) {
          const storeName = storeConfig.name;
          if (!legacyDb.objectStoreNames.contains(storeName)) continue;
          if (!db.objectStoreNames.contains(storeName)) continue;

          const legacyTx = legacyDb.transaction(storeName, "readonly");
          const legacyStore = legacyTx.objectStore(storeName);
          const [keys, values] = await Promise.all([
            this.promisifyRequest(legacyStore.getAllKeys()),
            this.promisifyRequest(legacyStore.getAll()),
          ]);
          if (!keys || keys.length === 0) continue;

          const targetTx = db.transaction(storeName, "readwrite");
          const targetStore = targetTx.objectStore(storeName);
          for (let i = 0; i < keys.length; i += 1) {
            const key = keys[i] as IDBValidKey;
            const value = values[i];
            if (targetStore.keyPath) {
              await this.promisifyRequest(targetStore.put(value));
            } else {
              await this.promisifyRequest(targetStore.put(value, key));
            }
          }
          await new Promise<void>((resolve, reject) => {
            targetTx.oncomplete = () => resolve();
            targetTx.onerror = () => reject(toIdbError(targetTx.error));
            targetTx.onabort = () => reject(toIdbError(targetTx.error));
          });
        }

        // Mark as migrated so a later open does not copy again.
        if (db.objectStoreNames.contains("settings")) {
          const markTx = db.transaction("settings", "readwrite");
          markTx
            .objectStore("settings")
            .put({ from: legacyName, at: Date.now() }, MIGRATION_MARKER_KEY);
          await new Promise<void>((resolve, reject) => {
            markTx.oncomplete = () => resolve();
            markTx.onerror = () => reject(toIdbError(markTx.error));
          });
        }
      } finally {
        legacyDb.close();
      }
    } catch (error) {
      console.warn(
        "[pi-for-office] Legacy IndexedDB migration skipped:",
        error,
      );
    }
  }

  private async getDB(): Promise<IDBDatabase> {
    if (!this.dbPromise) {
      this.dbPromise = new Promise((resolve, reject) => {
        const request = indexedDB.open(this.config.dbName, this.config.version);
        request.onerror = () => reject(toIdbError(request.error));
        request.onupgradeneeded = () => {
          const db = request.result;
          // Create object stores from config
          for (const storeConfig of this.config.stores) {
            if (db.objectStoreNames.contains(storeConfig.name)) continue;

            const params: IDBObjectStoreParameters = {};
            if (storeConfig.keyPath !== undefined)
              params.keyPath = storeConfig.keyPath;
            if (storeConfig.autoIncrement !== undefined) {
              params.autoIncrement = storeConfig.autoIncrement;
            }
            const store = db.createObjectStore(storeConfig.name, params);

            for (const indexConfig of storeConfig.indices ?? []) {
              const indexParams: IDBIndexParameters = {};
              if (indexConfig.unique !== undefined)
                indexParams.unique = indexConfig.unique;
              store.createIndex(
                indexConfig.name,
                indexConfig.keyPath,
                indexParams,
              );
            }
          }
        };
        request.onsuccess = () => {
          const db = request.result;
          if (this.migrateFrom && !this.migrationAttempted) {
            this.migrationAttempted = true;
            void this.migrateFromLegacyDatabase(db).then(
              () => resolve(db),
              () => resolve(db),
            );
            return;
          }
          resolve(db);
        };
      });
    }
    return this.dbPromise;
  }

  private promisifyRequest<T>(request: IDBRequest<T>): Promise<T> {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(toIdbError(request.error));
    });
  }

  /** Put honoring in-line (keyPath) vs out-of-line keys. */
  private putIntoStore(
    store: IDBObjectStore,
    key: string,
    value: DynamicValue,
  ): Promise<IDBValidKey> {
    if (store.keyPath) {
      return this.promisifyRequest(store.put(value));
    }
    return this.promisifyRequest(store.put(value, key));
  }

  async get<T = DynamicValue>(
    storeName: string,
    key: string,
  ): Promise<T | null> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    const result = await this.promisifyRequest<T | undefined>(
      store.get(key) as IDBRequest<T | undefined>,
    );
    return result ?? null;
  }

  async set<T = DynamicValue>(
    storeName: string,
    key: string,
    value: T,
  ): Promise<void> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readwrite").objectStore(storeName);
    await this.putIntoStore(store, key, value);
  }

  async delete(storeName: string, key: string): Promise<void> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readwrite").objectStore(storeName);
    await this.promisifyRequest(store.delete(key));
  }

  async keys(storeName: string, prefix?: string): Promise<string[]> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    if (prefix) {
      // Use IDBKeyRange for efficient prefix filtering
      const range = IDBKeyRange.bound(prefix, `${prefix}\uffff`, false, false);
      const keys = await this.promisifyRequest(store.getAllKeys(range));
      return keys.flatMap(normalizeIdbKey);
    }
    const keys = await this.promisifyRequest(store.getAllKeys());
    return keys.flatMap(normalizeIdbKey);
  }

  async getAllFromIndex<T = DynamicValue>(
    storeName: string,
    indexName: string,
    direction: "asc" | "desc" = "asc",
  ): Promise<T[]> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    const index = store.index(indexName);
    return new Promise((resolve, reject) => {
      const results: T[] = [];
      const request = index.openCursor(
        null,
        direction === "desc" ? "prev" : "next",
      );
      request.onsuccess = () => {
        const cursor = request.result;
        if (cursor) {
          results.push(cursor.value as T);
          cursor.continue();
        } else {
          resolve(results);
        }
      };
      request.onerror = () => reject(toIdbError(request.error));
    });
  }

  async clear(storeName: string): Promise<void> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readwrite").objectStore(storeName);
    await this.promisifyRequest(store.clear());
  }

  async has(storeName: string, key: string): Promise<boolean> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    const result = await this.promisifyRequest(store.getKey(key));
    return result !== undefined;
  }

  async transaction<T>(
    storeNames: string[],
    mode: "readonly" | "readwrite",
    operation: (tx: StorageTransaction) => Promise<T>,
  ): Promise<T> {
    const db = await this.getDB();
    const idbTx = db.transaction(storeNames, mode);
    const storageTx: StorageTransaction = {
      get: async <V = DynamicValue>(
        storeName: string,
        key: string,
      ): Promise<V | null> => {
        const store = idbTx.objectStore(storeName);
        const result = await this.promisifyRequest<V | undefined>(
          store.get(key) as IDBRequest<V | undefined>,
        );
        return result ?? null;
      },
      set: async <V = DynamicValue>(
        storeName: string,
        key: string,
        value: V,
      ): Promise<void> => {
        const store = idbTx.objectStore(storeName);
        await this.putIntoStore(store, key, value);
      },
      delete: async (storeName: string, key: string): Promise<void> => {
        const store = idbTx.objectStore(storeName);
        await this.promisifyRequest(store.delete(key));
      },
    };
    return operation(storageTx);
  }

  async getQuotaInfo(): Promise<{
    usage: number;
    quota: number;
    percent: number;
  }> {
    if (navigator.storage?.estimate) {
      const estimate = await navigator.storage.estimate();
      const usage = estimate.usage ?? 0;
      const quota = estimate.quota ?? 0;
      return {
        usage,
        quota,
        percent: quota ? (usage / quota) * 100 : 0,
      };
    }
    return { usage: 0, quota: 0, percent: 0 };
  }

  async requestPersistence(): Promise<boolean> {
    if (navigator.storage?.persist) {
      return await navigator.storage.persist();
    }
    return false;
  }
}
