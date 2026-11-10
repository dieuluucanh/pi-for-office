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

/** Shape of the migration marker (read back only for presence). */
type MigrationMarker = {
  from?: string;
  at?: number;
  status?: string;
  copied?: number;
  skipped?: number;
} | string | number | null;

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
   * database.
   *
   * Safety rules (why this is not a plain `put` loop any more):
   * - **insert-if-absent**: an existing target row always wins, so a retried or
   *   interrupted copy can never rewind newer sessions, the workbook's
   *   `latestSession` pointer, or the saved tab layout to a pre-rebrand snapshot;
   * - **one transaction**: the copy and the completion marker commit together,
   *   so a partially applied copy cannot happen (it either lands with its
   *   marker or not at all);
   * - **bounded open**: a legacy database locked by another connection must not
   *   hang taskpane boot.
   *
   * Never throws: a failed migration must not block the app from opening.
   *
   * NOTE: must not call `this.get`/`this.set` (they re-enter getDB and deadlock).
   */
  private async migrateFromLegacyDatabase(db: IDBDatabase): Promise<void> {
    const legacyName = this.migrateFrom;
    if (!legacyName) return;

    try {
      // Only Chromium/Edge webviews expose indexedDB.databases(); on
      // WKWebView (macOS) we can't enumerate safely, so we skip migration.
      if (!indexedDB.databases) return;

      const databases = await indexedDB.databases();
      const legacyExists = databases.some((entry) => entry.name === legacyName);
      if (!legacyExists) return;
      if (!db.objectStoreNames.contains("settings")) return;

      const marker = await this.promisifyRequest<MigrationMarker>(
        db
          .transaction("settings", "readonly")
          .objectStore("settings")
          .get(MIGRATION_MARKER_KEY) as IDBRequest<MigrationMarker>,
      );
      if (marker !== undefined && marker !== null) return;

      const legacyDb = await this.openDatabaseWithTimeout(legacyName);
      if (!legacyDb) return;

      try {
        const copyableStores = this.config.stores
          .map((storeConfig) => storeConfig.name)
          .filter(
            (storeName) =>
              legacyDb.objectStoreNames.contains(storeName) &&
              db.objectStoreNames.contains(storeName),
          );

        if (copyableStores.length === 0) {
          await this.writeMigrationMarker(db, legacyName, 0, 0);
          return;
        }

        // Read the legacy rows *before* opening the target transaction: a
        // transaction cannot survive interleaved awaits against another DB.
        const legacyRows = new Map<string, { keys: IDBValidKey[]; values: DynamicValue[] }>();
        for (const storeName of copyableStores) {
          const legacyStore = legacyDb
            .transaction(storeName, "readonly")
            .objectStore(storeName);
          const [keys, values] = await Promise.all([
            this.promisifyRequest<IDBValidKey[]>(legacyStore.getAllKeys()),
            this.promisifyRequest<DynamicValue[]>(legacyStore.getAll()),
          ]);
          legacyRows.set(storeName, { keys: keys ?? [], values: values ?? [] });
        }

        const storeNames = [...copyableStores, "settings"];
        const targetTx = db.transaction(storeNames, "readwrite");
        const committed = this.awaitTransaction(targetTx);

        let copied = 0;
        let skipped = 0;

        try {
          for (const storeName of copyableStores) {
            const rows = legacyRows.get(storeName);
            if (!rows) continue;

            const targetStore = targetTx.objectStore(storeName);
            for (let index = 0; index < rows.keys.length; index += 1) {
              const key = rows.keys[index] as IDBValidKey;
              const value = rows.values[index];

              const existing = await this.promisifyRequest(
                targetStore.getKey(key),
              );
              if (existing !== undefined) {
                skipped += 1;
                continue;
              }

              if (targetStore.keyPath) {
                await this.promisifyRequest(targetStore.put(value));
              } else {
                await this.promisifyRequest(targetStore.put(value, key));
              }
              copied += 1;
            }
          }

          // Marker commits atomically with the copied rows.
          targetTx.objectStore("settings").put(
            {
              from: legacyName,
              at: Date.now(),
              status: "complete",
              copied,
              skipped,
            },
            MIGRATION_MARKER_KEY,
          );

          await committed;
        } catch (error) {
          await committed.catch(() => undefined);
          throw error;
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

  /** Open a database, giving up (null) instead of hanging boot on `blocked`. */
  private openDatabaseWithTimeout(
    name: string,
    timeoutMs = 5000,
  ): Promise<IDBDatabase | null> {
    return new Promise((resolve) => {
      let settled = false;
      const finish = (db: IDBDatabase | null) => {
        if (settled) return;
        settled = true;
        resolve(db);
      };

      try {
        const request = indexedDB.open(name);
        request.onerror = () => finish(null);
        request.onblocked = () => finish(null);
        request.onsuccess = () => finish(request.result);
      } catch {
        finish(null);
        return;
      }

      globalThis.setTimeout(() => finish(null), timeoutMs);
    });
  }

  private async writeMigrationMarker(
    db: IDBDatabase,
    legacyName: string,
    copied: number,
    skipped: number,
  ): Promise<void> {
    const tx = db.transaction("settings", "readwrite");
    const committed = this.awaitTransaction(tx);
    tx.objectStore("settings").put(
      {
        from: legacyName,
        at: Date.now(),
        status: "complete",
        copied,
        skipped,
      },
      MIGRATION_MARKER_KEY,
    );
    await committed;
  }

  private async getDB(): Promise<IDBDatabase> {
    if (!this.dbPromise) {
      this.dbPromise = new Promise((resolve, reject) => {
        const request = indexedDB.open(this.config.dbName, this.config.version);
        request.onerror = () => reject(toIdbError(request.error));
        request.onblocked = () => {
          console.warn(
            "[pi-for-office] IndexedDB upgrade is blocked by another open task pane or tab. Close other Pi panes and reload.",
          );
        };
        request.onupgradeneeded = () => {
          const db = request.result;
          const upgradeTx = request.transaction;

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

          // Existing stores keep whatever schema they were created with. Add
          // any index the current config declares but the store lacks — an
          // older database without the `lastModified` index made every session
          // listing throw, which is invisible data loss for the user.
          if (!upgradeTx) return;
          for (const storeConfig of this.config.stores) {
            if (!db.objectStoreNames.contains(storeConfig.name)) continue;

            const store = upgradeTx.objectStore(storeConfig.name);
            for (const indexConfig of storeConfig.indices ?? []) {
              if (store.indexNames.contains(indexConfig.name)) continue;

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

  /**
   * Resolve when a transaction really commits.
   *
   * `IDBRequest.onsuccess` only means the *request* succeeded: the transaction
   * can still abort (quota, teardown), so awaiting the request alone reported
   * saves that were never durable.
   */
  private awaitTransaction(tx: IDBTransaction): Promise<void> {
    return new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(toIdbError(tx.error));
      tx.onabort = () => reject(toIdbError(tx.error));
    });
  }

  async set<T = DynamicValue>(
    storeName: string,
    key: string,
    value: T,
  ): Promise<void> {
    const db = await this.getDB();
    const tx = db.transaction(storeName, "readwrite");
    const committed = this.awaitTransaction(tx);
    try {
      await this.putIntoStore(tx.objectStore(storeName), key, value);
      await committed;
    } catch (error) {
      await committed.catch(() => undefined);
      throw error;
    }
  }

  async delete(storeName: string, key: string): Promise<void> {
    const db = await this.getDB();
    const tx = db.transaction(storeName, "readwrite");
    const committed = this.awaitTransaction(tx);
    try {
      await this.promisifyRequest(tx.objectStore(storeName).delete(key));
      await committed;
    } catch (error) {
      await committed.catch(() => undefined);
      throw error;
    }
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

  async getAll<T = DynamicValue>(storeName: string): Promise<T[]> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    const result = await this.promisifyRequest<T[]>(store.getAll() as IDBRequest<T[]>);
    return result ?? [];
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
    const tx = db.transaction(storeName, "readwrite");
    const committed = this.awaitTransaction(tx);
    try {
      await this.promisifyRequest(tx.objectStore(storeName).clear());
      await committed;
    } catch (error) {
      await committed.catch(() => undefined);
      throw error;
    }
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
    const committed = this.awaitTransaction(idbTx);
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

    try {
      const result = await operation(storageTx);
      await committed;
      return result;
    } catch (error) {
      // The transaction may already be aborting; surface the operation error.
      await committed.catch(() => undefined);
      throw error;
    }
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
