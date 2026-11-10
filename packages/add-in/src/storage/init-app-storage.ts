/**
 * Shared app storage initialization for taskpane + dialog.
 */

import { AppStorage, setAppStorage } from "./local/app-storage.js";
import { IndexedDBStorageBackend } from "./local/indexeddb-storage-backend.js";
import { CustomProvidersStore } from "./local/custom-providers-store.js";
import { ModelCatalogsStore } from "./local/model-catalogs-store.js";
import { ProviderKeysStore } from "./local/provider-keys-store.js";
import { SessionsStore } from "./local/sessions-store.js";
import { SettingsStore } from "./local/settings-store.js";

type InitializedAppStorage = {
  storage: AppStorage;
  settings: SettingsStore;
  providerKeys: ProviderKeysStore;
  sessions: SessionsStore;
  customProviders: CustomProvidersStore;
  modelCatalogs: ModelCatalogsStore;
  backend: IndexedDBStorageBackend;
};

/**
 * Current IndexedDB database name (post-rebrand).
 *
 * Renamed from "pi-for-excel" to "pi-for-office". Existing user data is copied
 * into this database once by IndexedDBStorageBackend's `migrateFrom` hook; the
 * legacy database is left untouched so nobody loses data if a rollback is needed.
 */
export const APP_DATABASE_NAME = "pi-for-office";

/**
 * IndexedDB schema version.
 *
 * v3 is the index-repair bump: opening an older database now adds any index the
 * current store config declares but the existing object store lacks (the
 * `lastModified` index was missing on databases created before it existed,
 * which made session listings throw).
 */
export const APP_DATABASE_VERSION = 3;

/** Pre-rebrand database name — copied from on first open, then left alone. */
export const LEGACY_DATABASE_NAME = "pi-for-excel";

export function initAppStorage(
  dbName: string = APP_DATABASE_NAME,
): InitializedAppStorage {
  const settings = new SettingsStore();
  const providerKeys = new ProviderKeysStore();
  const sessions = new SessionsStore();
  const customProviders = new CustomProvidersStore();
  const modelCatalogs = new ModelCatalogsStore();

  // One-time copy of pre-rebrand user data when using the production name.
  const backendOptions: { migrateFrom?: string } =
    dbName === APP_DATABASE_NAME ? { migrateFrom: LEGACY_DATABASE_NAME } : {};

  const backend = new IndexedDBStorageBackend(
    {
      dbName,
      version: APP_DATABASE_VERSION,
      stores: [
        settings.getConfig(),
        providerKeys.getConfig(),
        sessions.getConfig(),
        SessionsStore.getMetadataConfig(),
        customProviders.getConfig(),
        modelCatalogs.getConfig(),
      ],
    },
    backendOptions,
  );

  settings.setBackend(backend);
  providerKeys.setBackend(backend);
  sessions.setBackend(backend);
  customProviders.setBackend(backend);
  modelCatalogs.setBackend(backend);

  const storage = new AppStorage(
    settings,
    providerKeys,
    sessions,
    customProviders,
    modelCatalogs,
    backend,
  );
  setAppStorage(storage);

  return {
    storage,
    settings,
    providerKeys,
    sessions,
    customProviders,
    modelCatalogs,
    backend,
  };
}
