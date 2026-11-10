/**
 * One-time persistent-storage request.
 *
 * IndexedDB data on an origin that has not been granted persistent storage may
 * be evicted under storage pressure — in WebView2 (Excel on Windows) as well as
 * in a browser tab. The app already exposed `requestPersistence()` but nothing
 * ever called it, so every install was in the "best effort" bucket.
 *
 * The result is recorded so the diagnostics readout can state whether the store
 * is protected, and so an already-granted origin is never re-requested.
 */

import type { SettingsStore } from "./settings-store.js";
import type { StorageBackend } from "./types.js";
import { STORAGE_PERSISTENCE_MARKER_KEY } from "./diagnostics.js";

export interface PersistenceRequestResult {
  granted: boolean;
  /** False when a previous grant made the call unnecessary (or it failed). */
  attempted: boolean;
}

interface PersistenceMarker {
  granted?: DynamicValue;
}

function isGrantedMarker(value: DynamicValue): boolean {
  return (
    typeof value === "object" &&
    value !== null &&
    (value as PersistenceMarker).granted === true
  );
}

export async function ensurePersistentStorage(deps: {
  backend: Pick<StorageBackend, "requestPersistence">;
  settings: SettingsStore;
  now?: () => number;
}): Promise<PersistenceRequestResult> {
  try {
    const marker = await deps.settings.get(STORAGE_PERSISTENCE_MARKER_KEY);
    if (isGrantedMarker(marker)) {
      return { granted: true, attempted: false };
    }

    const granted = await deps.backend.requestPersistence();

    await deps.settings.set(STORAGE_PERSISTENCE_MARKER_KEY, {
      granted,
      at: new Date(deps.now?.() ?? Date.now()).toISOString(),
    });

    return { granted, attempted: true };
  } catch (error) {
    console.warn("[pi] Persistent storage request failed:", error);
    return { granted: false, attempted: false };
  }
}
