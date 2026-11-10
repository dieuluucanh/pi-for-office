import assert from "node:assert/strict";
import { test } from "node:test";

import { SessionsStore } from "../src/storage/local/sessions-store.ts";
import { MemoryStorage } from "./helpers/memory-storage.ts";
import {
  sessionMetadata,
  sessionRecord,
} from "./helpers/session-fixtures.ts";
import type { SessionData } from "../src/storage/local/types.ts";

const SESSIONS_CONFIG = [
  { name: "sessions", keyPath: "id", indices: [{ name: "lastModified", keyPath: "lastModified" }] },
  {
    name: "sessions-metadata",
    keyPath: "id",
    indices: [{ name: "lastModified", keyPath: "lastModified" }],
  },
];

function createStore(): { store: SessionsStore; backend: MemoryStorage } {
  const backend = new MemoryStorage({ stores: SESSIONS_CONFIG });
  const store = new SessionsStore();
  store.setBackend(backend);
  return { store, backend };
}

void test("getAllMetadataSafe lists indexed sessions newest first", async () => {
  const { store, backend } = createStore();
  backend.seed("sessions", "a", sessionRecord({ id: "a", lastModified: "2026-01-01T00:00:00.000Z" }));
  backend.seed("sessions", "b", sessionRecord({ id: "b", lastModified: "2026-03-01T00:00:00.000Z" }));
  backend.seed(
    "sessions-metadata",
    "a",
    sessionMetadata({ id: "a", lastModified: "2026-01-01T00:00:00.000Z" }),
  );
  backend.seed(
    "sessions-metadata",
    "b",
    sessionMetadata({ id: "b", lastModified: "2026-03-01T00:00:00.000Z" }),
  );

  const entries = await store.getAllMetadataSafe();
  assert.deepEqual(
    entries.map((entry) => entry.metadata.id),
    ["b", "a"],
  );
  assert.equal(entries[0]?.missingData, undefined);
});

void test("a missing index degrades to a full scan instead of hiding sessions", async () => {
  const { store, backend } = createStore();
  backend.seed("sessions", "a", sessionRecord({ id: "a", lastModified: "2026-01-01T00:00:00.000Z" }));
  backend.seed("sessions", "b", sessionRecord({ id: "b", lastModified: "2026-02-01T00:00:00.000Z" }));
  backend.failIndexReads = true;

  const entries = await store.getAllMetadataSafe();
  assert.equal(entries.length, 2);
  assert.ok(entries.every((entry) => entry.degraded === true));
  assert.deepEqual(
    entries.map((entry) => entry.metadata.id),
    ["b", "a"],
  );
});

void test("a record without a valid lastModified is still listed and ordered by creation", async () => {
  const { store, backend } = createStore();
  const invalid: SessionData = {
    ...sessionRecord({ id: "old", lastModified: "2026-01-01T00:00:00.000Z" }),
    lastModified: "not-a-date",
  };
  backend.seed("sessions", "old", invalid);
  backend.seed("sessions", "new", sessionRecord({ id: "new", lastModified: "2026-05-01T00:00:00.000Z" }));

  const entries = await store.getAllMetadataSafe();
  assert.deepEqual(
    entries.map((entry) => entry.metadata.id),
    ["new", "old"],
  );
  // The invalid timestamp is repaired from createdAt (not dropped).
  const old = entries.find((entry) => entry.metadata.id === "old");
  assert.equal(old?.metadata.lastModified, "2026-01-01T00:00:00.000Z");
});

void test("metadata rows without a full record are reported as missingData, not dropped", async () => {
  const { store, backend } = createStore();
  backend.seed(
    "sessions-metadata",
    "ghost",
    sessionMetadata({
      id: "ghost",
      lastModified: "2026-04-01T00:00:00.000Z",
      messageCount: 4,
    }),
  );

  const entries = await store.getAllMetadataSafe();
  assert.equal(entries.length, 1);
  assert.equal(entries[0]?.missingData, true);
});

void test("delete removes both the record and its metadata", async () => {
  const { store, backend } = createStore();
  backend.seed("sessions", "a", sessionRecord({ id: "a", lastModified: "2026-01-01T00:00:00.000Z" }));
  backend.seed(
    "sessions-metadata",
    "a",
    sessionMetadata({ id: "a", lastModified: "2026-01-01T00:00:00.000Z" }),
  );

  await store.delete("a");
  assert.equal(backend.size("sessions"), 0);
  assert.equal(backend.size("sessions-metadata"), 0);
});

void test("getLatestSessionId picks the newest record", async () => {
  const { store, backend } = createStore();
  backend.seed("sessions", "a", sessionRecord({ id: "a", lastModified: "2026-01-01T00:00:00.000Z" }));
  backend.seed("sessions", "b", sessionRecord({ id: "b", lastModified: "2026-06-01T00:00:00.000Z" }));

  assert.equal(await store.getLatestSessionId(), "b");
});

void test("save writes both stores through one transaction", async () => {
  const { store, backend } = createStore();
  const data = sessionRecord({ id: "x", lastModified: "2026-07-01T00:00:00.000Z" });
  const metadata = sessionMetadata({ id: "x", lastModified: data.lastModified });

  await store.save(data, metadata);
  assert.ok(backend.raw("sessions", "x"));
  assert.ok(backend.raw("sessions-metadata", "x"));
});
