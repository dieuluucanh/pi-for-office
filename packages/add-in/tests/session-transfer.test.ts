import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildSessionExport,
  importSessions,
  parseSessionTransfer,
  serializeSessionExport,
  SESSION_TRANSFER_SCHEMA,
} from "../src/storage/local/session-transfer.ts";
import { SessionsStore } from "../src/storage/local/sessions-store.ts";
import { MemoryStorage } from "./helpers/memory-storage.ts";
import {
  assistantMessage,
  sessionMetadata,
  sessionRecord,
  userMessage,
} from "./helpers/session-fixtures.ts";

const CONFIG = [
  { name: "sessions", keyPath: "id", indices: [{ name: "lastModified", keyPath: "lastModified" }] },
  {
    name: "sessions-metadata",
    keyPath: "id",
    indices: [{ name: "lastModified", keyPath: "lastModified" }],
  },
];

function createStore(): SessionsStore {
  const backend = new MemoryStorage({ stores: CONFIG });
  const store = new SessionsStore();
  store.setBackend(backend);
  return store;
}

async function seed(store: SessionsStore, id: string, title: string): Promise<void> {
  await store.save(
    sessionRecord({
      id,
      title,
      lastModified: "2026-01-02T00:00:00.000Z",
      messages: [userMessage(`${title} question`), assistantMessage(`${title} answer`)],
    }),
    sessionMetadata({ id, title, lastModified: "2026-01-02T00:00:00.000Z", messageCount: 2 }),
  );
}

void test("export → parse → import restores a session into an empty store", async () => {
  const source = createStore();
  await seed(source, "session-1", "Budget review");

  const exported = await buildSessionExport(source);
  assert.equal(exported.schema, SESSION_TRANSFER_SCHEMA);
  assert.equal(exported.sessions.length, 1);

  const parsed = parseSessionTransfer(serializeSessionExport(exported));
  assert.ok("file" in parsed);

  const target = createStore();
  const result = await importSessions(target, parsed.file);
  assert.equal(result.imported, 1);
  assert.equal(result.skipped, 0);

  const restored = await target.loadSession("session-1");
  assert.equal(restored?.title, "Budget review");
  assert.equal(restored?.messages.length, 2);
});

void test("import never overwrites an existing session id", async () => {
  const source = createStore();
  await seed(source, "shared-id", "Imported title");
  const exported = await buildSessionExport(source);

  const target = createStore();
  await seed(target, "shared-id", "Local title");

  const result = await importSessions(target, exported);
  assert.equal(result.imported, 1);

  const local = await target.loadSession("shared-id");
  assert.equal(local?.title, "Local title");

  const all = await target.getAllMetadataSafe();
  assert.equal(all.length, 2);
  assert.ok(all.some((entry) => entry.metadata.title === "Imported title"));
});

void test("rejects invalid JSON and foreign schemas with a readable error", () => {
  const notJson = parseSessionTransfer("{nope");
  assert.ok("error" in notJson);

  const wrongSchema = parseSessionTransfer(
    JSON.stringify({ schema: "other/v9", sessions: [] }),
  );
  assert.ok("error" in wrongSchema);
});

void test("malformed entries are skipped, valid ones still import", async () => {
  const payload = JSON.stringify({
    schema: SESSION_TRANSFER_SCHEMA,
    exportedAt: "2026-01-01T00:00:00.000Z",
    sessions: [
      { data: { id: "", messages: [] } },
      { data: { id: "no-messages" } },
      {
        data: sessionRecord({ id: "good", title: "Good session", lastModified: "2026-01-02T00:00:00.000Z" }),
        metadata: sessionMetadata({ id: "good", title: "Good session", lastModified: "2026-01-02T00:00:00.000Z" }),
      },
    ],
  });

  const parsed = parseSessionTransfer(payload);
  assert.ok("file" in parsed);
  assert.equal(parsed.file.sessions.length, 1);

  const target = createStore();
  const result = await importSessions(target, parsed.file);
  assert.equal(result.imported, 1);
});
