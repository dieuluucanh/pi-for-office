import assert from "node:assert/strict";
import { test } from "node:test";

import { SessionsStore } from "../src/storage/local/sessions-store.ts";
import { SettingsStore } from "../src/storage/local/settings-store.ts";
import { MemoryStorage } from "./helpers/memory-storage.ts";
import {
  repairSessionStorage,
  SESSION_REPAIR_VERSION,
  RECENTLY_CLOSED_KEY,
  TAB_LAYOUT_KEY_PREFIX,
} from "../src/storage/local/repair.ts";
import { STORAGE_REPAIR_MARKER_KEY } from "../src/storage/local/diagnostics.ts";
import { RECENTLY_CLOSED_SETTING_KEY } from "../src/taskpane/recently-closed-store.ts";
import { workbookTabLayoutKey } from "../src/taskpane/tab-layout.ts";
import { sessionMetadata, sessionRecord } from "./helpers/session-fixtures.ts";

const CONFIG = [
  { name: "settings" },
  { name: "sessions", keyPath: "id", indices: [{ name: "lastModified", keyPath: "lastModified" }] },
  {
    name: "sessions-metadata",
    keyPath: "id",
    indices: [{ name: "lastModified", keyPath: "lastModified" }],
  },
];

function createDeps(): {
  backend: MemoryStorage;
  sessions: SessionsStore;
  settings: SettingsStore;
} {
  const backend = new MemoryStorage({ stores: CONFIG });
  const settings = new SettingsStore();
  settings.setBackend(backend);
  const sessions = new SessionsStore();
  sessions.setBackend(backend);
  return { backend, sessions, settings };
}

function record(id: string, lastModified: string, messageCount = 1) {
  return sessionRecord({
    id,
    lastModified,
    messages: messageCount === 0 ? [] : undefined,
  });
}

function metadata(id: string, lastModified: string, messageCount = 1) {
  return sessionMetadata({ id, lastModified, messageCount });
}

void test("repair rebuilds missing metadata rows", async () => {
  const deps = createDeps();
  deps.backend.seed("sessions", "a", record("a", "2026-01-01T00:00:00.000Z"));

  const report = await repairSessionStorage(deps);
  assert.equal(report?.metadataRebuilt, 1);
  assert.ok(await deps.sessions.getMetadata("a"));

  const entries = await deps.sessions.getAllMetadataSafe();
  assert.equal(entries.length, 1);
});

void test("repair backfills unusable lastModified so the row is listable", async () => {
  const deps = createDeps();
  deps.backend.seed("sessions", "a", record("a", "2026-01-01T00:00:00.000Z"));
  deps.backend.seed("sessions-metadata", "a", metadata("a", "not-a-date"));

  const report = await repairSessionStorage(deps);
  assert.equal(report?.timestampsRepaired, 1);

  const repaired = await deps.sessions.getMetadata("a");
  assert.equal(repaired?.lastModified, "2026-01-01T00:00:00.000Z");
});

void test("repair removes only unreferenced blank records", async () => {
  const deps = createDeps();

  // Blank + unreferenced → removable.
  deps.backend.seed("sessions", "blank-free", record("blank-free", "2026-01-01T00:00:00.000Z", 0));
  deps.backend.seed("sessions-metadata", "blank-free", metadata("blank-free", "2026-01-01T00:00:00.000Z", 0));

  // Blank but referenced by the tab layout → kept.
  deps.backend.seed("sessions", "blank-in-tab", record("blank-in-tab", "2026-01-02T00:00:00.000Z", 0));
  deps.backend.seed("sessions-metadata", "blank-in-tab", metadata("blank-in-tab", "2026-01-02T00:00:00.000Z", 0));
  await deps.settings.set(workbookTabLayoutKey("wb-1"), {
    sessionIds: ["blank-in-tab"],
    activeSessionId: "blank-in-tab",
  });

  // Non-empty unreferenced → never removed.
  deps.backend.seed("sessions", "real", record("real", "2026-01-03T00:00:00.000Z", 2));
  deps.backend.seed("sessions-metadata", "real", metadata("real", "2026-01-03T00:00:00.000Z", 2));

  const report = await repairSessionStorage(deps);
  assert.equal(report?.blankDuplicatesRemoved, 1);

  assert.equal(await deps.sessions.loadSession("blank-free"), null);
  assert.ok(await deps.sessions.loadSession("blank-in-tab"));
  assert.ok(await deps.sessions.loadSession("real"));
});

void test("repair keeps a blank record referenced by the latest-session pointer", async () => {
  const deps = createDeps();
  deps.backend.seed("sessions", "blank-latest", record("blank-latest", "2026-01-01T00:00:00.000Z", 0));
  deps.backend.seed("sessions-metadata", "blank-latest", metadata("blank-latest", "2026-01-01T00:00:00.000Z", 0));
  await deps.settings.set("workbook.latestSession.v1.wb-1", "blank-latest");

  const report = await repairSessionStorage(deps);
  assert.equal(report?.blankDuplicatesRemoved, 0);
  assert.ok(await deps.sessions.loadSession("blank-latest"));
});

void test("dryRun reports without writing anything or setting a marker", async () => {
  const deps = createDeps();
  deps.backend.seed("sessions", "a", record("a", "2026-01-01T00:00:00.000Z"));
  deps.backend.seed("sessions", "blank", record("blank", "2026-01-01T00:00:00.000Z", 0));
  deps.backend.seed("sessions-metadata", "blank", metadata("blank", "2026-01-01T00:00:00.000Z", 0));

  const report = await repairSessionStorage({ ...deps, dryRun: true });
  assert.equal(report?.dryRun, true);
  assert.equal(report?.metadataRebuilt, 1);
  assert.equal(report?.blankDuplicatesRemoved, 1);

  assert.equal(await deps.sessions.getMetadata("a"), null);
  assert.ok(await deps.sessions.loadSession("blank"));
  assert.equal(await deps.settings.get(STORAGE_REPAIR_MARKER_KEY), null);
});

void test("the marker prevents a second run", async () => {
  const deps = createDeps();
  deps.backend.seed("sessions", "a", record("a", "2026-01-01T00:00:00.000Z"));

  const first = await repairSessionStorage(deps);
  assert.ok(first);

  const second = await repairSessionStorage(deps);
  assert.equal(second, null);

  const marker = await deps.settings.get(STORAGE_REPAIR_MARKER_KEY);
  assert.equal(
    (marker as { version?: number } | null)?.version,
    SESSION_REPAIR_VERSION,
  );
});

void test("repair literals still match the modules that own those keys", () => {
  assert.equal(TAB_LAYOUT_KEY_PREFIX, "workbook.tabLayout.v1.");
  assert.equal(workbookTabLayoutKey("x").startsWith(TAB_LAYOUT_KEY_PREFIX), true);
  assert.equal(RECENTLY_CLOSED_KEY, RECENTLY_CLOSED_SETTING_KEY);
});
