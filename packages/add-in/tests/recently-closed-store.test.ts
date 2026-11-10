import assert from "node:assert/strict";
import { test } from "node:test";

import { SettingsStore } from "../src/storage/local/settings-store.ts";
import { MemoryStorage } from "./helpers/memory-storage.ts";
import {
  createRecentlyClosedController,
  parseRecentlyClosedItems,
  RECENTLY_CLOSED_SETTING_KEY,
  serializeRecentlyClosedItems,
} from "../src/taskpane/recently-closed-store.ts";
import type { RecentlyClosedItem } from "../src/taskpane/recently-closed.ts";

function item(index: number, workbookId: string | null = "wb-1"): RecentlyClosedItem {
  return {
    id: `closed-${index}`,
    sessionId: `session-${index}`,
    title: `Session ${index}`,
    closedAt: new Date(2026, 0, index + 1).toISOString(),
    workbookId,
  };
}

function createSettings(): SettingsStore {
  const backend = new MemoryStorage({ stores: [{ name: "settings" }] });
  const settings = new SettingsStore();
  settings.setBackend(backend);
  return settings;
}

void test("hydrate restores newest-first order", async () => {
  const settings = createSettings();
  await settings.set(
    RECENTLY_CLOSED_SETTING_KEY,
    serializeRecentlyClosedItems([item(2), item(1), item(0)]),
  );

  const controller = createRecentlyClosedController(settings);
  await controller.hydrate();

  assert.deepEqual(
    controller.snapshot().map((entry) => entry.id),
    ["closed-2", "closed-1", "closed-0"],
  );
  assert.equal(controller.peekMostRecent()?.id, "closed-2");
});

void test("malformed persisted entries are dropped", () => {
  const parsed = parseRecentlyClosedItems({
    items: [
      item(1),
      { id: "bad" },
      null,
      { ...item(2), sessionId: "" },
    ],
  });

  assert.deepEqual(
    parsed.map((entry) => entry.id),
    ["closed-1"],
  );
});

void test("push persists and caps the ring buffer", async () => {
  const settings = createSettings();
  const controller = createRecentlyClosedController(settings, { limit: 3 });

  for (let index = 0; index < 5; index += 1) {
    controller.push(item(index));
  }
  await controller.flush();

  assert.equal(controller.size, 3);
  assert.deepEqual(
    controller.snapshot().map((entry) => entry.id),
    ["closed-4", "closed-3", "closed-2"],
  );

  const stored = await settings.get(RECENTLY_CLOSED_SETTING_KEY);
  assert.equal(parseRecentlyClosedItems(stored).length, 3);
});

void test("removing an entry persists the change", async () => {
  const settings = createSettings();
  const controller = createRecentlyClosedController(settings);

  controller.push(item(0));
  controller.push(item(1));
  await controller.flush();

  assert.equal(controller.removeById("closed-0")?.id, "closed-0");
  await controller.flush();

  const stored = parseRecentlyClosedItems(
    await settings.get(RECENTLY_CLOSED_SETTING_KEY),
  );
  assert.deepEqual(
    stored.map((entry) => entry.id),
    ["closed-1"],
  );
});

void test("a failed write never throws out of push", async () => {
  const settings = createSettings();
  const backend = new MemoryStorage({ stores: [{ name: "settings" }] });
  const failing = new MemoryStorage({ stores: [{ name: "settings" }] });
  failing.failWrites = true;
  settings.setBackend(failing);

  const warnings: string[] = [];
  const controller = createRecentlyClosedController(settings, {
    warn: (message) => warnings.push(message),
  });

  controller.push(item(0));
  await controller.flush();

  assert.equal(controller.size, 1);
  assert.equal(backend.size("settings"), 0);
  assert.ok(warnings.length >= 1);
});
