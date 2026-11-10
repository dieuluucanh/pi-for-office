/**
 * Persisted "recently closed" ring buffer.
 *
 * The stack used to be in-memory only, so `/reopen` and the resume overlay's
 * "Recently closed" section were empty after every taskpane reload — exactly
 * when a user needs them.
 *
 * Persistence is best-effort: a failed write must never block closing a tab.
 */

import type { SettingsStore } from "../storage/local/settings-store.js";
import {
  RecentlyClosedStack,
  type RecentlyClosedItem,
} from "./recently-closed.js";

export const RECENTLY_CLOSED_SETTING_KEY = "ui.recentlyClosed.v1";
const MAX_PERSISTED_ITEMS = 10;

export interface RecentlyClosedController {
  hydrate: () => Promise<void>;
  flush: () => Promise<void>;
  snapshot: () => readonly RecentlyClosedItem[];
  push: (item: RecentlyClosedItem) => void;
  popMostRecent: () => RecentlyClosedItem | null;
  removeById: (id: string) => RecentlyClosedItem | null;
  removeBySessionId: (sessionId: string) => RecentlyClosedItem | null;
  peekMostRecent: () => RecentlyClosedItem | null;
  readonly size: number;
}

function isRecentlyClosedItem(value: DynamicValue): value is RecentlyClosedItem {
  if (typeof value !== "object" || value === null) return false;
  const item = value as Partial<RecentlyClosedItem>;

  return (
    typeof item.id === "string" &&
    item.id.length > 0 &&
    typeof item.sessionId === "string" &&
    item.sessionId.length > 0 &&
    typeof item.title === "string" &&
    typeof item.closedAt === "string" &&
    (typeof item.workbookId === "string" || item.workbookId === null)
  );
}

/** Parse the persisted payload, dropping anything malformed. */
export function parseRecentlyClosedItems(
  value: DynamicValue,
): RecentlyClosedItem[] {
  const raw = Array.isArray(value)
    ? value
    : typeof value === "object" &&
        value !== null &&
        Array.isArray((value as { items?: DynamicValue }).items)
      ? (value as { items: DynamicValue[] }).items
      : [];

  const items: RecentlyClosedItem[] = [];
  for (const entry of raw) {
    if (!isRecentlyClosedItem(entry)) continue;
    items.push({
      id: entry.id,
      sessionId: entry.sessionId,
      title: entry.title,
      closedAt: entry.closedAt,
      workbookId: entry.workbookId,
    });
    if (items.length >= MAX_PERSISTED_ITEMS) break;
  }
  return items;
}

export function serializeRecentlyClosedItems(
  items: readonly RecentlyClosedItem[],
): { items: RecentlyClosedItem[] } {
  return { items: items.slice(0, MAX_PERSISTED_ITEMS) };
}

export function createRecentlyClosedController(
  settings: SettingsStore,
  opts?: { limit?: number; warn?: (message: string, error: DynamicValue) => void },
): RecentlyClosedController {
  const stack = new RecentlyClosedStack(opts?.limit ?? MAX_PERSISTED_ITEMS);
  const warn =
    opts?.warn ??
    ((message: string, error: DynamicValue) => {
      console.warn(message, error);
    });

  let writeChain: Promise<void> = Promise.resolve();

  const persist = (): void => {
    const payload = serializeRecentlyClosedItems(stack.snapshot());
    writeChain = writeChain
      .then(
        () => settings.set(RECENTLY_CLOSED_SETTING_KEY, payload),
        () => settings.set(RECENTLY_CLOSED_SETTING_KEY, payload),
      )
      .catch((error: DynamicValue) => {
        warn("[pi] Failed to persist recently closed items:", error);
      });
  };

  return {
    async hydrate(): Promise<void> {
      try {
        const stored = await settings.get(RECENTLY_CLOSED_SETTING_KEY);
        const items = parseRecentlyClosedItems(stored);
        // Push oldest-first so the newest ends up at the head of the stack.
        for (let index = items.length - 1; index >= 0; index -= 1) {
          const item = items[index];
          if (item) stack.push(item);
        }
      } catch (error) {
        warn("[pi] Failed to hydrate recently closed items:", error);
      }
    },

    flush(): Promise<void> {
      return writeChain;
    },

    snapshot: () => stack.snapshot(),

    push(item: RecentlyClosedItem): void {
      stack.push(item);
      persist();
    },

    popMostRecent(): RecentlyClosedItem | null {
      const removed = stack.popMostRecent();
      if (removed) persist();
      return removed;
    },

    removeById(id: string): RecentlyClosedItem | null {
      const removed = stack.removeById(id);
      if (removed) persist();
      return removed;
    },

    removeBySessionId(sessionId: string): RecentlyClosedItem | null {
      const removed = stack.removeBySessionId(sessionId);
      if (removed) persist();
      return removed;
    },

    peekMostRecent: () => stack.peekMostRecent(),

    get size(): number {
      return stack.size;
    },
  };
}
