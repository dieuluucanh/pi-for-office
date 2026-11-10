import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels } from "@earendil-works/pi-ai";

import { SessionsStore } from "../src/storage/local/sessions-store.ts";
import { SettingsStore } from "../src/storage/local/settings-store.ts";
import { MemoryStorage } from "./helpers/memory-storage.ts";
import { failOnUnexpectedStream } from "./fail-on-unexpected-stream.ts";
import { setupSessionPersistence } from "../src/taskpane/sessions.ts";
import { resolveSessionSaveTrigger } from "../src/taskpane/sessions.ts";
import {
  getLatestSessionForWorkbook,
  getSessionWorkbookId,
} from "../src/workbook/session-association.ts";
import {
  setCurrentSpreadsheetHost,
  resetCurrentSpreadsheetHostForTests,
} from "../src/host/current.ts";
import type {
  SpreadsheetHost,
  SpreadsheetHostReadyCallback,
  SpreadsheetHostReadyInfo,
} from "../src/host/types.ts";
import type { WorkbookContext } from "../src/host/workbook-context.ts";
import { settingsBackedSessionStorage } from "../src/host/session-storage.ts";
import type { SessionData } from "../src/storage/local/types.ts";
import {
  sessionRecord,
  TEST_MODEL,
  userMessage,
} from "./helpers/session-fixtures.ts";

const CONFIG = [
  { name: "settings" },
  { name: "sessions", keyPath: "id", indices: [{ name: "lastModified", keyPath: "lastModified" }] },
  {
    name: "sessions-metadata",
    keyPath: "id",
    indices: [{ name: "lastModified", keyPath: "lastModified" }],
  },
];

class MutableHost implements SpreadsheetHost {
  readonly kind = "browser" as const;
  readonly displayName = "Mutable test host";
  readonly sessionStorage = settingsBackedSessionStorage;
  context: WorkbookContext;
  documentScopedId: string | null;

  constructor(context: WorkbookContext, documentScopedId: string | null = null) {
    this.context = context;
    this.documentScopedId = documentScopedId;
  }

  whenReady(): Promise<SpreadsheetHostReadyInfo> {
    return Promise.resolve({
      kind: "browser",
      nativeHost: null,
      nativePlatform: null,
      reason: "browser",
    });
  }

  onReady(_callback: SpreadsheetHostReadyCallback): () => void {
    return () => undefined;
  }

  getWorkbookContext(): Promise<WorkbookContext> {
    return Promise.resolve(this.context);
  }

  getDocumentScopedId(): Promise<string | null> {
    return Promise.resolve(this.documentScopedId);
  }

  setDocumentScopedId(id: string): Promise<boolean> {
    this.documentScopedId = id;
    return Promise.resolve(true);
  }

  resolveThemeDark(): boolean | null {
    return null;
  }
}

interface Harness {
  backend: MemoryStorage;
  sessions: SessionsStore;
  settings: SettingsStore;
  agent: Agent;
  host: MutableHost;
  persistence: Awaited<ReturnType<typeof setupSessionPersistence>>;
  dispose: () => void;
}

async function createHarness(
  options: { workbookId?: string | null; documentScopedId?: string | null } = {},
): Promise<Harness> {
  const backend = new MemoryStorage({ stores: CONFIG });
  const settings = new SettingsStore();
  settings.setBackend(backend);
  const sessions = new SessionsStore();
  sessions.setBackend(backend);

  const host = new MutableHost(
    {
      workbookId: options.workbookId === undefined ? "url_sha256:wb-1" : options.workbookId,
      workbookName: "Plan.xlsx",
      source: "document.url",
    },
    options.documentScopedId ?? null,
  );
  setCurrentSpreadsheetHost(host);

  const agent = new Agent({ streamFn: failOnUnexpectedStream });
  const persistence = await setupSessionPersistence({
    agent,
    sessions,
    settings,
    models: createModels(),
  });

  return {
    backend,
    sessions,
    settings,
    agent,
    host,
    persistence,
    dispose: () => {
      persistence.dispose();
      resetCurrentSpreadsheetHostForTests();
    },
  };
}

function userMessageForTest(text: string): SessionData["messages"][number] {
  return userMessage(text);
}

void test("save trigger policy covers draft, agent, boundary and idle events", () => {
  const base = { firstAssistantSeen: false, recordExists: false };

  assert.equal(
    resolveSessionSaveTrigger({ ...base, eventType: "message_end", messageRole: "user" }),
    "draft",
  );
  assert.equal(
    resolveSessionSaveTrigger({
      ...base,
      recordExists: true,
      eventType: "message_end",
      messageRole: "user",
    }),
    "none",
  );
  assert.equal(
    resolveSessionSaveTrigger({ ...base, eventType: "message_end", messageRole: "assistant" }),
    "agent",
  );
  assert.equal(
    resolveSessionSaveTrigger({ ...base, eventType: "message_end", messageRole: "toolResult" }),
    "none",
  );
  assert.equal(
    resolveSessionSaveTrigger({
      ...base,
      firstAssistantSeen: true,
      eventType: "message_end",
      messageRole: "toolResult",
    }),
    "agent",
  );
  assert.equal(resolveSessionSaveTrigger({ ...base, eventType: "turn_end" }), "boundary");
  assert.equal(
    resolveSessionSaveTrigger({ ...base, eventType: "agent_end" }),
    "boundary_if_dirty",
  );
  assert.equal(resolveSessionSaveTrigger({ ...base, eventType: "message_start" }), "none");
});

void test("a forced blank save stores a tab placeholder but never claims 'latest'", async () => {
  const harness = await createHarness();
  try {
    await harness.persistence.saveSession({ force: true });

    const id = harness.persistence.getSessionId();
    assert.ok(await harness.sessions.loadSession(id));
    assert.equal(await getLatestSessionForWorkbook(harness.settings, "url_sha256:wb-1"), null);
  } finally {
    harness.dispose();
  }
});

void test("a forced blank save never overwrites an existing record", async () => {
  const harness = await createHarness();
  try {
    harness.agent.state.messages = [userMessage("keep me")];
    await harness.persistence.saveSession({ force: true });
    const id = harness.persistence.getSessionId();

    harness.agent.state.messages = [];
    await harness.persistence.saveSession({ force: true });

    const stored = await harness.sessions.loadSession(id);
    assert.equal(stored?.messages.length, 1);
  } finally {
    harness.dispose();
  }
});

void test("a user-only draft save is durable and links the session", async () => {
  const harness = await createHarness();
  try {
    harness.agent.state.messages = [userMessage("draft prompt")];
    await harness.persistence.saveSession({ force: true });

    const id = harness.persistence.getSessionId();
    const stored = await harness.sessions.loadSession(id);
    assert.equal(stored?.messages.length, 1);
    assert.equal(await getSessionWorkbookId(harness.settings, id), "url_sha256:wb-1");
    assert.equal(
      await getLatestSessionForWorkbook(harness.settings, "url_sha256:wb-1"),
      id,
    );
  } finally {
    harness.dispose();
  }
});

void test("loading a session keeps its stored id (no id remapping)", async () => {
  const harness = await createHarness();
  try {
    const data: SessionData = sessionRecord({
      id: "legacy-session-id",
      title: "Legacy",
      lastModified: "2026-01-02T00:00:00.000Z",
      messages: [userMessageForTest("hello")],
    });
    assert.equal(data.model.id, TEST_MODEL.id);

    await harness.persistence.applyLoadedSession(data);
    assert.equal(harness.persistence.getSessionId(), "legacy-session-id");
  } finally {
    harness.dispose();
  }
});

void test("save failures are surfaced instead of swallowed", async () => {
  const harness = await createHarness();
  try {
    harness.agent.state.messages = [userMessage("will fail")];
    harness.backend.failWrites = true;

    await harness.persistence.saveSession({ force: true });
    assert.ok(harness.persistence.getLastSaveError());

    harness.backend.failWrites = false;
    await harness.persistence.flush();
    assert.equal(harness.persistence.getLastSaveError(), null);
  } finally {
    harness.dispose();
  }
});

void test("serialized saves keep the newest title", async () => {
  const harness = await createHarness();
  try {
    harness.agent.state.messages = [userMessage("hello")];

    const first = harness.persistence.renameSession("First");
    const second = harness.persistence.renameSession("Second");
    await Promise.all([first, second]);
    await harness.persistence.flush();

    const id = harness.persistence.getSessionId();
    const stored = await harness.sessions.loadSession(id);
    assert.equal(stored?.title, "Second");
  } finally {
    harness.dispose();
  }
});

void test("sessions saved without an identity are linked once it appears", async () => {
  const harness = await createHarness({ workbookId: null });
  try {
    harness.agent.state.messages = [userMessage("before identity")];
    await harness.persistence.saveSession({ force: true });

    const id = harness.persistence.getSessionId();
    assert.equal(await getSessionWorkbookId(harness.settings, id), null);

    // The document is saved / Office becomes ready: identity now resolves.
    harness.host.context = {
      workbookId: "url_sha256:wb-2",
      workbookName: "Plan.xlsx",
      source: "document.url",
    };

    await harness.persistence.saveSession({ force: true });

    assert.equal(await getSessionWorkbookId(harness.settings, id), "url_sha256:wb-2");
    assert.equal(
      await getLatestSessionForWorkbook(harness.settings, "url_sha256:wb-2"),
      id,
    );
  } finally {
    harness.dispose();
  }
});

void test("the canonical document id is used for association when available", async () => {
  const harness = await createHarness({
    workbookId: "url_sha256:observed",
    documentScopedId: "doc_uuid:stable",
  });
  try {
    harness.agent.state.messages = [userMessage("canonical")];
    await harness.persistence.saveSession({ force: true });

    const id = harness.persistence.getSessionId();
    assert.equal(await getSessionWorkbookId(harness.settings, id), "doc_uuid:stable");
    assert.equal(
      await getLatestSessionForWorkbook(harness.settings, "doc_uuid:stable"),
      id,
    );
  } finally {
    harness.dispose();
  }
});
