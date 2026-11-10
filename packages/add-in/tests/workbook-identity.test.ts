import assert from "node:assert/strict";
import { test } from "node:test";

import {
  getWorkbookContextFromDocumentUrl,
} from "../src/host/workbook-context.ts";
import type {
  SpreadsheetHost,
  SpreadsheetHostReadyCallback,
  SpreadsheetHostReadyInfo,
} from "../src/host/types.ts";
import type { WorkbookContext } from "../src/host/workbook-context.ts";
import { settingsBackedSessionStorage } from "../src/host/session-storage.ts";
import {
  readWorkbookIdentityAlias,
  recordWorkbookIdentityAlias,
  resetWorkbookIdentityWriteCache,
  resolveWorkbookIdentity,
} from "../src/workbook/identity.ts";
import { SettingsStore } from "../src/storage/local/settings-store.ts";
import { MemoryStorage } from "./helpers/memory-storage.ts";

class FakeHost implements SpreadsheetHost {
  readonly kind = "browser" as const;
  readonly displayName = "Fake host";
  readonly sessionStorage = settingsBackedSessionStorage;
  private context: WorkbookContext;
  private documentScopedId: string | null;

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

  setContext(context: WorkbookContext): void {
    this.context = context;
  }
}

function createSettings(): SettingsStore {
  const backend = new MemoryStorage({ stores: [{ name: "settings" }] });
  const settings = new SettingsStore();
  settings.setBackend(backend);
  return settings;
}

async function hash(url: string): Promise<string | null> {
  return (await getWorkbookContextFromDocumentUrl(url)).workbookId;
}

void test("a file URL and the equivalent Windows path hash identically", async () => {
  const fromUrl = await hash("file:///C:/Users/alice/Budget.xlsx");
  const fromPath = await hash("C:\\Users\\alice\\Budget.xlsx");
  const lowercaseDrive = await hash("c:\\users\\alice\\Budget.xlsx".replace("users", "Users"));

  assert.ok(fromUrl);
  assert.equal(fromPath, fromUrl);
  assert.equal(lowercaseDrive, fromUrl);
});

void test("percent-encoded and decoded SharePoint paths hash identically", async () => {
  const encoded = await hash(
    "https://tenant.sharepoint.com/sites/team/Shared%20Documents/Budget.xlsx?web=1",
  );
  const decoded = await hash(
    "https://tenant.sharepoint.com/sites/team/Shared Documents/Budget.xlsx",
  );

  assert.ok(encoded);
  assert.equal(decoded, encoded);
});

void test("query/hash and trailing slashes do not change identity", async () => {
  const plain = await hash("https://example.com/reports/Forecast.xlsx");
  const noisy = await hash(
    "https://example.com/reports/Forecast.xlsx/?session=abc#sheet=2",
  );

  assert.equal(noisy, plain);
});

void test("UNC paths normalize stabily", async () => {
  const backslash = await hash("\\\\server\\share\\Budget.xlsx");
  const forward = await hash("//server/share/Budget.xlsx");

  assert.ok(backslash);
  assert.equal(forward, backslash);
});

void test("different workbooks keep different identities", async () => {
  assert.notEqual(
    await hash("https://example.com/reports/A.xlsx"),
    await hash("https://example.com/reports/B.xlsx"),
  );
});

void test("a document-scoped id wins over the observed URL hash and is aliased", async () => {
  resetWorkbookIdentityWriteCache();
  const settings = createSettings();
  const observed = await hash("file:///C:/Work/Plan.xlsx");
  assert.ok(observed);

  const host = new FakeHost(
    { workbookId: observed, workbookName: "Plan.xlsx", source: "document.url" },
    "doc_uuid:1111",
  );

  const identity = await resolveWorkbookIdentity({ host, settings });
  assert.equal(identity.canonicalWorkbookId, "doc_uuid:1111");
  assert.equal(identity.observedWorkbookId, observed);
  assert.equal(identity.aliasedFrom, observed);

  assert.equal(await readWorkbookIdentityAlias(settings, observed), "doc_uuid:1111");
});

void test("an existing alias resolves even when the host cannot read a document id", async () => {
  resetWorkbookIdentityWriteCache();
  const settings = createSettings();
  const observed = "url_sha256:observed";
  await recordWorkbookIdentityAlias(
    { settings },
    { observedId: observed, canonicalId: "doc_uuid:2222", workbookName: "Plan.xlsx" },
  );

  const host = new FakeHost(
    { workbookId: observed, workbookName: "Plan.xlsx", source: "document.url" },
    null,
  );
  const identity = await resolveWorkbookIdentity({ host, settings });

  assert.equal(identity.canonicalWorkbookId, "doc_uuid:2222");
  assert.equal(identity.aliasedFrom, observed);
});

void test("a self-alias is ignored and a second hop is not followed", async () => {
  resetWorkbookIdentityWriteCache();
  const settings = createSettings();

  await settings.set("workbook.identity.v1.alias.url_sha256:a", "url_sha256:b");
  await settings.set("workbook.identity.v1.alias.url_sha256:b", "url_sha256:c");

  assert.equal(await readWorkbookIdentityAlias(settings, "url_sha256:a"), "url_sha256:b");
  assert.equal(await readWorkbookIdentityAlias(settings, "url_sha256:b"), "url_sha256:c");
  assert.equal(await readWorkbookIdentityAlias(settings, "url_sha256:missing"), "url_sha256:missing");
});

void test("dryRun resolves without writing an alias", async () => {
  resetWorkbookIdentityWriteCache();
  const settings = createSettings();
  const host = new FakeHost(
    { workbookId: "url_sha256:observed", workbookName: "Plan.xlsx", source: "document.url" },
    "doc_uuid:3333",
  );

  const identity = await resolveWorkbookIdentity({ host, settings, dryRun: true });
  assert.equal(identity.canonicalWorkbookId, "doc_uuid:3333");
  assert.equal(await readWorkbookIdentityAlias(settings, "url_sha256:observed"), "url_sha256:observed");
});

void test("an unknown document resolves to a null canonical id", async () => {
  resetWorkbookIdentityWriteCache();
  const settings = createSettings();
  const host = new FakeHost({
    workbookId: null,
    workbookName: null,
    source: "unknown",
  });

  const identity = await resolveWorkbookIdentity({ host, settings });
  assert.equal(identity.canonicalWorkbookId, null);
  assert.equal(identity.observedWorkbookId, null);
  assert.equal(identity.aliasedFrom, null);
});
