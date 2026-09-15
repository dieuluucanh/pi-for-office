/**
 * Protocol v2 document identity mapping — host-neutral `WorkbookContext` →
 * `DocumentIdentity`. Covers Office URLs, WPS paths, unsaved and unknown docs.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { toDocumentIdentity } from "../src/bridge/pi-bridge-manager.ts";

void test("maps an Office document-url identity", () => {
  assert.deepEqual(
    toDocumentIdentity({
      workbookId: "url_sha256:abc123",
      workbookName: "MBI_2026.xlsx",
      source: "document.url",
    }),
    {
      id: "url_sha256:abc123",
      name: "MBI_2026.xlsx",
      source: "document.url",
    },
  );
});

void test("maps a WPS active-workbook path identity", () => {
  assert.deepEqual(
    toDocumentIdentity({
      workbookId: "wps_path_sha256:def456",
      workbookName: "Budget.et",
      source: "wps.activeWorkbook.fullName",
    }),
    {
      id: "wps_path_sha256:def456",
      name: "Budget.et",
      source: "wps.activeWorkbook.fullName",
    },
  );
});

void test("maps an unsaved/unknown document to a null id", () => {
  assert.deepEqual(
    toDocumentIdentity({
      workbookId: null,
      workbookName: null,
      source: "unknown",
    }),
    { id: null, name: null, source: "unknown" },
  );
});

void test("keeps a best-effort name when the id is unavailable", () => {
  assert.deepEqual(
    toDocumentIdentity({
      workbookId: null,
      workbookName: "Untitled 1.xlsx",
      source: "unknown",
    }),
    { id: null, name: "Untitled 1.xlsx", source: "unknown" },
  );
});
