import assert from "node:assert/strict";
import { test } from "node:test";

import { decideWorkbookContextRefresh } from "../src/taskpane/context-refresh-decision.ts";

void test("reinjects workbook context when prior injection was compacted away", () => {
  const decision = decideWorkbookContextRefresh({
    lastInjectedWorkbookId: "url_sha256:abc",
    lastInjectedBlueprintRevision: 4,
    currentWorkbookId: "url_sha256:abc",
    currentBlueprintRevision: 4,
    hasWorkbookContextMessage: false,
  });

  assert.equal(decision.refreshReason, "context_missing");
  assert.equal(decision.shouldBootstrap, false);
});

void test("bootstraps state without reinjection when restored history already has workbook context", () => {
  const decision = decideWorkbookContextRefresh({
    lastInjectedWorkbookId: undefined,
    lastInjectedBlueprintRevision: -1,
    currentWorkbookId: "url_sha256:abc",
    currentBlueprintRevision: 0,
    hasWorkbookContextMessage: true,
  });

  assert.equal(decision.refreshReason, null);
  assert.equal(decision.shouldBootstrap, true);
});

void test("prioritizes workbook switch over missing-context refresh", () => {
  const decision = decideWorkbookContextRefresh({
    lastInjectedWorkbookId: "url_sha256:abc",
    lastInjectedBlueprintRevision: 7,
    currentWorkbookId: "url_sha256:def",
    currentBlueprintRevision: 7,
    hasWorkbookContextMessage: false,
  });

  assert.equal(decision.refreshReason, "workbook_switched");
  assert.equal(decision.shouldBootstrap, false);
});
