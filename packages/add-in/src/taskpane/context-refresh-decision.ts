export type BlueprintRefreshReason =
  | "initial"
  | "workbook_switched"
  | "blueprint_invalidated"
  | "context_missing";

export interface WorkbookContextRefreshDecisionInput {
  lastInjectedWorkbookId: string | null | undefined;
  lastInjectedBlueprintRevision: number;
  currentWorkbookId: string | null;
  currentBlueprintRevision: number;
  hasWorkbookContextMessage: boolean;
}

export interface WorkbookContextRefreshDecision {
  refreshReason: BlueprintRefreshReason | null;
  shouldBootstrap: boolean;
}

export function decideWorkbookContextRefresh(
  input: WorkbookContextRefreshDecisionInput,
): WorkbookContextRefreshDecision {
  const {
    lastInjectedWorkbookId,
    lastInjectedBlueprintRevision,
    currentWorkbookId,
    currentBlueprintRevision,
    hasWorkbookContextMessage,
  } = input;

  if (lastInjectedWorkbookId !== undefined && lastInjectedWorkbookId !== currentWorkbookId) {
    return {
      refreshReason: "workbook_switched",
      shouldBootstrap: false,
    };
  }

  if (
    lastInjectedWorkbookId !== undefined &&
    lastInjectedBlueprintRevision !== currentBlueprintRevision
  ) {
    return {
      refreshReason: "blueprint_invalidated",
      shouldBootstrap: false,
    };
  }

  if (!hasWorkbookContextMessage) {
    return {
      refreshReason: lastInjectedWorkbookId === undefined ? "initial" : "context_missing",
      shouldBootstrap: false,
    };
  }

  if (lastInjectedWorkbookId === undefined) {
    return {
      refreshReason: null,
      shouldBootstrap: true,
    };
  }

  return {
    refreshReason: null,
    shouldBootstrap: false,
  };
}
