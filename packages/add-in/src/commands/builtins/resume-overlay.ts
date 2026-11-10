/**
 * Resume / session-history overlay.
 *
 * Design invariants (docs/session-persistence.md):
 * - **Nothing is hidden.** Sessions are grouped into this workbook / other
 *   workbooks / not linked, with badges. The old behaviour filtered foreign
 *   sessions out, so a workbook-identity change made recent sessions look lost.
 * - **Failures are visible.** A storage error renders an inline error with a
 *   retry action instead of rejecting the (fire-and-forget) open call.
 * - **Strict resume policy.** Opening a session linked to another workbook asks
 *   for confirmation; nothing is resumed automatically from here.
 * - **Repair without guessing.** "Move to this workbook" re-links a session
 *   explicitly; delete is confirm-gated; export/import is the escape hatch for
 *   origin changes and evicted storage.
 */

import type { SessionData } from "../../storage/local/types.js";
import { getAppStorage } from "../../storage/local/app-storage.js";
import type { SessionListEntry } from "../../storage/local/sessions-store.js";

import {
  getCrossWorkbookResumeConfirmMessage,
  getResumeTargetLabel,
  type ResumeDialogTarget,
} from "./resume-target.js";
import { formatRelativeDate } from "./overlay-relative-date.js";
import {
  filterSessionsByQuery,
  groupSessionsForWorkbook,
  shouldExpandOtherWorkbooks,
  type ResumeGroup,
  type ResumeGroupId,
} from "./resume-grouping.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { requestConfirmationDialog } from "../../ui/confirm-dialog.js";
import { RESUME_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";
import {
  forgetSessionWorkbook,
  readSessionWorkbookLinks,
  relinkSessionToWorkbook,
  setLatestSessionForWorkbook,
} from "../../workbook/session-association.js";
import { resolveWorkbookIdentity, type WorkbookIdentity } from "../../workbook/identity.js";
import { getCurrentSpreadsheetHost } from "../../host/index.js";
import {
  APP_DATABASE_NAME,
  APP_DATABASE_VERSION,
} from "../../storage/init-app-storage.js";
import { collectSessionStorageDiagnostics } from "../../storage/local/diagnostics.js";
import {
  buildSessionExport,
  importSessions,
  parseSessionTransfer,
  serializeSessionExport,
} from "../../storage/local/session-transfer.js";

const RESUME_ITEM_KIND_SESSION = "session";
const RESUME_ITEM_KIND_RECENTLY_CLOSED = "recently_closed";

export interface ResumeRecentlyClosedItem {
  id: string;
  sessionId: string;
  title: string;
  closedAt: string;
  workbookId: string | null;
}

export interface ResumeDialogOptions {
  defaultTarget?: ResumeDialogTarget;
  onOpenInNewTab: (sessionData: SessionData) => Promise<void>;
  onReplaceCurrent: (sessionData: SessionData) => Promise<void>;
  getRecentlyClosedItems?: () => readonly ResumeRecentlyClosedItem[];
  onReopenRecentlyClosed?: (item: ResumeRecentlyClosedItem) => Promise<boolean>;
  /** Notified after a session is deleted, so tab/recently-closed state can follow. */
  onSessionDeleted?: (sessionId: string) => void;
}

interface ResumeDialogData {
  entries: SessionListEntry[];
  links: Map<string, string>;
  identity: WorkbookIdentity;
  recentlyClosed: ResumeRecentlyClosedItem[];
}

function buildBadge(label: string, tone: "this" | "other" | "unlinked"): HTMLElement {
  const badge = document.createElement("span");
  badge.className = `pi-resume-badge pi-resume-badge--${tone}`;
  badge.textContent = label;
  return badge;
}

/** Download text without depending on the files workspace (works in WebView2). */
function downloadText(filename: string, text: string): void {
  const blob = new Blob([text], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.rel = "noopener";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function buildIconAction(label: string, iconText: string): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "pi-resume-item-action";
  button.title = label;
  button.setAttribute("aria-label", label);
  button.textContent = iconText;
  return button;
}

export async function showResumeDialog(opts: ResumeDialogOptions): Promise<void> {
  const storage = getAppStorage();

  const getRecentlyClosedItems = (): ResumeRecentlyClosedItem[] => {
    if (!opts.getRecentlyClosedItems) {
      return [];
    }
    return [...opts.getRecentlyClosedItems()];
  };

  // Toggle: a second invocation closes the open overlay.
  if (closeOverlayById(RESUME_OVERLAY_ID)) {
    return;
  }

  let data: ResumeDialogData;
  try {
    data = await loadData(getRecentlyClosedItems);
  } catch (error) {
    console.warn("[pi] Session history could not load:", error);
    showLoadErrorOverlay(error, opts);
    return;
  }

  if (data.entries.length === 0 && data.recentlyClosed.length === 0) {
    showToast(t("resume.noPreviousSessions"));
    return;
  }

  const workbookId = data.identity.canonicalWorkbookId;

  const dialog = createOverlayDialog({
    overlayId: RESUME_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--m pi-resume-dialog",
  });

  const closeOverlay = dialog.close;

  const { header } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: t("resume.closeLabel"),
    title: t("resume.title"),
    subtitle: t("resume.subtitle"),
  });

  // ── Target controls (open in new tab vs replace current) ──
  const targetControls = document.createElement("div");
  targetControls.className = "pi-resume-target-controls";

  const openInNewTabButton = document.createElement("button");
  openInNewTabButton.type = "button";
  openInNewTabButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-resume-target-btn";
  openInNewTabButton.textContent = t("resume.openInNewTab");

  const replaceCurrentButton = document.createElement("button");
  replaceCurrentButton.type = "button";
  replaceCurrentButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-resume-target-btn";
  replaceCurrentButton.textContent = t("resume.replaceCurrent");

  const targetHint = document.createElement("div");
  targetHint.className = "pi-resume-target-hint";

  let selectedTarget: ResumeDialogTarget = opts.defaultTarget ?? "new_tab";

  const syncTargetButtons = (): void => {
    const isNewTab = selectedTarget === "new_tab";
    openInNewTabButton.classList.toggle("is-active", isNewTab);
    replaceCurrentButton.classList.toggle("is-active", !isNewTab);
    openInNewTabButton.setAttribute("aria-pressed", String(isNewTab));
    replaceCurrentButton.setAttribute("aria-pressed", String(!isNewTab));
    targetHint.textContent = t("resume.defaultActionHint", {
      action: getResumeTargetLabel(selectedTarget),
    });
  };

  openInNewTabButton.addEventListener("click", () => {
    selectedTarget = "new_tab";
    syncTargetButtons();
  });
  replaceCurrentButton.addEventListener("click", () => {
    selectedTarget = "replace_current";
    syncTargetButtons();
  });

  targetControls.append(openInNewTabButton, replaceCurrentButton);
  targetHint.textContent = "";
  syncTargetButtons();

  // ── Search ──
  const searchRow = document.createElement("div");
  searchRow.className = "pi-resume-search";

  const searchInput = document.createElement("input");
  searchInput.type = "search";
  searchInput.className = "pi-resume-search__input";
  searchInput.placeholder = t("resume.searchPlaceholder");
  searchInput.setAttribute("aria-label", t("resume.searchAria"));
  searchRow.appendChild(searchInput);

  // ── Body ──
  const body = document.createElement("div");
  body.className = "pi-overlay-body pi-resume-body";

  const recentSection = document.createElement("section");
  recentSection.className = "pi-overlay-section pi-resume-section";
  recentSection.dataset.resumeSection = "recently-closed";

  const recentTitle = document.createElement("h3");
  recentTitle.className = "pi-overlay-section-title";
  recentTitle.textContent = t("resume.recentlyClosed");

  const recentHint = document.createElement("p");
  recentHint.className = "pi-overlay-hint pi-resume-section-hint";
  recentHint.textContent = t("resume.recentlyClosedHint");

  const recentList = document.createElement("div");
  recentList.className = "pi-resume-list pi-resume-list--recent";
  recentSection.append(recentTitle, recentHint, recentList);

  const groupHost = document.createElement("div");
  groupHost.className = "pi-resume-groups";
  body.append(recentSection, groupHost);

  // ── Footer: counts, diagnostics, export/import ──
  const footer = document.createElement("div");
  footer.className = "pi-resume-footer";

  const footerStatus = document.createElement("span");
  footerStatus.className = "pi-resume-footer__status";

  const diagnosticsToggle = document.createElement("button");
  diagnosticsToggle.type = "button";
  diagnosticsToggle.className = "pi-resume-footer__btn";
  diagnosticsToggle.textContent = t("resume.footer.storageDetails");

  const exportButton = document.createElement("button");
  exportButton.type = "button";
  exportButton.className = "pi-resume-footer__btn";
  exportButton.textContent = t("resume.footer.export");

  const importButton = document.createElement("button");
  importButton.type = "button";
  importButton.className = "pi-resume-footer__btn";
  importButton.textContent = t("resume.footer.import");

  const importInput = document.createElement("input");
  importInput.type = "file";
  importInput.accept = "application/json,.json";
  importInput.hidden = true;

  const diagnosticsPanel = document.createElement("pre");
  diagnosticsPanel.className = "pi-resume-diagnostics";
  diagnosticsPanel.hidden = true;

  footer.append(
    footerStatus,
    diagnosticsToggle,
    exportButton,
    importButton,
    importInput,
  );

  dialog.card.append(
    header,
    targetControls,
    targetHint,
    searchRow,
    body,
    footer,
    diagnosticsPanel,
  );

  // ── State ──
  let query = "";
  let manuallyExpanded: boolean | null = null;
  let recentlyClosedById = new Map<string, ResumeRecentlyClosedItem>();

  const groupTitle = (id: ResumeGroupId): string => {
    switch (id) {
      case "this_workbook":
        return t("resume.group.thisWorkbook");
      case "other_workbooks":
        return t("resume.group.otherWorkbooks");
      default:
        return t("resume.group.unlinked");
    }
  };

  const buildRecentlyClosedRow = (
    item: ResumeRecentlyClosedItem,
  ): HTMLButtonElement => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "pi-welcome-provider pi-resume-item pi-resume-item--recent";
    button.dataset.id = item.sessionId;
    button.dataset.recentId = item.id;
    button.dataset.resumeKind = RESUME_ITEM_KIND_RECENTLY_CLOSED;

    const title = document.createElement("span");
    title.className = "pi-resume-item__title";
    title.textContent = item.title || t("resume.defaultTitle");

    const meta = document.createElement("span");
    meta.className = "pi-resume-item__meta";
    meta.textContent = t("resume.recentlyClosedMeta", {
      date: formatRelativeDate(item.closedAt),
    });

    button.append(title, meta);
    return button;
  };

  const buildSessionRow = (entry: SessionListEntry): HTMLElement => {
    const session = entry.metadata;
    const linkedWorkbookId = data.links.get(session.id) ?? null;

    const row = document.createElement("div");
    row.className = "pi-resume-item-row";
    row.dataset.id = session.id;

    const main = document.createElement("button");
    main.type = "button";
    main.className = "pi-welcome-provider pi-resume-item";
    main.dataset.id = session.id;
    main.dataset.resumeKind = RESUME_ITEM_KIND_SESSION;

    const title = document.createElement("span");
    title.className = "pi-resume-item__title";
    title.textContent = session.title || t("resume.defaultTitle");

    const meta = document.createElement("span");
    meta.className = "pi-resume-item__meta";
    const metaParts = [
      t("resume.sessionMeta", {
        count: session.messageCount || 0,
        date: formatRelativeDate(session.lastModified),
      }),
    ];
    if (entry.missingData) metaParts.push(t("resume.missingData"));
    meta.textContent = metaParts.join(" · ");

    const badges = document.createElement("span");
    badges.className = "pi-resume-item__badges";
    if (!linkedWorkbookId) {
      badges.appendChild(buildBadge(t("resume.badge.unlinked"), "unlinked"));
    } else if (workbookId && linkedWorkbookId === workbookId) {
      badges.appendChild(buildBadge(t("resume.badge.thisWorkbook"), "this"));
    } else {
      badges.appendChild(buildBadge(t("resume.badge.otherWorkbook"), "other"));
    }

    main.append(title, meta, badges);
    row.appendChild(main);

    const actions = document.createElement("div");
    actions.className = "pi-resume-item-actions";

    const copyButton = buildIconAction(t("resume.action.copyId"), "⧉");
    copyButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      void navigator.clipboard
        ?.writeText(session.id)
        .then(() => {
          showToast(t("resume.toast.copiedId"));
        })
        .catch(() => {
          showToast(t("resume.toast.copyFailed"));
        });
    });
    actions.appendChild(copyButton);

    if (workbookId && linkedWorkbookId !== workbookId) {
      const moveButton = buildIconAction(t("resume.action.moveHere"), "⤵");
      moveButton.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        void moveSessionToCurrentWorkbook(session.id);
      });
      actions.appendChild(moveButton);
    }

    const deleteButton = buildIconAction(t("resume.action.delete"), "✕");
    deleteButton.classList.add("pi-resume-item-action--danger");
    deleteButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      void deleteSession(session.id, session.title || t("resume.defaultTitle"));
    });
    actions.appendChild(deleteButton);

    row.appendChild(actions);
    return row;
  };

  const buildGroupSection = (group: ResumeGroup): HTMLElement | null => {
    if (group.entries.length === 0) return null;

    const section = document.createElement("section");
    section.className = "pi-overlay-section pi-resume-section";
    section.dataset.resumeGroup = group.id;

    const heading = document.createElement("h3");
    heading.className = "pi-overlay-section-title pi-resume-group-title";
    heading.textContent = `${groupTitle(group.id)} (${group.entries.length})`;

    const list = document.createElement("div");
    list.className = "pi-resume-list";

    if (group.id === "other_workbooks") {
      const expanded =
        manuallyExpanded ?? shouldExpandOtherWorkbooks(currentGroups, workbookId);
      const toggle = document.createElement("button");
      toggle.type = "button";
      toggle.className = "pi-resume-group-toggle";
      toggle.textContent = expanded
        ? t("resume.group.collapse")
        : t("resume.group.expand");
      toggle.setAttribute("aria-expanded", String(expanded));
      toggle.addEventListener("click", () => {
        manuallyExpanded = !expanded;
        renderLists();
      });
      heading.appendChild(toggle);

      if (!expanded) {
        section.append(heading);
        return section;
      }
    }

    for (const entry of group.entries) {
      list.appendChild(buildSessionRow(entry));
    }

    section.append(heading, list);
    return section;
  };

  let currentGroups: ResumeGroup[] = [];

  const renderLists = (): void => {
    recentList.replaceChildren();
    groupHost.replaceChildren();

    const recentlyClosed = query.trim().length === 0 ? data.recentlyClosed : [];
    recentlyClosedById = new Map(recentlyClosed.map((item) => [item.id, item]));

    recentSection.hidden = recentlyClosed.length === 0;
    for (const item of recentlyClosed.slice(0, 6)) {
      recentList.appendChild(buildRecentlyClosedRow(item));
    }

    currentGroups = groupSessionsForWorkbook(data.entries, data.links, workbookId);

    const searchedGroups = currentGroups.map((group) => ({
      id: group.id,
      entries: filterSessionsByQuery(group.entries, query),
    }));
    // Expand/collapse decisions must follow the unfiltered groups so typing a
    // query cannot silently collapse the section the user is searching.
    const visibleGroups = searchedGroups.filter(
      (group) => group.entries.length > 0,
    );

    if (visibleGroups.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-overlay-empty pi-resume-list-empty";
      empty.textContent = t("resume.emptyFiltered");
      groupHost.appendChild(empty);
    } else {
      for (const group of visibleGroups) {
        const section = buildGroupSection(group);
        if (section) groupHost.appendChild(section);
      }
    }

    const total = data.entries.length;
    footerStatus.textContent = t("resume.footer.count", { count: total });
  };

  const refreshData = async (): Promise<void> => {
    try {
      data = await loadData(getRecentlyClosedItems);
    } catch (error) {
      console.warn("[pi] Session history refresh failed:", error);
      showToast(t("resume.error.load"));
      return;
    }
    manuallyExpanded = null;
    renderLists();
  };

  const moveSessionToCurrentWorkbook = async (sessionId: string): Promise<void> => {
    if (!workbookId) {
      showToast(t("resume.error.noWorkbook"));
      return;
    }

    const proceed = await requestConfirmationDialog({
      title: t("resume.confirm.moveTitle"),
      message: t("resume.confirm.moveMessage"),
      confirmLabel: t("resume.action.moveHere"),
      cancelLabel: t("confirm.cancel"),
      restoreFocusOnClose: false,
    });
    if (!proceed) return;

    try {
      await relinkSessionToWorkbook(storage.settings, sessionId, workbookId);
      const metadata = data.entries.find(
        (entry) => entry.metadata.id === sessionId,
      )?.metadata;
      if (metadata && (metadata.messageCount ?? 0) > 0) {
        await setLatestSessionForWorkbook(storage.settings, workbookId, sessionId);
      }
      showToast(t("resume.toast.moved"));
      await refreshData();
    } catch (error) {
      console.warn("[pi] Failed to move session:", error);
      showToast(t("resume.error.move"));
    }
  };

  const deleteSession = async (sessionId: string, title: string): Promise<void> => {
    const proceed = await requestConfirmationDialog({
      title: t("resume.confirm.deleteTitle"),
      message: t("resume.confirm.deleteMessage", { title }),
      confirmLabel: t("resume.action.delete"),
      cancelLabel: t("confirm.cancel"),
      restoreFocusOnClose: false,
      confirmButtonTone: "danger",
    });
    if (!proceed) return;

    try {
      await storage.sessions.delete(sessionId);
      await forgetSessionWorkbook(storage.settings, sessionId);
      opts.onSessionDeleted?.(sessionId);
      showToast(t("resume.toast.deleted"));
      await refreshData();
    } catch (error) {
      console.warn("[pi] Failed to delete session:", error);
      showToast(t("resume.error.delete"));
    }
  };

  // ── Events ──
  searchInput.addEventListener("input", () => {
    query = searchInput.value;
    renderLists();
  });

  diagnosticsToggle.addEventListener("click", () => {
    void (async () => {
      diagnosticsPanel.hidden = !diagnosticsPanel.hidden;
      if (diagnosticsPanel.hidden) return;

      diagnosticsPanel.textContent = t("resume.footer.loading");
      diagnosticsPanel.textContent = await collectSessionStorageDiagnostics({
        backend: storage.backend,
        settings: storage.settings,
        dbName: APP_DATABASE_NAME,
        dbVersion: APP_DATABASE_VERSION,
        canonicalWorkbookId: data.identity.canonicalWorkbookId,
        observedWorkbookId: data.identity.observedWorkbookId,
      });
    })();
  });

  exportButton.addEventListener("click", () => {
    void (async () => {
      try {
        const file = await buildSessionExport(storage.sessions);
        if (file.sessions.length === 0) {
          showToast(t("resume.error.nothingToExport"));
          return;
        }
        const stamp = new Date().toISOString().slice(0, 10);
        downloadText(
          `pi-sessions-${stamp}.json`,
          serializeSessionExport(file),
        );
        showToast(t("resume.toast.exported", { count: file.sessions.length }));
      } catch (error) {
        console.warn("[pi] Session export failed:", error);
        showToast(t("resume.error.export"));
      }
    })();
  });

  importButton.addEventListener("click", () => {
    importInput.click();
  });

  importInput.addEventListener("change", () => {
    const file = importInput.files?.[0];
    importInput.value = "";
    if (!file) return;

    void (async () => {
      try {
        const text = await file.text();
        const parsed = parseSessionTransfer(text);
        if ("error" in parsed) {
          showToast(parsed.error);
          return;
        }

        const result = await importSessions(storage.sessions, parsed.file);
        showToast(
          t("resume.toast.imported", {
            count: result.imported,
            skipped: result.skipped,
          }),
        );
        await refreshData();
      } catch (error) {
        console.warn("[pi] Session import failed:", error);
        showToast(t("resume.error.import"));
      }
    })();
  });

  recentList.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    const item = target.closest<HTMLElement>(".pi-resume-item");
    if (!item) return;

    const recentId = item.dataset.recentId;
    if (!recentId) {
      showToast(t("resume.sessionNotInRecentlyClosed"));
      renderLists();
      return;
    }

    const recentEntry = recentlyClosedById.get(recentId);
    if (!recentEntry) {
      showToast(t("resume.sessionNotInRecentlyClosed"));
      renderLists();
      return;
    }

    void (async () => {
      const reopened = await reopenRecentlyClosedEntry(recentEntry);
      if (reopened) {
        closeOverlay();
        return;
      }
      renderLists();
    })();
  });

  groupHost.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    const item = target.closest<HTMLElement>(".pi-resume-item");
    if (!item) return;

    const id = item.dataset.id;
    if (!id) return;

    void (async () => {
      const targetMode = selectedTarget;

      const linkedWorkbookId = data.links.get(id) ?? null;
      if (workbookId && linkedWorkbookId && linkedWorkbookId !== workbookId) {
        const proceed = await requestConfirmationDialog({
          title: t("resume.confirm.anotherWorkbookTitle"),
          message: getCrossWorkbookResumeConfirmMessage(targetMode),
          confirmLabel: t("resume.confirm.resumeAnyway"),
          cancelLabel: t("confirm.cancel"),
          restoreFocusOnClose: false,
        });
        if (!proceed) return;
      }

      const sessionData = await storage.sessions.loadSession(id);
      if (!sessionData) {
        showToast(t("resume.sessionNotFound"));
        await refreshData();
        return;
      }

      if (targetMode === "replace_current") {
        await opts.onReplaceCurrent(sessionData);
      } else {
        await opts.onOpenInNewTab(sessionData);
      }

      closeOverlay();
      const resumedMode =
        targetMode === "replace_current" ? "current tab" : "new tab";
      showToast(
        t("resume.toast.resumedIn", {
          mode: resumedMode,
          title: sessionData.title || t("session.untitled"),
        }),
      );
    })();
  });

  const reopenRecentlyClosedEntry = async (
    item: ResumeRecentlyClosedItem,
  ): Promise<boolean> => {
    if (opts.onReopenRecentlyClosed) {
      return opts.onReopenRecentlyClosed(item);
    }

    const sessionData = await storage.sessions.loadSession(item.sessionId);
    if (!sessionData) {
      showToast(t("resume.couldNotReopen"));
      return false;
    }

    await opts.onOpenInNewTab(sessionData);
    showToast(
      t("resume.toast.reopened", {
        title: sessionData.title || t("session.untitled"),
      }),
    );
    return true;
  };

  renderLists();
  dialog.mount();
}

/** Read everything the overlay needs, in one pass per store. */
async function loadData(
  getRecentlyClosedItems: () => ResumeRecentlyClosedItem[],
): Promise<ResumeDialogData> {
  const storage = getAppStorage();

  const [entries, links, identity] = await Promise.all([
    storage.sessions.getAllMetadataSafe(),
    readSessionWorkbookLinks(storage.settings),
    resolveWorkbookIdentity({
      host: getCurrentSpreadsheetHost(),
      settings: storage.settings,
      // The dialog must not write aliases just by being opened.
      dryRun: true,
    }),
  ]);

  return { entries, links, identity, recentlyClosed: getRecentlyClosedItems() };
}

/**
 * Storage failure while opening history: show the real problem with a retry
 * instead of rejecting into the void (the caller opens this fire-and-forget).
 */
function showLoadErrorOverlay(
  error: DynamicValue,
  opts: ResumeDialogOptions,
): void {
  const dialog = createOverlayDialog({
    overlayId: RESUME_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--m pi-resume-dialog",
  });

  const { header } = createOverlayHeader({
    onClose: dialog.close,
    closeLabel: t("resume.closeLabel"),
    title: t("resume.title"),
    subtitle: t("resume.error.subtitle"),
  });

  const body = document.createElement("div");
  body.className = "pi-overlay-body pi-resume-body";

  const message = document.createElement("p");
  message.className = "pi-overlay-hint";
  message.textContent = t("resume.error.load");

  const detail = document.createElement("pre");
  detail.className = "pi-resume-diagnostics";
  detail.textContent = error instanceof Error ? error.message : String(error);

  const retry = document.createElement("button");
  retry.type = "button";
  retry.className = "pi-overlay-btn pi-overlay-btn--ghost";
  retry.textContent = t("resume.error.retry");
  retry.addEventListener("click", () => {
    dialog.close();
    void showResumeDialog(opts);
  });

  body.append(message, detail, retry);
  dialog.card.append(header, body);
  dialog.mount();
}
