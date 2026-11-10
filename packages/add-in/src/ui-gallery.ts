/**
 * UI Gallery — standalone page for visually verifying components.
 *
 * No Office.js dependency. Loads the same CSS as the real taskpane and
 * renders mock components for agent-browser screenshot verification.
 *
 * Access at: http://localhost:3141/src/ui-gallery.html
 *
 * Each section has a data-gallery attribute for targeted screenshots:
 *   agent-browser screenshot --selector '[data-gallery="tool-cards"]'
 */

// Boot with the same CSS + patches as the real taskpane.
// This imports theme.css and installs Lit/marked/theme patches.
import "./boot.js";

// Register web components we render
import "./ui/register-components.js";

import { escapeHtml, setSafeInnerHTML } from "./utils/html.js";
import { createDismissControl } from "./ui/dismiss-control.js";

const galleryRoot = document.getElementById("gallery-root");
if (!galleryRoot) throw new Error("Missing #gallery-root");
const root: HTMLElement = galleryRoot;
root.style.cssText = `
  max-width: 380px;
  margin: 0 auto;
  padding: 16px;
  font-family: var(--font-sans);
  background: var(--background);
  color: var(--foreground);
  min-height: 100vh;
`;

function section(id: string, title: string): HTMLDivElement {
  const el = document.createElement("div");
  el.setAttribute("data-gallery", id);
  el.style.cssText = "margin-bottom: 32px;";

  const heading = document.createElement("h3");
  heading.textContent = title;
  heading.style.cssText = `
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--muted-foreground);
    margin: 0 0 12px;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--alpha-8);
  `;

  el.appendChild(heading);
  root.appendChild(el);
  return el;
}

/* ── 1. Overlay Badges ───────────────────────────────── */

const badgeSection = section("badges", "Overlay Badges");
const badgeRow = document.createElement("div");
badgeRow.style.cssText = "display: flex; gap: 8px; flex-wrap: wrap;";

for (const tone of ["muted", "ok", "warn", "info"]) {
  const badge = document.createElement("span");
  badge.className = `pi-overlay-badge pi-overlay-badge--${tone}`;
  badge.textContent =
    tone === "muted"
      ? "Read only"
      : tone === "ok"
        ? "Connected"
        : tone === "warn"
          ? "Warning"
          : "Info";
  badgeRow.appendChild(badge);
}
badgeSection.appendChild(badgeRow);

/* ── 2. File Item Rows (with badge) ──────────────────── */

const fileItemSection = section("file-items", "File List Items");

function createMockFileItem(
  name: string,
  meta: string,
  badgeLabel?: string,
): HTMLButtonElement {
  const row = document.createElement("button");
  row.type = "button";
  row.className = "pi-files-item pi-files-item--muted";
  row.style.cssText = "width: 100%;";

  const icon = document.createElement("span");
  icon.className = "pi-files-item__icon";
  icon.textContent = "📄";

  const info = document.createElement("div");
  info.className = "pi-files-item__info";

  const nameRow = document.createElement("div");
  nameRow.className = "pi-files-item__name-row";

  const nameEl = document.createElement("span");
  nameEl.className = "pi-files-item__name";
  nameEl.textContent = name;
  nameRow.appendChild(nameEl);

  const metaEl = document.createElement("span");
  metaEl.className = "pi-files-item__meta";
  metaEl.textContent = meta;

  info.append(nameRow, metaEl);
  row.append(icon, info);

  // Badge is a row-level right-aligned column — same x on every row.
  if (badgeLabel) {
    const badge = document.createElement("span");
    badge.className = "pi-overlay-badge pi-overlay-badge--muted";
    badge.textContent = badgeLabel;
    row.appendChild(badge);
  }

  const arrow = document.createElement("span");
  arrow.className = "pi-files-item__arrow";
  arrow.textContent = "›";

  row.appendChild(arrow);
  return row;
}

fileItemSection.appendChild(
  createMockFileItem(
    "cache-observability-baselines.md",
    "Pi documentation · 2.98 KB",
    "Read only",
  ),
);
fileItemSection.appendChild(
  createMockFileItem(
    "context-management-policy.md",
    "Pi documentation · 12.0 KB",
    "Read only",
  ),
);
fileItemSection.appendChild(
  createMockFileItem("quarterly-report.xlsx", "1.2 MB · Uploaded · 2h ago"),
);

/* ── 3. Tool Cards ───────────────────────────────────── */

const toolCardSection = section("tool-cards", "Tool Cards");

function createMockToolCard(
  state: string,
  action: string,
  detail: string,
): HTMLDivElement {
  const card = document.createElement("div");
  card.className = "pi-tool-card";
  card.setAttribute("data-state", state);
  card.setAttribute("data-tool-name", "fill_formula");

  const header = document.createElement("div");
  header.className = "pi-tool-card__header";

  const toggle = document.createElement("div");
  toggle.className = "pi-tool-card__toggle pi-tool-card__toggle--static";

  const main = document.createElement("span");
  main.className = "pi-tool-card__toggle-main";

  const title = document.createElement("span");
  title.className = "pi-tool-card__title";
  setSafeInnerHTML(
    title,
    `<strong>${escapeHtml(action)}</strong> <span class="pi-tool-card__detail-text">${escapeHtml(detail)}</span>`,
    "UI gallery mock tool-card markup with escaped demo labels",
  );

  main.appendChild(title);
  toggle.appendChild(main);
  header.appendChild(toggle);
  card.appendChild(header);

  return card;
}

toolCardSection.appendChild(
  createMockToolCard("complete", "Filled", "'Cash Flow'!D10:L10 — 9 changes"),
);
toolCardSection.appendChild(
  createMockToolCard("complete", "Filled", "'Cash Flow'!D13:L13 — 9 changes"),
);
toolCardSection.appendChild(
  createMockToolCard("complete", "Filled", "'Cash Flow'!D14:L14 — 9 changes"),
);
toolCardSection.appendChild(
  createMockToolCard("error", "Fill", "'Cash Flow'!D15:L15 — error"),
);

/* ── 4. Tool Card Group ──────────────────────────────── */

const groupSection = section("tool-groups", "Grouped Tool Cards");

const group = document.createElement("div");
group.className = "pi-tool-group";

// Real groups always carry a header (see tool-grouping.ts).
const groupHeader = document.createElement("button");
groupHeader.type = "button";
groupHeader.className = "pi-tool-group__header";
groupHeader.setAttribute("aria-expanded", "true");
const groupChevron = document.createElement("span");
groupChevron.className = "pi-tool-group__chevron";
groupChevron.textContent = "▸";
const groupLabel = document.createElement("span");
groupLabel.className = "pi-tool-group__label";
groupLabel.textContent = "5 fill operations";
groupHeader.append(groupChevron, groupLabel);
groupHeader.addEventListener("click", () => {
  const collapsed = group.classList.toggle("pi-tool-group--collapsed");
  groupHeader.setAttribute("aria-expanded", collapsed ? "false" : "true");
  // Shipped CSS hides direct-child <tool-message>; the gallery uses div
  // stand-ins, so mirror the collapse manually.
  for (const child of group.children) {
    if (child !== groupHeader) {
      (child instanceof HTMLElement ? child : null)?.style.setProperty(
        "display",
        collapsed ? "none" : "",
      );
    }
  }
});
group.appendChild(groupHeader);

for (let i = 10; i <= 14; i++) {
  // Simulate tool-message wrapping — grouped-card CSS matches any
  // descendant .pi-tool-card, so a plain div stands in for <tool-message>.
  const wrapper = document.createElement("div");
  const card = createMockToolCard(
    "complete",
    "Filled",
    `'Cash Flow'!D${i}:L${i} — 9 changes`,
  );
  wrapper.appendChild(card);
  group.appendChild(wrapper);
}
groupSection.appendChild(group);

/* ── 5. Changes Diff Table ───────────────────────────── */

const diffSection = section("diff-table", "Cell Changes Diff Table");

const diffWrap = document.createElement("div");
diffWrap.className = "pi-tool-card__section";
setSafeInnerHTML(
  diffWrap,
  `
  <div class="pi-tool-card__section-label">Changes (9)</div>
  <div class="pi-tool-card__diff">
    <table class="pi-tool-card__diff-table">
      <thead>
        <tr><th>Cell</th><th>Before</th><th>After</th></tr>
      </thead>
      <tbody>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">D10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$125,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =C19+C24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$130,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =D19+D24</div>
          </td>
        </tr>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">E10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$130,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =D19+D24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$135,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =E19+E24</div>
          </td>
        </tr>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">F10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$135,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =E19+E24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$140,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =F19+F24</div>
          </td>
        </tr>
      </tbody>
    </table>
  </div>
`,
  "static UI gallery diff-table fixture markup",
);
diffSection.appendChild(diffWrap);

/* ── 6. Text Preview (file detail) ───────────────────── */

const previewSection = section("text-preview", "File Text Preview");

const sampleMarkdown = `# Context Management Policy

**Status:** Active policy (2026-02-12)
**Scope:** How Pi for Office builds and manages context

## Why this exists

We optimize for **answer quality and reliability** across multi-turn sessions. In practice, quality drops when we blindly stuff context or let it grow unbounded.

## Core principles

1. **Minimal viable context** — include only what improves this turn.
2. **Freshness over volume** — recent state > historical state.
3. **Bounded growth** — auto-compact before hitting limits.`;

const previewCaptionMd = document.createElement("div");
previewCaptionMd.style.cssText =
  "font-size: 10px; color: var(--muted-foreground); margin: 0 0 4px; opacity: 0.8;";
previewCaptionMd.textContent = "Markdown file — rendered";
previewSection.appendChild(previewCaptionMd);

const mdPreview = document.createElement("div");
mdPreview.className =
  "pi-files-detail-preview pi-files-detail-preview--markdown";
const mdPreviewBlock = document.createElement(
  "markdown-block",
) as HTMLElement & { content: string };
mdPreviewBlock.content = sampleMarkdown;
mdPreview.appendChild(mdPreviewBlock);
const mdFade = document.createElement("div");
mdFade.className = "pi-files-detail-preview__fade";
mdPreview.appendChild(mdFade);
previewSection.appendChild(mdPreview);

const previewCaptionText = document.createElement("div");
previewCaptionText.style.cssText =
  "font-size: 10px; color: var(--muted-foreground); margin: 12px 0 4px; opacity: 0.8;";
previewCaptionText.textContent = "Plain text file — wrapped, no line numbers";
previewSection.appendChild(previewCaptionText);

const textPreview = document.createElement("div");
textPreview.className = "pi-files-detail-preview pi-files-detail-preview--text";
const textPreviewBody = document.createElement("div");
textPreviewBody.className = "pi-files-detail-preview__body";
textPreviewBody.textContent = `region,quarter,revenue,cost,margin
EMEA,Q1 2026,1204000,861000,0.285
EMEA,Q2 2026,1287500,903200,0.298
APAC,Q1 2026,644800,517900,0.197
APAC,Q2 2026,701300,548100,0.218`;
textPreview.appendChild(textPreviewBody);
previewSection.appendChild(textPreview);

/* ── 7. Action Buttons ───────────────────────────────── */

const buttonsSection = section("buttons", "Overlay Buttons");

const btnRow = document.createElement("div");
btnRow.className = "pi-files-detail-actions";

for (const [label, cls] of [
  ["Open ↗", "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact"],
  ["Download", "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact"],
  ["Delete", "pi-overlay-btn pi-overlay-btn--danger pi-overlay-btn--compact"],
] as const) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = cls;
  btn.textContent = label;
  btnRow.appendChild(btn);
}
buttonsSection.appendChild(btnRow);

/* ── 8. Toast variants ───────────────────────────────
 * Rendered through a real `.pi-toast-host` (position overridden so it can sit
 * in the gallery flow). Every toast carries the × control, and the long
 * message + action pair is the regression case from a ~350px taskpane. */

const toastSection = section("toasts", "Toast Notifications");

function galleryToast(opts: {
  message: string;
  variant?: "info" | "error";
  action?: string;
}): HTMLDivElement {
  const toast = document.createElement("div");
  toast.className = "pi-toast visible";
  if (opts.variant === "error") toast.classList.add("pi-toast--error");
  if (opts.action) toast.classList.add("pi-toast--action");
  toast.setAttribute("role", opts.variant === "error" ? "alert" : "status");
  toast.setAttribute("aria-atomic", "true");

  const content = document.createElement("div");
  content.className = "pi-toast__content";

  const message = document.createElement("span");
  message.className = "pi-toast__message";
  message.textContent = opts.message;
  content.appendChild(message);

  if (opts.action) {
    const action = document.createElement("button");
    action.type = "button";
    action.className = "pi-toast__action";
    action.textContent = opts.action;
    content.appendChild(action);
  }

  const close = createDismissControl({
    onDismiss: () => {
      toast.classList.remove("visible");
    },
    className: "pi-toast__close",
  });

  toast.append(content, close);
  return toast;
}

function toastHost(): HTMLDivElement {
  const host = document.createElement("div");
  host.className = "pi-toast-host";
  host.style.cssText = "position: static; margin-bottom: 10px;";
  return host;
}

const toastStack = toastHost();
toastStack.style.maxWidth = "360px";
toastStack.append(
  galleryToast({ message: "Tab name reset" }),
  galleryToast({ message: "Closed Chat 3", action: "Undo" }),
  galleryToast({ message: "Could not save", variant: "error" }),
);
toastSection.appendChild(toastStack);

const toastLongLabel = document.createElement("div");
toastLongLabel.textContent = "Long message + action (pane-width regression)";
toastLongLabel.style.cssText =
  "font-size: 11px; color: var(--muted-foreground); margin-bottom: 6px;";
toastSection.appendChild(toastLongLabel);

const toastLongHost = toastHost();
toastLongHost.style.maxWidth = "330px";
toastLongHost.style.alignItems = "stretch";
toastLongHost.appendChild(
  galleryToast({
    message: "No session to restore for this workbook — open session history to pick one.",
    action: "Open session history",
  }),
);
toastSection.appendChild(toastLongHost);

const toastNarrowLabel = document.createElement("div");
toastNarrowLabel.textContent = "Narrow pane (220px) — action wraps below";
toastNarrowLabel.style.cssText =
  "font-size: 11px; color: var(--muted-foreground); margin: 10px 0 6px;";
toastSection.appendChild(toastNarrowLabel);

const toastNarrowHost = toastHost();
toastNarrowHost.style.maxWidth = "220px";
toastNarrowHost.style.alignItems = "stretch";
toastNarrowHost.appendChild(
  galleryToast({ message: "Session moved to this workbook", action: "Undo" }),
);
toastSection.appendChild(toastNarrowHost);

/* ── 9. Markdown rendering (font test) ───────────────── */

const mdSection = section("markdown", "Markdown Rendering (font consistency)");

const mdBlock = document.createElement("markdown-block") as HTMLElement & {
  content: string;
};
mdBlock.content = `The formula is \`=IF(C$4-Assumptions!$B$10+1=Assumptions!$B$49,...)\` — C4 = calendaryear (2025 for Year 1, 2031 for Year 7).

Assumptions!B10 = 2025 (start year) − Assumptions!B$49 = 7

So for Year 7 (column I, calendar year 2031): 2031 – 2025 + 1 = 7 ✓`;

mdSection.appendChild(mdBlock);

/* ── 10. Activity Block (Phase 2 proposal — mockup only) ─ */
// Condenses a run of tool calls between assistant messages into one
// collapsible block. Proposal CSS lives in this <style> tag until the
// direction is approved; it is NOT shipped in the taskpane bundle.

const activitySection = section(
  "activity-block",
  "Activity Block (Phase 2 proposal)",
);

const activityStyles = document.createElement("style");
activityStyles.textContent = `
  /* Proposal: pi-activity — condensed tool-run block (mockup only) */
  .pi-activity {
    border: var(--pill-brand-border);
    background: var(--pill-brand-bg);
    border-radius: var(--pill-radius);
    box-shadow: var(--pill-shadow);
    overflow: hidden;
    margin-bottom: 12px;
  }
  .pi-activity--working {
    border: var(--pill-border);
    background: var(--pill-bg);
  }
  .pi-activity__header {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
    padding: var(--pill-padding-y) var(--pill-padding-x);
    border: none;
    background: none;
    cursor: pointer;
    font-family: var(--font-sans);
    font-size: var(--text-base);
    color: var(--muted-foreground);
    text-align: left;
    transition: color var(--duration-fast);
  }
  .pi-activity__header:hover { color: var(--foreground); }
  .pi-activity__chevron {
    font-size: var(--text-sm);
    opacity: 0.55;
    flex-shrink: 0;
    width: 12px;
    transition: transform var(--duration-fast);
  }
  .pi-activity__header:hover .pi-activity__chevron { opacity: 1; }
  .pi-activity--open .pi-activity__chevron { transform: rotate(90deg); }
  .pi-activity__summary {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    min-width: 0;
  }
  .pi-activity__count {
    margin-left: auto;
    flex-shrink: 0;
    font-size: var(--text-sm);
    opacity: 0.7;
  }
  .pi-activity__count em {
    font-style: normal;
    color: var(--destructive);
  }
  .pi-activity__steps {
    border-top: 1px solid var(--brand-alpha-12);
    padding: 4px 0 6px;
  }
  .pi-activity__step {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
    padding: 3px var(--pill-padding-x);
    border: none;
    background: none;
    cursor: pointer;
    text-align: left;
    font-family: var(--font-sans);
    font-size: var(--text-sm);
    color: var(--muted-foreground);
  }
  .pi-activity__step:hover { color: var(--foreground); }
  .pi-activity__step-status {
    flex-shrink: 0;
    width: 12px;
    font-size: var(--text-xs);
    color: var(--pi-brand);
    opacity: 0.8;
  }
  .pi-activity__step-status--error { color: var(--destructive); }
  .pi-activity__step-label {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    min-width: 0;
  }
  .pi-activity__step-label strong { font-weight: 400; }
  .pi-activity__step-detail { opacity: 0.7; }
  .pi-activity__step--error .pi-activity__step-label { color: var(--destructive); }
  .pi-activity__drilldown {
    padding: 2px var(--pill-padding-x) 8px calc(var(--pill-padding-x) + 20px);
  }
`;
activitySection.appendChild(activityStyles);

const activityNote = document.createElement("p");
activityNote.style.cssText =
  "font-size: 11px; color: var(--muted-foreground); margin: 0 0 12px;";
activityNote.textContent =
  "Proposal: one block per tool run instead of one card per call. States: live (working), collapsed (default when done), expanded with drill-in.";
activitySection.appendChild(activityNote);

type MockStep = { verb: string; detail: string; error?: boolean };

const mockSteps: MockStep[] = [
  { verb: "Read", detail: "'Cash Flow'!A1:N40" },
  { verb: "Filled", detail: "'Cash Flow'!D10:L10 — 9 changes" },
  { verb: "Filled", detail: "'Cash Flow'!D13:L13 — 9 changes" },
  { verb: "Fill", detail: "'Cash Flow'!D15:L15 — #REF! error", error: true },
  { verb: "Formatted", detail: "'Cash Flow'!D10:L15 — currency" },
];

function createActivityBlock(
  state: "working" | "collapsed" | "open",
): HTMLDivElement {
  const block = document.createElement("div");
  block.className =
    state === "working"
      ? "pi-activity pi-activity--working"
      : state === "open"
        ? "pi-activity pi-activity--open"
        : "pi-activity";

  const header = document.createElement("button");
  header.type = "button";
  header.className = "pi-activity__header";

  if (state === "working") {
    const label = document.createElement("span");
    // Reuses the real thinking-block shimmer + spinner affordance.
    label.className = "pi-thinking-label--streaming pi-activity__summary";
    label.textContent = "Working — formatting 'Cash Flow'!D10:L15";
    const count = document.createElement("span");
    count.className = "pi-activity__count";
    count.textContent = "step 5";
    header.append(label, count);
  } else {
    const chevron = document.createElement("span");
    chevron.className = "pi-activity__chevron";
    chevron.textContent = "▸";
    const summary = document.createElement("span");
    summary.className = "pi-activity__summary";
    summary.textContent = "Worked for 8s";
    const count = document.createElement("span");
    count.className = "pi-activity__count";
    setSafeInnerHTML(
      count,
      `5 steps · <em>1 issue</em>`,
      "UI gallery mock activity count with static demo markup",
    );
    header.append(chevron, summary, count);
  }
  block.appendChild(header);

  if (state === "open") {
    const steps = document.createElement("div");
    steps.className = "pi-activity__steps";
    mockSteps.forEach((step, i) => {
      const row = document.createElement("button");
      row.type = "button";
      row.className = step.error
        ? "pi-activity__step pi-activity__step--error"
        : "pi-activity__step";
      const status = document.createElement("span");
      status.className = step.error
        ? "pi-activity__step-status pi-activity__step-status--error"
        : "pi-activity__step-status";
      status.textContent = step.error ? "✕" : "✓";
      const label = document.createElement("span");
      label.className = "pi-activity__step-label";
      setSafeInnerHTML(
        label,
        `<strong>${escapeHtml(step.verb)}</strong> <span class="pi-activity__step-detail">${escapeHtml(step.detail)}</span>`,
        "UI gallery mock activity step with escaped demo labels",
      );
      row.append(status, label);
      steps.appendChild(row);

      // Demonstrate drill-in: the error step expands to its full tool card.
      if (i === 3) {
        const drilldown = document.createElement("div");
        drilldown.className = "pi-activity__drilldown";
        drilldown.appendChild(
          createMockToolCard(
            "error",
            "Fill",
            "'Cash Flow'!D15:L15 — #REF! error",
          ),
        );
        steps.appendChild(drilldown);
      }
    });
    block.appendChild(steps);
  }

  return block;
}

for (const [label, state] of [
  ["Live (streaming)", "working"],
  ["Done — collapsed (default)", "collapsed"],
  ["Done — expanded with drill-in", "open"],
] as const) {
  const caption = document.createElement("div");
  caption.style.cssText =
    "font-size: 10px; color: var(--muted-foreground); margin: 0 0 4px; opacity: 0.8;";
  caption.textContent = label;
  activitySection.appendChild(caption);
  activitySection.appendChild(createActivityBlock(state));
}

console.log("[ui-gallery] Rendered all sections");

/* ── 11. Dialog cards (fluid sizing) ──────────────────── */

const dialogSection = section("dialogs", "Dialog Cards (fluid sizing)");

const DIALOG_SIZES: Array<{ label: string; className: string; note: string }> = [
  {
    label: "--s (confirm / text input)",
    className: "pi-overlay-card--s",
    note: "min(100%, 480px)",
  },
  {
    label: "--m (resume history / files)",
    className: "pi-overlay-card--m",
    note: "min(100%, 720px)",
  },
  {
    label: "--l (settings shell)",
    className: "pi-overlay-card--l pi-set-shell",
    note: "min(100%, 960px)",
  },
];

function createDialogCardMock(sizeClass: string): HTMLDivElement {
  const card = document.createElement("div");
  card.className = `pi-welcome-card pi-overlay-card ${sizeClass}`;

  const header = document.createElement("div");
  header.className = "pi-overlay-header";
  const titleWrap = document.createElement("div");
  titleWrap.className = "pi-overlay-title-wrap";
  const title = document.createElement("h2");
  title.className = "pi-overlay-title";
  title.textContent = "Session history";
  const subtitle = document.createElement("p");
  subtitle.className = "pi-overlay-subtitle";
  subtitle.textContent = "Pick a session to resume in this workbook.";
  titleWrap.append(title, subtitle);
  header.append(
    titleWrap,
    createDismissControl({ onDismiss: () => {}, label: "Close dialog" }),
  );

  const body = document.createElement("div");
  body.className = "pi-overlay-body";
  for (let index = 1; index <= 5; index += 1) {
    const row = document.createElement("div");
    row.className = "pi-resume-item-row";
    const main = document.createElement("button");
    main.type = "button";
    main.className = "pi-welcome-provider pi-resume-item";
    const rowTitle = document.createElement("span");
    rowTitle.className = "pi-resume-item__title";
    rowTitle.textContent = `Quarterly forecast ${index}`;
    const meta = document.createElement("span");
    meta.className = "pi-resume-item__meta";
    meta.textContent = "48 messages · 2 hours ago";
    main.append(rowTitle, meta);
    row.appendChild(main);
    body.appendChild(row);
  }

  const footer = document.createElement("div");
  footer.className = "pi-resume-footer";
  const status = document.createElement("span");
  status.className = "pi-resume-footer__status";
  status.textContent = "10 sessions";
  footer.appendChild(status);
  for (const label of ["Storage details", "Export", "Import"]) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "pi-resume-footer__btn";
    button.textContent = label;
    footer.appendChild(button);
  }

  card.append(header, body, footer);
  return card;
}

for (const size of DIALOG_SIZES) {
  const caption = document.createElement("div");
  caption.style.cssText =
    "font-size: 10px; color: var(--muted-foreground); margin: 0 0 4px; opacity: 0.8;";
  caption.textContent = `${size.label} — ${size.note}`;
  dialogSection.appendChild(caption);

  // 360px frame: the real Office taskpane width.
  const frame = document.createElement("div");
  frame.className = "pi-welcome-overlay";
  frame.style.cssText =
    "position: relative; inset: auto; width: 360px; height: 460px; margin: 0 0 16px; padding: var(--pi-dialog-gutter); box-sizing: border-box;";
  frame.appendChild(createDialogCardMock(size.className));
  dialogSection.appendChild(frame);
}

/* ── 12. Notices (dismiss affordances) ───────────────── */

const noticeSection = section("notices", "Notices (dismiss affordances)");

const errorHost = document.createElement("div");
const errorFrame = document.createElement("div");

errorFrame.style.cssText = "margin-bottom: 12px;";
errorFrame.appendChild(errorHost);
noticeSection.appendChild(errorFrame);

const errorBannerModule = await import("./ui/error-banner.js");
errorBannerModule.mountErrorBanner(errorHost, {
  message: "LLM error: rate limited — retrying in 8s.",
  actions: [
    { label: "Retry", onClick: () => {} },
    { label: "New session", variant: "cancel", onClick: () => {} },
  ],
});

const proxyModule = await import("./ui/proxy-banner.js");
const proxyBanner = proxyModule.createProxyBanner();
proxyBanner.update("not-detected");
noticeSection.appendChild(proxyBanner.root);

const popoverModule = await import("./taskpane/status-popovers.js");
const { header: popoverHeader } = popoverModule.createStatusPopoverHeader("Thinking level");
const popoverMock = document.createElement("div");
popoverMock.className = "pi-status-popover";
popoverMock.style.cssText = "position: static; margin-bottom: 12px; width: 100%;";
const popoverDescription = document.createElement("p");
popoverDescription.className = "pi-status-popover__description";
popoverDescription.textContent =
  "How deeply Pi reasons before answering — higher is slower but more thorough.";
popoverMock.append(popoverHeader, popoverDescription);
noticeSection.appendChild(popoverMock);

const disclosureClose = document.createElement("div");
disclosureClose.className = "pi-disclosure-bar";
const disclosureHead = document.createElement("div");
disclosureHead.className = "pi-disclosure-bar__head";
const disclosureText = document.createElement("div");
disclosureText.className = "pi-disclosure-bar__text";
disclosureText.textContent =
  "Pi can search the web, use extensions, and connect to external services.";
disclosureHead.append(
  disclosureText,
  createDismissControl({
    onDismiss: () => disclosureClose.remove(),
    className: "pi-disclosure-bar__close",
  }),
);
disclosureClose.appendChild(disclosureHead);
noticeSection.appendChild(disclosureClose);

/* ── 13. Layout probe (measured numbers for verification) ─ */

const probeSection = section("layout-probe", "Layout Probe (current viewport)");
const probeOutput = document.createElement("pre");
probeOutput.style.cssText =
  "font: 11px/1.4 var(--font-mono); white-space: pre-wrap; background: var(--alpha-4); padding: 8px; border-radius: var(--radius-sm); margin: 0;";
probeOutput.textContent = "measuring…";
probeSection.appendChild(probeOutput);

function measureCard(sizeClass: string): { width: number; height: number } {
  const overlay = document.createElement("div");
  overlay.className = "pi-welcome-overlay";
  overlay.style.cssText = "visibility: hidden; pointer-events: none; z-index: -1;";
  const card = createDialogCardMock(sizeClass);
  overlay.appendChild(card);
  document.body.appendChild(overlay);
  const rect = card.getBoundingClientRect();
  overlay.remove();
  return { width: Math.round(rect.width), height: Math.round(rect.height) };
}

function measureToast(): { width: number; height: number; messageWidth: number } {
  const host = document.createElement("div");
  host.className = "pi-toast-host";
  host.style.cssText = "position: static;";
  const toast = galleryToast({
    message:
      "No session to restore for this workbook — open session history to pick one.",
    action: "Open session history",
  });
  host.appendChild(toast);
  document.body.appendChild(host);
  const rect = toast.getBoundingClientRect();
  const messageRect = toast.querySelector(".pi-toast__message")?.getBoundingClientRect();
  host.remove();
  return {
    width: Math.round(rect.width),
    height: Math.round(rect.height),
    messageWidth: Math.round(messageRect?.width ?? 0),
  };
}

// Let layout settle before measuring, then print machine-readable numbers.
requestAnimationFrame(() => {
  const viewport = { width: window.innerWidth, height: window.innerHeight };
  const cards = DIALOG_SIZES.map((size) => {
    const measured = measureCard(size.className);
    return {
      size: size.className,
      ...measured,
      widthPct: Math.round((measured.width / viewport.width) * 100),
      heightPct: Math.round((measured.height / viewport.height) * 100),
    };
  });
  const toast = measureToast();
  const report = { viewport, cards, toast };
  probeOutput.textContent = JSON.stringify(report, null, 2);
  console.log("[ui-gallery] layout-probe", JSON.stringify(report));
});
