/**
 * `#` document mentions for the Pi TUI prompt editor.
 *
 * When the user types `#` while composing a prompt, this provider lists the
 * Office documents currently attached to the bridge and inserts `#<name>` (or
 * `#<document id>` when two open documents share a name). The server accepts a
 * leading `#` on the `document` tool argument, so the inserted mention is
 * directly usable as the target.
 *
 * The provider is stacked on top of Pi's built-in autocomplete provider via
 * `ctx.ui.addAutocompleteProvider(...)`: it only intercepts its own `#…` token
 * and delegates everything else (notably `@` file references, `/` commands and
 * Tab path completion) untouched.
 *
 * Deliberately free of `@earendil-works/pi-tui` imports: that package is not a
 * dependency of this extension, so the handful of provider/item fields we use
 * are described structurally here. `fuzzyFilter` is replaced by a small
 * case-insensitive substring filter, which is predictable for short document
 * names.
 */

import type { OfficeHostApp } from "./protocol.js";

/** The trigger character for document mentions. */
export const DOCUMENT_MENTION_TRIGGER = "#";

/** Maximum number of documents offered in the dropdown. */
export const MAX_DOCUMENT_MENTIONS = 20;

/** One document offered by the mention dropdown. */
export interface DocumentMentionTarget {
  host: OfficeHostApp;
  /** Display label for the host app, e.g. "Excel". */
  hostLabel: string;
  /** Display name, when the host can provide one. */
  name: string | null;
  /** Stable document id hash, or null for unsaved documents. */
  documentId: string | null;
  /** Server routing key (stable id, or an ephemeral per-connection id). */
  targetId: string;
  /** True when this document is the focused one for its host. */
  focused: boolean;
  /** How many attached windows show this same document. */
  windows: number;
}

/** Minimal structural view of Pi's autocomplete item (`AutocompleteItem`). */
export interface MentionItem {
  value: string;
  label: string;
  description?: string;
}

/** Minimal structural view of Pi's autocomplete provider (`AutocompleteProvider`). */
export interface MentionProvider {
  triggerCharacters?: string[];
  getSuggestions(
    lines: readonly string[],
    cursorLine: number,
    cursorCol: number,
    options: { signal: AbortSignal; force?: boolean },
  ): Promise<{ prefix: string; items: MentionItem[] } | null>;
  applyCompletion(
    lines: readonly string[],
    cursorLine: number,
    cursorCol: number,
    item: MentionItem,
    prefix: string,
  ): { lines: string[]; cursorLine: number; cursorCol: number };
  shouldTriggerFileCompletion?(
    lines: readonly string[],
    cursorLine: number,
    cursorCol: number,
  ): boolean;
}

/**
 * Extract the document-mention token before the cursor, or undefined when the
 * cursor is not inside one. Matches `#` at the start of the line or after
 * whitespace, so `foo#bar` and URL fragments are not treated as mentions.
 */
export function extractDocumentToken(
  line: string,
  cursorCol: number,
): string | undefined {
  const before = line.slice(0, cursorCol);
  const match = before.match(/(?:^|[ \t])#([^\s#]*)$/u);
  return match?.[1];
}

/** Short display form of a document id hash (`url_sha256:abcd…` → `abcd…`). */
export function shortDocumentId(id: string): string {
  const raw = id.includes(":") ? id.slice(id.indexOf(":") + 1) : id;
  return raw.slice(0, 8);
}

/**
 * Build the dropdown items for a query.
 *
 * A document's mention value is `#<name>` when its name is unique among the
 * attached documents (stable across sessions and readable in the transcript),
 * otherwise `#<id>` so the mention stays unambiguous. Focused documents sort
 * first; an empty query lists every document.
 */
export function documentMentionItems(
  targets: readonly DocumentMentionTarget[],
  query: string,
): MentionItem[] {
  if (targets.length === 0) return [];

  const nameCounts = new Map<string, number>();
  for (const target of targets) {
    if (target.name === null) continue;
    const key = target.name.toLowerCase();
    nameCounts.set(key, (nameCounts.get(key) ?? 0) + 1);
  }

  const needle = query.toLowerCase();
  return [...targets]
    .sort((a, b) => {
      if (a.focused !== b.focused) return a.focused ? -1 : 1;
      return (a.name ?? "").localeCompare(b.name ?? "");
    })
    .filter((target) => {
      if (needle === "") return true;
      const id = target.documentId ?? target.targetId;
      return (
        (target.name ?? "").toLowerCase().includes(needle) ||
        id.toLowerCase().includes(needle)
      );
    })
    .slice(0, MAX_DOCUMENT_MENTIONS)
    .map((target) => {
      const id = target.documentId ?? target.targetId;
      const name = target.name;
      const uniqueName =
        name !== null && (nameCounts.get(name.toLowerCase()) ?? 0) === 1;
      const value = uniqueName && name !== null ? name : id;
      const notes = [
        target.hostLabel,
        shortDocumentId(id),
        target.focused ? "focused" : null,
        target.windows > 1 ? `${target.windows} windows` : null,
      ].filter((note): note is string => note !== null);
      return {
        value: `${DOCUMENT_MENTION_TRIGGER}${value}`,
        label: `${DOCUMENT_MENTION_TRIGGER}${target.name ?? shortDocumentId(id)}`,
        description: notes.join(" · "),
      };
    });
}

/**
 * Wrap the built-in autocomplete provider with `#` document mentions.
 *
 * `current` is Pi's provider at the time the wrapper is installed.
 * `getTargets` is called on every keystroke so the dropdown always reflects the
 * documents attached right now.
 */
export function createDocumentMentionProvider(
  current: MentionProvider,
  getTargets: () => readonly DocumentMentionTarget[],
): MentionProvider {
  return {
    triggerCharacters: [DOCUMENT_MENTION_TRIGGER],
    async getSuggestions(lines, cursorLine, cursorCol, options) {
      const token = extractDocumentToken(lines[cursorLine] ?? "", cursorCol);
      if (token === undefined) {
        return current.getSuggestions(lines, cursorLine, cursorCol, options);
      }
      const targets = getTargets();
      if (options.signal.aborted || targets.length === 0) {
        return current.getSuggestions(lines, cursorLine, cursorCol, options);
      }
      const items = documentMentionItems(targets, token);
      if (items.length === 0) {
        return current.getSuggestions(lines, cursorLine, cursorCol, options);
      }
      return { prefix: `${DOCUMENT_MENTION_TRIGGER}${token}`, items };
    },
    applyCompletion(lines, cursorLine, cursorCol, item, prefix) {
      if (!prefix.startsWith(DOCUMENT_MENTION_TRIGGER)) {
        return current.applyCompletion(
          lines,
          cursorLine,
          cursorCol,
          item,
          prefix,
        );
      }
      // Replace the `#token` before the cursor with the selected mention.
      const line = lines[cursorLine] ?? "";
      const before = line.slice(0, cursorCol - prefix.length);
      const after = line.slice(cursorCol);
      const next = [...lines];
      next[cursorLine] = `${before}${item.value}${after}`;
      return {
        lines: next,
        cursorLine,
        cursorCol: before.length + item.value.length,
      };
    },
    shouldTriggerFileCompletion(lines, cursorLine, cursorCol) {
      return (
        current.shouldTriggerFileCompletion?.(lines, cursorLine, cursorCol) ??
        true
      );
    },
  };
}
