/**
 * Active office tool reconciliation — pure functions that decide which
 * `office_<host>_<op>` tools Pi exposes based on the currently attached panes.
 *
 * Kept free of Pi API calls so it is unit-testable without a live Pi session.
 *
 * Rules:
 *  - Office tool names are identified from the shared catalog.
 *  - A pane advertising `ops` exposes exactly those ops' tools.
 *  - A legacy pane (no `ops`) exposes only the v1 op set.
 *  - Tools for hosts that are no longer attached are deactivated.
 *  - Non-office tools in the current active set are preserved untouched.
 */

import {
 CATALOG_VERSION,
 LEGACY_V1_OPS,
 OFFICE_CATALOG_BY_OP,
 OFFICE_TOOL_NAMES,
 officeToolName,
} from "./protocol.js";
import type { OfficeHostApp } from "./protocol.js";

/** What Pi knows about a pane's capabilities (server-normalized). */
export interface PaneCapability {
 host: OfficeHostApp;
 /** Ops this pane advertises, or null for legacy panes (v1 set only). */
 ops: readonly string[] | null;
}

/** The catalog version the pane's ops were derived from, when advertised. */
export type { CATALOG_VERSION };

/** Office tool names from the catalog, as a set for O(1) membership. */
const OFFICE_TOOL_NAME_SET: ReadonlySet<string> = new Set(OFFICE_TOOL_NAMES);

/** True when the name is one of the catalog's office_* tools. */
export function isOfficeToolName(name: string): boolean {
 return OFFICE_TOOL_NAME_SET.has(name);
}

/**
 * The pi tool names enabled by the given panes (host-scoped, deterministic
 * order). Legacy panes without an `ops` list get the v1 op set only.
 */
export function activeOfficeToolNames(
 panes: readonly PaneCapability[],
): string[] {
 const names = new Set<string>();
 for (const pane of panes) {
  const ops = pane.ops ?? LEGACY_V1_OPS;
  for (const op of ops) {
   const entry = OFFICE_CATALOG_BY_OP.get(op);
   if (entry && entry.host === pane.host) {
    names.add(officeToolName(entry.host, entry.op));
   }
  }
 }
 return [...names].sort((a, b) => a.localeCompare(b));
}

/**
 * Compute the next active tool list: drop all office tools, then re-add the
 * office tools for the currently attached panes, preserving the caller's
 * existing non-office tools. Deterministic and idempotent.
 */
export function reconcileOfficeToolActivation(
 currentActive: readonly string[],
 panes: readonly PaneCapability[],
): string[] {
 const next = currentActive.filter((name) => !isOfficeToolName(name));
 const toAdd = activeOfficeToolNames(panes);
 return [...new Set([...next, ...toAdd])].sort((a, b) => {
  const aOffice = isOfficeToolName(a) ? 1 : 0;
  const bOffice = isOfficeToolName(b) ? 1 : 0;
  if (aOffice !== bOffice) return aOffice - bOffice;
  return a.localeCompare(b);
 });
}
