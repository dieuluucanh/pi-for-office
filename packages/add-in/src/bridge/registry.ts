/**
 * Bridge op registry — resolves a `<host>.<op>` id to its pane-side executor.
 *
 * Builds the full registry once; a task-pane instance typically uses only the
 * ops for its own host app, but keeping every host's ops available costs
 * nothing and makes the pane resilient to Pi-side routing changes.
 */

import type { OfficeHostApp } from "@dieulc/pi-office-protocol";
import { EXCEL_OPS } from "./excel-ops.js";
import { WORD_OPS } from "./word-ops.js";
import { POWERPOINT_OPS } from "./ppt-ops.js";
import { executeOfficeOp, type OfficeOpExecutor, type OfficeOpOutcome } from "./ops.js";

/** Every bridge op this pane can execute, keyed by "<host>.<op>". */
export const ALL_BRIDGE_OPS: ReadonlyMap<string, OfficeOpExecutor> = new Map<string, OfficeOpExecutor>([
  ...EXCEL_OPS.entries(),
  ...WORD_OPS.entries(),
  ...POWERPOINT_OPS.entries(),
]);

/** The ops a given host app may execute (used to reject cross-host misuse). */
export function opsForHost(host: OfficeHostApp): ReadonlyMap<string, OfficeOpExecutor> {
  switch (host) {
    case "excel":
      return EXCEL_OPS;
    case "word":
      return WORD_OPS;
    case "powerpoint":
      return POWERPOINT_OPS;
  }
}

/** Host-scoped dispatch — the registry is injected for testability. */
export function dispatchOfficeOp(
  host: OfficeHostApp,
  op: string,
  args: Record<string, unknown>,
  registry: ReadonlyMap<string, OfficeOpExecutor> = opsForHost(host),
): Promise<OfficeOpOutcome> {
  return executeOfficeOp(registry, op, args);
}
