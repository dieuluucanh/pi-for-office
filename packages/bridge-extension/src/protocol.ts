/**
 * Re-export of the shared bridge protocol + office op catalog.
 *
 * The single source of truth lives in `@dieulc/pi-office-protocol` so the Pi
 * extension (server) and the add-in task pane (client) stay in lockstep. The
 * catalog is re-exported from its subpath (the main entry keeps Node-loadable
 * raw TS without a `.js`→`.ts` rewrite).
 */
export * from "@dieulc/pi-office-protocol";
export * from "@dieulc/pi-office-protocol/office-catalog";
