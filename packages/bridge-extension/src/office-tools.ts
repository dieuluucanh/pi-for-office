/**
 * Office tool catalog adapter — thin wrapper over the shared catalog.
 *
 * The authoritative op definitions (names, schemas, descriptions) live in
 * `@dieulc/pi-office-protocol` (`office-catalog.ts`). This module only maps
 * catalog entries to the descriptors this extension registers with Pi, so the
 * Pi side can never drift from the add-in's pane-side registry.
 */

import { Type, type TSchema } from "typebox";
import {
 CATALOG_VERSION,
 HOST_APP_LABEL,
 OFFICE_CATALOG,
 OFFICE_TARGET_PARAM,
 hostForToolName,
} from "./protocol.js";
import type { OfficeHostApp } from "./protocol.js";

export interface OfficeToolDescriptor {
 /** Host app that can execute this op. */
 host: OfficeHostApp;
 /** Payload op id, namespaced by host: "excel.read_range". */
 op: string;
 /** Pi-registered tool name, e.g. "office_excel_read_range". */
 name: string;
 label: string;
 description: string;
 promptSnippet?: string;
 promptGuidelines?: string[];
 parameters: TSchema;
}

/** The catalog version this extension's tool surface was built from. */
export const TOOL_CATALOG_VERSION = CATALOG_VERSION;

/**
 * Optional per-tool target document selector (protocol v2). Added to every
 * bridge descriptor and stripped by the server before the pane sees the call.
 */
export const DOCUMENT_TARGET_PARAM = Type.Optional(
 Type.String({
  description:
   'Target document: "active" (default) for the focused document, a document ' +
   "id from office_list_documents, or a unique document name. Required when " +
   "several documents of this host are open and none is focused. A " +
   '"#name" mention from the TUI is accepted.',
 }),
);

/**
 * Add the optional `document` target property to a catalog parameter schema.
 *
 * The shared catalog stays target-free so the add-in's local/browser tools keep
 * their current schemas; only the bridge's descriptors gain the selector. Every
 * other schema option (including TypeBox's internal kind/optional markers) is
 * copied through unchanged; `required` is unaffected because the new property
 * is optional.
 */
export function withDocumentTarget(schema: TSchema): TSchema {
 // SAFETY: catalog schemas are TypeBox object schemas created at runtime; the
 // keyed-record widening lets us copy all schema options and replace only
 // `properties`.
 const obj = schema as unknown as Record<string, unknown>;
 const properties = obj.properties as Record<string, TSchema> | undefined;
 if (obj.type !== "object" || properties === undefined) return schema;
 return {
  ...obj,
  properties: {
   [OFFICE_TARGET_PARAM]: DOCUMENT_TARGET_PARAM,
   ...properties,
  },
 } as TSchema;
}

/** All descriptors built from the shared catalog (deterministic order). */
export const OFFICE_TOOL_DESCRIPTORS: readonly OfficeToolDescriptor[] =
 OFFICE_CATALOG.map((entry) => ({
  host: entry.host,
  op: entry.op,
  name: entry.name,
  label: entry.label,
  description: entry.description,
  ...(entry.promptSnippet === undefined
   ? null
   : { promptSnippet: entry.promptSnippet }),
  ...(entry.promptGuidelines === undefined
   ? null
   : { promptGuidelines: entry.promptGuidelines }),
  parameters: withDocumentTarget(entry.parameters),
 }));

/** Index by op id for fast lookup. */
export const OFFICE_TOOL_BY_OP: ReadonlyMap<string, OfficeToolDescriptor> =
 new Map(OFFICE_TOOL_DESCRIPTORS.map((d) => [`${d.host}.${d.op}`, d]));

/** All tool names registered by this extension. */
export const OFFICE_TOOL_NAMES: readonly string[] = OFFICE_TOOL_DESCRIPTORS.map(
 (d) => d.name,
);

/** The office host this tool name drives, or null when unknown. (from catalog) */
export { HOST_APP_LABEL, hostForToolName };
