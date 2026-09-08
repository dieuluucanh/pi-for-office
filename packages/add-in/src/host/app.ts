/**
 * Office host-app detection — the *application* (Excel / Word / PowerPoint),
 * distinct from the runtime kind (office / wps / browser) in `detection.ts`.
 *
 * A single task-pane manifest now targets all three hosts (see manifest.xml),
 * so at boot the add-in must learn which app it is running inside to pick the
 * right tool set and context model.
 */

export type OfficeApp = "excel" | "word" | "powerpoint" | "powerbi" | "other";

/** Map a raw Office host string (e.g. from `Office.onReady` info or
 *  `Office.context.host`) to a canonical app id. Unknown/null → "other"/null. */
export function parseOfficeApp(raw: string | null | undefined): OfficeApp | null {
  if (typeof raw !== "string" || raw.trim().length === 0) return null;
  const host = raw.trim().toLowerCase();
  switch (host) {
    case "excel":
      return "excel";
    case "word":
      return "word";
    case "powerpoint":
      return "powerpoint";
    case "power bi":
    case "powerbi":
      // Not an Office-add-in host today (Power BI uses the visual SDK), but
      // reserved so future Power BI add-in hosts are recognised.
      return "powerbi";
    default:
      return "other";
  }
}

export function officeAppLabel(app: OfficeApp | null): string {
  switch (app) {
    case "excel":
      return "Excel";
    case "word":
      return "Word";
    case "powerpoint":
      return "PowerPoint";
    case "powerbi":
      return "Power BI";
    case "other":
      return "Other";
    case null:
      return "Unknown";
  }
}

/** Synchronous detection from the Office global (guard-style, no DOM types). */
export function detectOfficeAppFromGlobals(scope: object = globalThis): OfficeApp | null {
  const office = Reflect.get(scope, "Office");
  if (typeof office !== "object" || office === null) return null;
  const context = Reflect.get(office, "context");
  if (typeof context !== "object" || context === null) return null;
  const host = Reflect.get(context, "host");
  return parseOfficeApp(typeof host === "string" ? host : null);
}
