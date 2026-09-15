/**
 * Runtime Office.js capability detection for the Word tools.
 *
 * Office.js features are gated by *requirement sets* (`WordApi 1.3`,
 * `WordApiDesktop 1.3`, …). Some sets are platform-specific: `WordApiDesktop`
 * exists only in Word on Windows and Mac, so an API in that set throws on Word
 * for the web.
 *
 * Nothing in this add-in checked any of that. The consequence was not only
 * opaque runtime errors but also a description that claimed formatting was
 * "fully supported" — which is how the agent ended up telling a user that Word
 * cannot highlight text, when `font.highlightColor` is WordApi 1.1.
 *
 * This module is the single place that answers "can this host do X?".
 *
 * Microsoft guidance: the `<Set>` element of an add-in-only manifest cannot
 * declare platform-specific requirement sets, so a runtime check is the only
 * compliant mechanism for this repo (see `packages/add-in/manifest.xml`).
 * https://learn.microsoft.com/en-us/office/dev/add-ins/develop/platform-specific-requirement-sets
 */

/** Result of probing one capability.
 *
 * - `supported` — the host advertises the requirement set.
 * - `unsupported` — the host is known to lack it.
 * - `unknown` — the host exposes no requirements API at all (WPS, or an
 *   Office.js build predating `isSetSupported`). Callers must not treat this
 *   as unsupported: attempt the call and surface the real error.
 */
export type CapabilityVerdict = "supported" | "unsupported" | "unknown";

/** One host capability the Word tool surface depends on. */
export interface CapabilitySpec {
  /** Stable id, e.g. `word.hyperlink.native`. */
  id: string;
  /** Human label used in tool output. */
  label: string;
  /** Requirement set that provides the API. */
  set: string;
  /**
   * Minimum version within `set`. A **string** on purpose — Microsoft warns the
   * JS parser cannot distinguish the numeric `1.1` from `1.10`.
   */
  minVersion: string;
  /** What to do instead when the host does not support it. */
  remedy: string;
}

/**
 * Word capabilities the tool surface depends on, ordered from universal
 * (WordApi 1.1) to desktop-only (WordApiDesktop 1.3).
 */
export const WORD_CAPABILITIES: readonly CapabilitySpec[] = [
  {
    id: "word.font.basic",
    label: "font formatting (bold, italic, underline, size, name, color)",
    set: "WordApi",
    minVersion: "1.1",
    remedy: "Requires Word.",
  },
  {
    id: "word.font.highlight",
    label: "text highlight color (#RRGGBB)",
    set: "WordApi",
    minVersion: "1.1",
    remedy: "Requires Word.",
  },
  {
    id: "word.search.replace",
    label: "find and replace",
    set: "WordApi",
    minVersion: "1.1",
    remedy: "Requires Word.",
  },
  {
    id: "word.ooxml.insert",
    label: "raw OOXML insertion",
    set: "WordApi",
    minVersion: "1.1",
    remedy: "Requires Word.",
  },
  {
    id: "word.image.inline",
    label: "inline images",
    set: "WordApi",
    minVersion: "1.2",
    remedy: "Update Word to a build supporting WordApi 1.2.",
  },
  {
    id: "word.table.insert",
    label: "table insertion",
    set: "WordApi",
    minVersion: "1.3",
    remedy: "Update Word to a build supporting WordApi 1.3.",
  },
  {
    id: "word.range.getRange",
    label: "range navigation (`getRange`) for multi-run insertion",
    set: "WordApi",
    minVersion: "1.3",
    remedy: "Update Word to a build supporting WordApi 1.3.",
  },
  {
    id: "word.hyperlink.native",
    label: "native hyperlink insertion",
    set: "WordApiDesktop",
    minVersion: "1.3",
    remedy:
      "Desktop-only (Windows/Mac). On Word for the web insert the link as OOXML instead.",
  },
  {
    id: "word.shading.range",
    label: "range/paragraph shading (true background fill)",
    set: "WordApiDesktop",
    minVersion: "1.3",
    remedy:
      "Desktop-only (Windows/Mac). On Word for the web use text highlight instead.",
  },
];

/**
 * Excel capabilities the tool surface depends on.
 *
 * Versions verified against the installed `@types/office-js` d.ts rather than
 * assumed — an earlier audit guessed several of these wrong (it claimed `merge`
 * needed 1.16 and shape access 1.16; both are far older).
 */
export const EXCEL_CAPABILITIES: readonly CapabilitySpec[] = [
  {
    id: "excel.sheet.visibility",
    label: "hiding and unhiding worksheets",
    set: "ExcelApi",
    minVersion: "1.2",
    remedy: "Update Excel to a build supporting ExcelApi 1.2.",
  },
  {
    id: "excel.conditionalFormat.types",
    label:
      "condition-based formats (data bars, color scales, icon sets, top/bottom, preset and text rules)",
    set: "ExcelApi",
    minVersion: "1.6",
    remedy: "Update Excel to a build supporting ExcelApi 1.6.",
  },
  {
    id: "excel.range.image",
    label: "rendering a range or chart as an image",
    set: "ExcelApi",
    minVersion: "1.7",
    remedy: "Update Excel to a build supporting ExcelApi 1.7.",
  },
  {
    id: "excel.view.gridlines",
    label: "worksheet gridline and heading visibility",
    set: "ExcelApi",
    minVersion: "1.8",
    remedy: "Update Excel to a build supporting ExcelApi 1.8.",
  },
  {
    id: "excel.font.effects",
    label: "strikethrough, subscript/superscript and indent level",
    set: "ExcelApi",
    minVersion: "1.9",
    remedy: "Update Excel to a build supporting ExcelApi 1.9.",
  },
  {
    id: "excel.rangeAreas",
    label: "multi-area range operations",
    set: "ExcelApi",
    minVersion: "1.9",
    remedy: "Update Excel to a build supporting ExcelApi 1.9.",
  },
  {
    id: "excel.autoFill",
    label: "autofill",
    set: "ExcelApi",
    minVersion: "1.9",
    remedy: "Update Excel to a build supporting ExcelApi 1.9.",
  },
  {
    id: "excel.shape",
    label: "shapes",
    set: "ExcelApi",
    minVersion: "1.9",
    remedy: "Update Excel to a build supporting ExcelApi 1.9.",
  },
  {
    id: "excel.comments",
    label: "cell comments and threaded replies",
    set: "ExcelApi",
    minVersion: "1.10",
    remedy: "Update Excel to a build supporting ExcelApi 1.10.",
  },
];

/**
 * PowerPoint capabilities the tool surface depends on.
 *
 * Note the split inside bullets: `visible` is 1.4 but `style`/`type` are 1.10,
 * so bullet style must be gated separately from bullet visibility.
 */
export const POWERPOINT_CAPABILITIES: readonly CapabilitySpec[] = [
  {
    id: "powerpoint.shapes",
    label: "shape access",
    set: "PowerPointApi",
    minVersion: "1.3",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.3.",
  },
  {
    id: "powerpoint.textFormat",
    label: "text frame, font and paragraph formatting",
    set: "PowerPointApi",
    minVersion: "1.4",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.4.",
  },
  {
    id: "powerpoint.shape.fill",
    label: "shape fill colour, transparency and clear",
    set: "PowerPointApi",
    minVersion: "1.4",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.4.",
  },
  {
    id: "powerpoint.shape.line",
    label: "shape line colour, weight, style and visibility",
    set: "PowerPointApi",
    minVersion: "1.4",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.4.",
  },
  {
    id: "powerpoint.font.underlineStyles",
    label: "underline styles beyond single",
    set: "PowerPointApi",
    minVersion: "1.4",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.4.",
  },
  {
    id: "powerpoint.bullet.visible",
    label: "bullet visibility",
    set: "PowerPointApi",
    minVersion: "1.4",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.4.",
  },
  {
    id: "powerpoint.slide.selection",
    label: "selecting slides in the editing surface",
    set: "PowerPointApi",
    minVersion: "1.5",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.5.",
  },
  {
    id: "powerpoint.font.effects",
    label: "strikethrough, small caps, all caps and super/subscript",
    set: "PowerPointApi",
    minVersion: "1.8",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.8.",
  },
  {
    id: "powerpoint.bullet.style",
    label: "bullet style and type",
    set: "PowerPointApi",
    minVersion: "1.10",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.10.",
  },
  {
    id: "powerpoint.paragraph.indent",
    label: "paragraph indent level",
    set: "PowerPointApi",
    minVersion: "1.10",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.10.",
  },
  {
    id: "powerpoint.slide.background",
    label: "slide background fill",
    set: "PowerPointApi",
    minVersion: "1.10",
    remedy: "Update PowerPoint to a build supporting PowerPointApi 1.10.",
  },
];

/** Hosts that have a capability registry. */
export type CapabilityHost = "word" | "excel" | "powerpoint";

/**
 * Capability ids that exist only to be *reported* by the probe ops.
 *
 * Every id here is a requirement set the current tool surface does not gate
 * on — either because it is universal (`WordApi 1.1`, present on every host
 * that can run the add-in) or because no parameter needs it. They still appear
 * in the capability reports so the model can answer "can this host do X?"
 * truthfully instead of guessing.
 *
 * The consistency guard fails any id that is neither gated by a caller nor
 * listed here, so a declared-but-unreachable capability cannot linger.
 */
export const PROBE_ONLY_CAPABILITY_IDS: ReadonlySet<string> = new Set([
  "word.font.basic",
  "word.font.highlight",
  "word.search.replace",
  "word.ooxml.insert",
]);

/** The subset of `Office.context.requirements` this module depends on. */
interface RequirementsLike {
  isSetSupported(set: string, minVersion: string): boolean;
}

/**
 * The global object viewed as an open property bag, so the Office global can be
 * probed without assuming it exists (WPS and browser-only runs lack it).
 */
// SAFETY: `globalThis` is an object whose properties are probed by name. The
// double assertion is needed because the DOM lib types `globalThis` without an
// index signature; every property is re-validated by `asBag` before use, so no
// unchecked value escapes this module.
const GLOBAL_SCOPE = globalThis as unknown as Record<string, unknown>;

/** Narrow an unknown value to a property bag, or null. */
function asBag(value: unknown): Record<string, unknown> | null {
  return typeof value === "object" && value !== null
    ? (value as Record<string, unknown>)
    : null;
}

/** Read `Office.context` without assuming the Office global exists. */
function officeContextFrom(
  scope: Record<string, unknown>,
): Record<string, unknown> | null {
  const office = asBag(scope["Office"]);
  if (office === null) return null;
  return asBag(office["context"]);
}

/** Resolve the requirements API, or null when the host does not provide it. */
function requirementsFrom(
  scope: Record<string, unknown>,
): RequirementsLike | null {
  const context = officeContextFrom(scope);
  if (context === null) return null;
  const requirements = asBag(context["requirements"]);
  if (requirements === null) return null;
  const isSetSupported = requirements["isSetSupported"];
  if (typeof isSetSupported !== "function") return null;
  return {
    isSetSupported: isSetSupported as RequirementsLike["isSetSupported"],
  };
}

/**
 * The host platform as reported by Office (`"PC"`, `"Mac"`, `"OfficeOnline"`,
 * `"iOS"`, `"Android"`), or null when unknown.
 *
 * Used only to *name* the platform in messages — never to infer capability.
 * `isSetSupported` is the decision-maker, because a platform-specific set can
 * span more than one platform (WordApiDesktop covers Windows *and* Mac).
 */
export function currentPlatform(
  scope: Record<string, unknown> = GLOBAL_SCOPE,
): string | null {
  const context = officeContextFrom(scope);
  if (context === null) return null;
  const platform = context["platform"];
  return typeof platform === "string" ? platform : null;
}

/** Probe one capability against the current host. */
export function probeCapability(
  spec: CapabilitySpec,
  scope: Record<string, unknown> = GLOBAL_SCOPE,
): CapabilityVerdict {
  const requirements = requirementsFrom(scope);
  if (requirements === null) return "unknown";
  try {
    return requirements.isSetSupported(spec.set, spec.minVersion)
      ? "supported"
      : "unsupported";
  } catch {
    // A malformed/unknown set name must not make the probe itself fail.
    return "unknown";
  }
}

/**
 * All capability specs across every host, for id lookup.
 *
 * Ids are host-prefixed (`word.`, `excel.`, `powerpoint.`), so one flat list is
 * unambiguous and `capabilitySpec()` resolves whichever host a caller names.
 */
const ALL_CAPABILITIES: readonly CapabilitySpec[] = [
  ...WORD_CAPABILITIES,
  ...EXCEL_CAPABILITIES,
  ...POWERPOINT_CAPABILITIES,
];

/** Capability registry per host. */
const CAPABILITIES_BY_HOST: Readonly<
  Record<CapabilityHost, readonly CapabilitySpec[]>
> = {
  word: WORD_CAPABILITIES,
  excel: EXCEL_CAPABILITIES,
  powerpoint: POWERPOINT_CAPABILITIES,
};

/** Look up a capability spec by its stable id, on any host. */
export function capabilitySpec(id: string): CapabilitySpec | undefined {
  return ALL_CAPABILITIES.find((spec) => spec.id === id);
}

/** Every capability spec declared for a host, in registry order. */
export function capabilitiesForHost(
  host: CapabilityHost,
): readonly CapabilitySpec[] {
  return CAPABILITIES_BY_HOST[host];
}

/**
 * Render the fail-loud message for a capability the host does not support:
 * what is missing, the requirement set that provides it, the platform we are
 * actually on, and what to do instead.
 *
 * Callers must NOT use this for an `"unknown"` verdict — that means the host
 * exposes no requirements API, so the call should be attempted and its real
 * error surfaced instead.
 */
export function capabilityGapMessage(
  spec: CapabilitySpec,
  scope: Record<string, unknown> = GLOBAL_SCOPE,
): string {
  return (
    `${spec.label} needs ${spec.set} ${spec.minVersion}, which this host does not support ` +
    `(platform: ${platformLabel(currentPlatform(scope))}). ${spec.remedy}`
  );
}

/**
 * Fail-loud guard for a capability the caller cannot work around.
 *
 * Returns null when the host may attempt the call — supported, or an "unknown"
 * verdict (no requirements API, e.g. WPS) where the Office.js call surfaces its
 * own error. Returns the capability/platform/remedy message only when the host
 * is known to lack the set, so callers never silently no-op.
 */
export function requireCapability(
  id: string,
  scope: Record<string, unknown> = GLOBAL_SCOPE,
): string | null {
  const spec = capabilitySpec(id);
  if (spec === undefined) return null;
  return probeCapability(spec, scope) === "unsupported"
    ? capabilityGapMessage(spec, scope)
    : null;
}

/**
 * True only when the host is known to lack a capability.
 *
 * Callers that *have* a fallback path use this to choose between the native
 * call and the fallback. An "unknown" verdict returns false, so the native call
 * is attempted and its real error surfaces rather than being pre-empted.
 */
export function capabilityUnsupported(
  id: string,
  scope: Record<string, unknown> = GLOBAL_SCOPE,
): boolean {
  const spec = capabilitySpec(id);
  if (spec === undefined) return false;
  return probeCapability(spec, scope) === "unsupported";
}

/** A probed capability and its verdict. */
export interface CapabilityEntry {
  spec: CapabilitySpec;
  verdict: CapabilityVerdict;
}

/** Full capability snapshot for the current host. */
export interface CapabilityReport {
  platform: string | null;
  requirementsApiAvailable: boolean;
  entries: readonly CapabilityEntry[];
}

/** Probe every capability registered for a host. */
export function probeCapabilities(
  host: CapabilityHost,
  scope: Record<string, unknown> = GLOBAL_SCOPE,
): CapabilityReport {
  return {
    platform: currentPlatform(scope),
    requirementsApiAvailable: requirementsFrom(scope) !== null,
    entries: capabilitiesForHost(host).map((spec) => ({
      spec,
      verdict: probeCapability(spec, scope),
    })),
  };
}

/** Probe every Word capability (thin alias kept for existing callers). */
export function probeWordCapabilities(
  scope: Record<string, unknown> = GLOBAL_SCOPE,
): CapabilityReport {
  return probeCapabilities("word", scope);
}

/** Friendly platform names for the report. */
const PLATFORM_LABEL: Readonly<Record<string, string>> = {
  PC: "Windows desktop",
  Mac: "Mac desktop",
  OfficeOnline: "Office on the web",
  iOS: "iPad/iPhone",
  Android: "Android",
};

/** Describe the platform for humans, keeping the raw value. */
export function platformLabel(platform: string | null): string {
  if (platform === null) return "unknown";
  const label = PLATFORM_LABEL[platform];
  return label === undefined ? platform : `${label} (${platform})`;
}

/**
 * Render a capability report as compact text for a tool result.
 *
 * Lists what *is* available and, separately, what is not — with the exact
 * requirement set and the remedy — so the model can explain a limitation
 * precisely instead of asserting one.
 */
export function formatCapabilityReport(report: CapabilityReport): string {
  const lines: string[] = [
    "Word capabilities",
    `Platform: ${platformLabel(report.platform)}`,
  ];

  if (!report.requirementsApiAvailable) {
    lines.push(
      "Requirements API: unavailable on this host, so per-API support cannot be probed.",
      "",
      "Do not assume a capability is missing. Attempt the operation and report the host's",
      "own error message if it fails.",
    );
    return lines.join("\n");
  }

  lines.push("Requirements API: available", "", "Available on this host:");

  const supported = report.entries.filter((e) => e.verdict === "supported");
  const unsupported = report.entries.filter((e) => e.verdict === "unsupported");
  const unknown = report.entries.filter((e) => e.verdict === "unknown");

  if (supported.length === 0) lines.push("- (none probed)");
  for (const { spec } of supported) {
    lines.push(`- ${spec.label} — ${spec.set} ${spec.minVersion}`);
  }

  if (unsupported.length > 0) {
    lines.push("", "NOT available on this host:");
    for (const { spec } of unsupported) {
      lines.push(
        `- ${spec.label} — needs ${spec.set} ${spec.minVersion}. ${spec.remedy}`,
      );
    }
  }

  if (unknown.length > 0) {
    lines.push("", "Could not determine:");
    for (const { spec } of unknown) {
      lines.push(
        `- ${spec.label} — probing ${spec.set} ${spec.minVersion} failed`,
      );
    }
  }

  return lines.join("\n");
}
