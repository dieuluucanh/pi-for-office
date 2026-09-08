/**
 * format-builder — turns a preset name + params into an Excel format string.
 *
 * Generates the 3-section format: positive;negative;zero
 * Respects house-style conventions (parens, dash zeros, accounting padding).
 */

import type { NumberFormatConventions, NumberPreset } from "./types.js";
import {
  DEFAULT_CURRENCY_SYMBOL,
  DEFAULT_FORMAT_CONVENTIONS,
  PRESET_DEFAULT_DP,
} from "./format-defaults.js";

export interface FormatBuildResult {
  /** The Excel number format string. */
  format: string;
  /** Warnings generated during building. */
  warnings: string[];
}

/**
 * Build an Excel format string from a preset name + optional overrides.
 *
 * @param preset  - One of the 6 format presets.
 * @param dp      - Override decimal places (null = use preset default).
 * @param symbol  - Override currency symbol (only for "currency" preset).
 * @param conventions - House-style overrides (defaults to default format conventions).
 */
export function buildFormatString(
  preset: NumberPreset,
  dp?: number | null,
  symbol?: string | null,
  conventions: NumberFormatConventions = DEFAULT_FORMAT_CONVENTIONS,
): FormatBuildResult {
  const warnings: string[] = [];

  // ── Text: trivial ──────────────────────────────────────────────────
  if (preset === "text") {
    if (dp != null) {
      warnings.push("number_format_dp ignored — not applicable to text preset.");
    }
    if (symbol) {
      warnings.push("currency_symbol ignored — only applies to currency preset.");
    }
    return { format: "@", warnings };
  }

  // ── Type-check warnings ────────────────────────────────────────────
  if (preset === "integer" && dp != null && dp > 0) {
    warnings.push(
      `integer preset has 0dp by definition; use "number" with dp:${dp} instead.`,
    );
    // Still honour the request — build with the requested dp.
  }
  if (symbol && preset !== "currency") {
    warnings.push("currency_symbol ignored — only applies to currency preset.");
  }

  // ── Resolve dp ─────────────────────────────────────────────────────
  const defaultDp = PRESET_DEFAULT_DP[preset];
  const resolvedDp = dp ?? defaultDp ?? 0;

  // ── Build the core number pattern ──────────────────────────────────
  const intPart = conventions.thousandsSeparator ? "#,##0" : "0";
  const decPart = resolvedDp > 0 ? "." + "0".repeat(resolvedDp) : "";
  const core = intPart + decPart;

  // ── Suffix (percent / ratio) ───────────────────────────────────────
  const suffix = preset === "percent" ? "%" : preset === "ratio" ? "x" : "";

  // ── Padding chars (for zero alignment) ─────────────────────────────
  // Zero section needs _ padding for each trailing char so `--` aligns with numbers.
  const padChars = conventions.accountingPadding
    ? "_)" // always pad for the `)` that negatives have
    : "";
  const suffixPad = suffix
    ? `_${suffix}` // pad for the suffix char (% or x)
    : "";

  // ── Currency prefix ────────────────────────────────────────────────
  const currSymbol = preset === "currency" ? (symbol ?? DEFAULT_CURRENCY_SYMBOL) : "";
  const currPrefix = currSymbol ? `${currSymbol}* ` : "";

  // ── Assemble sections ──────────────────────────────────────────────
  const positive = `${currPrefix}${core}${suffix}${padChars}`;

  let negative: string;
  if (conventions.negativeStyle === "parens") {
    negative = `${currPrefix}(${core}${suffix})`;
  } else {
    negative = `${currPrefix}-${core}${suffix}${padChars}`;
  }

  let zero: string;
  if (conventions.zeroStyle === "dash") {
    zero = `${currPrefix}--${suffixPad}${padChars}`;
  } else if (conventions.zeroStyle === "single-dash") {
    zero = `${currPrefix}-${suffixPad}${padChars}`;
  } else if (conventions.zeroStyle === "blank") {
    zero = "";
  } else {
    // "zero" — show literal 0 with same format as positive
    zero = positive;
  }

  const format = `${positive};${negative};${zero}`;
  return { format, warnings };
}

/**
 * Check if a string is a known preset name.
 */
export function isPresetName(value: string): value is NumberPreset {
  return (
    value === "number" ||
    value === "integer" ||
    value === "currency" ||
    value === "percent" ||
    value === "ratio" ||
    value === "text"
  );
}
