import type { NumberFormatConventions, NumberPreset } from "./types.js";

export const DEFAULT_FORMAT_CONVENTIONS: NumberFormatConventions = {
  negativeStyle: "parens",
  thousandsSeparator: true,
  zeroStyle: "dash",
  accountingPadding: true,
};

export const DEFAULT_CURRENCY_SYMBOL = "$";

export const PRESET_DEFAULT_DP: Record<NumberPreset, number | null> = {
  number: 2,
  integer: 0,
  currency: 2,
  percent: 1,
  ratio: 1,
  text: null,
};
