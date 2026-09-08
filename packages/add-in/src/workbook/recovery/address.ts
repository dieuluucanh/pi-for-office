/** Address helpers shared by recovery modules. */

export function localAddressPart(address: string): string {
  const trimmed = address.trim();
  const separatorIndex = trimmed.lastIndexOf("!");
  if (separatorIndex < 0) {
    return trimmed;
  }

  return trimmed.slice(separatorIndex + 1);
}

export function quoteSheetName(sheetName: string): string {
  const escaped = sheetName.replace(/'/g, "''");
  const needsQuote = /[\s'!]/.test(sheetName);
  return needsQuote ? `'${escaped}'` : sheetName;
}

export function qualifyAddressWithSheet(sheetName: string, address: string): string {
  const local = localAddressPart(address);
  return `${quoteSheetName(sheetName)}!${local}`;
}

export function firstCellAddress(address: string): string {
  const local = localAddressPart(address);
  const firstArea = local.split(",")[0] ?? local;
  const first = firstArea.split(":")[0] ?? firstArea;
  return first.trim();
}

export function splitRangeList(range: string): string[] {
  return range
    .split(/[;,]/)
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
}
