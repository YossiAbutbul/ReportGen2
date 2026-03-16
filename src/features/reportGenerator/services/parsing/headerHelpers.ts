export function isMeaningfulHeader(value: unknown): boolean {
  if (value == null) return false;
  const text = String(value).trim();
  if (!text) return false;

  const lower = text.toLowerCase();
  return lower !== "cell factor";
}

export function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

export function isBlankRow(row: unknown[]): boolean {
  return row.every((cell) => {
    if (cell == null) return true;
    return String(cell).trim() === "";
  });
}