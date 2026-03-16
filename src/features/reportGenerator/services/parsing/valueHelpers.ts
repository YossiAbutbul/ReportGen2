export function toNumber(value: unknown): number | undefined {
  if (value == null || value === "") return undefined;

  const num =
    typeof value === "number"
      ? value
      : Number(String(value).replace(/,/g, "").trim());

  return Number.isFinite(num) ? num : undefined;
}

export function cleanText(value: unknown): string {
  if (value == null) return "";
  return String(value).trim();
}