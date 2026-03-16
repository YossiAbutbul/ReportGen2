export function toNumber(value: unknown): number | undefined {
  if (value == null || value === "") return undefined;

  const numericValue =
    typeof value === "number"
      ? value
      : Number(String(value).replace(/,/g, ""));

  return Number.isFinite(numericValue) ? numericValue : undefined;
}

export function cleanText(value: unknown): string {
  if (value == null) return "";
  return String(value).trim();
}

export function formatNumber(value?: number, digits = 3): string {
  if (value == null || Number.isNaN(value)) return "-";

  return Number(value)
    .toFixed(digits)
    .replace(/\.0+$/, "")
    .replace(/(\.\d*?)0+$/, "$1");
}

export function formatFrequency(value?: number): string {
  if (value == null) return "-";
  return `${formatNumber(value, 1)}MHz`;
}

export function todayAsDDMMYYYY(): string {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, "0");
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const yyyy = String(now.getFullYear());

  return `${dd}.${mm}.${yyyy}`;
}
