import { normalizeHeader } from "./headerHelpers";

export type HeaderIndexes = {
  unitIdIndex?: number;
  unitTypeIndex?: number;
  frequencyIndex?: number;
  trpIndex?: number;
  maxPeakIndex?: number;
  graphIndex?: number;
  photoIndex?: number;
};

export function getHeaderIndexes(headerRow: unknown[]): HeaderIndexes | undefined {
  const normalizedHeaders = headerRow.map((cell) => normalizeHeader(cell));

  const findIndexByHeader = (matcher: (header: string) => boolean) => {
    const found = normalizedHeaders.findIndex(matcher);
    return found >= 0 ? found : undefined;
  };

  const indexes: HeaderIndexes = {
    unitIdIndex: findIndexByHeader(
      (h) => h.includes("unit id") || h === "unit" || h.includes("serial")
    ),
    unitTypeIndex: findIndexByHeader(
      (h) => h.includes("unit type") || h.includes("type")
    ),
    frequencyIndex: findIndexByHeader((h) => h.includes("frequency")),
    trpIndex: findIndexByHeader(
      (h) => h === "trp" || h.startsWith("trp ") || h.includes(" trp")
    ),
    maxPeakIndex: findIndexByHeader(
      (h) => h.includes("max peak") || h.includes("peak")
    ),
    graphIndex: findIndexByHeader(
      (h) => h.includes("3d graph") || h.includes("graph")
    ),
    photoIndex: findIndexByHeader(
      (h) => h.includes("photo") || h.includes("image")
    ),
  };

  const hasCoreHeader =
    indexes.unitIdIndex !== undefined &&
    (
      indexes.frequencyIndex !== undefined ||
      indexes.trpIndex !== undefined ||
      indexes.maxPeakIndex !== undefined
    );

  return hasCoreHeader ? indexes : undefined;
}