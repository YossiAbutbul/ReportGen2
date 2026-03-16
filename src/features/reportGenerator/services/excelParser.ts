import * as XLSX from "xlsx";

import type { MatrixCell, ResultRow, SheetMatrix, SummaryData } from "../types";
import { cleanText, toNumber } from "../utils/format";

function isMeaningfulHeader(value: unknown): boolean {
  if (value == null) return false;

  const text = String(value).trim();
  if (!text) return false;

  return text.toLowerCase() !== "cell factor";
}

function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function getCellTextOrLink(
  sheet: XLSX.WorkSheet,
  rowIndex: number,
  columnIndex?: number
): string {
  if (columnIndex == null) return "";

  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  const cell = sheet[cellAddress] as
    | (XLSX.CellObject & { l?: { Target?: string } })
    | undefined;

  const textValue = cleanText(cell?.v);
  if (textValue) return textValue;

  return cleanText(cell?.l?.Target);
}

function extractRowsFromMatrix(
  sheet: XLSX.WorkSheet,
  matrix: SheetMatrix
): ResultRow[] {
  if (!matrix.length) return [];

  const headerRow = matrix[0] ?? [];
  const validIndexes = headerRow
    .map((cell, index) => ({ cell, index }))
    .filter(({ cell }) => isMeaningfulHeader(cell))
    .map(({ index }) => index);

  if (!validIndexes.length) return [];

  const normalizedHeaders = validIndexes.map((index) =>
    normalizeHeader(headerRow[index])
  );

  const findIndexByHeader = (matcher: (header: string) => boolean) => {
    const found = normalizedHeaders.findIndex(matcher);
    return found >= 0 ? validIndexes[found] : undefined;
  };

  const unitIdIndex = findIndexByHeader(
    (header) =>
      header.includes("unit id") ||
      header === "unit" ||
      header.includes("serial")
  );
  const frequencyIndex = findIndexByHeader((header) =>
    header.includes("frequency")
  );
  const trpIndex = findIndexByHeader(
    (header) =>
      header === "trp" ||
      header.startsWith("trp ") ||
      header.includes(" trp")
  );
  const maxPeakIndex = findIndexByHeader(
    (header) => header.includes("max peak") || header.includes("peak")
  );
  const graphIndex = findIndexByHeader(
    (header) => header.includes("3d graph") || header.includes("graph")
  );
  const photoIndex = findIndexByHeader(
    (header) =>
      header === "3d photo" ||
      header.startsWith("3d photo ") ||
      header.includes(" 3d photo")
  );

  const parsedRows: ResultRow[] = [];
  let currentUnitId = "";

  for (let rowIndex = 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];

    const unitIdCell = getCellTextOrLink(sheet, rowIndex, unitIdIndex);
    const frequencyValue = toNumber(
      frequencyIndex != null ? row[frequencyIndex] : undefined
    );
    const trpValue = toNumber(trpIndex != null ? row[trpIndex] : undefined);
    const maxPeakValue = toNumber(
      maxPeakIndex != null ? row[maxPeakIndex] : undefined
    );
    const graphValue = getCellTextOrLink(sheet, rowIndex, graphIndex);
    const photoValue = getCellTextOrLink(sheet, rowIndex, photoIndex);

    if (unitIdCell) currentUnitId = unitIdCell;

    const hasMeasurements =
      frequencyValue != null ||
      trpValue != null ||
      maxPeakValue != null ||
      Boolean(graphValue) ||
      Boolean(photoValue);

    if (!currentUnitId || !hasMeasurements) continue;

    const looksLikeCellFactorRow =
      trpValue == null &&
      maxPeakValue == null &&
      !graphValue &&
      !photoValue &&
      frequencyValue != null &&
      !unitIdCell;

    if (looksLikeCellFactorRow) continue;

    parsedRows.push({
      unitId: currentUnitId,
      frequencyMHz: frequencyValue,
      trp: trpValue,
      maxPeak: maxPeakValue,
      graphValue,
      photoValue,
    });
  }

  return parsedRows;
}

function summarizeRows(rows: ResultRow[]): SummaryData {
  const uniqueUnitIds = Array.from(
    new Set(rows.map((row) => row.unitId).filter(Boolean))
  );

  const uniqueFrequencies = Array.from(
    new Set(
      rows
        .map((row) => row.frequencyMHz)
        .filter((value): value is number => value != null)
    )
  ).sort((a, b) => a - b);

  return { rows, uniqueUnitIds, uniqueFrequencies };
}

export function parseWorkbook(file: File): Promise<SummaryData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = event.target?.result;
        if (!data) throw new Error("Could not read the file contents.");

        const workbook = XLSX.read(data, { type: "array" });
        const allRows: ResultRow[] = [];

        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          const matrix = XLSX.utils.sheet_to_json<MatrixCell[]>(sheet, {
            header: 1,
            raw: true,
            blankrows: false,
            defval: null,
          }) as SheetMatrix;

          allRows.push(...extractRowsFromMatrix(sheet, matrix));
        });

        resolve(summarizeRows(allRows));
      } catch (error) {
        reject(
          error instanceof Error
            ? error
            : new Error("Failed to parse workbook.")
        );
      }
    };

    reader.onerror = () => reject(new Error("Failed to read the Excel file."));
    reader.readAsArrayBuffer(file);
  });
}
