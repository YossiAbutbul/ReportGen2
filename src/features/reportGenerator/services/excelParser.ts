import * as XLSX from "xlsx";

import {
  isBlankRow,
  normalizeHeader,
} from "./parsing/headerHelpers";

import {
  extractEmbeddedPhotos,
  extractRichValueImages,
  extractSheetVmMappings,
  getCellRichImage,
  getEmbeddedPhotoForRow,
} from "./parsing/imageHelpers";

import type { MatrixCell, ResultRow, SheetMatrix, SummaryData } from "../types";
import { cleanText, toNumber } from "../utils/format";
import { getHeaderIndexes, type HeaderIndexes } from "./parsing/sectionHelpers";


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

  if (cell?.t === "e") {
    return cleanText(cell?.l?.Target);
  }

  const textValue = cleanText(cell?.v);
  if (textValue) return textValue;

  return cleanText(cell?.l?.Target);
}

function getSectionUnitType(
  matrix: SheetMatrix,
  sectionStartIndex: number,
  headerRowIndex: number,
  fallbackValue: string
): string {
  for (let rowIndex = headerRowIndex - 1; rowIndex >= sectionStartIndex; rowIndex -= 1) {
    const row = matrix[rowIndex] ?? [];
    const values = row.map((cell) => cleanText(cell)).filter(Boolean);

    if (!values.length) continue;

    return values[0] || fallbackValue;
  }

  return fallbackValue;
}

function extractSectionRows(
  sheet: XLSX.WorkSheet,
  matrix: SheetMatrix,
  headerRowIndex: number,
  endRowIndexExclusive: number,
  headerIndexes: HeaderIndexes,
  unitType: string,
  embeddedPhotos?: Map<string, string>,
  richValueImages?: Map<string, string>,
  vmMappings?: Map<string, string>
): ResultRow[] {
  const imagePreviewIndex = headerIndexes.photoIndex ?? headerIndexes.graphIndex;
  const parsedRows: ResultRow[] = [];
  let currentUnitId = "";

  for (let rowIndex = headerRowIndex + 1; rowIndex < endRowIndexExclusive; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];

    if (getHeaderIndexes(row)) {
      currentUnitId = "";
      continue;
    }

    const unitIdCell = getCellTextOrLink(sheet, rowIndex, headerIndexes.unitIdIndex);
    const frequencyValue = toNumber(
      headerIndexes.frequencyIndex != null
        ? row[headerIndexes.frequencyIndex]
        : undefined
    );
    const trpValue = toNumber(
      headerIndexes.trpIndex != null ? row[headerIndexes.trpIndex] : undefined
    );
    const maxPeakValue = toNumber(
      headerIndexes.maxPeakIndex != null ? row[headerIndexes.maxPeakIndex] : undefined
    );
    const graphValue = getCellTextOrLink(sheet, rowIndex, headerIndexes.graphIndex);
    const photoValue =
      getCellRichImage(rowIndex, imagePreviewIndex, richValueImages, vmMappings) ||
      getCellTextOrLink(sheet, rowIndex, imagePreviewIndex) ||
      getEmbeddedPhotoForRow(embeddedPhotos, rowIndex, imagePreviewIndex);

    if (normalizeHeader(unitIdCell) === "unit id") {
      currentUnitId = "";
      continue;
    }

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
      unitType,
      frequencyMHz: frequencyValue,
      trp: trpValue,
      maxPeak: maxPeakValue,
      graphValue,
      photoValue,
    });
  }

  return parsedRows;
}

function extractRowsFromMatrix(
  sheet: XLSX.WorkSheet,
  sheetName: string,
  matrix: SheetMatrix,
  embeddedPhotos?: Map<string, string>,
  richValueImages?: Map<string, string>,
  vmMappings?: Map<string, string>
): ResultRow[] {
  if (!matrix.length) return [];

  const parsedRows: ResultRow[] = [];
  let rowIndex = 0;

  while (rowIndex < matrix.length) {
    while (rowIndex < matrix.length && isBlankRow(matrix[rowIndex])) {
      rowIndex += 1;
    }

    if (rowIndex >= matrix.length) break;

    const sectionStartIndex = rowIndex;
    let sectionEndIndexExclusive = rowIndex;

    while (
      sectionEndIndexExclusive < matrix.length &&
      !isBlankRow(matrix[sectionEndIndexExclusive])
    ) {
      sectionEndIndexExclusive += 1;
    }

    let headerRowIndex = -1;
    let headerIndexes: HeaderIndexes | null = null;

    for (let index = sectionStartIndex; index < sectionEndIndexExclusive; index += 1) {
      const nextHeaderIndexes = getHeaderIndexes(matrix[index] ?? []);
      if (nextHeaderIndexes) {
        headerRowIndex = index;
        headerIndexes = nextHeaderIndexes;
        break;
      }
    }

    if (headerRowIndex >= 0 && headerIndexes) {
      const unitType = getSectionUnitType(
        matrix,
        sectionStartIndex,
        headerRowIndex,
        sheetName
      );

      parsedRows.push(
        ...extractSectionRows(
          sheet,
          matrix,
          headerRowIndex,
          sectionEndIndexExclusive,
          headerIndexes,
          unitType,
          embeddedPhotos,
          richValueImages,
          vmMappings
        )
      );
    }

    rowIndex = sectionEndIndexExclusive + 1;
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

  const uniqueUnitTypes = Array.from(
    new Set(rows.map((row) => row.unitType).filter(Boolean))
  );

  return { rows, uniqueUnitIds, uniqueFrequencies, uniqueUnitTypes };
}

export function parseWorkbook(file: File): Promise<SummaryData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async (event) => {
      try {
        const data = event.target?.result;
        if (!data) throw new Error("Could not read the file contents.");

        const workbookData = data as ArrayBuffer;
        const workbook = XLSX.read(workbookData, { type: "array" });
        const embeddedPhotos = await extractEmbeddedPhotos(
          workbookData,
          workbook.SheetNames
        );
        const richValueImages = await extractRichValueImages(workbookData);
        const sheetVmMappings = await extractSheetVmMappings(
          workbookData,
          workbook.SheetNames
        );
        const allRows: ResultRow[] = [];

        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          const matrix = XLSX.utils.sheet_to_json<MatrixCell[]>(sheet, {
            header: 1,
            raw: true,
            blankrows: true,
            defval: null,
          }) as SheetMatrix;

          allRows.push(
            ...extractRowsFromMatrix(
              sheet,
              sheetName,
              matrix,
              embeddedPhotos.get(sheetName),
              richValueImages,
              sheetVmMappings.get(sheetName)
            )
          );
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
