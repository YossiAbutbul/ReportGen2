import ExcelJS from "exceljs";

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
  sheet: ExcelJS.Worksheet,
  rowIndex: number,
  columnIndex?: number
): string {
  if (columnIndex == null) return "";

  const cell = sheet.getRow(rowIndex + 1).getCell(columnIndex + 1);
  const cellValue = cell.value;

  if (cellValue && typeof cellValue === "object") {
    if ("hyperlink" in cellValue) {
      return cleanText(cellValue.text || cellValue.hyperlink);
    }

    if ("text" in cellValue) {
      return cleanText(cellValue.text);
    }

    if ("result" in cellValue) {
      return cleanText(cellValue.result);
    }

    if ("error" in cellValue) {
      return "";
    }
  }

  const textValue = cleanText(cell.text);
  if (textValue) return textValue;

  return cleanText(cellValue);
}

function worksheetToMatrix(sheet: ExcelJS.Worksheet): SheetMatrix {
  const matrix: SheetMatrix = [];
  let maxColumnCount = 0;

  sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const actualCellCount = row.cellCount;
    if (actualCellCount > maxColumnCount) {
      maxColumnCount = actualCellCount;
    }

    const normalizedRow: MatrixCell[] = [];

    for (let columnIndex = 1; columnIndex <= actualCellCount; columnIndex += 1) {
      const cellValue = row.getCell(columnIndex).value;
      normalizedRow.push(cellValue == null ? null : (cellValue as MatrixCell));
    }

    matrix[rowNumber - 1] = normalizedRow;
  });

  for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    if (row.length < maxColumnCount) {
      matrix[rowIndex] = [...row, ...Array(maxColumnCount - row.length).fill(null)];
    }
  }

  return matrix;
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
  sheet: ExcelJS.Worksheet,
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
  sheet: ExcelJS.Worksheet,
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
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(workbookData);

        const sheetNames = workbook.worksheets.map((sheet) => sheet.name);

        const embeddedPhotos = await extractEmbeddedPhotos(workbookData, sheetNames);
        const richValueImages = await extractRichValueImages(workbookData);
        const sheetVmMappings = await extractSheetVmMappings(workbookData, sheetNames);

        const allRows: ResultRow[] = [];

        workbook.worksheets.forEach((sheet) => {
          const matrix = worksheetToMatrix(sheet);

          allRows.push(
            ...extractRowsFromMatrix(
              sheet,
              sheet.name,
              matrix,
              embeddedPhotos.get(sheet.name),
              richValueImages,
              sheetVmMappings.get(sheet.name)
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