import JSZip from "jszip";
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

  if (cell?.t === "e") {
    return cleanText(cell?.l?.Target);
  }

  const textValue = cleanText(cell?.v);
  if (textValue) return textValue;

  return cleanText(cell?.l?.Target);
}

function getAttributeByLocalName(element: Element, localName: string): string {
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.localName === localName) return attribute.value;
  }

  return "";
}

function getFirstChildByLocalName(
  parent: ParentNode,
  localName: string
): Element | null {
  return (
    Array.from(parent.childNodes).find(
      (node): node is Element =>
        node.nodeType === Node.ELEMENT_NODE &&
        (node as Element).localName === localName
    ) ?? null
  );
}

function getDescendantsByLocalName(
  parent: Document | Element,
  localName: string
): Element[] {
  return Array.from(parent.getElementsByTagName("*")).filter(
    (node): node is Element => node.localName === localName
  );
}

function resolveZipPath(basePath: string, relativePath: string): string {
  const baseParts = basePath.split("/").slice(0, -1);
  const relativeParts = relativePath.split("/");
  const resolved = [...baseParts];

  relativeParts.forEach((part) => {
    if (!part || part === ".") return;

    if (part === "..") {
      resolved.pop();
      return;
    }

    resolved.push(part);
  });

  return resolved.join("/");
}

function getImageMimeType(filePath: string): string {
  const extension = filePath.split(".").pop()?.toLowerCase();

  switch (extension) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    case "svg":
      return "image/svg+xml";
    default:
      return "application/octet-stream";
  }
}

function toBlobUrl(buffer: ArrayBuffer, mimeType: string): string {
  const blob = new Blob([buffer], { type: mimeType });
  return URL.createObjectURL(blob);
}

async function extractRichValueImages(
  workbookData: ArrayBuffer
): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(workbookData);
  const metadataFile = zip.file("xl/metadata.xml");
  const richValueFile = zip.file("xl/richData/rdrichvalue.xml");
  const richValueRelFile = zip.file("xl/richData/richValueRel.xml");
  const richValueRelsFile = zip.file("xl/richData/_rels/richValueRel.xml.rels");

  if (!metadataFile || !richValueFile || !richValueRelFile || !richValueRelsFile) {
    return new Map<string, string>();
  }

  const metadataDoc = new DOMParser().parseFromString(
    await metadataFile.async("string"),
    "application/xml"
  );
  const richValueDoc = new DOMParser().parseFromString(
    await richValueFile.async("string"),
    "application/xml"
  );
  const richValueRelDoc = new DOMParser().parseFromString(
    await richValueRelFile.async("string"),
    "application/xml"
  );
  const richValueRelsDoc = new DOMParser().parseFromString(
    await richValueRelsFile.async("string"),
    "application/xml"
  );

  const richValueRelationshipIds = getDescendantsByLocalName(richValueRelDoc, "rel").map(
    (element) => getAttributeByLocalName(element, "id")
  );
  const relationshipTargets = new Map<string, string>();

      getDescendantsByLocalName(richValueRelsDoc, "Relationship").forEach((element) => {
    const id = getAttributeByLocalName(element, "Id");
    const target = element.getAttribute("Target") ?? "";
    if (id && target) {
      relationshipTargets.set(
        id,
        resolveZipPath("xl/richData/richValueRel.xml", target)
      );
    }
  });

  const richValueImageByIndex = new Map<number, string>();
  const richValues = getDescendantsByLocalName(richValueDoc, "rv");

  for (const [index, richValue] of richValues.entries()) {
    const values = getDescendantsByLocalName(richValue, "v").map(
      (element) => Number(element.textContent ?? "")
    );
    const relationshipIndex = values[0];
    if (!Number.isFinite(relationshipIndex)) continue;

    const relationshipId = richValueRelationshipIds[relationshipIndex];
    const imagePath = relationshipTargets.get(relationshipId);
    if (!imagePath) continue;

    const imageFile = zip.file(imagePath);
    if (!imageFile) continue;

    const imageBuffer = await imageFile.async("arraybuffer");
    richValueImageByIndex.set(
      index,
      toBlobUrl(imageBuffer, getImageMimeType(imagePath))
    );
  }

  const vmToImage = new Map<string, string>();
  const futureMetadataEntries = getDescendantsByLocalName(metadataDoc, "bk");
  const valueMetadataEntries = getDescendantsByLocalName(metadataDoc, "valueMetadata")[0];
  const valueMetadataBooks = valueMetadataEntries
    ? Array.from(valueMetadataEntries.childNodes).filter(
        (node): node is Element =>
          node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === "bk"
      )
    : [];

  valueMetadataBooks.forEach((book, valueMetadataIndex) => {
    const rcElement = getDescendantsByLocalName(book, "rc")[0];
    const futureMetadataIndex = Number(rcElement?.getAttribute("v") ?? "");
    const futureMetadataBook = futureMetadataEntries[futureMetadataIndex];
    const rvbElement = futureMetadataBook
      ? getDescendantsByLocalName(futureMetadataBook, "rvb")[0]
      : undefined;
    const richValueIndex = Number(rvbElement?.getAttribute("i") ?? "");
    const imageSource = richValueImageByIndex.get(richValueIndex);

    if (imageSource) {
      vmToImage.set(String(valueMetadataIndex + 1), imageSource);
    }
  });

  return vmToImage;
}

async function extractSheetVmMappings(
  workbookData: ArrayBuffer,
  sheetNames: string[]
): Promise<Map<string, Map<string, string>>> {
  const zip = await JSZip.loadAsync(workbookData);
  const vmMappingsBySheet = new Map<string, Map<string, string>>();

  await Promise.all(
    sheetNames.map(async (sheetName, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${sheetIndex + 1}.xml`;
      const sheetFile = zip.file(sheetPath);
      if (!sheetFile) return;

      const sheetDoc = new DOMParser().parseFromString(
        await sheetFile.async("string"),
        "application/xml"
      );
      const cellMappings = new Map<string, string>();

      getDescendantsByLocalName(sheetDoc, "c").forEach((cell) => {
        const cellRef = cell.getAttribute("r") ?? "";
        const vmValue = cell.getAttribute("vm") ?? "";

        if (cellRef && vmValue) {
          cellMappings.set(cellRef, vmValue);
        }
      });

      if (cellMappings.size) {
        vmMappingsBySheet.set(sheetName, cellMappings);
      }
    })
  );

  return vmMappingsBySheet;
}

function getCellRichImage(
  rowIndex: number,
  columnIndex: number | undefined,
  richValueImages: Map<string, string> | undefined,
  vmMappings: Map<string, string> | undefined
): string {
  if (columnIndex == null || !richValueImages?.size || !vmMappings?.size) return "";

  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  const vmValue = vmMappings.get(cellAddress);
  if (!vmValue) return "";

  return richValueImages.get(vmValue) ?? "";
}

async function extractEmbeddedPhotos(
  workbookData: ArrayBuffer,
  sheetNames: string[]
): Promise<Map<string, Map<string, string>>> {
  const zip = await JSZip.loadAsync(workbookData);
  const imageCache = new Map<string, string>();
  const photosBySheet = new Map<string, Map<string, string>>();

  await Promise.all(
    sheetNames.map(async (sheetName, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${sheetIndex + 1}.xml`;
      const sheetFile = zip.file(sheetPath);
      if (!sheetFile) return;

      const sheetXml = await sheetFile.async("string");
      const sheetDoc = new DOMParser().parseFromString(sheetXml, "application/xml");
      const drawingElement = getDescendantsByLocalName(sheetDoc, "drawing")[0];
      if (!drawingElement) return;

      const drawingRelId = getAttributeByLocalName(drawingElement, "id");
      if (!drawingRelId) return;

      const sheetRelsPath = `xl/worksheets/_rels/sheet${sheetIndex + 1}.xml.rels`;
      const sheetRelsFile = zip.file(sheetRelsPath);
      if (!sheetRelsFile) return;

      const sheetRelsDoc = new DOMParser().parseFromString(
        await sheetRelsFile.async("string"),
        "application/xml"
      );
      const drawingTarget = getDescendantsByLocalName(
        sheetRelsDoc,
        "Relationship"
      ).find(
        (relationship) => getAttributeByLocalName(relationship, "Id") === drawingRelId
      )?.getAttribute("Target");

      if (!drawingTarget) return;

      const drawingPath = resolveZipPath(sheetPath, drawingTarget);
      const drawingFile = zip.file(drawingPath);
      if (!drawingFile) return;

      const drawingDoc = new DOMParser().parseFromString(
        await drawingFile.async("string"),
        "application/xml"
      );

      const drawingRelsPath = resolveZipPath(
        drawingPath,
        `./_rels/${drawingPath.split("/").pop()}.rels`
      );
      const drawingRelsFile = zip.file(drawingRelsPath);
      if (!drawingRelsFile) return;

      const drawingRelsDoc = new DOMParser().parseFromString(
        await drawingRelsFile.async("string"),
        "application/xml"
      );
      const drawingRelationships = getDescendantsByLocalName(
        drawingRelsDoc,
        "Relationship"
      );

      const cellPhotoMap = new Map<string, string>();
      const anchors = [
        ...getDescendantsByLocalName(drawingDoc, "twoCellAnchor"),
        ...getDescendantsByLocalName(drawingDoc, "oneCellAnchor"),
      ];

      for (const anchor of anchors) {
        const fromElement = getFirstChildByLocalName(anchor, "from");
        if (!fromElement) continue;

        const rowValue = Number(
          getFirstChildByLocalName(fromElement, "row")?.textContent ?? ""
        );
        const columnValue = Number(
          getFirstChildByLocalName(fromElement, "col")?.textContent ?? ""
        );

        if (!Number.isFinite(rowValue) || !Number.isFinite(columnValue)) continue;

        const blipElement = getDescendantsByLocalName(anchor, "blip")[0];
        const imageRelId = blipElement
          ? getAttributeByLocalName(blipElement, "embed")
          : "";
        if (!imageRelId) continue;

        const imageTarget = drawingRelationships.find(
          (relationship) => getAttributeByLocalName(relationship, "Id") === imageRelId
        )?.getAttribute("Target");

        if (!imageTarget) continue;

        const imagePath = resolveZipPath(drawingPath, imageTarget);
        let imageSource = imageCache.get(imagePath);

        if (!imageSource) {
          const imageFile = zip.file(imagePath);
          if (!imageFile) continue;

          const imageBuffer = await imageFile.async("arraybuffer");
          imageSource = toBlobUrl(imageBuffer, getImageMimeType(imagePath));
          imageCache.set(imagePath, imageSource);
        }

        cellPhotoMap.set(`${rowValue}:${columnValue}`, imageSource);
      }

      if (cellPhotoMap.size) {
        photosBySheet.set(sheetName, cellPhotoMap);
      }
    })
  );

  return photosBySheet;
}

function getEmbeddedPhotoForRow(
  embeddedPhotos: Map<string, string> | undefined,
  rowIndex: number,
  photoColumnIndex?: number
): string {
  if (!embeddedPhotos?.size) return "";

  if (photoColumnIndex != null) {
    const exactMatch = embeddedPhotos.get(`${rowIndex}:${photoColumnIndex}`);
    if (exactMatch) return exactMatch;
  }

  for (const [key, value] of embeddedPhotos) {
    const [photoRowIndex] = key.split(":").map(Number);
    if (photoRowIndex === rowIndex) return value;
  }

  return "";
}

function extractRowsFromMatrix(
  sheet: XLSX.WorkSheet,
  matrix: SheetMatrix,
  embeddedPhotos?: Map<string, string>,
  richValueImages?: Map<string, string>,
  vmMappings?: Map<string, string>
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
  const imagePreviewIndex = photoIndex ?? graphIndex;

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
    const photoValue =
      getCellRichImage(rowIndex, imagePreviewIndex, richValueImages, vmMappings) ||
      getCellTextOrLink(sheet, rowIndex, imagePreviewIndex) ||
      getEmbeddedPhotoForRow(embeddedPhotos, rowIndex, imagePreviewIndex);

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
            blankrows: false,
            defval: null,
          }) as SheetMatrix;

          allRows.push(
            ...extractRowsFromMatrix(
              sheet,
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
