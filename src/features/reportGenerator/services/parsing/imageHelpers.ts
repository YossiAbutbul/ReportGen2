import JSZip from "jszip";
import * as XLSX from "xlsx";
import {
  getAttributeByLocalName,
  getDescendantsByLocalName,
  getFirstChildByLocalName,
} from "../xml/xmlHelpers";
import {
  getImageMimeType,
  resolveZipPath,
  toBlobUrl,
} from "../xml/zipHelpers";

export async function extractRichValueImages(
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

export async function extractSheetVmMappings(
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

export function getCellRichImage(
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

export async function extractEmbeddedPhotos(
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

export function getEmbeddedPhotoForRow(
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