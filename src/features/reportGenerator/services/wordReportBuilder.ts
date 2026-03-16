import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  Packer,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlignTable,
  WidthType,
} from "docx";

import aradLogoUrl from "../../../assets/arad-logo-from-sample.jpeg";
import testSetupSvgUrl from "../../../assets/generic-trp-test-setup.svg";
import type { ReportBuildParams } from "../types";
import { formatFrequency, formatNumber } from "../utils/format";

const brandBlue = "002060";
const accentBlue = "D9E8F6";
const borderColor = "9FB6CE";
const lightBorderColor = "D7E1EC";
const testSetupWidthPx = 601;
const testSetupHeightPx = 389;

async function loadAssetBuffer(assetUrl: string): Promise<ArrayBuffer> {
  const response = await fetch(assetUrl);

  if (!response.ok) {
    throw new Error(`Failed to load report asset: ${assetUrl}`);
  }

  return response.arrayBuffer();
}

async function loadSvgImageOptions(assetUrl: string) {
  const svgResponse = await fetch(assetUrl);

  if (!svgResponse.ok) {
    throw new Error(`Failed to load report asset: ${assetUrl}`);
  }

  const svgText = await svgResponse.text();
  const svgData = new TextEncoder().encode(svgText);

  const svgBlob = new Blob([svgText], { type: "image/svg+xml" });
  const blobUrl = URL.createObjectURL(svgBlob);

  try {
    const image = await new Promise<HTMLImageElement>((resolve, reject) => {
      const nextImage = new Image();
      nextImage.onload = () => resolve(nextImage);
      nextImage.onerror = () =>
        reject(new Error(`Failed to decode SVG asset: ${assetUrl}`));
      nextImage.src = blobUrl;
    });

    const canvas = document.createElement("canvas");
    canvas.width = Math.max(image.naturalWidth, 1200);
    canvas.height = Math.max(image.naturalHeight, 700);

    const context = canvas.getContext("2d");
    if (!context) {
      throw new Error("Failed to create canvas for SVG fallback.");
    }

    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);

    const pngBlob = await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (blob) resolve(blob);
        else reject(new Error("Failed to create PNG fallback for SVG."));
      }, "image/png");
    });

    const pngData = await pngBlob.arrayBuffer();

    return {
      type: "svg" as const,
      data: svgData,
      fallback: {
        type: "png" as const,
        data: pngData,
      },
      width: canvas.width,
      height: canvas.height,
    };
  } finally {
    URL.revokeObjectURL(blobUrl);
  }
}

function paragraph(
  text: string,
  options?: {
    bold?: boolean;
    color?: string;
    heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel];
    spacingAfter?: number;
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
    size?: number;
    italics?: boolean;
    pageBreakBefore?: boolean;
  }
) {
  return new Paragraph({
    heading: options?.heading,
    alignment: options?.alignment,
    pageBreakBefore: options?.pageBreakBefore,
    spacing: { after: options?.spacingAfter ?? 140 },
    children: [
      new TextRun({
        text,
        bold: options?.bold,
        color: options?.color,
        size: options?.size,
        italics: options?.italics,
      }),
    ],
  });
}

function sectionHeading(
  text: string,
  level: (typeof HeadingLevel)[keyof typeof HeadingLevel] = HeadingLevel.HEADING_1,
  pageBreakBefore = false
) {
  return new Paragraph({
    heading: level,
    pageBreakBefore,
    spacing: { before: 120, after: 220 },
    children: [
      new TextRun({
        text,
        bold: true,
        color: brandBlue,
        size: level === HeadingLevel.HEADING_1 ? 32 : 26,
      }),
    ],
  });
}

function makeCell(
  text: string,
  options?: {
    header?: boolean;
    widthPct?: number;
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    verticalAlign?: (typeof VerticalAlignTable)[keyof typeof VerticalAlignTable];
  }
) {
  return new TableCell({
    width: options?.widthPct
      ? { size: options.widthPct, type: WidthType.PERCENTAGE }
      : undefined,
    verticalAlign: options?.verticalAlign ?? VerticalAlignTable.CENTER,
    shading: options?.header
      ? { fill: accentBlue, type: ShadingType.CLEAR, color: "auto" }
      : undefined,
    margins: {
      top: 100,
      bottom: 100,
      left: 120,
      right: 120,
    },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      left: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      right: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
    },
    children: text.split("\n").map(
      (line) =>
        new Paragraph({
          alignment: options?.align ?? AlignmentType.LEFT,
          spacing: { after: 0 },
          children: [
            new TextRun({
              text: line || " ",
              bold: !!options?.header,
              color: options?.header ? brandBlue : "000000",
              size: 21,
            }),
          ],
        })
    ),
  });
}

function makeLabelValueTable(rows: Array<[string, string]>) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      left: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      right: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      insideHorizontal: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: lightBorderColor,
      },
      insideVertical: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: lightBorderColor,
      },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          makeCell("Parameter", { header: true, widthPct: 35 }),
          makeCell("Value", { header: true, widthPct: 65 }),
        ],
      }),
      ...rows.map(
        ([label, value]) =>
          new TableRow({
            children: [makeCell(label), makeCell(value || "-")],
          })
      ),
    ],
  });
}

function makeResultsTable(params: ReportBuildParams) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      left: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      right: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      insideHorizontal: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: lightBorderColor,
      },
      insideVertical: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: lightBorderColor,
      },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          makeCell("Unit ID", {
            header: true,
            widthPct: 34,
            align: AlignmentType.CENTER,
          }),
          makeCell("Frequency", {
            header: true,
            widthPct: 18,
            align: AlignmentType.CENTER,
          }),
          makeCell("TRP", {
            header: true,
            widthPct: 16,
            align: AlignmentType.CENTER,
          }),
          makeCell("Max Peak", {
            header: true,
            widthPct: 16,
            align: AlignmentType.CENTER,
          }),
          makeCell("3D Graph", {
            header: true,
            widthPct: 16,
            align: AlignmentType.CENTER,
          }),
        ],
      }),
      ...params.rows.map(
        (row) =>
          new TableRow({
            children: [
              makeCell(row.unitId, { align: AlignmentType.CENTER }),
              makeCell(formatFrequency(row.frequencyMHz), {
                align: AlignmentType.CENTER,
              }),
              makeCell(formatNumber(row.trp), {
                align: AlignmentType.CENTER,
              }),
              makeCell(formatNumber(row.maxPeak), {
                align: AlignmentType.CENTER,
              }),
              makeCell(row.graphValue || "-", {
                align: AlignmentType.CENTER,
              }),
            ],
          })
      ),
    ],
  });
}

function createHeader(logoData: ArrayBuffer) {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 0 },
        children: [
          new ImageRun({
            type: "jpg",
            data: logoData,
            transformation: {
              width: 140,
              height: 55,
            },
          }),
        ],
      }),
    ],
  });
}

function createCoverHeader(logoData: ArrayBuffer) {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 0 },
        children: [
          new ImageRun({
            type: "jpg",
            data: logoData,
            transformation: {
              width: 140,
              height: 55,
            },
          }),
        ],
      }),
    ],
  });
}

function createFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
        children: [
          new TextRun({
            children: [PageNumber.CURRENT],
            color: "4472C4",
            allCaps: true,
            size: 20,
          }),
        ],
      }),
    ],
  });
}

export async function buildDocx(params: ReportBuildParams) {
  const logoData = await loadAssetBuffer(aradLogoUrl);
  const testSetupGraphic = await loadSvgImageOptions(testSetupSvgUrl);

  const frequencyLines = params.frequencies.length
    ? params.frequencies.map((value) => `LoRa ${formatFrequency(value)}`)
    : ["LoRa -"];

  const measurementParametersTable = makeLabelValueTable([
    ["Frequency", frequencyLines.join("\n")],
    ["Tested Power", params.testedPower || "-"],
  ]);

  const testParametersTable = makeLabelValueTable([
    ["FW Version", params.fwVersion || "-"],
    ["HW Version", params.hwVersion || "-"],
    ["Unit IDs", params.unitIds.length ? params.unitIds.join("\n") : "-"],
  ]);

  const resultsTable = makeResultsTable(params);
  const coverHeader = createCoverHeader(logoData);
  const defaultHeader = createHeader(logoData);
  const defaultFooter = createFooter();

  const doc = new Document({
    creator: params.author,
    title: params.title,
    description: "Generated test summary",
    styles: {
      default: {
        document: {
          run: {
            font: "Calibri",
            size: 22,
          },
          paragraph: {
            spacing: {
              line: 276,
            },
          },
        },
      },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            font: "Calibri",
            bold: true,
            color: brandBlue,
            size: 32,
          },
          paragraph: {
            spacing: {
              before: 120,
              after: 220,
            },
          },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            font: "Calibri",
            bold: true,
            color: brandBlue,
            size: 26,
          },
          paragraph: {
            spacing: {
              before: 120,
              after: 180,
            },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
              header: 706,
              footer: 706,
            },
          },
          titlePage: true,
        },
        headers: {
          first: coverHeader,
          default: defaultHeader,
        },
        footers: {
          first: new Footer({ children: [] }),
          default: defaultFooter,
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 220 },
            children: [
              new ImageRun({
                type: "jpg",
                data: logoData,
                transformation: {
                  width: 420,
                  height: 165,
                },
              }),
            ],
          }),
          paragraph("", { spacingAfter: 420 }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 260 },
            children: [
              new TextRun({
                text: params.title,
                bold: true,
                color: brandBlue,
                size: 160,
              }),
            ],
          }),
          paragraph(`By: ${params.author}`, {
            alignment: AlignmentType.CENTER,
            size: 28,
            spacingAfter: 120,
          }),
          paragraph(params.dateText, {
            alignment: AlignmentType.CENTER,
            size: 28,
            spacingAfter: 0,
          }),
          sectionHeading("Test Setup:", HeadingLevel.HEADING_1, true),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 260 },
            children: [
              new ImageRun({
                type: testSetupGraphic.type,
                data: testSetupGraphic.data,
                fallback: testSetupGraphic.fallback,
                transformation: {
                  width: testSetupWidthPx,
                  height: testSetupHeightPx,
                },
              }),
            ],
          }),
          sectionHeading("Scope of Testing", HeadingLevel.HEADING_1, true),
          ...params.scopeOfTesting
            .split(/\r?\n/)
            .map((line) => line.trim())
            .filter(Boolean)
            .map(
              (line, index) =>
                new Paragraph({
                  spacing: { after: 120 },
                  children: [
                    new TextRun({
                      text: `${index + 1}. ${line}`,
                      size: 24,
                    }),
                  ],
                })
            ),
          sectionHeading("Test Parameters:", HeadingLevel.HEADING_1),
          sectionHeading("Measurement Parameters:", HeadingLevel.HEADING_2),
          measurementParametersTable,
          paragraph("", { spacingAfter: 180 }),
          sectionHeading("Unit IDs:", HeadingLevel.HEADING_2),
          testParametersTable,
          sectionHeading("Radiated Results", HeadingLevel.HEADING_1, true),
          resultsTable,
          sectionHeading("Notes:", HeadingLevel.HEADING_1, true),
          paragraph(""),
          paragraph(""),
          paragraph(""),
        ],
      },
    ],
  });

  return Packer.toBlob(doc);
}
