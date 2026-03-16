import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

import { TEMPLATE_SCOPE } from "../constants";
import type { ReportBuildParams } from "../types";
import { formatFrequency, formatNumber } from "../utils/format";

const headerFill = "EAF1FB";
const borderColor = "BFC7D5";

function paragraph(
  text: string,
  options?: {
    bold?: boolean;
    heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel];
    spacingAfter?: number;
  }
) {
  return new Paragraph({
    heading: options?.heading,
    spacing: { after: options?.spacingAfter ?? 140 },
    children: [new TextRun({ text, bold: options?.bold })],
  });
}

function makeCell(
  text: string,
  options?: {
    header?: boolean;
    widthPct?: number;
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
  }
) {
  return new TableCell({
    width: options?.widthPct
      ? { size: options.widthPct, type: WidthType.PERCENTAGE }
      : undefined,
    shading: options?.header
      ? { fill: headerFill, type: ShadingType.CLEAR, color: "auto" }
      : undefined,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      left: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
      right: { style: BorderStyle.SINGLE, size: 1, color: borderColor },
    },
    children: [
      new Paragraph({
        alignment: options?.align ?? AlignmentType.LEFT,
        children: [new TextRun({ text, bold: !!options?.header })],
      }),
    ],
  });
}

export async function buildDocx(params: ReportBuildParams) {
  const frequencyLines = params.frequencies.length
    ? params.frequencies.map((value) => `LoRa ${formatFrequency(value)}`)
    : ["LoRa -"];

  const testParametersTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          makeCell("Measurement Parameters", { header: true, widthPct: 50 }),
          makeCell("", { header: true, widthPct: 50 }),
        ],
      }),
      new TableRow({
        children: [makeCell("Frequency"), makeCell(frequencyLines.join("\n"))],
      }),
      new TableRow({
        children: [makeCell("Tested Power"), makeCell(params.testedPower || "-")],
      }),
    ],
  });

  const resultsTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          makeCell("Unit ID", { header: true, widthPct: 35 }),
          makeCell("Frequency", {
            header: true,
            widthPct: 20,
            align: AlignmentType.CENTER,
          }),
          makeCell("TRP", {
            header: true,
            widthPct: 15,
            align: AlignmentType.CENTER,
          }),
          makeCell("Max Peak", {
            header: true,
            widthPct: 15,
            align: AlignmentType.CENTER,
          }),
          makeCell("3D Graph", {
            header: true,
            widthPct: 15,
            align: AlignmentType.CENTER,
          }),
        ],
      }),
      ...params.rows.map(
        (row) =>
          new TableRow({
            children: [
              makeCell(row.unitId),
              makeCell(formatFrequency(row.frequencyMHz), {
                align: AlignmentType.CENTER,
              }),
              makeCell(formatNumber(row.trp), {
                align: AlignmentType.CENTER,
              }),
              makeCell(formatNumber(row.maxPeak), {
                align: AlignmentType.CENTER,
              }),
              makeCell(row.graphValue || "", {
                align: AlignmentType.CENTER,
              }),
            ],
          })
      ),
    ],
  });

  const doc = new Document({
    creator: params.author,
    title: params.title,
    description: "Generated test summary",
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
            spacing: { after: 220 },
            children: [new TextRun({ text: params.title, bold: true, size: 34 })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
            children: [new TextRun({ text: `By: ${params.author}`, italics: true })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 420 },
            children: [new TextRun({ text: params.dateText })],
          }),
          paragraph("Contents", {
            bold: true,
            heading: HeadingLevel.HEADING_1,
            spacingAfter: 180,
          }),
          paragraph("Test Setup"),
          paragraph("Scope of Testing"),
          paragraph("Test Parameters"),
          paragraph("Measurement Parameters"),
          paragraph("Unit IDs"),
          paragraph("Radiated Results", { spacingAfter: 260 }),
          paragraph("Test Setup", {
            bold: true,
            heading: HeadingLevel.HEADING_1,
            spacingAfter: 180,
          }),
          paragraph(""),
          paragraph("Scope of Testing", {
            bold: true,
            heading: HeadingLevel.HEADING_1,
            spacingAfter: 180,
          }),
          ...TEMPLATE_SCOPE.map(
            (item, index) =>
              new Paragraph({
                spacing: { after: 120 },
                children: [new TextRun({ text: `${index + 1}. ${item}` })],
              })
          ),
          paragraph("Test Parameters", {
            bold: true,
            heading: HeadingLevel.HEADING_1,
            spacingAfter: 180,
          }),
          testParametersTable,
          new Paragraph({ spacing: { after: 220 } }),
          new Paragraph({
            spacing: { after: 100 },
            children: [new TextRun({ text: "FW Version:", bold: true })],
          }),
          paragraph(params.fwVersion || "-"),
          new Paragraph({
            spacing: { after: 100 },
            children: [new TextRun({ text: "HW Version:", bold: true })],
          }),
          paragraph(params.hwVersion || "-"),
          new Paragraph({
            spacing: { after: 100 },
            children: [new TextRun({ text: "Unit IDs:", bold: true })],
          }),
          ...(params.unitIds.length
            ? params.unitIds.map(
                (id) =>
                  new Paragraph({
                    spacing: { after: 80 },
                    children: [new TextRun({ text: `• ${id}` })],
                  })
              )
            : [paragraph("-")]),
          new Paragraph({ spacing: { after: 220 } }),
          paragraph("Radiated Results", {
            bold: true,
            heading: HeadingLevel.HEADING_1,
            spacingAfter: 180,
          }),
          resultsTable,
        ],
      },
    ],
  });

  return Packer.toBlob(doc);
}
