/*
 * Revision Change Narrative (.docx) generator.
 *
 * Builds a standalone Word document from the editable revision summary and downloads it.
 * This never touches the specification document — it produces a new file only.
 */

/* global URL, document */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
} from "docx";
import { RevisionSummary } from "./revision";

const FOREST = "12413C";
const HEADER_TEXT = "FFFFFF";
const GRID = "D7DEE5";

function border() {
  return { style: BorderStyle.SINGLE, size: 4, color: GRID };
}

function cellBorders() {
  return { top: border(), bottom: border(), left: border(), right: border() };
}

function headerCell(text: string, widthPct: number): TableCell {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { fill: FOREST },
    borders: cellBorders(),
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: [
      new Paragraph({
        children: [new TextRun({ text, bold: true, color: HEADER_TEXT, size: 18 })],
      }),
    ],
  });
}

function bodyCell(text: string, widthPct: number): TableCell {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    borders: cellBorders(),
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: [new Paragraph({ children: [new TextRun({ text: text || "", size: 18 })] })],
  });
}

function metaLine(label: string, value: string): Paragraph {
  return new Paragraph({
    spacing: { after: 60 },
    children: [
      new TextRun({ text: label + ":  ", bold: true, size: 20 }),
      new TextRun({ text: value || "—", size: 20 }),
    ],
  });
}

function today(): string {
  const d = new Date();
  return d.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });
}

export async function exportNarrativeDocx(
  summary: RevisionSummary,
  docBase: string
): Promise<void> {
  const widths = [8, 22, 14, 34, 22]; // #, Location, Type, Description, Spec reference

  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell("#", widths[0]),
      headerCell("Location", widths[1]),
      headerCell("Type", widths[2]),
      headerCell("Description of Change", widths[3]),
      headerCell("Specification Reference", widths[4]),
    ],
  });

  const bodyRows = summary.rows.map(
    (r, i) =>
      new TableRow({
        children: [
          bodyCell(String(i + 1), widths[0]),
          bodyCell(r.location, widths[1]),
          bodyCell(r.type, widths[2]),
          bodyCell(r.description, widths[3]),
          bodyCell(r.specText, widths[4]),
        ],
      })
  );

  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...bodyRows],
  });

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({
                text: "Revision Change Narrative",
                bold: true,
                color: FOREST,
                size: 40,
              }),
            ],
          }),
          new Paragraph({ spacing: { after: 120 }, children: [] }),
          metaLine("Project", summary.project),
          metaLine("Section", summary.section),
          metaLine("Revision", summary.revision),
          metaLine("Date", summary.date || today()),
          metaLine("Prepared by", summary.preparedBy),
          new Paragraph({ spacing: { after: 160 }, children: [] }),
          table,
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const safeRev = (summary.revision || "revision")
    .replace(/[^A-Za-z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "");
  const a = document.createElement("a");
  a.href = url;
  a.download = `${docBase}-revision-change-narrative-${safeRev}.docx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
