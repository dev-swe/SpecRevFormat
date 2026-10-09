/*
 * Submittal Review (.docx) generator.
 *
 * Builds a standalone "Submittal Review" Word document from extracted SubmittalItems and
 * downloads it — the add-in equivalent of what `spec_to_submittal.py` wrote from the
 * firm's template. Mirrors the layout of "Submittal Review - Valves.docx":
 *   - a header block (Project, SWE Project No., Review Date, Reviewed By, submittal nos.)
 *   - an ACTION CODES legend
 *   - the 6-column review table (Item No. / Action Code / Last Submit / Spec Section /
 *     Item Name / Comments) with bold 10 pt data rows and two-line Spec Section / Item
 *     Name cells.
 *
 * Uses the `docx` library + a blob download, matching core/docx.ts (the Revision
 * Change Narrative generator). It never touches any open document.
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
  AlignmentType,
  BorderStyle,
  HeadingLevel,
} from "docx";
import { SubmittalItem } from "./submittal";
import { SpecLinkPayload, embedSpecLinkInBlob } from "./specLink";

const FOREST = "12413C";
const GRID = "D7DEE5";

/** Header fields for the Submittal Review cover block (all optional). */
export interface SubmittalMeta {
  project: string;
  sweProjectNo: string;
  reviewDate: string; // blank -> today
  reviewedBy: string;
  sweSubmittalNo: string; // also used as the "Last Submit" column value
  contractorSubmittalNo: string;
}

// Column widths from the firm template, in inches -> DXA (1 in = 1440 dxa).
const COL_IN = [0.54, 0.7, 0.76, 0.81, 1.34, 2.34];
const COL_DXA = COL_IN.map((w) => Math.round(w * 1440));
const DATA_PT = 20; // 10 pt, in half-points

function border() {
  return { style: BorderStyle.SINGLE, size: 4, color: GRID };
}
function cellBorders() {
  return { top: border(), bottom: border(), left: border(), right: border() };
}

function headerCell(text: string, dxa: number): TableCell {
  return new TableCell({
    width: { size: dxa, type: WidthType.DXA },
    shading: { fill: FOREST },
    borders: cellBorders(),
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [
      new Paragraph({
        children: [new TextRun({ text, bold: true, color: "FFFFFF", size: DATA_PT })],
      }),
    ],
  });
}

/** A data cell whose text may contain "\n" for a second (line-broken) line. */
function dataCell(text: string, dxa: number): TableCell {
  const lines = (text || "").split("\n");
  const runs: TextRun[] = [];
  lines.forEach((line, i) => {
    runs.push(new TextRun({ text: line, bold: true, size: DATA_PT, break: i ? 1 : 0 }));
  });
  return new TableCell({
    width: { size: dxa, type: WidthType.DXA },
    borders: cellBorders(),
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({ children: runs })],
  });
}

function metaLine(label: string, value: string): Paragraph {
  return new Paragraph({
    spacing: { after: 40 },
    children: [
      new TextRun({ text: label + "\t", bold: true, size: DATA_PT }),
      new TextRun({ text: value || "", size: DATA_PT }),
    ],
  });
}

const ACTION_CODES: [string, string][] = [
  ["A", "No Exceptions Taken"],
  ["B", "Make Corrections Noted*"],
  ["C", "Revise and Resubmit"],
  ["D", "Submittal Not Received"],
  ["E", "Incomplete, Not Reviewed"],
  ["F", "Rejected"],
];

function today(): string {
  return new Date().toLocaleDateString(undefined, {
    year: "numeric",
    month: "numeric",
    day: "numeric",
  });
}

function buildDoc(items: SubmittalItem[], meta: SubmittalMeta): Document {
  const headers = [
    "Item No.",
    "Action Code",
    "Last Submit",
    "Spec Section",
    "Item Name",
    "Comments",
  ];
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => headerCell(h, COL_DXA[i])),
  });

  const bodyRows = items.map((it) => {
    const specCell = it.section + (it.outline ? "\n" + it.outline : "");
    const nameCell = it.name + (it.qualifier ? "\n" + it.qualifier : "");
    const values = ["", "", meta.sweSubmittalNo || "", specCell, nameCell, ""];
    return new TableRow({ children: values.map((v, i) => dataCell(v, COL_DXA[i])) });
  });

  const table = new Table({ rows: [headerRow, ...bodyRows] });

  // ACTION CODES legend as a simple labeled list.
  const legend: Paragraph[] = [
    new Paragraph({
      spacing: { before: 120, after: 40 },
      children: [new TextRun({ text: "ACTION CODES", bold: true, size: DATA_PT })],
    }),
    ...ACTION_CODES.map(
      ([code, desc]) =>
        new Paragraph({
          spacing: { after: 20 },
          children: [
            new TextRun({ text: code + "\t", bold: true, size: DATA_PT }),
            new TextRun({ text: desc, size: DATA_PT }),
          ],
        })
    ),
    new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({
          text: "*No exceptions taken and a resubmittal is not required if correction notes are applied.",
          italics: true,
          size: 18,
        }),
      ],
    }),
  ];

  return new Document({
    sections: [
      {
        children: [
          new Paragraph({
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({ text: "SUBMITTAL REVIEW", bold: true, color: FOREST, size: 36 }),
            ],
          }),
          new Paragraph({ spacing: { after: 120 }, children: [] }),
          metaLine("Project Name:", meta.project),
          metaLine("SWE Project No.:", meta.sweProjectNo),
          metaLine("Review Date:", meta.reviewDate || today()),
          metaLine("Reviewed By:", meta.reviewedBy),
          metaLine("SWE Submittal No.:", meta.sweSubmittalNo),
          metaLine("Contractor Submittal No.:", meta.contractorSubmittalNo),
          ...legend,
          table,
        ],
      },
    ],
  });
}

/** Build and download the Submittal Review .docx. Returns the file name used. */
export async function exportSubmittalDocx(
  items: SubmittalItem[],
  meta: SubmittalMeta,
  fileBase: string,
  specLink?: SpecLinkPayload
): Promise<string> {
  let blob = await Packer.toBlob(buildDoc(items, meta));
  if (specLink) blob = await embedSpecLinkInBlob(blob, specLink);
  const url = URL.createObjectURL(blob);
  const safe = (fileBase || "Submittal Review").replace(/[^A-Za-z0-9 ._-]+/g, "-").trim();
  const fileName = `${safe}.docx`;
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return fileName;
}
