/*
 * Template-based Submittal Review generator.
 *
 * To match the firm's house style exactly (Tahoma default, Heading 3 title, the cover
 * disclaimer, Heading 5 field lines, the bordered ACTION CODES table, forest header
 * shading), this does NOT rebuild the document from scratch. Instead it loads the real
 * template `assets/submittal-template.docx` (bundled with the add-in), edits it in the
 * browser with JSZip, and downloads the result — the same template-injection approach
 * the Python `spec_to_submittal.py` used.
 *
 * The template holds the full cover plus the review table with its header row and ONE
 * fully-styled prototype data row. For each extracted item we clone that row, fill the
 * Last Submit / Spec Section / Item Name cells (leaving Item No. to auto-number and
 * Action Code / Comments blank for the reviewer), then remove the prototype.
 */

/* global fetch, DOMParser, XMLSerializer, Blob, URL, document, Element, Document */

import JSZip from "jszip";
import { SubmittalItem } from "./submittal";
import { SubmittalMeta } from "./submittalDocx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

/** Default location of the bundled template, relative to taskpane.html. */
export const TEMPLATE_URL = "assets/submittal-template.docx";

function wEl(doc: Document, name: string): Element {
  return doc.createElementNS(W, "w:" + name);
}

function textNode(doc: Document, text: string): Element {
  const t = wEl(doc, "t");
  t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.textContent = text;
  return t;
}

/** The concatenated <w:t> text of a paragraph or cell. */
function elementText(el: Element): string {
  const ts = el.getElementsByTagNameNS(W, "t");
  let s = "";
  for (let i = 0; i < ts.length; i++) s += ts[i].textContent || "";
  return s;
}

/** Next sibling that is an element node (avoids relying on nextElementSibling). */
function nextElement(el: Element): Element | null {
  let n = el.nextSibling;
  while (n && n.nodeType !== 1) n = n.nextSibling;
  return (n as Element) || null;
}

/** First child element with the given local name, or null. */
function firstChild(el: Element, local: string): Element | null {
  for (let i = 0; i < el.childNodes.length; i++) {
    const n = el.childNodes[i] as Element;
    if (n.nodeType === 1 && n.localName === local) return n;
  }
  return null;
}

/**
 * Replace a cell paragraph's runs with the given lines, reusing the prototype run's
 * <w:rPr> so font/bold/size are preserved. An empty `lines` clears the cell.
 */
function setCellLines(doc: Document, cell: Element, lines: string[]): void {
  const p = cell.getElementsByTagNameNS(W, "p")[0];
  if (!p) return;
  const runs = p.getElementsByTagNameNS(W, "r");
  const protoRpr = runs.length ? firstChild(runs[0] as Element, "rPr") : null;
  // remove existing runs (keep <w:pPr>)
  for (let i = runs.length - 1; i >= 0; i--) {
    const r = runs[i] as Element;
    r.parentNode?.removeChild(r);
  }
  const nonEmpty = lines.filter((l) => l !== undefined && l !== null);
  nonEmpty.forEach((line, i) => {
    const r = wEl(doc, "r");
    if (protoRpr) r.appendChild(protoRpr.cloneNode(true));
    if (i > 0) r.appendChild(wEl(doc, "br"));
    r.appendChild(textNode(doc, line));
    p.appendChild(r);
  });
}

/** Paragraph text including tab/break whitespace (runs walked in order). */
function fullText(p: Element): string {
  const runs = p.getElementsByTagNameNS(W, "r");
  let s = "";
  for (let i = 0; i < runs.length; i++) {
    const r = runs[i] as Element;
    for (let j = 0; j < r.childNodes.length; j++) {
      const n = r.childNodes[j] as Element;
      if (n.nodeType !== 1) continue;
      if (n.localName === "t") s += n.textContent || "";
      else if (n.localName === "tab") s += "\t";
      else if (n.localName === "br") s += "\n";
    }
  }
  return s;
}

/** Set a Heading-5 field paragraph's value (after its label/tab). */
function setFieldValue(doc: Document, p: Element, value: string): void {
  const existing = fullText(p);
  const needsSpace = existing.length > 0 && !/[\t \n]$/.test(existing);
  const r = wEl(doc, "r");
  r.appendChild(textNode(doc, (needsSpace ? " " : "") + value));
  p.appendChild(r);
}

function fillFields(doc: Document, meta: SubmittalMeta): void {
  const fields: [RegExp, string][] = [
    [/^Project Name/i, meta.project],
    [/^SWE Project No/i, meta.sweProjectNo],
    [/^Review Date/i, meta.reviewDate || today()],
    [/^Reviewed By/i, meta.reviewedBy],
    [/^SWE Submittal No/i, meta.sweSubmittalNo],
    [/^Contractor Submittal No/i, meta.contractorSubmittalNo],
  ];
  const paras = doc.getElementsByTagNameNS(W, "p");
  for (let i = 0; i < paras.length; i++) {
    const p = paras[i] as Element;
    const txt = elementText(p).trim();
    for (const [re, val] of fields) {
      if (val && re.test(txt)) {
        setFieldValue(doc, p, val);
        break;
      }
    }
  }
}

function today(): string {
  return new Date().toLocaleDateString(undefined, {
    year: "numeric",
    month: "numeric",
    day: "numeric",
  });
}

/** Locate the data table's header <w:tr> (the one containing "Spec Section"). */
function findHeaderRow(doc: Document): Element | null {
  const rows = doc.getElementsByTagNameNS(W, "tr");
  for (let i = 0; i < rows.length; i++) {
    if (elementText(rows[i] as Element).includes("Spec Section")) return rows[i] as Element;
  }
  return null;
}

function fillTable(doc: Document, items: SubmittalItem[], meta: SubmittalMeta): void {
  const header = findHeaderRow(doc);
  if (!header)
    throw new Error("Template is missing the submittal table (no 'Spec Section' header).");
  const proto = nextElement(header);
  if (!proto || proto.localName !== "tr") {
    throw new Error("Template is missing its prototype data row.");
  }
  const table = header.parentNode as Element;

  for (const it of items) {
    const row = proto.cloneNode(true) as Element;
    const cells = row.getElementsByTagNameNS(W, "tc");
    // 0 Item No. (auto-numbered — leave as-is), 1 Action Code (clear),
    // 2 Last Submit, 3 Spec Section, 4 Item Name, 5 Comments (clear).
    if (cells[1]) setCellLines(doc, cells[1] as Element, []);
    if (cells[2])
      setCellLines(doc, cells[2] as Element, meta.sweSubmittalNo ? [meta.sweSubmittalNo] : []);
    if (cells[3])
      setCellLines(doc, cells[3] as Element, it.outline ? [it.section, it.outline] : [it.section]);
    if (cells[4])
      setCellLines(doc, cells[4] as Element, it.qualifier ? [it.name, it.qualifier] : [it.name]);
    if (cells[5]) setCellLines(doc, cells[5] as Element, []);
    table.insertBefore(row, proto);
  }
  table.removeChild(proto); // drop the prototype
}

/** Fetch the bundled template bytes. */
export async function fetchTemplate(url: string = TEMPLATE_URL): Promise<ArrayBuffer> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Could not load the submittal template (${res.status}).`);
  return res.arrayBuffer();
}

/** Build the Submittal Review by injecting rows into the template bytes. */
export async function buildFromTemplate(
  templateBytes: ArrayBuffer,
  items: SubmittalItem[],
  meta: SubmittalMeta
): Promise<Blob> {
  const zip = await JSZip.loadAsync(templateBytes);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error("Template is not a valid .docx (no document.xml).");
  const xml = await docFile.async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");

  fillFields(doc, meta);
  fillTable(doc, items, meta);

  const out = new XMLSerializer().serializeToString(doc);
  zip.file("word/document.xml", out);
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

/** Fetch the template, inject rows, and download the result. Returns the file name. */
export async function exportSubmittalFromTemplate(
  items: SubmittalItem[],
  meta: SubmittalMeta,
  fileBase: string,
  templateUrl: string = TEMPLATE_URL
): Promise<string> {
  const bytes = await fetchTemplate(templateUrl);
  const blob = await buildFromTemplate(bytes, items, meta);
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
