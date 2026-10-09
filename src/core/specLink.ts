/*
 * Spec link: association between a Submittal Review .docx and its source specification.
 *
 * At export we embed a small JSON payload (each item's source Article heading + the full
 * Article block text) into the review .docx as a Word Custom XML Data part. Custom XML
 * parts are round-tripped by Word, so when the review is reopened the add-in reads the
 * payload back from the document's own bytes (via getFileAsync + JSZip — no Office.js
 * custom-XML API, and no need for the original spec file) and can show the source section
 * for any item.
 */

/* global DOMParser, crypto, Blob */

import JSZip from "jszip";
import { SubmittalItem } from "./submittal";

const NS = "urn:swe:speclink";
const CT_PROPS = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const REL_CUSTOMXML =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const REL_CUSTOMXMLPROPS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";

export interface SpecLinkArticle {
  title: string; // "2.2 BALL VALVES"
  text: string; // full Article block
}
export interface SpecLinkItem {
  section: string;
  outline: string;
  itemName: string;
  artIndex: number; // index into articles[]
}
export interface SpecLinkPayload {
  version: number;
  specName: string;
  articles: SpecLinkArticle[];
  items: SpecLinkItem[];
}

/** Build the payload from extracted items, de-duplicating shared Article blocks. */
export function buildSpecLinkPayload(items: SubmittalItem[], specName: string): SpecLinkPayload {
  const articles: SpecLinkArticle[] = [];
  const indexByKey = new Map<string, number>();
  const outItems: SpecLinkItem[] = items.map((it) => {
    const title = it.specTitle || "";
    const text = it.specText || "";
    const key = title + "\u0000" + text;
    let idx = indexByKey.get(key);
    if (idx === undefined) {
      idx = articles.length;
      articles.push({ title, text });
      indexByKey.set(key, idx);
    }
    const itemName = it.qualifier ? `${it.name} ${it.qualifier}` : it.name;
    return { section: it.section, outline: it.outline, itemName, artIndex: idx };
  });
  return { version: 1, specName, articles, items: outItems };
}

/* ------------------------------- embedding ------------------------------- */

function xmlEscape(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function guid(): string {
  try {
    if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  } catch {
    // fall through
  }
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}

/** Add our Custom XML Data part (+ props, rels, content-type, document rel) to a zip. */
export async function addSpecLinkToZip(zip: JSZip, payload: SpecLinkPayload): Promise<void> {
  let n = 1;
  while (zip.file(`customXml/item${n}.xml`)) n += 1;

  const json = JSON.stringify(payload);
  const itemXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<specLink xmlns="${NS}">${xmlEscape(json)}</specLink>`;
  const propsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<ds:datastoreItem ds:itemID="{${guid()}}" ` +
    `xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">` +
    `<ds:schemaRefs><ds:schemaRef ds:uri="${NS}"/></ds:schemaRefs></ds:datastoreItem>`;
  const relsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL_CUSTOMXMLPROPS}" Target="itemProps${n}.xml"/></Relationships>`;

  zip.file(`customXml/item${n}.xml`, itemXml);
  zip.file(`customXml/itemProps${n}.xml`, propsXml);
  zip.file(`customXml/_rels/item${n}.xml.rels`, relsXml);

  await addContentTypeParts(zip, n);
  await addDocumentRel(zip, `../customXml/item${n}.xml`);
}

async function addContentTypeParts(zip: JSZip, n: number): Promise<void> {
  const file = zip.file("[Content_Types].xml");
  if (!file) throw new Error("Invalid .docx: missing [Content_Types].xml");
  let xml = await file.async("string");
  // Ensure a Default for the "xml" extension exists (customXml items use it).
  if (!/<Default[^>]*Extension="xml"/.test(xml)) {
    xml = xml.replace(
      /<Types([^>]*)>/,
      `<Types$1><Default Extension="xml" ContentType="application/xml"/>`
    );
  }
  const override = `<Override PartName="/customXml/itemProps${n}.xml" ContentType="${CT_PROPS}"/>`;
  if (!xml.includes(`/customXml/itemProps${n}.xml`)) {
    xml = xml.replace("</Types>", `${override}</Types>`);
  }
  zip.file("[Content_Types].xml", xml);
}

async function addDocumentRel(zip: JSZip, target: string): Promise<void> {
  const path = "word/_rels/document.xml.rels";
  const file = zip.file(path);
  if (!file) throw new Error("Invalid .docx: missing document.xml.rels");
  let xml = await file.async("string");
  // Unique relationship id not colliding with existing rIdN.
  let id = "rId1000";
  let k = 1000;
  while (xml.includes(`Id="${id}"`)) {
    k += 1;
    id = `rId${k}`;
  }
  const rel = `<Relationship Id="${id}" Type="${REL_CUSTOMXML}" Target="${target}"/>`;
  xml = xml.replace("</Relationships>", `${rel}</Relationships>`);
  zip.file(path, xml);
}

/** Add the spec link to an already-built .docx blob (used by the fallback builder). */
export async function embedSpecLinkInBlob(blob: Blob, payload: SpecLinkPayload): Promise<Blob> {
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  await addSpecLinkToZip(zip, payload);
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

/* -------------------------------- reading -------------------------------- */

/** Read the spec-link payload from a .docx's bytes, or null if none is embedded. */
export async function readSpecLink(
  data: ArrayBuffer | Uint8Array | Blob
): Promise<SpecLinkPayload | null> {
  const zip = await JSZip.loadAsync(data as ArrayBuffer);
  const names = Object.keys(zip.files).filter((f) => /^customXml\/item\d+\.xml$/.test(f));
  for (const name of names) {
    const xml = await zip.file(name)!.async("string");
    if (!xml.includes(NS) || !xml.includes("specLink")) continue;
    const dom = new DOMParser().parseFromString(xml, "application/xml");
    const root = dom.documentElement;
    const jsonText = root ? root.textContent || "" : "";
    try {
      return JSON.parse(jsonText) as SpecLinkPayload;
    } catch {
      return null;
    }
  }
  return null;
}
