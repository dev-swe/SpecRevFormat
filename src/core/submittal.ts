/*
 * Submittal item extraction.
 *
 * TypeScript port of the `spec_to_submittal.py` extractor. Reads a MasterSpec-style
 * section (.docx) entirely in the browser (JSZip — no server, no Python) and pulls the
 * PART 2 - PRODUCTS articles into submittal-review rows:
 *
 *     Spec Section        Item Name
 *     22 05 23            Ball Valves
 *     2.2 A               (NPS < 3)
 *
 * Office.js can only see the one open document, so batch/combine works by letting the
 * user pick several .docx files in the task pane; each is parsed here and the rows are
 * merged and sorted by `combineAndSort`.
 *
 * Parsing mirrors the Python docx parser: it classifies each paragraph by its CSI
 * SectionFormat paragraph style (SCT / PRT / ART / PR1…) and computes the outline
 * numbers (2.2 A) from the style sequence, because the printed A./1./a. numbers are
 * auto-generated list numbers, not literal text.
 */

/* global DOMParser, Element, Blob, File */

import JSZip from "jszip";

export interface SubmittalItem {
  section: string; // "22 05 23"
  outline: string; // "2.2 A"
  name: string; // "Ball Valves"
  qualifier: string; // "(NPS < 3)"  (may be "")
  source: string; // originating file name
}

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Articles in Part 2 that are not products (skipped unless allParts).
const NON_PRODUCT_ART = /GENERAL REQUIREMENTS/i;

// Acronyms kept uppercase when title-casing an article name.
const ACRONYMS = new Set([
  "NPS",
  "PVC",
  "CPVC",
  "HDPE",
  "PP",
  "PVDF",
  "PEX",
  "ABS",
  "AWWA",
  "ASME",
  "PTFE",
  "EPDM",
  "CWP",
  "SWP",
]);

/* ------------------------------- text helpers ------------------------------- */

// Built from char codes so the source carries no literal irregular whitespace or
// look-alike hyphens: U+2010/U+2011 hyphen variants and U+00A0 non-breaking space.
const ODD_HYPHENS = new RegExp("[" + String.fromCharCode(0x2010, 0x2011) + "]", "g");
const NBSP = new RegExp(String.fromCharCode(0x00a0), "g");

/** Normalize the non-breaking hyphens/spaces MasterSpec uses. */
export function clean(text: string): string {
  return text.replace(ODD_HYPHENS, "-").replace(NBSP, " ").trim();
}

/** Pull "22 05 23" from "SECTION 22 05 23 - ..." (hyphens/spaces vary), or any 6 digits. */
export function findSectionNumber(text: string): string {
  const m = /SECTION\s+(\d\d)\D?(\d\d)\D?(\d\d)/i.exec(text);
  if (m) return `${m[1]} ${m[2]} ${m[3]}`;
  const g = /(\d\d)\D?(\d\d)\D?(\d\d)/.exec(text);
  return g ? `${g[1]} ${g[2]} ${g[3]}` : "";
}

export function titlecaseName(name: string): string {
  return clean(name)
    .split(/\s+/)
    .map((w) => {
      const stripped = w.replace(/^[.,:;]+|[.,:;]+$/g, "");
      if (ACRONYMS.has(stripped.toUpperCase())) return w.toUpperCase();
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    })
    .join(" ");
}

/** "NPS 3 and Smaller" -> "(NPS < 3)"; "NPS 4 and Larger" -> "(NPS > 4)". */
export function makeQualifier(pr1Text: string): string {
  const t = clean(pr1Text)
    .replace(/:+\s*$/, "")
    .trim();
  let m = /^NPS\s+([0-9/\-.]+)\s+and\s+(Smaller|Larger)/i.exec(t);
  if (m) {
    const op = m[2].toLowerCase() === "smaller" ? "<" : ">";
    return `(NPS ${op} ${m[1]})`;
  }
  m = /^NPS\s+([0-9/\-.]+)\s+(?:to|through|-)\s*NPS\s+([0-9/\-.]+)/i.exec(t);
  if (m) return `(NPS ${m[1]}-${m[2]})`;
  return t ? `(${t})` : "";
}

export function looksLikeSize(pr1Text: string): boolean {
  return /^\s*NPS\b/i.test(clean(pr1Text));
}

/* ------------------------------- docx parsing ------------------------------- */

/** Build a styleId -> style name map from styles.xml (names fall back to the id). */
function styleIdToName(stylesXml: string | null): Map<string, string> {
  const map = new Map<string, string>();
  if (!stylesXml) return map;
  const dom = new DOMParser().parseFromString(stylesXml, "application/xml");
  const styles = dom.getElementsByTagNameNS(W_NS, "style");
  for (let i = 0; i < styles.length; i++) {
    const id = styles[i].getAttributeNS(W_NS, "styleId") || "";
    const nameEl = styles[i].getElementsByTagNameNS(W_NS, "name")[0];
    const name = nameEl ? nameEl.getAttributeNS(W_NS, "val") || "" : "";
    if (id) map.set(id, name || id);
  }
  return map;
}

/** Collapse a <w:p>'s runs into plain text. */
function paragraphText(p: Element): string {
  const parts: string[] = [];
  const texts = p.getElementsByTagNameNS(W_NS, "t");
  for (let i = 0; i < texts.length; i++) parts.push(texts[i].textContent || "");
  return parts.join("");
}

/** The paragraph's style name (resolved through styles.xml), uppercased, or "". */
function paragraphStyle(p: Element, idToName: Map<string, string>): string {
  const pPr = p.getElementsByTagNameNS(W_NS, "pPr")[0];
  if (!pPr) return "";
  const pStyle = pPr.getElementsByTagNameNS(W_NS, "pStyle")[0];
  if (!pStyle) return "";
  const id = pStyle.getAttributeNS(W_NS, "val") || "";
  return (idToName.get(id) || id).toUpperCase();
}

const LEVEL: Record<string, string | number> = {
  SCT: "SCT",
  PRT: "PRT",
  ART: "ART",
  PR1: 1,
  PR2: 2,
  PR3: 3,
  PR4: 4,
  PR5: 5,
};

/**
 * Read a .docx File/Blob/ArrayBuffer and return its extracted submittal rows.
 *
 * `fullSubheading` controls the Item Name's second line: when false (default) it is the
 * abbreviated qualifier ("NPS 3 and Smaller" -> "(NPS < 3)"); when true it is the complete
 * subheading text that follows the A./B. letter ("NPS 3 and Smaller"), verbatim.
 */
export async function extractFromDocx(
  data: ArrayBuffer | Uint8Array | Blob,
  sourceName: string,
  allParts: boolean,
  fullSubheading = false
): Promise<SubmittalItem[]> {
  const zip = await JSZip.loadAsync(data as ArrayBuffer);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error(`${sourceName}: not a Word .docx (no document.xml).`);
  const documentXml = await docFile.async("string");
  const stylesFile = zip.file("word/styles.xml");
  const stylesXml = stylesFile ? await stylesFile.async("string") : null;
  const idToName = styleIdToName(stylesXml);

  const dom = new DOMParser().parseFromString(documentXml, "application/xml");
  const body = dom.getElementsByTagNameNS(W_NS, "body")[0] || dom.documentElement;
  const paras = body.getElementsByTagNameNS(W_NS, "p");

  let section = "";
  let part = 0; // 1=GENERAL, 2=PRODUCTS, 3=EXECUTION
  let art = 0;
  let letter = 0; // PR1 counter within an article
  let curArtName = "";
  let curIsProduct = false;
  const items: SubmittalItem[] = [];

  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    const style = paragraphStyle(p, idToName);
    const level = LEVEL[style];
    const text = clean(paragraphText(p));
    if (!text) continue;

    if (style === "SCT" || (!section && text.toUpperCase().includes("SECTION"))) {
      section = findSectionNumber(text) || section;
    } else if (level === "PRT") {
      part += 1;
      art = 0;
    } else if (level === "ART") {
      art += 1;
      letter = 0;
      curArtName = titlecaseName(text);
      curIsProduct = part === 2 && !NON_PRODUCT_ART.test(text);
    } else if (level === 1) {
      letter += 1;
      if (!(allParts || curIsProduct)) continue;
      const outline = `${part}.${art} ${String.fromCharCode(64 + letter)}`;
      const qualifier = fullSubheading
        ? clean(text).replace(/:+\s*$/, "")
        : looksLikeSize(text)
          ? makeQualifier(text)
          : "";
      items.push({ section, outline, name: curArtName, qualifier, source: sourceName });
    }
  }

  if (!section) section = findSectionNumber(sourceName);
  for (const it of items) if (!it.section) it.section = section;
  return items;
}

/** Read a user-picked File (.docx) into submittal rows. */
export async function extractFromFile(
  file: File,
  allParts: boolean,
  fullSubheading = false
): Promise<SubmittalItem[]> {
  const lower = file.name.toLowerCase();
  if (!lower.endsWith(".docx")) {
    throw new Error(`${file.name}: only .docx is supported in the add-in (PDF needs pdf.js).`);
  }
  const buf = await file.arrayBuffer();
  return extractFromDocx(buf, file.name, allParts, fullSubheading);
}

/* ------------------------------- combine / sort ------------------------------- */

function sortKey(it: SubmittalItem): [number[], number, number, string] {
  const sect = (it.section.match(/\d+/g) || ["9999"]).map((n) => parseInt(n, 10));
  const m = /(\d+)\.(\d+)\s*([A-Z]?)/.exec(it.outline);
  const part = m ? parseInt(m[1], 10) : 9999;
  const art = m ? parseInt(m[2], 10) : 9999;
  const letter = m ? m[3] : "";
  return [sect, part, art, letter];
}

/** Merge rows from several specs and order by section, then part.article, then PR1 letter. */
export function combineAndSort(groups: SubmittalItem[][]): SubmittalItem[] {
  const all = ([] as SubmittalItem[]).concat(...groups);
  return all.sort((a, b) => {
    const ka = sortKey(a);
    const kb = sortKey(b);
    const la = ka[0];
    const lb = kb[0];
    for (let i = 0; i < Math.max(la.length, lb.length); i++) {
      const d = (la[i] ?? 0) - (lb[i] ?? 0);
      if (d) return d;
    }
    if (ka[1] !== kb[1]) return ka[1] - kb[1];
    if (ka[2] !== kb[2]) return ka[2] - kb[2];
    return ka[3] < kb[3] ? -1 : ka[3] > kb[3] ? 1 : 0;
  });
}
