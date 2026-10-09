/*
 * Highlight the extracted product elements in the open document.
 *
 * The preview (core/submittal.ts) reads the document's bytes, but to mark the source we
 * walk the live document with Office.js and apply the SAME classification: within
 * PART 2 - PRODUCTS, each product Article heading (e.g. "BALL VALVES") and each of its
 * PR1 size subheadings (e.g. "NPS 3 and Smaller") — the two lines that become a table
 * row — get a highlight. `clearProductHighlights` walks the same set and removes it.
 *
 * Classification helpers are shared with the byte parser so the highlighted elements
 * match the previewed rows.
 */

/* global Word */

import { clean } from "./submittal";

// Highlight applied to matched paragraphs (Word highlight color name or #RRGGBB).
const HIGHLIGHT = "Yellow";

// Mirrors core/submittal.ts.
const NON_PRODUCT_ART = /GENERAL REQUIREMENTS/i;
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

function setHighlight(p: Word.Paragraph, apply: boolean): void {
  p.font.set({ highlightColor: apply ? HIGHLIGHT : null });
}

/**
 * Walk the open document and apply (or clear) the highlight on the same product
 * Article headings + PR1 size lines the extractor pulls. Returns how many paragraphs
 * were marked.
 */
async function markProductItems(allParts: boolean, apply: boolean): Promise<number> {
  let count = 0;
  await Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load("text,style");
    await context.sync();

    let part = 0;
    let curIsProduct = false;
    let curArt: Word.Paragraph | null = null;
    let curArtMarked = false;

    for (const p of paras.items) {
      const text = clean(p.text || "");
      if (!text) continue;
      const style = (p.style || "").toUpperCase();
      const level = LEVEL[style];

      if (level === "PRT") {
        part += 1;
      } else if (level === "ART") {
        curIsProduct = part === 2 && !NON_PRODUCT_ART.test(text);
        curArt = p;
        curArtMarked = false;
      } else if (level === 1) {
        if (!(allParts || curIsProduct)) continue;
        if (curArt && !curArtMarked) {
          setHighlight(curArt, apply);
          curArtMarked = true;
          count += 1;
        }
        setHighlight(p, apply);
        count += 1;
      }
    }
    await context.sync();
  });
  return count;
}

/** Highlight the extracted product elements in the open document. */
export function highlightProductItems(allParts: boolean): Promise<number> {
  return markProductItems(allParts, true);
}

/** Remove the highlight from the same product elements. */
export function clearProductHighlights(allParts: boolean): Promise<number> {
  return markProductItems(allParts, false);
}
