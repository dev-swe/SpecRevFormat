/*
 * Architect template definitions.
 *
 * Ported from the original VBA macro `modTemplates.bas` (ArchitectTemplate type +
 * GetTemplates). Colors are stored as "#RRGGBB" hex strings so they map directly to
 * both an <input type="color"> in the UI and Office.js `font.color`, which accepts
 * hex. (The VBA used RGB() Longs; that conversion is no longer needed.)
 */

export interface ArchitectTemplate {
  name: string;
  bodyFont: string;
  bodySize: number;
  bodyColor: string;
  additionColor: string;
  deletionColor: string;
  headingFont: string;
  headingSize: number;
  headingColor: string;
  /** Paragraph SpaceAfter, in points (reformat mode only). */
  paraSpacingPt: number;
  /**
   * Paragraph style names that mark a heading, e.g. the CSI SectionFormat styles
   * SCT / PRT / ART. A paragraph using one of these is treated as a heading even when
   * its text doesn't match a heading pattern (article numbers like "1.01" are
   * auto-generated list numbers, not text, so they never match by text alone).
   * Matched case-insensitively, in addition to Word built-in "Heading N" styles and
   * the text patterns.
   */
  headingStyles: string[];
  /**
   * When true, the template does NOT reformat the document. Instead it scans for
   * bold runs (additions) and bold+strikethrough runs (deletions) and appends a tag
   * (e.g. "(ASI-1)") after each. Mirrors the VBA `AppendTag` flag.
   */
  appendTag: boolean;
}

/** Default CSI SectionFormat title styles treated as headings. */
export const DEFAULT_HEADING_STYLES: string[] = ["SCT", "PRT", "ART"];

/**
 * The seed templates, matching the four defined in `modTemplates.bas`.
 * Only "Rowell Brokaw Architects" is a real, tag-append template; B/C/D are editable
 * starting points. Users can add / edit / delete these in the task pane; the set is
 * persisted (see store.ts), so these defaults are only used on first run.
 */
export const DEFAULT_TEMPLATES: ArchitectTemplate[] = [
  {
    name: "Rowell Brokaw Architects (RBA)",
    bodyFont: "Arial",
    bodySize: 10,
    bodyColor: "#000000",
    additionColor: "#0000FF", // RGB(0,0,255)
    deletionColor: "#FF0000", // RGB(255,0,0)
    headingFont: "Arial",
    headingSize: 12,
    headingColor: "#000000",
    paraSpacingPt: 6,
    headingStyles: ["SCT", "PRT", "ART"],
    appendTag: true,
  },
  {
    name: "Architect B",
    bodyFont: "Times New Roman",
    bodySize: 10,
    bodyColor: "#000000",
    additionColor: "#007000", // RGB(0,112,0)
    deletionColor: "#C00000", // RGB(192,0,0)
    headingFont: "Times New Roman",
    headingSize: 12,
    headingColor: "#000000",
    paraSpacingPt: 6,
    headingStyles: ["SCT", "PRT", "ART"],
    appendTag: false,
  },
  {
    name: "Architect C",
    bodyFont: "Calibri",
    bodySize: 11,
    bodyColor: "#000000",
    additionColor: "#0000C8", // RGB(0,0,200)
    deletionColor: "#C80000", // RGB(200,0,0)
    headingFont: "Calibri",
    headingSize: 13,
    headingColor: "#000000",
    paraSpacingPt: 8,
    headingStyles: ["SCT", "PRT", "ART"],
    appendTag: false,
  },
  {
    name: "Architect D",
    bodyFont: "Courier New",
    bodySize: 10,
    bodyColor: "#000000",
    additionColor: "#0000B4", // RGB(0,0,180)
    deletionColor: "#B40000", // RGB(180,0,0)
    headingFont: "Courier New",
    headingSize: 12,
    headingColor: "#000000",
    paraSpacingPt: 6,
    headingStyles: ["SCT", "PRT", "ART"],
    appendTag: false,
  },
];

/** A fresh, blank template for the "Add template" flow. */
export function newTemplate(): ArchitectTemplate {
  return {
    name: "New architect",
    bodyFont: "Arial",
    bodySize: 10,
    bodyColor: "#000000",
    additionColor: "#0000FF",
    deletionColor: "#FF0000",
    headingFont: "Arial",
    headingSize: 12,
    headingColor: "#000000",
    paraSpacingPt: 6,
    headingStyles: [...DEFAULT_HEADING_STYLES],
    appendTag: false,
  };
}
