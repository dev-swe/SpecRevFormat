/*
 * "Spec narrative contribution" export.
 *
 * Serializes the revision summary into the shared JSON contract that the pyRevit
 * Track Revisions tool imports, so specification-change narratives are compiled into
 * the SAME architect Revision Change Narrative .docx as the affected-sheet narratives.
 *
 * Track Revisions is the compiler (it has file-system access to the project folder and
 * the architect template); this add-in, running in the Office.js sandbox, can only
 * produce this JSON for the user to import there. Keyed by `revision` (e.g. "ASI-1") so
 * both tools collate onto the same revision. The specification document is not modified.
 */

/* global Blob, URL, document */

import { RevisionSummary } from "./revision";

/** One specification change → the {{SPEC_*}} tokens in the architect template. */
export interface SpecNarrativeEntry {
  location: string; // article/heading  -> {{SPEC_ARTICLE}}
  type: string; // Addition/Deletion/Revision
  description: string; // narrative      -> {{SPEC_NARRATIVE}}
  specText: string; // spec reference text
}

export interface SpecNarrativeContribution {
  kind: "spec-narrative";
  schemaVersion: number;
  revision: string; // matches a Revit revision number, e.g. "ASI-1"
  project: { name: string; number: string };
  source: string; // spec document file name
  section: string; // "SECTION 220500 - …"  -> {{SPEC_SECTION}}
  date: string;
  preparedBy: string;
  entries: SpecNarrativeEntry[];
}

function hasContent(r: { location: string; description: string; specText: string }): boolean {
  return (r.location.trim() + r.description.trim() + r.specText.trim()).length > 0;
}

/** Build the contribution object from the current summary. */
export function buildContribution(
  summary: RevisionSummary,
  source: string
): SpecNarrativeContribution {
  return {
    kind: "spec-narrative",
    schemaVersion: 1,
    revision: summary.revision,
    project: { name: summary.project, number: "" },
    source,
    section: summary.section,
    date: summary.date,
    preparedBy: summary.preparedBy,
    entries: summary.rows.filter(hasContent).map((r) => ({
      location: r.location,
      type: r.type,
      description: r.description,
      specText: r.specText,
    })),
  };
}

/** Download the contribution as a JSON file for import into Track Revisions. */
export function downloadContribution(
  summary: RevisionSummary,
  source: string,
  docBase: string
): void {
  const data = buildContribution(summary, source);
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const safeRev = (summary.revision || "revision")
    .replace(/[^A-Za-z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "");
  const a = document.createElement("a");
  a.href = url;
  a.download = `${docBase}-spec-narrative-${safeRev}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
