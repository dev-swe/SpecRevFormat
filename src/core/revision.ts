/*
 * Revision-summary data model + persistence.
 *
 * The summary is stored in the DOCUMENT's own settings (Office.context.document.settings),
 * so it travels inside the spec file and is scoped to that document — NOT in the machine's
 * localStorage. Storing settings never alters the document body; only the export produces a
 * separate Revision Change Narrative .docx (see docx.ts). The spec text is never modified.
 */

/* global Office */

import { LogEntry } from "./format";

export interface RevisionRow {
  location: string; // article / heading the change sits under
  type: string; // "Addition" | "Deletion" | "Revision" | free text
  description: string; // user-written narrative of the change
  specText: string; // the change text pulled from the spec (reference only)
}

export interface RevisionSummary {
  project: string;
  section: string; // e.g. "SECTION 220500 - GENERAL PLUMBING PROVISIONS"
  revision: string; // e.g. "ASI-1"
  date: string;
  preparedBy: string;
  rows: RevisionRow[];
}

const SETTINGS_KEY = "specFormatter.revisionSummary.v1";

export function emptyRow(): RevisionRow {
  return { location: "", type: "Addition", description: "", specText: "" };
}

export function emptySummary(): RevisionSummary {
  return { project: "", section: "", revision: "", date: "", preparedBy: "", rows: [] };
}

/** Build editable rows from a scan's log entries (seeds the narrative column empty). */
export function rowsFromEntries(entries: LogEntry[]): RevisionRow[] {
  return entries.map((e) => ({
    location: e.location,
    type: e.type === "deletion" ? "Deletion" : "Addition",
    description: "",
    specText: e.text,
  }));
}

/** The nearest section string from a scan, for prefilling the Section field. */
export function sectionFromEntries(entries: LogEntry[]): string {
  const withSection = entries.find((e) => e.section);
  return withSection ? withSection.section : "";
}

function isSummary(x: unknown): x is RevisionSummary {
  if (!x || typeof x !== "object") return false;
  const s = x as Record<string, unknown>;
  return typeof s.project === "string" && Array.isArray(s.rows);
}

/** Read the summary stored in this document's settings (synchronous; cached at init). */
export function loadSummary(): RevisionSummary {
  try {
    const raw = Office.context.document.settings.get(SETTINGS_KEY);
    if (isSummary(raw)) {
      // normalize rows so older/partial data loads cleanly
      const rows = (raw.rows as unknown[])
        .filter((r) => r && typeof r === "object")
        .map((r) => {
          const row = r as Record<string, unknown>;
          return {
            location: String(row.location ?? ""),
            type: String(row.type ?? ""),
            description: String(row.description ?? ""),
            specText: String(row.specText ?? ""),
          } as RevisionRow;
        });
      return { ...emptySummary(), ...raw, rows };
    }
  } catch {
    // settings unavailable — return empty
  }
  return emptySummary();
}

/** Persist the summary into the document's settings (writes into the .docx on save). */
export function saveSummary(summary: RevisionSummary): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.settings.set(SETTINGS_KEY, summary);
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message || "Could not save the revision summary."));
      });
    } catch (e) {
      reject(e instanceof Error ? e : new Error(String(e)));
    }
  });
}
