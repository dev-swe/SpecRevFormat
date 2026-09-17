/*
 * Revision-log export.
 *
 * Turns the structured LogEntry rows produced by a tag-append run into a CSV the user
 * can open in Excel and paste into an ASI/CCD cover document. Columns:
 *   Tag, Type, Section, Location, Text
 */

/* global Blob, URL, document */

import { LogEntry } from "./format";

function csvCell(value: string): string {
  const s = value ?? "";
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

/** Build the CSV text for a set of entries (exported for testing/reuse). */
export function buildRevisionCsv(entries: LogEntry[], tagText: string): string {
  const header = ["Tag", "Type", "Section", "Location", "Text"];
  const rows = [header];
  for (const e of entries) {
    rows.push([
      tagText,
      e.type === "deletion" ? "Deletion" : "Addition",
      e.section,
      e.location,
      e.text,
    ]);
  }
  // CRLF line endings + BOM so Excel opens it cleanly.
  return "﻿" + rows.map((r) => r.map(csvCell).join(",")).join("\r\n");
}

/** Trigger a download of the revision log as a CSV file. */
export function downloadRevisionLog(entries: LogEntry[], tagText: string, docBase: string): void {
  const csv = buildRevisionCsv(entries, tagText);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const safeTag = (tagText || "revision").replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  const a = document.createElement("a");
  a.href = url;
  a.download = `${docBase}-revision-log-${safeTag}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
