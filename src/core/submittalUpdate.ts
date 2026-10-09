/*
 * Apply a revision diff to the Submittal Review table open in Word, in place.
 *
 * Preserves reviewer columns (Action Code / Last Submit / Comments). Per the confirmed
 * conventions: deleted items stay but go gray + strikethrough with a Comments note; changed
 * items go bold + black with a re-review note; renumbered items get their outline updated;
 * carried-over (unchanged) items go gray; added items are appended bold. Rows from sections
 * not covered by the revised files are left untouched. Then the embedded spec link is
 * refreshed to the revised content.
 */

/* global Word */

import { RevisionDiff, DiffRow, itemKey } from "./specDiff";
import { SPEC_LINK_NS } from "./specLink";

const GRAY = "#808080";
const BLACK = "#000000";

export interface ApplyResult {
  added: number;
  deleted: number;
  changed: number;
  renumbered: number;
  unchanged: number;
  orphanRows: number;
}

const LINE_BREAKS = new RegExp("[" + String.fromCharCode(13, 10, 11) + "]+");
function splitLines(s: string): string[] {
  return String(s || "").split(LINE_BREAKS);
}

/** Set a cell's font: Tahoma 10, with color/bold/strikethrough. */
function styleCell(cell: Word.TableCell, color: string, bold: boolean, strike: boolean): void {
  const f = cell.body.font;
  f.name = "Tahoma";
  f.size = 10;
  f.color = color;
  f.bold = bold;
  f.strikeThrough = strike;
}

function styleRow(
  table: Word.Table,
  ri: number,
  color: string,
  bold: boolean,
  strike: boolean
): void {
  for (let ci = 0; ci < 6; ci++) styleCell(table.getCell(ri, ci), color, bold, strike);
}

/** Rewrite a two-line cell (line1 + line break + line2). */
function setTwoLine(cell: Word.TableCell, line1: string, line2: string): void {
  cell.body.clear();
  cell.body.insertText(line1, Word.InsertLocation.start);
  if (line2) {
    cell.body.insertBreak(Word.BreakType.line, Word.InsertLocation.end);
    cell.body.insertText(line2, Word.InsertLocation.end);
  }
}

/** Append a note to the Comments cell, separated if it already has text. */
function appendComment(table: Word.Table, ri: number, existing: string, note: string): void {
  const cell = table.getCell(ri, 5);
  const prefix = existing && existing.trim() ? " — " : "";
  cell.body.insertText(prefix + note, Word.InsertLocation.end);
}

/** Find the submittal table and read all cell text in one pass. */
async function findTable(
  context: Word.RequestContext
): Promise<{ table: Word.Table; values: string[][] }> {
  const tables = context.document.body.tables;
  // eslint-disable-next-line office-addins/no-navigational-load
  tables.load("items");
  await context.sync();
  tables.items.forEach((t) => t.load("values"));
  await context.sync();
  for (const t of tables.items) {
    const v = t.values;
    if (
      v &&
      v[0] &&
      String(v[0][0] || "")
        .trim()
        .toLowerCase()
        .startsWith("item")
    ) {
      return { table: t, values: v };
    }
  }
  throw new Error("Couldn't find the submittal table in this document.");
}

/** Apply the diff to the open review table. Returns what was changed. */
export async function applyRevisionToTable(
  diff: RevisionDiff,
  revSubmittalNo: string
): Promise<ApplyResult> {
  const result: ApplyResult = {
    added: 0,
    deleted: 0,
    changed: 0,
    renumbered: 0,
    unchanged: 0,
    orphanRows: 0,
  };
  const scope = new Set(diff.sections);
  const rowByKey = new Map<string, DiffRow>();
  for (const r of diff.rows) {
    if (r.status !== "added") rowByKey.set(itemKey(r.section, r.name, r.qualifier), r);
  }
  const added = diff.rows.filter((r) => r.status === "added");

  await Word.run(async (context) => {
    const { table, values } = await findTable(context);

    for (let ri = 1; ri < values.length; ri++) {
      const section = (splitLines(values[ri][3])[0] || "").trim();
      if (!scope.has(section)) continue; // other spec/section — untouched
      const nameParts = splitLines(values[ri][4]);
      const name = (nameParts[0] || "").trim();
      const qualifier = (nameParts[1] || "").trim();
      const dr = rowByKey.get(itemKey(section, name, qualifier));
      if (!dr) {
        result.orphanRows += 1;
        continue; // reviewer-added or unmatched row — leave as-is
      }
      const comments = values[ri][5] || "";
      if (dr.status === "deleted") {
        appendComment(table, ri, comments, "Deleted in revision");
        styleRow(table, ri, GRAY, false, true);
        result.deleted += 1;
      } else if (dr.status === "changed") {
        if (dr.newOutline) setTwoLine(table.getCell(ri, 3), section, dr.newOutline);
        appendComment(table, ri, comments, "Spec revised — re-review");
        styleRow(table, ri, BLACK, true, false);
        result.changed += 1;
      } else if (dr.status === "renumbered") {
        if (dr.newOutline) setTwoLine(table.getCell(ri, 3), section, dr.newOutline);
        styleRow(table, ri, GRAY, false, false);
        result.renumbered += 1;
      } else {
        styleRow(table, ri, GRAY, false, false); // unchanged carried-over
        result.unchanged += 1;
      }
    }

    if (added.length) {
      const newValues = added.map((r) => [
        "",
        "",
        revSubmittalNo || "",
        r.newOutline ? `${r.section}\n${r.newOutline}` : r.section,
        r.qualifier ? `${r.name}\n${r.qualifier}` : r.name,
        "",
      ]);
      const start = values.length;
      table.addRows(Word.InsertLocation.end, newValues.length, newValues);
      await context.sync();
      for (let j = 0; j < added.length; j++) {
        const ri = start + j;
        const r = added[j];
        setTwoLine(table.getCell(ri, 3), r.section, r.newOutline || "");
        setTwoLine(table.getCell(ri, 4), r.name, r.qualifier);
        styleRow(table, ri, BLACK, true, false);
      }
      result.added = added.length;
    }

    await context.sync();
  });

  return result;
}

/** Replace the embedded spec-link custom XML part with refreshed content. */
export async function reembedSpecLink(xml: string): Promise<void> {
  await Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(SPEC_LINK_NS);
    // eslint-disable-next-line office-addins/no-navigational-load
    parts.load("items");
    await context.sync();
    parts.items.forEach((p) => p.delete());
    await context.sync();
    context.document.customXmlParts.add(xml);
    await context.sync();
  });
}
