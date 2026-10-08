/*
 * Insert the compiled submittal table into the open Word document.
 *
 * The download path (core/submittalTemplate.ts) reproduces the firm's standalone
 * Submittal Review file. This path instead drops the review table inline at the cursor,
 * for when the reviewer is already working inside a submittal document. Office.js can't
 * carry the full template styling into an arbitrary document, so it applies the essentials:
 * Tahoma 10, bold body, two-line Spec Section / Item Name cells, and the forest header row.
 */

/* global Word */

import { SubmittalItem } from "./submittal";
import { SubmittalMeta } from "./submittalDocx";

const FOREST = "#12413C";
const HEADERS = ["Item No.", "Action Code", "Last Submit", "Spec Section", "Item Name", "Comments"];

/** Insert the table after the current selection. Returns the row count inserted. */
export async function insertSubmittalTable(
  items: SubmittalItem[],
  meta: SubmittalMeta
): Promise<number> {
  const values: string[][] = [
    HEADERS,
    ...items.map((it) => [
      "",
      "",
      meta.sweSubmittalNo || "",
      it.outline ? `${it.section}\n${it.outline}` : it.section,
      it.qualifier ? `${it.name}\n${it.qualifier}` : it.name,
      "",
    ]),
  ];

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const table = selection.insertTable(
      values.length,
      HEADERS.length,
      Word.InsertLocation.after,
      values
    );

    // Base font + bold body, matching the template's data rows.
    const range = table.getRange();
    range.font.name = "Tahoma";
    range.font.size = 10;
    range.font.bold = true;

    // Forest header row with white text (getCell avoids a navigational load).
    for (let c = 0; c < HEADERS.length; c++) {
      const cell = table.getCell(0, c);
      cell.shadingColor = FOREST;
      cell.body.font.color = "#FFFFFF";
      cell.body.font.bold = true;
    }

    table.select();
    await context.sync();
  });

  return items.length;
}
