/*
 * Revision diff: compare a revised specification's extracted items against the baseline
 * embedded in a linked Submittal Review, and classify each item.
 *
 * Identity keys on section + product name + qualifier (NOT outline, which shifts when items
 * are inserted). The diff is scoped to the sections present in the revised files; baseline
 * items in other sections are carried over untouched. Pure — no Office.js, fully testable.
 */

import { SubmittalItem } from "./submittal";
import { SpecLinkPayload, SpecLinkItem } from "./specLink";

export type DiffStatus = "added" | "deleted" | "changed" | "renumbered" | "unchanged";

export interface DiffRow {
  status: DiffStatus;
  section: string;
  name: string;
  qualifier: string;
  itemName: string;
  oldOutline?: string; // baseline outline (deleted/matched)
  newOutline?: string; // revised outline (added/matched)
  revisedItem?: SubmittalItem; // present for added/changed/renumbered/unchanged (for re-embed)
}

export interface RevisionDiff {
  rows: DiffRow[];
  counts: Record<DiffStatus, number>;
  sections: string[]; // sections covered by the revised files
  merged: SubmittalItem[]; // full item set after the revision (for re-embedding the link)
}

/** Identity key for matching an item across revisions (and to a table row). */
export function itemKey(section: string, name: string, qualifier: string): string {
  return `${section.trim()}\u0000${name.trim().toLowerCase()}\u0000${qualifier.trim().toLowerCase()}`;
}
const key = itemKey;

/** The baseline item rebuilt as a SubmittalItem (so carried-over items re-embed cleanly). */
function baselineAsItem(it: SpecLinkItem, baseline: SpecLinkPayload): SubmittalItem {
  const art = baseline.articles[it.artIndex];
  return {
    section: it.section,
    outline: it.outline,
    name: it.name,
    qualifier: it.qualifier,
    source: baseline.specName,
    specTitle: art ? art.title : "",
    specText: art ? art.text : "",
  };
}

export function diffRevision(
  baseline: SpecLinkPayload,
  revisedItems: SubmittalItem[]
): RevisionDiff {
  const revisedSections = new Set(revisedItems.map((it) => it.section));
  const sections = Array.from(revisedSections);

  const baselineInScope = new Map<string, SpecLinkItem>();
  const carriedOver: SubmittalItem[] = []; // baseline items in sections NOT being revised
  for (const it of baseline.items) {
    if (revisedSections.has(it.section)) {
      baselineInScope.set(key(it.section, it.name, it.qualifier), it);
    } else {
      carriedOver.push(baselineAsItem(it, baseline));
    }
  }

  const rows: DiffRow[] = [];
  const merged: SubmittalItem[] = [...carriedOver];
  const matchedKeys = new Set<string>();

  // Walk revised items: added / changed / renumbered / unchanged.
  for (const rev of revisedItems) {
    const k = key(rev.section, rev.name, rev.qualifier);
    const base = baselineInScope.get(k);
    const itemName = rev.qualifier ? `${rev.name} ${rev.qualifier}` : rev.name;
    if (!base) {
      rows.push({
        status: "added",
        section: rev.section,
        name: rev.name,
        qualifier: rev.qualifier,
        itemName,
        newOutline: rev.outline,
        revisedItem: rev,
      });
      merged.push(rev);
      continue;
    }
    matchedKeys.add(k);
    const baseText = baseline.articles[base.artIndex]?.text || "";
    let status: DiffStatus;
    if ((rev.specText || "") !== baseText) status = "changed";
    else if (rev.outline !== base.outline) status = "renumbered";
    else status = "unchanged";
    rows.push({
      status,
      section: rev.section,
      name: rev.name,
      qualifier: rev.qualifier,
      itemName,
      oldOutline: base.outline,
      newOutline: rev.outline,
      revisedItem: rev,
    });
    merged.push(rev);
  }

  // Baseline items in scope with no revised match -> deleted (kept in merged for the record).
  for (const [k, base] of Array.from(baselineInScope.entries())) {
    if (matchedKeys.has(k)) continue;
    rows.push({
      status: "deleted",
      section: base.section,
      name: base.name,
      qualifier: base.qualifier,
      itemName: base.itemName,
      oldOutline: base.outline,
    });
    merged.push(baselineAsItem(base, baseline));
  }

  const counts: Record<DiffStatus, number> = {
    added: 0,
    deleted: 0,
    changed: 0,
    renumbered: 0,
    unchanged: 0,
  };
  for (const r of rows) counts[r.status] += 1;

  return { rows, counts, sections, merged };
}
