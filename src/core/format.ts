/*
 * The Word/Office.js formatting engine.
 *
 * Ports `FormatSpecification` from `modFormatter.bas`. Two modes, chosen by the
 * template's `appendTag` flag:
 *   - tag-append: insert a bold tag after each continuous bold (addition) run and each
 *     continuous bold+strikethrough (deletion) run, skipping headings.
 *   - reformat: apply heading vs. body fonts/sizes/colors, recolor addition/deletion
 *     runs, and set paragraph spacing.
 *
 * A single read-only scan (loadParagraphInfo + buildTokens) powers both the dry-run
 * `previewSpecification` and the applying `formatSpecification`, so the preview reflects
 * exactly what formatting will do. Run detection is at word granularity via Range.split
 * (see runs.ts). All edits happen inside one Word.run/context.sync batch (one Ctrl+Z).
 */

/* global Word */

import { ArchitectTemplate } from "./templates";
import { isHeading } from "./headings";
import { classifyWord, computeTagRuns, Token } from "./runs";

/** A serializable revision-log row (no Word handles), for preview + CSV export. */
export interface LogEntry {
  type: "addition" | "deletion";
  text: string;
  section: string;
  location: string;
}

export interface FormatResult {
  templateName: string;
  mode: "tag" | "reformat";
  tagCount: number;
  entries: LogEntry[];
}

export interface PreviewResult {
  templateName: string;
  mode: "tag" | "reformat";
  paragraphCount: number;
  headingCount: number;
  additionRuns: number;
  deletionRuns: number;
  entries: LogEntry[];
}

interface ParaInfo {
  para: Word.Paragraph;
  text: string;
  heading: boolean;
  words?: Word.RangeCollection; // populated for non-heading paragraphs
}

/** First line of a heading's text (SCT titles span two lines via a soft break). */
function firstLine(text: string): string {
  return text
    .split(/[\r\n\v]/)[0]
    .replace(/\t/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/** Read-only passes: load every paragraph's text/style + split non-heading words. */
async function loadParagraphInfo(
  context: Word.RequestContext,
  tmpl: ArchitectTemplate
): Promise<ParaInfo[]> {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  const infos: ParaInfo[] = paragraphs.items.map((para) => {
    para.load("text, style, styleBuiltIn");
    return { para, text: "", heading: false };
  });
  await context.sync();

  for (const info of infos) {
    info.text = info.para.text;
    info.heading = isHeading(
      info.para.text,
      info.para.style,
      info.para.styleBuiltIn,
      tmpl.headingStyles
    );
  }

  for (const info of infos) {
    if (info.heading) continue;
    const range = info.para.getRange(Word.RangeLocation.content);
    const words = range.split(
      [" "],
      false /* multiParagraphs */,
      true /* trimDelimiters */,
      true /* trimSpacing */
    );
    words.load("items/text, items/font/bold, items/font/strikeThrough");
    info.words = words;
  }
  await context.sync();

  return infos;
}

/** Build the ordered token stream, tracking section/heading context for the log. */
function buildTokens(infos: ParaInfo[]): Token<Word.Range>[] {
  const tokens: Token<Word.Range>[] = [];
  let ctxSection = "";
  let ctxLocation = "";

  for (const info of infos) {
    if (info.heading) {
      const line = firstLine(info.text);
      if (line) {
        ctxLocation = line;
        if (/^section\b/i.test(line)) ctxSection = line;
      }
      tokens.push({ kind: "flush" });
      continue;
    }

    const wordItems = info.words ? info.words.items : [];
    if (wordItems.length === 0) {
      tokens.push({ kind: "flush" });
      continue;
    }

    for (const word of wordItems) {
      tokens.push({
        kind: "word",
        bold: word.font.bold === true,
        strike: word.font.strikeThrough === true,
        ref: word,
        text: word.text,
        section: ctxSection,
        location: ctxLocation,
      });
    }
  }
  return tokens;
}

/** Dry run: report what would change, without modifying the document. */
export async function previewSpecification(tmpl: ArchitectTemplate): Promise<PreviewResult> {
  return Word.run(async (context) => {
    const infos = await loadParagraphInfo(context, tmpl);
    const runs = computeTagRuns(buildTokens(infos));
    const entries: LogEntry[] = runs.map((r) => ({
      type: r.type,
      text: r.text,
      section: r.section,
      location: r.location,
    }));
    return {
      templateName: tmpl.name,
      mode: tmpl.appendTag ? "tag" : "reformat",
      paragraphCount: infos.length,
      headingCount: infos.filter((i) => i.heading).length,
      additionRuns: entries.filter((e) => e.type === "addition").length,
      deletionRuns: entries.filter((e) => e.type === "deletion").length,
      entries,
    };
  });
}

/** Apply formatting. Returns the revision log (tag mode) alongside the result. */
export async function formatSpecification(
  tmpl: ArchitectTemplate,
  tagText: string
): Promise<FormatResult> {
  return Word.run(async (context) => {
    const infos = await loadParagraphInfo(context, tmpl);
    const tokens = buildTokens(infos);

    let tagCount = 0;
    let entries: LogEntry[] = [];

    if (tmpl.appendTag) {
      const runs = computeTagRuns(tokens);
      entries = runs.map((r) => ({
        type: r.type,
        text: r.text,
        section: r.section,
        location: r.location,
      }));
      const tag = " " + tagText;
      for (const run of runs) {
        const inserted = run.ref.insertText(tag, Word.InsertLocation.after);
        inserted.font.bold = true;
        inserted.font.strikeThrough = false;
        tagCount++;
      }
    } else {
      applyReformat(infos, tmpl);
    }

    await context.sync();

    return {
      templateName: tmpl.name,
      mode: tmpl.appendTag ? "tag" : "reformat",
      tagCount,
      entries,
    };
  });
}

/** Reformat mode: heading vs. body fonts, paragraph spacing, per-run recoloring. */
function applyReformat(infos: ParaInfo[], tmpl: ArchitectTemplate): void {
  for (const info of infos) {
    if (info.heading) {
      applyHeadingStyle(info.para, tmpl);
      continue;
    }
    info.para.spaceAfter = tmpl.paraSpacingPt;
    const paraRange = info.para.getRange(Word.RangeLocation.whole);
    paraRange.font.name = tmpl.bodyFont;
    paraRange.font.size = tmpl.bodySize;

    const wordItems = info.words ? info.words.items : [];
    for (const word of wordItems) {
      const cls = classifyWord(word.font.bold === true, word.font.strikeThrough === true);
      applyRunColor(word, cls, tmpl);
    }
  }
}

function applyHeadingStyle(para: Word.Paragraph, tmpl: ArchitectTemplate): void {
  const range = para.getRange(Word.RangeLocation.whole);
  range.font.name = tmpl.headingFont;
  range.font.size = tmpl.headingSize;
  range.font.color = tmpl.headingColor;
  range.font.bold = true;
  para.spaceAfter = tmpl.paraSpacingPt;
}

function applyRunColor(
  word: Word.Range,
  cls: ReturnType<typeof classifyWord>,
  tmpl: ArchitectTemplate
): void {
  if (cls === "addition") {
    word.font.color = tmpl.additionColor;
    word.font.bold = true;
  } else if (cls === "deletion") {
    word.font.color = tmpl.deletionColor;
    word.font.bold = true;
    word.font.strikeThrough = true;
  } else {
    word.font.color = tmpl.bodyColor;
  }
}
