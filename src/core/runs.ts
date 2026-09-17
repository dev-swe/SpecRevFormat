/*
 * Pure run-classification logic, ported from `modFormatter.bas`.
 *
 * The VBA walked the document character-by-character, grouping equal (bold, strike)
 * spans into "runs". In this port we work at WORD granularity (a token per word), which
 * reproduces the macro's behavior for real spec redlines, where whole words/phrases are
 * bolded. Each word carries its bold/strike state plus its text and location context;
 * this module turns a token stream into structured runs used for both tag insertion and
 * the revision-log export.
 *
 * Classification (matches the VBA):
 *   addition = bold && !strike   -> AdditionColor, bold
 *   deletion = bold &&  strike   -> DeletionColor, bold, strikethrough
 *   body     = !bold             -> BodyColor  (strike-without-bold falls here, as in VBA)
 */

export type WordClass = "addition" | "deletion" | "body";

export function classifyWord(bold: boolean, strike: boolean): WordClass {
  if (bold && !strike) return "addition";
  if (bold && strike) return "deletion";
  return "body";
}

/**
 * A token in document order. Headings and empty non-heading paragraphs act as
 * section terminators (an open bold/strike section is "flushed" before them), exactly
 * as in the VBA where a heading or a zero-length paragraph closes the current section.
 * A `word` token carries its text and the nearest section/heading context for the log;
 * `ref` is an opaque handle (e.g. a Word.Range) the caller uses to place the tag.
 */
export type Token<TRef> =
  | {
      kind: "word";
      bold: boolean;
      strike: boolean;
      ref: TRef;
      text: string;
      section: string;
      location: string;
    }
  | { kind: "flush" };

/** A continuous addition or deletion run, with where to tag it and its context. */
export interface TagRun<TRef> {
  ref: TRef; // end-of-run handle: where the tag is inserted
  type: "addition" | "deletion";
  text: string; // the run's combined text
  section: string; // nearest "SECTION #####" heading
  location: string; // nearest heading/article above the run
}

/**
 * Group the token stream into continuous addition and deletion runs. Faithful to
 * `AppendTagsToDocument`: a section persists across paragraph boundaries and is only
 * closed by a body word, an empty paragraph, a heading, or end-of-document; switching
 * directly addition<->deletion closes the first. The end-of-run `ref` is where a tag is
 * placed. Used for both tag insertion (map to `ref`) and the revision log (full record).
 */
export function computeTagRuns<TRef>(tokens: Token<TRef>[]): TagRun<TRef>[] {
  const runs: TagRun<TRef>[] = [];
  let section: "none" | "addition" | "deletion" = "none";
  let words: string[] = [];
  let lastRef: TRef | null = null;
  let ctxSection = "";
  let ctxLocation = "";

  const close = () => {
    if (section !== "none" && lastRef !== null) {
      runs.push({
        ref: lastRef,
        type: section,
        text: words.join(" ").replace(/\s+/g, " ").trim(),
        section: ctxSection,
        location: ctxLocation,
      });
    }
    section = "none";
    words = [];
    lastRef = null;
  };

  for (const t of tokens) {
    if (t.kind === "flush") {
      close();
      continue;
    }
    const cls = classifyWord(t.bold, t.strike);
    if (cls === "addition") {
      if (section === "deletion") close();
      section = "addition";
      words.push(t.text);
      lastRef = t.ref;
      ctxSection = t.section;
      ctxLocation = t.location;
    } else if (cls === "deletion") {
      if (section === "addition") close();
      section = "deletion";
      words.push(t.text);
      lastRef = t.ref;
      ctxSection = t.section;
      ctxLocation = t.location;
    } else {
      close();
    }
  }
  close();
  return runs;
}
