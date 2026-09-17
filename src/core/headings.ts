/*
 * Heading detection.
 *
 * Ported from `IsHeading` / `MatchesHeadingPattern` in the VBA `modFormatter.bas`.
 *
 * Note on VBA `Like`: only `? * # [ ]` are wildcards; `.` is a LITERAL period. So the
 * original patterns translate to regex as follows:
 *   "PART [0-9]*"            -> /^PART \d/
 *   "SECTION [0-9]*"         -> /^SECTION \d/
 *   "[0-9].[0-9][0-9]*"      -> /^\d\.\d\d/      (e.g. "1.01", "2.03 SUBMITTALS")
 *   "[0-9].[0-9][0-9].[0-9]*"-> /^\d\.\d\d\.\d/  (subset of the above; kept for clarity)
 *   second char "." + "[A-Z].*" -> /^[A-Z]\./    (e.g. "A. Wall assembly")
 *
 * The construction-spec styles in these documents (SCT/PRT/ART/PR1…) are NOT Word
 * built-in "Heading N" styles, so detection leans on the text patterns above. A Word
 * built-in heading style is still honored, matching the original.
 */

const HEADING_TEXT_PATTERNS: RegExp[] = [
  /^PART \d/,
  /^SECTION \d/,
  /^\d\.\d\d/,
  /^\d\.\d\d\.\d/,
  /^[A-Z]\./,
];

/** Normalize a paragraph's raw text the way the VBA did before pattern matching. */
export function normalizeHeadingText(raw: string): string {
  return raw.replace(/\t/g, " ").replace(/\r/g, "").replace(/\n/g, "").trim();
}

/** True if the trimmed/normalized paragraph text matches a heading pattern. */
export function matchesHeadingPattern(text: string): boolean {
  const txt = normalizeHeadingText(text);
  if (txt.length === 0) return false;
  return HEADING_TEXT_PATTERNS.some((re) => re.test(txt));
}

/**
 * True if a paragraph is a heading, given its text and its Word style name(s).
 * `styleName` is `paragraph.style`; `builtInStyle` is `paragraph.styleBuiltIn`
 * (locale-independent, e.g. "Heading1"). A paragraph counts as a heading if any of:
 *   - its style (custom or built-in) starts with "Heading", or
 *   - its style matches one of `headingStyles` (e.g. CSI SCT/PRT/ART), case-insensitive, or
 *   - its text matches a heading pattern.
 * `headingStyles` is per-template and editable in the UI.
 */
export function isHeading(
  text: string,
  styleName?: string,
  builtInStyle?: string,
  headingStyles: string[] = []
): boolean {
  if (styleName && styleName.startsWith("Heading")) return true;
  if (builtInStyle && builtInStyle.startsWith("Heading")) return true;
  if (styleMatchesList(styleName, headingStyles)) return true;
  if (styleMatchesList(builtInStyle, headingStyles)) return true;
  return matchesHeadingPattern(text);
}

function styleMatchesList(style: string | undefined, headingStyles: string[]): boolean {
  if (!style) return false;
  const s = style.trim().toLowerCase();
  return headingStyles.some((h) => h.trim().toLowerCase() === s);
}
