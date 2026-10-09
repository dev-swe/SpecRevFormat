/*
 * Specification Section dialog.
 *
 * Shown as an Office Dialog (a real popup window) so the spec section has room, instead of
 * being cramped in the task pane. On ready it tells the parent it is listening, then renders
 * whatever section payload the parent sends (and re-renders if the parent sends another).
 */

/* global Office, document, HTMLElement, Element */

interface SectionMessage {
  section: string;
  outline: string;
  itemName: string;
  title: string;
  text: string;
}

interface Line {
  level: number; // 0 article heading, 1 PR1, 2 PR2, 3 PR3, 4 deeper
  marker: string; // "A." / "1." / "a." / ""
  body: string;
}

/**
 * Parse a reconstructed outline line back into {level, marker, body}. The extractor
 * encodes level by indentation: article & PR1 at column 0 (article starts "N.N "),
 * PR2 at 4 spaces, PR3 at 8, deeper at 12.
 */
function parseLine(raw: string): Line {
  const indent = raw.length - raw.replace(/^\s+/, "").length;
  const t = raw.trim();
  if (indent === 0 && /^\d+\.\d+\s/.test(t)) return { level: 0, marker: "", body: t };
  if (indent === 0) {
    const m = /^([A-Za-z]\.)\s+([\s\S]*)$/.exec(t);
    if (m) return { level: 1, marker: m[1], body: m[2] };
    return { level: 1, marker: "", body: t };
  }
  const level = indent >= 12 ? 4 : indent >= 8 ? 3 : 2;
  const m = /^([A-Za-z0-9]+\.)\s+([\s\S]*)$/.exec(t);
  if (m) return { level, marker: m[1], body: m[2] };
  return { level, marker: "", body: t };
}

/** Append the body text, bolding a leading "Property:" label when present. */
function appendBody(row: Element, body: string): void {
  const m = /^([^:]{1,48}):\s*([\s\S]*)$/.exec(body);
  if (m) {
    const strong = document.createElement("strong");
    strong.textContent = m[1] + ":";
    row.appendChild(strong);
    if (m[2]) row.appendChild(document.createTextNode(" " + m[2]));
  } else {
    row.appendChild(document.createTextNode(body));
  }
}

function render(data: SectionMessage): void {
  const crumb = document.getElementById("crumb") as HTMLElement;
  const title = document.getElementById("title") as HTMLElement;
  const body = document.getElementById("body") as HTMLElement;
  crumb.textContent = `${data.section} ${data.outline} — ${data.itemName}`.trim();
  title.textContent = data.title || "Specification Section";
  document.title = data.title || "Specification Section";

  body.textContent = "";
  if (!data.text) {
    body.innerHTML = '<span class="empty">No section text was found for this item.</span>';
    return;
  }
  for (const raw of data.text.split("\n")) {
    if (!raw.trim()) continue;
    const line = parseLine(raw);
    const row = document.createElement("div");
    row.className = "lvl" + line.level;
    if (line.marker) {
      const mk = document.createElement("span");
      mk.className = "mk";
      mk.textContent = line.marker;
      row.appendChild(mk);
      row.appendChild(document.createTextNode(" "));
    }
    appendBody(row, line.body);
    body.appendChild(row);
  }
}

Office.onReady(() => {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    try {
      render(JSON.parse((arg as { message: string }).message) as SectionMessage);
    } catch {
      // ignore malformed messages
    }
  });
  // Tell the parent we're ready to receive the section payload.
  Office.context.ui.messageParent("ready");
});
