/*
 * Template persistence.
 *
 * Templates are stored in the task pane's localStorage (per machine/user), which
 * survives across documents and Word restarts. Export/import to a JSON file lets a
 * small team share one template set (replacing the old "edit modTemplates.bas" step).
 *
 * All reads/writes are wrapped in try/catch: in some Office webview contexts storage can
 * be unavailable, in which case we fall back to the in-memory defaults for the session.
 */

/* global window, document, Blob, URL, File, FileReader */

import { ArchitectTemplate, DEFAULT_HEADING_STYLES, DEFAULT_TEMPLATES } from "./templates";

const STORAGE_KEY = "specFormatter.templates.v1";

function isTemplate(x: unknown): x is Partial<ArchitectTemplate> {
  if (!x || typeof x !== "object") return false;
  const t = x as Record<string, unknown>;
  return (
    typeof t.name === "string" &&
    typeof t.bodyFont === "string" &&
    typeof t.bodySize === "number" &&
    typeof t.bodyColor === "string" &&
    typeof t.additionColor === "string" &&
    typeof t.deletionColor === "string" &&
    typeof t.headingFont === "string" &&
    typeof t.headingSize === "number" &&
    typeof t.headingColor === "string" &&
    typeof t.paraSpacingPt === "number" &&
    typeof t.appendTag === "boolean"
  );
}

/**
 * Fill in fields added after a template was first saved, so older stored/imported
 * templates load cleanly. Currently: default `headingStyles` when missing/invalid.
 */
function normalizeTemplate(t: Partial<ArchitectTemplate>): ArchitectTemplate {
  const headingStyles =
    Array.isArray(t.headingStyles) && t.headingStyles.every((s) => typeof s === "string")
      ? t.headingStyles
      : [...DEFAULT_HEADING_STYLES];
  return { ...(t as ArchitectTemplate), headingStyles };
}

/** Parse an unknown value into a clean template array, or null if it isn't valid. */
export function parseTemplates(value: unknown): ArchitectTemplate[] | null {
  if (!Array.isArray(value)) return null;
  const clean = value.filter(isTemplate).map(normalizeTemplate);
  return clean.length > 0 ? clean : null;
}

export function loadTemplates(): ArchitectTemplate[] {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = parseTemplates(JSON.parse(raw));
      if (parsed) return parsed;
    }
  } catch {
    // storage unavailable or corrupt — fall through to defaults
  }
  // Seed defaults on first run.
  const seed = DEFAULT_TEMPLATES.map((t) => ({ ...t }));
  saveTemplates(seed);
  return seed;
}

export function saveTemplates(templates: ArchitectTemplate[]): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
  } catch {
    // best-effort; session continues with in-memory copy
  }
}

/** Trigger a download of the current templates as a JSON file. */
export function exportTemplates(templates: ArchitectTemplate[]): void {
  const blob = new Blob([JSON.stringify(templates, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "spec-formatter-templates.json";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/** Read a user-selected JSON file and return its templates, or throw on invalid input. */
export function importTemplatesFromFile(file: File): Promise<ArchitectTemplate[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Could not read the file."));
    reader.onload = () => {
      try {
        const parsed = parseTemplates(JSON.parse(String(reader.result)));
        if (!parsed) {
          reject(new Error("That file does not contain valid Spec Formatter templates."));
          return;
        }
        resolve(parsed);
      } catch {
        reject(new Error("That file is not valid JSON."));
      }
    };
    reader.readAsText(file);
  });
}
