/*
 * Task pane UI controller for Spec Formatter.
 * Wires the picker + template editor to the core formatting engine.
 */

/* global Office, document, Event, HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement, setTimeout, clearTimeout, File, URL, location */

import { ArchitectTemplate, newTemplate } from "../core/templates";
import {
  loadTemplates,
  saveTemplates,
  exportTemplates,
  importTemplatesFromFile,
} from "../core/store";
import { formatSpecification, previewSpecification, LogEntry, PreviewResult } from "../core/format";
import { downloadRevisionLog } from "../core/revlog";
import {
  downloadBackup,
  readDocumentBytes,
  documentIsSaved,
  documentBaseName,
  documentFileName,
} from "../core/backup";
import {
  RevisionSummary,
  emptyRow,
  emptySummary,
  loadSummary,
  saveSummary,
  rowsFromEntries,
  sectionFromEntries,
} from "../core/revision";
import { exportNarrativeDocx } from "../core/docx";
import { downloadContribution } from "../core/contribution";
import { SubmittalItem, extractFromFile, extractFromDocx, combineAndSort } from "../core/submittal";
import { exportSubmittalDocx, SubmittalMeta } from "../core/submittalDocx";
import { exportSubmittalFromTemplate } from "../core/submittalTemplate";
import { insertSubmittalTable } from "../core/submittalInsert";
import { highlightProductItems, clearProductHighlights } from "../core/submittalHighlight";
import { buildSpecLinkPayload, readSpecLink, SpecLinkPayload } from "../core/specLink";

let templates: ArchitectTemplate[] = [];
let editingIndex: number | null = null; // index being edited, or null when adding
let lastLog: { entries: LogEntry[]; tagText: string } | null = null; // for CSV export
let summary: RevisionSummary = emptySummary(); // revision-summary table (per document)
let saveTimer: ReturnType<typeof setTimeout> | undefined;
let submittalFiles: File[] = []; // specs picked for the submittal-table compiler
let extractedItems: SubmittalItem[] = []; // rows from the open document, for editing/export

function $<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) return;
  $("sideload-msg").style.display = "none";
  $("app-body").style.display = "block";

  templates = loadTemplates();
  renderTemplateSelect();
  renderTemplateList();
  updateTagRow();
  if (!documentIsSaved()) $("not-saved-warning").style.display = "block";

  $<HTMLSelectElement>("template-select").addEventListener("change", () => {
    updateTagRow();
    resetPreview();
  });
  $<HTMLInputElement>("tag-text").addEventListener("input", resetPreview);
  $("preview-btn").addEventListener("click", onPreview);
  $("format-btn").addEventListener("click", onFormat);
  $("backup-btn").addEventListener("click", onBackup);
  $("download-log-btn").addEventListener("click", onDownloadLog);

  $("add-template-btn").addEventListener("click", onAddTemplate);
  $("export-btn").addEventListener("click", () => exportTemplates(templates));
  $("import-btn").addEventListener("click", () => $<HTMLInputElement>("import-file").click());
  $<HTMLInputElement>("import-file").addEventListener("change", onImportFile);
  $("editor-cancel").addEventListener("click", closeEditor);
  $("editor-save").addEventListener("click", onEditorSave);

  $("ex-extract-btn").addEventListener("click", onExtractOpenDoc);
  $("ex-clear-btn").addEventListener("click", onClearHighlights);
  $("ex-export-btn").addEventListener("click", onExportEdited);
  $("sub-pick-btn").addEventListener("click", () => $<HTMLInputElement>("sub-files").click());
  $<HTMLInputElement>("sub-files").addEventListener("change", onSubmittalFilesPicked);
  $("sub-compile-btn").addEventListener("click", () => onSubmittalAction("download"));
  $("sub-insert-btn").addEventListener("click", () => onSubmittalAction("insert"));

  initTabs();
  initRevision();
  initSpecLink();
});

/* --------------------- Specification link (review → spec) --------------------- */

async function initSpecLink(): Promise<void> {
  try {
    const bytes = await readDocumentBytes();
    const payload = await readSpecLink(bytes);
    renderSpecLink(payload);
  } catch {
    renderSpecLink(null);
  }
}

function renderSpecLink(payload: SpecLinkPayload | null): void {
  const card = $("speclink-card");
  if (!payload || payload.items.length === 0) {
    card.dataset.hidden = "true";
    activateTab(currentTab);
    return;
  }
  $("speclink-note").textContent =
    `This review is linked to ${payload.specName}. Click an item to see its specification section.`;
  const list = $("speclink-list");
  list.innerHTML = payload.items
    .map(
      (it, i) =>
        `<button type="button" class="speclink-item" data-i="${i}">` +
        `${esc(it.outline)} — ${esc(it.itemName)}</button>`
    )
    .join("");
  list
    .querySelectorAll<HTMLElement>(".speclink-item")
    .forEach((btn) =>
      btn.addEventListener("click", () => showSpecSection(payload, Number(btn.dataset.i)))
    );
  $("speclink-detail").style.display = "none";
  card.dataset.hidden = "false";
  // Surface the link: a linked review is open, so show the Submittal review tab.
  activateTab("submittal");
}

interface SectionMessage {
  section: string;
  outline: string;
  itemName: string;
  title: string;
  text: string;
}

let sectionDialog: Office.Dialog | null = null;
let pendingSection: SectionMessage | null = null;

function showSpecSection(payload: SpecLinkPayload, i: number): void {
  const it = payload.items[i];
  const art = payload.articles[it.artIndex];
  const data: SectionMessage = {
    section: it.section,
    outline: it.outline,
    itemName: it.itemName,
    title: art ? art.title : it.itemName,
    text: art ? art.text : "",
  };
  openSectionDialog(data);
}

/** Open (or update) the popup dialog showing the section; fall back to the inline panel. */
function openSectionDialog(data: SectionMessage): void {
  pendingSection = data;
  // If a dialog is already open, just push the new section into it.
  if (sectionDialog) {
    sectionDialog.messageChild(JSON.stringify(data));
    return;
  }
  const url = new URL("specsection.html", location.href).href;
  try {
    Office.context.ui.displayDialogAsync(url, { height: 60, width: 45 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        renderSectionInline(data);
        return;
      }
      sectionDialog = result.value;
      sectionDialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        // The dialog signals it's ready; send the current section.
        if (sectionDialog && pendingSection)
          sectionDialog.messageChild(JSON.stringify(pendingSection));
      });
      sectionDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        sectionDialog = null; // dialog closed (by user or error)
      });
    });
  } catch {
    renderSectionInline(data);
  }
}

/** Fallback: show the section inline in the pane (if the Dialog API is unavailable). */
function renderSectionInline(data: SectionMessage): void {
  const detail = $("speclink-detail");
  detail.innerHTML =
    `<p class="preview-summary">${esc(data.section)} ${esc(data.outline)} — ${esc(
      data.itemName
    )}</p>` + `<pre class="speclink-pre">${esc(data.text || "(section text not found)")}</pre>`;
  detail.style.display = "block";
}

/* -------------------------------- Tabs -------------------------------- */

let currentTab = "format";

function initTabs(): void {
  const tabs = document.querySelectorAll<HTMLElement>(".tab");
  tabs.forEach((tab) =>
    tab.addEventListener("click", () => activateTab(tab.dataset.tab || "format"))
  );
  activateTab("format");
}

function activateTab(name: string): void {
  currentTab = name;
  document.querySelectorAll<HTMLElement>(".tab").forEach((tab) => {
    tab.classList.toggle("is-active", tab.dataset.tab === name);
  });
  document.querySelectorAll<HTMLElement>(".card[data-group]").forEach((card) => {
    const show = card.dataset.group === name && card.dataset.hidden !== "true";
    card.style.display = show ? "" : "none";
  });
}

/* --------------------- Extract from open document --------------------- */

function setExStatus(message: string, kind: "info" | "error" | "success" = "info"): void {
  const el = $("ex-status");
  el.textContent = message;
  el.className = "status " + kind;
}

async function onExtractOpenDoc(): Promise<void> {
  const btn = $<HTMLInputElement>("ex-extract-btn");
  const clearBtn = $<HTMLInputElement>("ex-clear-btn");
  const allParts = $<HTMLInputElement>("ex-allparts").checked;
  const fullSub = $<HTMLInputElement>("ex-fullsub").checked;
  const highlight = $<HTMLInputElement>("ex-highlight").checked;
  btn.disabled = true;
  clearBtn.disabled = true;
  setExStatus("Reading the open document…");
  try {
    const bytes = await readDocumentBytes();
    const name = documentFileName();
    const items = await extractFromDocx(bytes, name, allParts, fullSub);
    extractedItems = items;
    renderEditablePreview($("ex-preview"), items);
    $("ex-export-row").style.display = items.length ? "" : "none";
    if (items.length === 0) {
      setExStatus(
        "No product items found. Is this a MasterSpec section (SCT/PRT/ART styles)? Try “Include every Article”.",
        "error"
      );
      return;
    }
    let msg = `Extracted ${items.length} item(s) from ${name}. Edit any Item Name, then export.`;
    if (highlight) {
      const marked = await highlightProductItems(allParts);
      msg += ` Highlighted ${marked} element(s) in the document.`;
    }
    setExStatus(msg, "success");
  } catch (err) {
    setExStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    btn.disabled = false;
    clearBtn.disabled = false;
  }
}

async function onClearHighlights(): Promise<void> {
  const btn = $<HTMLInputElement>("ex-extract-btn");
  const clearBtn = $<HTMLInputElement>("ex-clear-btn");
  const allParts = $<HTMLInputElement>("ex-allparts").checked;
  btn.disabled = true;
  clearBtn.disabled = true;
  setExStatus("Clearing highlights…");
  try {
    const cleared = await clearProductHighlights(allParts);
    setExStatus(`Cleared highlights from ${cleared} element(s).`, "success");
  } catch (err) {
    setExStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    btn.disabled = false;
    clearBtn.disabled = false;
  }
}

/** The Item Name shown (and edited) for a row: name plus any qualifier, on one line. */
function itemNameText(it: SubmittalItem): string {
  return it.qualifier ? `${it.name} ${it.qualifier}` : it.name;
}

/** Render the extracted rows with an editable Item Name column. */
function renderEditablePreview(panel: HTMLElement, items: SubmittalItem[]): void {
  if (items.length === 0) {
    panel.innerHTML = "";
    panel.style.display = "none";
    return;
  }
  const rows = items
    .map(
      (it, i) =>
        `<tr><td>${esc(it.section)} ${esc(it.outline)}</td>` +
        `<td><input class="ex-name-input" type="text" data-idx="${i}" value="${esc(
          itemNameText(it)
        )}" /></td></tr>`
    )
    .join("");
  panel.innerHTML =
    `<p class="preview-summary">${items.length} item(s) — edit Item Name as needed.</p>` +
    `<div class="preview-table-wrap"><table class="preview-table"><thead><tr>` +
    `<th>Spec Section</th><th>Item Name</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  panel.style.display = "block";
}

/** Read the current (possibly edited) Item Name values back onto the items. */
function editedItems(): SubmittalItem[] {
  const inputs = $("ex-preview").querySelectorAll<HTMLInputElement>("input.ex-name-input");
  return extractedItems.map((it, i) => {
    const input = Array.from(inputs).find((el) => Number(el.dataset.idx) === i);
    const name = input ? input.value.trim() : itemNameText(it);
    // The edited text becomes the whole Item Name (single line); qualifier folds in.
    return { ...it, name, qualifier: "" };
  });
}

async function onExportEdited(): Promise<void> {
  if (extractedItems.length === 0) {
    setExStatus("Extract a document first.", "error");
    return;
  }
  const btn = $<HTMLInputElement>("ex-export-btn");
  btn.disabled = true;
  setExStatus("Building the submittal review…");
  try {
    const items = editedItems();
    const section = items[0].section;
    const fileBase = section ? `Submittal Review - ${section}` : "Submittal Review";
    const meta = {
      project: "",
      sweProjectNo: "",
      reviewDate: "",
      reviewedBy: "",
      sweSubmittalNo: "",
      contractorSubmittalNo: "",
    };
    const specLink = buildSpecLinkPayload(items, documentFileName());
    let name: string;
    try {
      name = await exportSubmittalFromTemplate(items, meta, fileBase, specLink);
    } catch {
      name = await exportSubmittalDocx(items, meta, fileBase, specLink);
    }
    setExStatus(`Exported ${items.length} item(s) to ${name} (linked to the spec).`, "success");
  } catch (err) {
    setExStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    btn.disabled = false;
  }
}

/* ----------------------- Compile submittal table ----------------------- */

function setSubStatus(message: string, kind: "info" | "error" | "success" = "info"): void {
  const el = $("sub-status");
  el.textContent = message;
  el.className = "status " + kind;
}

function onSubmittalFilesPicked(e: Event): void {
  const input = e.target as HTMLInputElement;
  submittalFiles = input.files ? Array.from(input.files) : [];
  const list = $("sub-files-list");
  const hasFiles = submittalFiles.length > 0;
  $<HTMLInputElement>("sub-compile-btn").disabled = !hasFiles;
  $<HTMLInputElement>("sub-insert-btn").disabled = !hasFiles;
  if (!hasFiles) {
    list.textContent = "";
    return;
  }
  list.textContent =
    `${submittalFiles.length} file(s): ` + submittalFiles.map((f) => f.name).join(", ");
  $("sub-preview").style.display = "none";
  setSubStatus("");
}

function submittalMeta(): SubmittalMeta {
  return {
    project: $<HTMLInputElement>("sub-project").value.trim(),
    sweProjectNo: $<HTMLInputElement>("sub-projectno").value.trim(),
    reviewDate: "",
    reviewedBy: $<HTMLInputElement>("sub-reviewedby").value.trim(),
    sweSubmittalNo: $<HTMLInputElement>("sub-submittal").value.trim(),
    contractorSubmittalNo: $<HTMLInputElement>("sub-contractorno").value.trim(),
  };
}

type SubmittalMode = "download" | "insert";

async function onSubmittalAction(mode: SubmittalMode): Promise<void> {
  if (submittalFiles.length === 0) {
    setSubStatus("Choose at least one .docx specification first.", "error");
    return;
  }
  const compileBtn = $<HTMLInputElement>("sub-compile-btn");
  const insertBtn = $<HTMLInputElement>("sub-insert-btn");
  const allParts = $<HTMLInputElement>("sub-allparts").checked;
  const fullSub = $<HTMLInputElement>("sub-fullsub").checked;
  compileBtn.disabled = true;
  insertBtn.disabled = true;
  setSubStatus(`Scanning ${submittalFiles.length} spec(s)…`);
  try {
    const groups = [];
    const errors: string[] = [];
    for (const file of submittalFiles) {
      try {
        groups.push(await extractFromFile(file, allParts, fullSub));
      } catch (err) {
        errors.push(err instanceof Error ? err.message : String(err));
      }
    }
    const items = combineAndSort(groups);
    renderSubmittalPreview(items, errors);
    if (items.length === 0) {
      setSubStatus(
        "No product items found. Check that these are MasterSpec sections (SCT/PRT/ART styles), or tick “Include every Article”.",
        "error"
      );
      return;
    }
    const meta = submittalMeta();
    const note = errors.length ? ` (${errors.length} file(s) skipped)` : "";

    if (mode === "insert") {
      await insertSubmittalTable(items, meta);
      setSubStatus(
        `Inserted ${items.length} item(s) from ${groups.length} spec(s) into the document.${note}`,
        "success"
      );
    } else {
      const fileBase =
        submittalFiles.length === 1 && items[0].section
          ? `Submittal Review - ${items[0].section}`
          : "Submittal Review - Combined";
      const specName =
        submittalFiles.length === 1
          ? submittalFiles[0].name
          : `${submittalFiles.length} specifications`;
      const specLink = buildSpecLinkPayload(items, specName);
      // Prefer the firm template (exact house style); fall back to the built-in layout
      // if the bundled template can't be fetched (e.g. offline cache miss).
      let name: string;
      try {
        name = await exportSubmittalFromTemplate(items, meta, fileBase, specLink);
      } catch {
        name = await exportSubmittalDocx(items, meta, fileBase, specLink);
      }
      setSubStatus(
        `Compiled ${items.length} item(s) from ${groups.length} spec(s). Downloaded ${name}.${note}`,
        "success"
      );
    }
  } catch (err) {
    setSubStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    compileBtn.disabled = false;
    insertBtn.disabled = false;
  }
}

function renderSubmittalPreview(items: ReturnType<typeof combineAndSort>, errors: string[]): void {
  renderItemsPreview($("sub-preview"), items, errors);
}

/** Render extracted rows into a preview panel (Spec Section | Item Name). */
function renderItemsPreview(
  panel: HTMLElement,
  items: ReturnType<typeof combineAndSort>,
  errors: string[] = []
): void {
  const rows = items
    .map(
      (it) =>
        `<tr><td>${esc(it.section)} ${esc(it.outline)}</td><td>${esc(it.name)} ${esc(
          it.qualifier
        )}</td></tr>`
    )
    .join("");
  const errHtml = errors.length ? `<p class="warning">${errors.map(esc).join("<br/>")}</p>` : "";
  const summaryHtml = `<p class="preview-summary">${items.length} item(s) found.</p>`;
  const tableHtml = rows
    ? `<div class="preview-table-wrap"><table class="preview-table"><thead><tr><th>Spec Section</th><th>Item Name</th></tr></thead><tbody>${rows}</tbody></table></div>`
    : "";
  panel.innerHTML = summaryHtml + tableHtml + errHtml;
  panel.style.display = "block";
}

function esc(s: string): string {
  return s.replace(/[&<>"]/g, (c) =>
    c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : "&quot;"
  );
}

/* ----------------------------- Format panel ----------------------------- */

function selectedTemplate(): ArchitectTemplate {
  const idx = $<HTMLSelectElement>("template-select").selectedIndex;
  return templates[Math.max(0, idx)];
}

function updateTagRow(): void {
  $("tag-row").style.display = selectedTemplate()?.appendTag ? "block" : "none";
}

function setStatus(message: string, kind: "info" | "error" | "success" = "info"): void {
  const el = $("status");
  el.textContent = message;
  el.className = "status " + kind;
}

async function onFormat(): Promise<void> {
  const tmpl = selectedTemplate();
  if (!tmpl) return;

  const tagText = $<HTMLInputElement>("tag-text").value.trim();
  if (tmpl.appendTag && tagText.length === 0) {
    setStatus("Enter tag text (e.g. “(ASI-1)”) before formatting.", "error");
    return;
  }

  const btn = $<HTMLInputElement>("format-btn");
  btn.disabled = true;
  setStatus("Formatting…");
  try {
    const result = await formatSpecification(tmpl, tagText);
    let msg = `Formatted with “${result.templateName}”.`;
    if (result.mode === "tag") msg += ` ${result.tagCount} “${tagText}” tag(s) inserted.`;
    msg += " Press Ctrl+Z to undo.";
    setStatus(msg, "success");
    setLog(result.entries, tagText);
  } catch (e) {
    setStatus("Formatting failed: " + errorMessage(e), "error");
  } finally {
    btn.disabled = false;
  }
}

async function onPreview(): Promise<void> {
  const tmpl = selectedTemplate();
  if (!tmpl) return;
  const tagText = $<HTMLInputElement>("tag-text").value.trim();

  const btn = $<HTMLInputElement>("preview-btn");
  btn.disabled = true;
  setStatus("Scanning…");
  try {
    const result = await previewSpecification(tmpl);
    renderPreview(result, tagText);
    setStatus("Preview ready — nothing changed yet.", "info");
    setLog(result.entries, tagText);
  } catch (e) {
    setStatus("Preview failed: " + errorMessage(e), "error");
  } finally {
    btn.disabled = false;
  }
}

function onDownloadLog(): void {
  if (!lastLog || lastLog.entries.length === 0) return;
  downloadRevisionLog(lastLog.entries, lastLog.tagText, documentBaseName());
}

/** Remember the latest scan/format entries and toggle the CSV download button. */
function setLog(entries: LogEntry[], tagText: string): void {
  if (entries.length > 0) {
    lastLog = { entries, tagText };
    $("log-row").style.display = "flex";
  } else {
    lastLog = null;
    $("log-row").style.display = "none";
  }
}

function resetPreview(): void {
  $("preview-panel").style.display = "none";
  $("preview-panel").textContent = "";
  $("log-row").style.display = "none";
  lastLog = null;
}

function renderPreview(result: PreviewResult, tagText: string): void {
  const panel = $("preview-panel");
  panel.textContent = "";

  const summary = document.createElement("p");
  summary.className = "preview-summary";
  if (result.mode === "tag") {
    const n = result.additionRuns + result.deletionRuns;
    summary.textContent =
      `Would insert ${n} tag${n === 1 ? "" : "s"}` +
      (tagText ? ` (“${tagText}”)` : "") +
      ` — ${result.additionRuns} addition${result.additionRuns === 1 ? "" : "s"}, ` +
      `${result.deletionRuns} deletion${result.deletionRuns === 1 ? "" : "s"}. Nothing changed yet.`;
  } else {
    summary.textContent =
      `Would reformat ${result.paragraphCount} paragraph${result.paragraphCount === 1 ? "" : "s"} ` +
      `(${result.headingCount} heading${result.headingCount === 1 ? "" : "s"}); recolor ` +
      `${result.additionRuns} addition${result.additionRuns === 1 ? "" : "s"} and ` +
      `${result.deletionRuns} deletion${result.deletionRuns === 1 ? "" : "s"}. Nothing changed yet.`;
  }
  panel.appendChild(summary);

  if (result.entries.length > 0) {
    panel.appendChild(buildPreviewTable(result.entries));
  }
  panel.style.display = "block";
}

function buildPreviewTable(entries: LogEntry[]): HTMLElement {
  const MAX = 100;
  const wrap = document.createElement("div");
  wrap.className = "preview-table-wrap";
  const table = document.createElement("table");
  table.className = "preview-table";

  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  for (const h of ["Type", "Location", "Text"]) {
    const th = document.createElement("th");
    th.textContent = h;
    hr.appendChild(th);
  }
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  entries.slice(0, MAX).forEach((e) => {
    const tr = document.createElement("tr");
    const type = document.createElement("td");
    type.textContent = e.type === "deletion" ? "Deletion" : "Addition";
    type.className = e.type === "deletion" ? "cell-del" : "cell-add";
    const loc = document.createElement("td");
    loc.textContent = e.location || "—";
    const txt = document.createElement("td");
    txt.textContent = truncate(e.text, 80);
    tr.appendChild(type);
    tr.appendChild(loc);
    tr.appendChild(txt);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrap.appendChild(table);

  if (entries.length > MAX) {
    const more = document.createElement("p");
    more.className = "help";
    more.textContent = `…and ${entries.length - MAX} more (all included in the CSV).`;
    wrap.appendChild(more);
  }
  return wrap;
}

function truncate(s: string, n: number): string {
  return s.length > n ? s.slice(0, n - 1).trimEnd() + "…" : s;
}

async function onBackup(): Promise<void> {
  const btn = $<HTMLInputElement>("backup-btn");
  btn.disabled = true;
  setStatus("Preparing a copy…");
  try {
    const fileName = await downloadBackup();
    setStatus(`Saved a copy as “${fileName}” (check your downloads).`, "success");
  } catch (e) {
    setStatus("Could not save a copy: " + errorMessage(e), "error");
  } finally {
    btn.disabled = false;
  }
}

/* --------------------------- Revision summary --------------------------- */

function initRevision(): void {
  summary = loadSummary();
  setInput("rev-project", summary.project);
  setInput("rev-revision", summary.revision);
  setInput("rev-date", summary.date);
  setInput("rev-preparedBy", summary.preparedBy);
  setInput("rev-section", summary.section);
  renderRevisionRows();

  bindMeta("rev-project", (v) => (summary.project = v));
  bindMeta("rev-revision", (v) => (summary.revision = v));
  bindMeta("rev-date", (v) => (summary.date = v));
  bindMeta("rev-preparedBy", (v) => (summary.preparedBy = v));
  bindMeta("rev-section", (v) => (summary.section = v));

  $("rev-add-btn").addEventListener("click", () => {
    summary.rows.push(emptyRow());
    renderRevisionRows();
    scheduleSave();
  });
  $("rev-load-btn").addEventListener("click", onLoadRevisionFromScan);
  $("rev-export-btn").addEventListener("click", onExportNarrative);
  $("rev-export-json-btn").addEventListener("click", onExportContribution);
}

function bindMeta(id: string, apply: (v: string) => void): void {
  $<HTMLInputElement>(id).addEventListener("input", () => {
    apply(getInput(id));
    scheduleSave();
  });
}

const REV_TYPES = ["Addition", "Deletion", "Revision"];

function renderRevisionRows(): void {
  const host = $("rev-rows");
  host.textContent = "";
  $("rev-empty").style.display = summary.rows.length === 0 ? "block" : "none";

  summary.rows.forEach((row, i) => {
    const card = document.createElement("div");
    card.className = "rev-row";

    const head = document.createElement("div");
    head.className = "rev-row-head";

    const loc = document.createElement("input");
    loc.className = "control rev-loc";
    loc.placeholder = "Location (article)";
    loc.value = row.location;
    loc.addEventListener("input", () => {
      row.location = loc.value;
      scheduleSave();
    });

    const type = document.createElement("select");
    type.className = "control select rev-type";
    const options = REV_TYPES.slice();
    if (row.type && !options.includes(row.type)) options.push(row.type);
    options.forEach((t) => {
      const o = document.createElement("option");
      o.value = t;
      o.textContent = t;
      type.appendChild(o);
    });
    type.value = row.type || "Addition";
    type.addEventListener("change", () => {
      row.type = type.value;
      scheduleSave();
    });

    const del = document.createElement("button");
    del.className = "btn btn-link danger rev-del";
    del.type = "button";
    del.textContent = "×";
    del.title = "Delete row";
    del.addEventListener("click", () => {
      summary.rows.splice(i, 1);
      renderRevisionRows();
      scheduleSave();
    });

    head.appendChild(loc);
    head.appendChild(type);
    head.appendChild(del);

    const desc = document.createElement("textarea");
    desc.className = "control rev-desc";
    desc.rows = 2;
    desc.placeholder = "Description of change (narrative)";
    desc.value = row.description;
    desc.addEventListener("input", () => {
      row.description = desc.value;
      autoGrow(desc);
      scheduleSave();
    });

    const specLabel = document.createElement("span");
    specLabel.className = "rev-spec-label";
    specLabel.textContent = "Spec text (reference)";
    const spec = document.createElement("textarea");
    spec.className = "control rev-spec";
    spec.rows = 2;
    spec.value = row.specText;
    spec.addEventListener("input", () => {
      row.specText = spec.value;
      autoGrow(spec);
      scheduleSave();
    });

    card.appendChild(head);
    card.appendChild(desc);
    card.appendChild(specLabel);
    card.appendChild(spec);
    host.appendChild(card);
    autoGrow(desc);
    autoGrow(spec);
  });
}

/** Size a textarea to fit its content (capped by CSS max-height, then it scrolls). */
function autoGrow(el: HTMLTextAreaElement): void {
  el.style.height = "auto";
  el.style.height = el.scrollHeight + 2 + "px";
}

async function onLoadRevisionFromScan(): Promise<void> {
  const tmpl = selectedTemplate();
  if (!tmpl) return;
  const btn = $<HTMLInputElement>("rev-load-btn");
  btn.disabled = true;
  setRevStatus("Scanning…");
  try {
    const result = await previewSpecification(tmpl);
    const scanned = rowsFromEntries(result.entries);
    const existing = new Set(summary.rows.map((r) => r.location + " " + r.specText));
    let added = 0;
    for (const r of scanned) {
      const key = r.location + " " + r.specText;
      if (!existing.has(key)) {
        summary.rows.push(r);
        existing.add(key);
        added++;
      }
    }
    if (!summary.section) {
      const sec = sectionFromEntries(result.entries);
      if (sec) {
        summary.section = sec;
        setInput("rev-section", sec);
      }
    }
    if (!summary.revision) {
      const tag = $<HTMLInputElement>("tag-text").value.trim();
      if (tag) {
        summary.revision = tag;
        setInput("rev-revision", tag);
      }
    }
    renderRevisionRows();
    scheduleSave();
    setRevStatus(`Added ${added} row(s) from scan. The specification was not changed.`, "success");
  } catch (e) {
    setRevStatus("Scan failed: " + errorMessage(e), "error");
  } finally {
    btn.disabled = false;
  }
}

async function onExportNarrative(): Promise<void> {
  if (summary.rows.length === 0) {
    setRevStatus("Add at least one row before exporting.", "error");
    return;
  }
  const btn = $<HTMLInputElement>("rev-export-btn");
  btn.disabled = true;
  setRevStatus("Building narrative…");
  try {
    await exportNarrativeDocx(summary, documentBaseName());
    setRevStatus("Revision Change Narrative downloaded (.docx).", "success");
  } catch (e) {
    setRevStatus("Export failed: " + errorMessage(e), "error");
  } finally {
    btn.disabled = false;
  }
}

function onExportContribution(): void {
  if (summary.rows.length === 0) {
    setRevStatus("Add at least one row before exporting.", "error");
    return;
  }
  try {
    downloadContribution(summary, documentFileName(), documentBaseName());
    setRevStatus("Track Revisions JSON downloaded — import it in the Revit tool.", "success");
  } catch (e) {
    setRevStatus("Export failed: " + errorMessage(e), "error");
  }
}

function scheduleSave(): void {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    saveSummary(summary).catch((e) => setRevStatus("Could not save: " + errorMessage(e), "error"));
  }, 500);
}

function setRevStatus(message: string, kind: "info" | "error" | "success" = "info"): void {
  const el = $("rev-status");
  el.textContent = message;
  el.className = "status " + kind;
}

/* ---------------------------- Templates panel ---------------------------- */

function renderTemplateSelect(): void {
  const sel = $<HTMLSelectElement>("template-select");
  const prev = sel.selectedIndex;
  sel.innerHTML = "";
  templates.forEach((t, i) => {
    const opt = document.createElement("option");
    opt.value = String(i);
    opt.textContent = t.name + (t.appendTag ? "  · tag" : "");
    sel.appendChild(opt);
  });
  sel.selectedIndex = prev >= 0 && prev < templates.length ? prev : 0;
}

function renderTemplateList(): void {
  const list = $("template-list");
  list.innerHTML = "";
  templates.forEach((t, i) => {
    const row = document.createElement("div");
    row.className = "template-row";

    const info = document.createElement("div");
    info.className = "template-info";
    const name = document.createElement("div");
    name.className = "template-name";
    name.textContent = t.name;
    const meta = document.createElement("div");
    meta.className = "template-meta";
    meta.textContent = t.appendTag ? "Tag-append mode" : `Reformat · ${t.bodyFont} ${t.bodySize}pt`;
    info.appendChild(name);
    info.appendChild(meta);

    const actions = document.createElement("div");
    actions.className = "template-actions";
    const edit = document.createElement("button");
    edit.className = "btn btn-link";
    edit.textContent = "Edit";
    edit.addEventListener("click", () => openEditor(i));
    const del = document.createElement("button");
    del.className = "btn btn-link danger";
    del.textContent = "Delete";
    del.addEventListener("click", () => onDeleteTemplate(i));
    actions.appendChild(edit);
    actions.appendChild(del);

    row.appendChild(info);
    row.appendChild(actions);
    list.appendChild(row);
  });
}

function onAddTemplate(): void {
  templates.push(newTemplate());
  persist();
  openEditor(templates.length - 1);
}

function onDeleteTemplate(index: number): void {
  if (templates.length <= 1) {
    setStatus("Keep at least one template.", "error");
    return;
  }
  templates.splice(index, 1);
  if (editingIndex === index) closeEditor();
  persist();
}

function openEditor(index: number): void {
  editingIndex = index;
  const t = templates[index];
  $("editor-title").textContent = `Edit “${t.name}”`;
  setInput("ed-name", t.name);
  setChecked("ed-appendTag", t.appendTag);
  setInput("ed-bodyFont", t.bodyFont);
  setInput("ed-bodySize", String(t.bodySize));
  setInput("ed-headingFont", t.headingFont);
  setInput("ed-headingSize", String(t.headingSize));
  setInput("ed-paraSpacingPt", String(t.paraSpacingPt));
  setInput("ed-headingStyles", t.headingStyles.join(", "));
  setInput("ed-bodyColor", t.bodyColor);
  setInput("ed-headingColor", t.headingColor);
  setInput("ed-additionColor", t.additionColor);
  setInput("ed-deletionColor", t.deletionColor);
  $("editor-error").style.display = "none";
  $("template-editor").style.display = "block";
  $("templates-card").setAttribute("open", "");
  $("template-editor").scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function closeEditor(): void {
  editingIndex = null;
  $("template-editor").style.display = "none";
}

function onEditorSave(): void {
  if (editingIndex === null) return;
  const name = getInput("ed-name").trim();
  if (name.length === 0) {
    showEditorError("Name is required.");
    return;
  }
  const bodySize = Number(getInput("ed-bodySize"));
  const headingSize = Number(getInput("ed-headingSize"));
  const paraSpacingPt = Number(getInput("ed-paraSpacingPt"));
  if ([bodySize, headingSize, paraSpacingPt].some((n) => !isFinite(n) || n < 0)) {
    showEditorError("Sizes and spacing must be valid numbers.");
    return;
  }

  templates[editingIndex] = {
    name,
    appendTag: getChecked("ed-appendTag"),
    bodyFont: getInput("ed-bodyFont").trim() || "Arial",
    bodySize,
    headingFont: getInput("ed-headingFont").trim() || "Arial",
    headingSize,
    paraSpacingPt,
    headingStyles: parseStyleList(getInput("ed-headingStyles")),
    bodyColor: getInput("ed-bodyColor"),
    headingColor: getInput("ed-headingColor"),
    additionColor: getInput("ed-additionColor"),
    deletionColor: getInput("ed-deletionColor"),
  };
  persist();
  closeEditor();
  setStatus(`Saved template “${name}”.`, "success");
}

async function onImportFile(ev: Event): Promise<void> {
  const input = ev.target as HTMLInputElement;
  const file = input.files && input.files[0];
  input.value = ""; // allow re-importing the same file
  if (!file) return;
  try {
    templates = await importTemplatesFromFile(file);
    persist();
    setStatus(`Imported ${templates.length} template(s).`, "success");
  } catch (e) {
    setStatus("Import failed: " + errorMessage(e), "error");
  }
}

/* ------------------------------- helpers -------------------------------- */

function persist(): void {
  saveTemplates(templates);
  renderTemplateSelect();
  renderTemplateList();
  updateTagRow();
}

function setInput(id: string, value: string): void {
  $<HTMLInputElement>(id).value = value;
}
function getInput(id: string): string {
  return $<HTMLInputElement>(id).value;
}
function setChecked(id: string, value: boolean): void {
  $<HTMLInputElement>(id).checked = value;
}
function getChecked(id: string): boolean {
  return $<HTMLInputElement>(id).checked;
}
function showEditorError(message: string): void {
  const el = $("editor-error");
  el.textContent = message;
  el.style.display = "block";
}
function parseStyleList(value: string): string[] {
  return value
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}
function errorMessage(e: unknown): string {
  if (e instanceof Error) return e.message;
  return String(e);
}
