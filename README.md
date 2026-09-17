# Spec Formatter (Word add-in)

A Microsoft Word task-pane add-in that formats construction specification documents to an
architect's standard. It is a port of the original VBA macro (`modFormatter.bas`,
`modTemplates.bas`, `modPicker.bas`) into a modern Office.js add-in — one install, a real
UI, no per-document VBA importing.

## What it does

Pick an **architect template** in the task pane, then click **Format Specification**. Two
modes, driven by each template's *tag-append* flag (identical rules to the macro):

- **Tag-append mode** (e.g. *Rowell Brokaw Architects*): scans the document, treats each
  continuous **bold** run as an *addition* and each **bold + strikethrough** run as a
  *deletion*, skips headings, and appends a bold tag (e.g. `(ASI-1)`) after each run.
- **Reformat mode**: applies heading vs. body fonts/sizes/colors, recolors
  addition/deletion runs, and sets paragraph spacing.

**Templates are editable in the UI** (Templates panel): add / edit / delete, with color
pickers and a tag-append toggle. The set is saved in the task pane's local storage and can
be shared via **Export / Import** (a JSON file).

### Safety

- **Save a copy first** downloads a `<name>_original.docx` backup to your Downloads folder
  (the browser-based add-in sandbox can't write next to the original file the way the macro
  did, so this is the equivalent).
- All edits run in a single batch, so **one Ctrl+Z** reverts a formatting run.

## Run it (dev / sideload)

Prerequisites: Node 22, Word on Windows or Mac.

```bash
npm install
npm start
```

`npm start` builds, starts the HTTPS dev server on `https://localhost:3000`, and sideloads
the add-in into Word. Look for the **Spec Formatter** button on the Home ribbon. To stop:

```bash
npm stop
```

Manual sideload (alternative): run `npm run dev-server`, then in Word use
*Home ▸ Add-ins ▸ More Add-ins ▸ My Add-ins ▸ Upload My Add-in* and pick
`appPackage/manifest.json`.

## Deploy (production)

Distribution = hosting the built files on real HTTPS + pushing the manifest firm-wide.
It does **not** change the Office.js security sandbox: a deployed add-in still cannot read
the `10 CAD` network folder — the JSON-contribution bridge to Track Revisions stays.

Target: **Azure Static Web Apps** (hosting) + **Central Deployment** (M365 admin center).

1. **Build + package** with the production host baked in (must end in `/`):

   ```bash
   PROD_URL=https://<your-swa-name>.azurestaticapps.net/ npm run build:package
   ```

   This writes `dist/` (the web files, with all `localhost:3000` URLs rewritten to `PROD_URL`)
   and `appPackage/build/spec-formatter-<version>.zip` (the app package: prod `manifest.json`
   + `assets/color.png` + `assets/outline.png`). The packager refuses to build if the manifest
   still points at localhost, and warns if `PROD_URL` was left as the placeholder.

2. **Host `dist/` on Azure Static Web Apps.** Easiest is the SWA CLI:

   ```bash
   npx @azure/static-web-apps-cli deploy ./dist --env production
   ```

   (or connect the repo to an SWA resource for CI). `staticwebapp.config.json` sets caching and
   MIME types and keeps `manifest.json` uncached. Use the resulting `*.azurestaticapps.net` URL as
   `PROD_URL` above (or a custom domain). Real cert required — the Office webview rejects
   self-signed certs.

3. **Deploy the manifest via Central Deployment.** In the **Microsoft 365 admin center →
   Settings → Integrated apps → Upload custom apps**, upload
   `appPackage/build/spec-formatter-<version>.zip`, then assign it to the users/groups who
   should get it. It appears on the Word Home ribbon (Windows/Mac/web) automatically — no
   per-user sideloading. Roll out an update by bumping `version` in the manifest, rebuilding,
   re-hosting `dist/`, and re-uploading the package.

To validate the manifest at any time: `npm run validate`.

### Beta testing (one-click, no admin)

Before firm-wide Central Deployment, hand a few testers a one-click installer. This uses a
classic **XML** manifest and a per-user **shared-folder catalog** (no admin, no PowerShell).

1. **Host `dist/` at a beta URL.** Any free HTTPS static host works (Netlify, Cloudflare Pages,
   Azure SWA free) — no IT needed. Build + assemble the tester package with that host baked in:

   ```bash
   PROD_URL=https://<your-beta-host>/ npm run build:beta
   ```

   Deploy `dist/` to that host, and you'll get `appPackage/beta-package/` containing
   `manifest.xml`, `Install-Spec-Formatter-Beta.cmd`, an uninstaller, and `BETA-README.txt`.

2. **Put `beta-package/` on a NETWORK SHARE** (a `\\server\share` path) your testers can reach.
   Office add-in catalogs **must be a UNC network share** — a local folder (`C:\...`) is rejected
   with *"Please add or enable add-in catalogs from the Trust Center"*. The installer registers
   *its own folder* as the catalog, so run it from the share.

3. **Testers run `Install-Spec-Formatter-Beta.cmd`** once **from the share** — it registers the
   catalog automatically (no Trust Center needed). Restart Word, then
   *Home ▸ Add-ins ▸ **Advanced…** ▸ Shared Folder ▸ Spec Formatter (Beta) ▸ Add*. The Home-ribbon
   button appears. (The new Add-ins panel only shows Store/Developer add-ins up front; the
   **Advanced…** link opens the dialog with the Shared Folder section — this is the non-obvious step.)
   `BETA-README.txt` has the tester-facing version.

Manual alternative (no script): *File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Trusted
Add-in Catalogs*, paste the share's **UNC path**, **Add catalog**, tick **Show in Menu**, OK, restart
Word, then add it via *Advanced… ▸ Shared Folder*.

Troubleshooting: if an old preview button is stuck on the ribbon, remove it via *File ▸ Options ▸
Customize Ribbon*. Requirements: Word desktop on Windows, on the network. Ship an update by
re-hosting `dist/` (push to `main`) and, if the manifest changed, re-sharing `beta-package/`.

## Verify

Use the sample redline in `C:\Users\elutz\Desktop\Conspectus\`:
`220500 General Plumbing Provisions TC (4).docx` (its `..._format.docx` sibling is the
macro's own output — a golden reference). Run **Rowell Brokaw** with tag `(ASI-1)` and
confirm each bold run and each bold+strike run gets a trailing bold `(ASI-1)`.

## Source layout

- `src/core/templates.ts` — `ArchitectTemplate` type + seed templates (ported from `modTemplates.bas`).
- `src/core/headings.ts` — heading detection (ported from `IsHeading`/`MatchesHeadingPattern`).
- `src/core/runs.ts` — pure run classification + tag-insertion state machine (from `AppendTagsToDocument`).
- `src/core/format.ts` — the Office.js engine that applies both modes to the live document.
- `src/core/store.ts` — template persistence + JSON export/import.
- `src/core/backup.ts` — the "save a copy first" download.
- `src/taskpane/` — the task pane UI (HTML/CSS + controller).
- `appPackage/manifest.json` — the add-in manifest.

## Heading detection

A paragraph is treated as a heading if **any** of these hold:

- its paragraph **style** is one of the template's **heading styles** (default `SCT`, `PRT`,
  `ART` — the CSI SectionFormat title styles), editable per template in the UI;
- its style is a Word built-in `Heading N`; or
- its **text** matches a pattern: `SECTION #`, `PART #`, `#.##`, or `A.`.

The style-name check matters because in these specs the article numbers (`1.01`, `A.`) are
**auto-generated list numbers**, not literal text, so text patterns alone only catch the
`SECTION #####` line. Matching on the CSI styles makes reformat mode correctly style all the
title lines. (The original macro only did the text-pattern + built-in-Heading checks.)

## Known limitation

- Run detection is at **word granularity**; a word with mixed bold/strike is treated by its
  dominant state. Real spec redlines bold whole words/phrases, so this matches in practice.
