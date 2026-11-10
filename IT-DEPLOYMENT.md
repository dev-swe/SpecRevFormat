# Spec Formatter — Production Deployment (for IT)

**What it is:** A Microsoft Word add-in (Office.js web add-in) that formats construction
specification documents. It is designed for Word on Windows, Mac, and the web, and has been
tested on **Word for Windows (desktop)** only. It only reads/writes the **currently open
document** — no mailbox, no Microsoft Graph, no company data, no stored credentials, and no
telemetry. Its only network requests are for its own files from the host URL and the Office.js
library from Microsoft's CDN (see *Network requirements*).

**Why this request:** The beta was distributed as a *shared-folder add-in catalog*, which requires
each user to register a trusted catalog locally. That works for accounts with elevated permissions
but is blocked for standard users. Production should use **Centralized Deployment**, which IT
controls and which needs **zero per-user setup**.

---

## How the pieces fit together

There are two separate things to deploy. Microsoft 365 does **not** host the add-in's code.

| Piece | What it is | Where it lives | Who updates it |
|---|---|---|---|
| **App package** (`.zip`) | The manifest plus two icons. It tells Word the add-in exists, where its pages are, and who gets the ribbon button. | **Microsoft 365 admin center** (Integrated apps) | IT uploads it |
| **Web files** (`dist/`) | The task pane pages, JavaScript, CSS, ribbon icons, and one Word template (`submittal-template.docx`). Word loads these every time the pane opens. | **An HTTPS web site** (see *Hosting*) | Evie, by merging to `main` (automatic deploy) |

The manifest contains the web site's URL, so the package must be built **for the final host URL**.
The manifest copy that is published on the web site is only a build input. Word uses the copy held
in the admin center.

---

## Recommended: Centralized Deployment (Microsoft 365 admin center)

Deploy once; the add-in appears on the Word ribbon automatically for assigned users.

**You'll need**
- A Microsoft 365 admin role that can deploy add-ins (**Global Administrator**, or a role with
  Integrated Apps / add-in deployment rights).
- The **app package** `spec-formatter-<version>.zip` (provided by Evie — contains the manifest + icons).
- The web files hosted at an **HTTPS URL** (see *Hosting* below). The site must be live **before**
  users receive the add-in.

**Steps**
1. Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**.
2. Choose the app type **Office Add-in** and upload the provided **app package (.zip)**.
3. Assign to a **pilot group / users** and accept the requested permission (Document read/write).
4. Review and **Deploy**.
5. Users fully close and relaunch Word → **Spec Formatter** appears on the **Home** ribbon. No user
   action is required. Propagation can take several hours; it is not instant.

Reference: Microsoft — "Deploy and manage Office Add-ins in the Microsoft 365 admin center."

**Status of testing:** the manifest has been sideloaded through Microsoft 365 in Word for Windows and
was accepted and ran correctly. It has **not yet been tested through the admin center upload**, so
please start with a small pilot group. If the admin center rejects the package, tell Evie the exact
message; a classic XML manifest (`beta/manifest.template.xml`) can be supplied instead.

---

## Hosting the web files

The add-in's files (HTML/JS/CSS/icons) must be served over **HTTPS with a valid certificate**; the
manifest points at that URL.

- **Beta** currently runs from GitHub Pages (external, temporary).
- **Production:** please host on company-controlled infrastructure — e.g. **Azure Static Web Apps**,
  Azure Storage static website, or an internal IIS/HTTPS site. Evie can supply the built files
  (`dist/`); once you provide the site URL, the package is rebuilt to point at it.

**Hosting requirements**
- **HTTPS** with a certificate **trusted by users' PCs** (publicly trusted, or issued by a company CA
  that PCs already trust). A self-signed certificate shows "The content is blocked because it isn't
  signed by a valid security certificate."
- **Reachable by every user.** An internal-only site works only for people on the network or VPN.
- **Anonymous access.** Word loads the pages without a sign-in. Do not put the site behind SSO,
  Conditional Access, or an authenticating proxy.
- **Static files only.** No server code, database, or API is needed.
- **MIME types:** serve `.html`, `.js`, `.css`, `.json`, `.png`, and `.docx`
  (`application/vnd.openxmlformats-officedocument.wordprocessingml.document`). On IIS, add the
  `.docx` type if missing, or the template download fails.
- **Same origin.** Everything is served from one host name, so no CORS configuration is needed.
- **Caching:** the JavaScript and CSS file names change with every build, so they can be cached
  long-term. The HTML pages and `manifest.json` should be cached for a short time (about 10 minutes
  or less) so updates reach users promptly. For Azure Static Web Apps the repository's
  `staticwebapp.config.json` already sets `no-store` on `manifest.json` and long-term caching on
  `/assets/*`; it does not set a cache time for the HTML pages, so confirm they are not cached for long.
- **Stable URL.** Changing the host later means rebuilding the package and re-uploading it
  (see *If the host URL changes*).

### Network requirements

Allow outbound HTTPS from users' PCs to:
- the **hosting site** chosen above, and
- **`appsforoffice.microsoft.com`** — Microsoft's CDN for the Office.js library, which every Office
  add-in loads.

The add-in makes no other network requests and sends no data anywhere. All processing happens inside
Word on the user's PC.

**Once IT provides the final hosting URL, Evie will supply the final app package built to point at it.**

---

## Key facts

| | |
|---|---|
| App type | Office Add-in (unified / JSON manifest) |
| App ID | `c1b06178-4084-4a0e-8f1e-7acddec19930` |
| Permission | `Document.ReadWrite` (open document only) |
| Platforms | Designed for Word for Windows, Mac, and Web; tested on Word for Windows desktop |
| Ribbon | One button, **Spec Formatter**, on the **Home** tab |
| Data stored | Architect templates and revision-summary fields, in each user's local browser storage for the add-in (per user, per PC, per host URL) |
| Telemetry | None |
| Downloads | The add-in can download files to the user's Downloads folder (document backup copies, exported reviews, CSV logs). It cannot write anywhere else |
| Updates | Web changes: re-host files, no admin action. Manifest changes: re-upload the package (see below) |

---

## Rollout plan (suggested)

1. **Pilot:** assign to 3–5 users. Ask them to format a copy of a real specification and confirm the
   ribbon button, the task pane, formatting, and the backup download.
2. **Wider group:** expand to the engineering group once the pilot is clean.
3. **Retire the beta** (see *Replacing the beta*).
4. **Everyone:** assign to the rest of the intended users.

Users can be removed or the add-in withdrawn at any time: Integrated apps → select
**Spec Formatter** → **Remove** (or change the assigned users). Users keep any documents they
already formatted; formatting is saved in the document, not the add-in.

---

## Keeping it up to date

There are two kinds of updates.

| What changed | What happens | Admin center action |
|---|---|---|
| **Code or screens** (features, fixes, wording) | Evie merges the change to `main`; the web files are re-deployed automatically. Users get it the next time the task pane opens (allow up to about 10 minutes; a Word restart guarantees it). | **None** |
| **The manifest** (ribbon button, name, icons, permissions, host URL) | Evie raises the `version` in the manifest, builds a new package, and sends it to IT. | **Update:** Integrated apps → **Spec Formatter** → **Update** → upload the new `.zip` |

There is no staging environment: a change that reaches `main` reaches every user. Changes are tested
in Word on a sideloaded copy and reviewed in a pull request before they are merged.

**Rolling back**
- **Code:** revert the pull request on `main`. The site redeploys the previous version.
- **Manifest:** upload the previous package `.zip` through **Update**.

---

## If the host URL changes

Moving from one host (for example GitHub Pages) to another (for example Azure Static Web Apps) requires:
1. Deploying the web files to the new host and confirming they load in a browser.
2. Evie building a new package for the new URL (`PROD_URL=<new URL> npm run build:package`).
3. IT uploading it through **Update**.

**Users' saved templates do not move.** Templates are stored per host name, so after the switch each
user starts with the default templates. Before switching, ask users who made custom templates to use
**Templates → Export** (a JSON file) and **Import** it afterwards.

---

## Replacing the beta

The beta (Spec Formatter (Beta)) is installed from a shared-folder catalog and uses a **different app
ID** from production. A user who has both will see **two buttons**. When production reaches a
tester, ask them to run `Uninstall-Spec-Formatter-Beta.cmd` from the beta folder, then restart Word.

---

## Alternative

If Centralized Deployment isn't available, the manifest can instead be published to the
**tenant / SharePoint App Catalog** (org-managed) — also no per-user setup. Same hosting
requirement applies.

---

## Open items before production

- **Final hosting URL** (and who owns the site, its certificate renewal, and deployment access).
- **Privacy and terms pages.** The manifest's website, privacy, and terms URLs currently point at the
  add-in's home page. Replace with real pages if your admin center or policy requires them.
- **Admin center upload** has not yet been tested end to end (see *Status of testing*).

---

*Contact: Evie Lutz — for the app package, the web files to host, and any questions.*
