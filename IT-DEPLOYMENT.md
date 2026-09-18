# Spec Formatter — Production Deployment (for IT)

**What it is:** A Microsoft Word add-in (Office.js web add-in) that formats construction
specification documents. It runs inside Word on Windows, Mac, and the web. It only reads/writes
the **currently open document** — no mailbox, no Microsoft Graph, no company data, no stored
credentials, and no external calls beyond loading its own files from the host URL.

**Why this request:** The beta was distributed as a *shared-folder add-in catalog*, which requires
each user to register a trusted catalog locally. That works for accounts with elevated permissions
but is blocked for standard users. Production should use **Centralized Deployment**, which IT
controls and which needs **zero per-user setup**.

---

## Recommended: Centralized Deployment (Microsoft 365 admin center)

Deploy once; the add-in appears on the Word ribbon automatically for assigned users.

**You'll need**
- A Microsoft 365 admin role that can deploy add-ins (**Global Administrator**, or a role with
  Integrated Apps / add-in deployment rights).
- The **app package** `spec-formatter-<version>.zip` (provided by Evie — contains the manifest + icons).
- The web files hosted at an **HTTPS URL** (see *Hosting* below).

**Steps**
1. Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**.
2. Upload the provided **app package (.zip)**.
3. Assign to a **pilot group / users** and accept the requested permission (Document read/write).
4. Users relaunch Word → **Spec Formatter** appears on the Home ribbon. No user action required.

Reference: Microsoft — "Deploy and manage Office Add-ins in the Microsoft 365 admin center."

---

## Hosting the web files

The add-in's files (HTML/JS/icons) must be served over **HTTPS with a valid certificate**; the
manifest points at that URL.
- **Beta** currently runs from GitHub Pages (external, temporary).
- **Production:** please host on company-controlled infrastructure — e.g. **Azure Static Web Apps**,
  Azure Storage static website, or an internal IIS/HTTPS site. Evie can supply the built files
  (`dist/`); once you provide the site URL, the package is rebuilt to point at it.

**Once IT provides the final hosting URL, Evie will supply the final app package built to point at it.**

---

## Key facts

| | |
|---|---|
| App type | Office Add-in (unified / JSON manifest) |
| App ID | `c1b06178-4084-4a0e-8f1e-7acddec19930` |
| Permission | `Document.ReadWrite` (open document only) |
| Platforms | Word for Windows, Mac, and Web |
| Updates | Re-host files (and re-upload the package if the manifest changed); users update automatically |

**Alternative:** if Centralized Deployment isn't available, the manifest can instead be published to
the **tenant / SharePoint App Catalog** (org-managed) — also no per-user setup. Same hosting
requirement applies.

---

*Contact: Evie Lutz — for the app package, the web files to host, and any questions.*
