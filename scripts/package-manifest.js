/* eslint-disable no-undef */
/*
 * Builds the Microsoft 365 app package (.zip) for Central Deployment.
 *
 * Reads the PRODUCTION build output in dist/ (so the manifest already carries the
 * PROD_URL host, not localhost) and zips the unified manifest together with the two
 * package-relative icons it references. Upload the resulting zip in the Microsoft 365
 * admin center → Integrated Apps → Upload custom apps.
 *
 * Run after `npm run build`:  node scripts/package-manifest.js   (or `npm run package`)
 */

const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");

const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist");
const outDir = path.join(root, "appPackage", "build");

// Files that go INTO the app package. The manifest's top-level `icons` are
// package-relative (assets/color.png, assets/outline.png); everything else in the
// manifest is an https URL fetched from the host, so it is NOT packaged.
const entries = [
  { disk: path.join(dist, "manifest.json"), zip: "manifest.json" },
  { disk: path.join(dist, "assets", "color.png"), zip: "assets/color.png" },
  { disk: path.join(dist, "assets", "outline.png"), zip: "assets/outline.png" },
];

function fail(msg) {
  console.error("package-manifest: " + msg);
  process.exit(1);
}

if (!fs.existsSync(dist)) {
  fail('dist/ not found — run "npm run build" first.');
}

// Guard against shipping a manifest still pointing at localhost.
const manifestText = fs.readFileSync(entries[0].disk, "utf8");
if (manifestText.includes("localhost:3000")) {
  fail(
    'dist/manifest.json still references localhost:3000 — build in production mode ' +
      "with PROD_URL set (npm run build)."
  );
}
if (manifestText.includes("REPLACE-ME.azurestaticapps.net")) {
  console.warn(
    "package-manifest: WARNING — manifest host is the REPLACE-ME placeholder. " +
      "Set PROD_URL to your real Azure Static Web Apps URL and rebuild before deploying."
  );
}

const version = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8")).version || "0.0.0";
const zip = new AdmZip();
for (const e of entries) {
  if (!fs.existsSync(e.disk)) fail("missing build file: " + e.disk);
  const dir = path.posix.dirname(e.zip);
  zip.addLocalFile(e.disk, dir === "." ? "" : dir);
}

fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, `spec-formatter-${version}.zip`);
zip.writeZip(outFile);
console.log("Wrote app package: " + path.relative(root, outFile));
console.log("Upload it in M365 admin center → Integrated Apps → Upload custom apps.");
