/* eslint-disable no-undef */
/*
 * Assembles the shareable BETA install folder.
 *
 * Fills the host into beta/manifest.template.xml (from PROD_URL/BETA_URL, same value
 * webpack bakes into the hosted files) and copies the installer + readme, producing
 * appPackage/beta-package/. Zip that folder and put it on a shared drive; testers run
 * Install-Spec-Formatter-Beta.cmd from there. Host the built dist/ at the same URL.
 *
 * Run after `npm run build`:  node scripts/build-beta.js   (or `npm run build:beta`)
 */

const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const betaDir = path.join(root, "beta");
const outDir = path.join(root, "appPackage", "beta-package");

let host = process.env.PROD_URL || process.env.BETA_URL || "https://REPLACE-ME.azurestaticapps.net/";
if (!host.endsWith("/")) host += "/";
if (host.includes("REPLACE-ME")) {
  console.warn(
    "build-beta: WARNING — host is the REPLACE-ME placeholder. Set PROD_URL (or BETA_URL) " +
      "to the URL where you host dist/, then rebuild before sharing."
  );
}

const template = fs.readFileSync(path.join(betaDir, "manifest.template.xml"), "utf8");
const manifest = template.replace(/\{\{HOST\}\}/g, host);

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, "manifest.xml"), manifest, "utf8");
for (const name of [
  "Install-Spec-Formatter-Beta.cmd",
  "Uninstall-Spec-Formatter-Beta.cmd",
  "BETA-README.txt",
]) {
  fs.copyFileSync(path.join(betaDir, name), path.join(outDir, name));
}

console.log("Beta package ready: " + path.relative(root, outDir));
console.log("  host baked into manifest.xml: " + host);
console.log("Next: host dist/ at that URL, then zip beta-package/ and share it.");
