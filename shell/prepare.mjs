// Gathers what the window carries with it, into src-tauri/resources: the program's source (it runs from source, so there is nothing to
// build), the one package it needs at run time (acorn, which checks the source before the server serves it), and the Node that runs it,
// so that a computer needs nothing installed.
//
//   resources/app/{package.json,package-lock.json,src,program,public,scripts/*.mjs the server needs,node_modules/acorn}
//   resources/node/node.exe
//
// Run by `npm run build` and `npm run dev`. The folder it fills is not kept in the repository.

import { cpSync, existsSync, mkdirSync, rmSync, statSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repo = join(here, "..");
const out = join(here, "src-tauri", "resources");

rmSync(out, { recursive: true, force: true });
const copy = (from, to = from) => {
  const source = join(repo, from);
  if (!existsSync(source)) throw new Error(`${from} is not where it should be: ${source}`);
  cpSync(source, join(out, "app", to), { recursive: true });
};
copy("package.json");
copy("package-lock.json");
copy("src");
copy("program");
// The public folder is the add-in's pane. Its file window (picker.html) and the PDF library it uses (public/vendor/pdf*) are for Word's dialog,
// which the program does not have: they are not carried, and the program's server does not serve them.
const addInOnly = (path) => /[\\/]picker\.html$/.test(path) || /[\\/]vendor[\\/](pdf\.|pdfjs-)/.test(path);
cpSync(join(repo, "public"), join(out, "app", "public"), { recursive: true, filter: (source) => !addInOnly(source) });
// Every script the server loads, and the ones those load (tests/shell-prepare.test.mjs checks that this list is complete).
for (const script of ["project-server.mjs", "build-snapshot.mjs", "http-handler.mjs", "security-headers.mjs", "updates.mjs"]) copy(join("scripts", script));
copy(join("node_modules", "acorn"));
copy("LICENSE");
copy("NOTICE");

// The release this window is (OPENSTAGE_GUIDE_RELEASE, such as v0.1.0-alpha.1; tools/build_installer.py sets it), written beside the program's files for
// Check for updates to compare with the published releases. A window built without one has none, and Check for updates then says which release is
// the newest.
const release = (process.env.OPENSTAGE_GUIDE_RELEASE ?? "").trim();
if (release) {
  if (!/^v\d{1,6}\.\d{1,6}\.\d{1,6}(?:-(?:alpha|beta|rc)\.\d{1,6})?$/.test(release)) throw new Error(`OPENSTAGE_GUIDE_RELEASE is not a release tag like v0.1.0-alpha.1: ${release}`);
  writeFileSync(join(out, "app", "release.txt"), `${release}
`);
}

// The Node that is running this is the Node the program was written against, and it is the one carried.
mkdirSync(join(out, "node"), { recursive: true });
const nodeName = process.platform === "win32" ? "node.exe" : "node";
cpSync(process.execPath, join(out, "node", nodeName));

const size = (path) => (statSync(path).size / 1024 / 1024).toFixed(1);
console.log(`Gathered the program and Node ${process.version} (${size(join(out, "node", nodeName))} MB) into ${out}`);
