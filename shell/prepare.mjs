// Gathers what the window carries with it, into src-tauri/resources: the program's source (it runs from source, so there is nothing to
// build), the one package it needs at run time (acorn, which checks the source before the server serves it), and the Node that runs it,
// so that a computer needs nothing installed.
//
//   resources/app/{package.json,package-lock.json,src,program,public,scripts/*.mjs the server needs,node_modules/acorn}
//   resources/node/node.exe
//
// Run by `npm run build` and `npm run dev`. The folder it fills is not kept in the repository.

import { cpSync, existsSync, mkdirSync, rmSync, statSync } from "node:fs";
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
copy("public");
for (const script of ["project-server.mjs", "build-snapshot.mjs", "http-handler.mjs"]) copy(join("scripts", script));
copy(join("node_modules", "acorn"));
copy("LICENSE");
copy("NOTICE");

// The Node that is running this is the Node the program was written against, and it is the one carried.
mkdirSync(join(out, "node"), { recursive: true });
const nodeName = process.platform === "win32" ? "node.exe" : "node";
cpSync(process.execPath, join(out, "node", nodeName));

const size = (path) => (statSync(path).size / 1024 / 1024).toFixed(1);
console.log(`Gathered the program and Node ${process.version} (${size(join(out, "node", nodeName))} MB) into ${out}`);
