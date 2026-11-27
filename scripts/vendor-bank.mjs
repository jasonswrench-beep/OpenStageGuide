// The shared core's bank reader, @openstage/bank (openstage-core, tag bank-v0.1.0, Apache-2.0), carried in src/openstage-bank so the pane
// and the program can load it. The program serves only its own .mjs files, and the server refuses any import in them that is not a
// literal relative path to another of them (scripts/build-snapshot.mjs). The package's files are copied with the three changes that rule
// needs, and nothing else:
//   - each file is named .mjs, and the imports between them say so;
//   - zip.js's fallback to node:zlib, for a place with no DecompressionStream, is taken out: Node 20 and every browser the program runs in
//     have DecompressionStream, so it is never reached;
//   - a first line says where the file came from.
// Run by `npm run assets` (scripts/prepare.mjs); tests/vendor-bank.test.mjs checks that the copy in src/openstage-bank is what this makes
// from the installed package. To take a newer release: install it, run `npm run assets`, and run the tests.
import fs from 'node:fs';
import path from 'node:path';

export const BANK_PACKAGE='@openstage/bank';
export const BANK_PACKAGE_VERSION='0.1.0';
export const BANK_TAG='bank-v0.1.0';
export const BANK_FILES=Object.freeze(['index','text','zip']);
export const VENDOR_DIR='src/openstage-bank';

const ZLIB_FALLBACK=`  const zlib = await import("node:zlib");
  try {
    return new Uint8Array(zlib.inflateRawSync(data, { maxOutputLength: limit }));
  } catch (e) {
    throw new ZipError(e.code === "ERR_BUFFER_TOO_LARGE" ? \`unpacks to more than \${limit} bytes\` : e.message);
  }
`;
const NO_FALLBACK=`  // (OpenStage Guide's copy: the fallback to node:zlib is taken out, since the program serves no import that is not a file of its own.)
  throw new ZipError("there is no DecompressionStream here to unpack it with");
`;

const once=(text,find,replace,what)=>{
  const at=text.indexOf(find);
  if(at<0||text.indexOf(find,at+1)>=0)throw new Error(`${BANK_PACKAGE} has changed: ${what} is not there exactly once. Read the new version before carrying it.`);
  return text.slice(0,at)+replace+text.slice(at+find.length);
};

/** The files of src/openstage-bank as they are made from the package's src folder: [{name,text}]. */
export function vendoredBank(packageDir){
  const manifest=JSON.parse(fs.readFileSync(path.join(packageDir,'package.json'),'utf8'));
  if(manifest.name!==BANK_PACKAGE||manifest.version!==BANK_PACKAGE_VERSION)throw new Error(`Expected ${BANK_PACKAGE} ${BANK_PACKAGE_VERSION} in ${packageDir}, found ${manifest.name} ${manifest.version}.`);
  return BANK_FILES.map(base=>{
    let text=fs.readFileSync(path.join(packageDir,'src',`${base}.js`),'utf8');
    if(text.includes('\r'))throw new Error(`${base}.js has CR line endings; the copy is kept with LF.`);
    const imports=[...text.matchAll(/ from "\.\/(index|text|zip)\.js";/g)].length;
    text=text.replace(/ from "\.\/(index|text|zip)\.js";/g,' from "./$1.mjs";');
    if(base==='index'&&imports!==3)throw new Error(`${BANK_PACKAGE} has changed: index.js imports ${imports} of its own files, not 3.`);
    if(base==='zip')text=once(text,ZLIB_FALLBACK,NO_FALLBACK,'the fallback to node:zlib in zip.js');
    const header=`// Copied from ${BANK_PACKAGE} ${BANK_PACKAGE_VERSION} (openstage-core, tag ${BANK_TAG}, Apache-2.0) by scripts/vendor-bank.mjs. Do not edit: see that file.\n`;
    return {name:`${base}.mjs`,text:header+text};
  });
}

/** Writes src/openstage-bank from the installed package. */
export function writeVendoredBank(root){
  const files=vendoredBank(path.join(root,'node_modules',...BANK_PACKAGE.split('/')));
  const dir=path.join(root,VENDOR_DIR);
  fs.mkdirSync(dir,{recursive:true});
  for(const f of files)fs.writeFileSync(path.join(dir,f.name),f.text);
  return files.map(f=>path.posix.join(VENDOR_DIR,f.name));
}
