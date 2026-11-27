// The shared core's bank reader is carried in src/openstage-bank (scripts/vendor-bank.mjs says why and how). These hold that copy to the
// installed release, so it cannot drift or be edited by hand, and check that the program's server will serve it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {vendoredBank,BANK_PACKAGE,BANK_PACKAGE_VERSION,BANK_TAG,BANK_FILES,VENDOR_DIR} from '../scripts/vendor-bank.mjs';
import {createBuildSnapshot} from '../scripts/build-snapshot.mjs';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const packageDir=path.join(root,'node_modules',...BANK_PACKAGE.split('/'));

test('src/openstage-bank is the installed @openstage/bank, copied by scripts/vendor-bank.mjs and not changed since',()=>{
  const made=vendoredBank(packageDir);
  assert.deepEqual(fs.readdirSync(path.join(root,VENDOR_DIR)).sort(),BANK_FILES.map(f=>`${f}.mjs`).sort(),'nothing else is in the folder');
  for(const f of made)assert.equal(fs.readFileSync(path.join(root,VENDOR_DIR,f.name),'utf8'),f.text,`${f.name} is what npm run assets makes`);
});

test('the package is the core\'s published release, named in package.json by its tag',()=>{
  const pkg=JSON.parse(fs.readFileSync(path.join(root,'package.json'),'utf8'));
  assert.equal(pkg.devDependencies[BANK_PACKAGE],`https://github.com/jasonswrench-beep/openstage-core/releases/download/${BANK_TAG}/openstage-bank-${BANK_PACKAGE_VERSION}.tgz`);
  assert.equal(pkg.dependencies?.[BANK_PACKAGE],undefined,'nothing loads it from node_modules at run time');
});

test('the copy differs from the package only by its file names, the fallback to node:zlib, and a first line',()=>{
  for(const f of vendoredBank(packageDir)){
    const original=fs.readFileSync(path.join(packageDir,'src',f.name.replace(/\.mjs$/,'.js')),'utf8');
    const [first,...rest]=f.text.split('\n');
    assert.match(first,/^\/\/ Copied from @openstage\/bank 0\.1\.0 \(openstage-core, tag bank-v0\.1\.0, Apache-2\.0\)/);
    const back=rest.join('\n').replace(/ from "\.\/(index|text|zip)\.mjs";/g,' from "./$1.js";');
    if(f.name!=='zip.mjs'){assert.equal(back,original);continue;}
    assert.doesNotMatch(back,/\bimport\s*\(/,'no import() is left');
    const removed=original.split('\n').filter(l=>!back.split('\n').includes(l));
    assert.deepEqual(removed.map(l=>l.trim()),['const zlib = await import("node:zlib");','try {','return new Uint8Array(zlib.inflateRawSync(data, { maxOutputLength: limit }));','} catch (e) {','throw new ZipError(e.code === "ERR_BUFFER_TOO_LARGE" ? `unpacks to more than ${limit} bytes` : e.message);'],'only the fallback to node:zlib is taken out');
  }
});

test('a newer package whose zip.js has changed is refused until someone reads it',()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-bank-'));
  try{
    fs.mkdirSync(path.join(dir,'src'));
    fs.copyFileSync(path.join(packageDir,'package.json'),path.join(dir,'package.json'));
    for(const f of BANK_FILES)fs.copyFileSync(path.join(packageDir,'src',`${f}.js`),path.join(dir,'src',`${f}.js`));
    fs.writeFileSync(path.join(dir,'src','zip.js'),fs.readFileSync(path.join(dir,'src','zip.js'),'utf8').replace('await import("node:zlib")','await import("node:zlib" + "")'));
    assert.throws(()=>vendoredBank(dir),/@openstage\/bank has changed: the fallback to node:zlib in zip\.js is not there exactly once/);
    fs.writeFileSync(path.join(dir,'package.json'),JSON.stringify({name:BANK_PACKAGE,version:'0.2.0'}));
    assert.throws(()=>vendoredBank(dir),/Expected @openstage\/bank 0\.1\.0/);
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});

test('the program\'s server takes the copy: its imports are its own files, and it is served with the rest of the source',async()=>{
  const snapshot=await createBuildSnapshot(root);
  for(const f of BANK_FILES)assert.ok(snapshot.read(`/src/openstage-bank/${f}.mjs`),`${f}.mjs is served`);
  assert.match(snapshot.read('/src/guide-bank.mjs').toString(),/from '\.\/openstage-bank\/index\.mjs'/);
});
