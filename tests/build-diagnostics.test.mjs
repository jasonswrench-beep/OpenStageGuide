import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {createHash} from 'node:crypto';
import {createBuildSnapshot} from '../scripts/build-snapshot.mjs';
import {createHandler} from '../scripts/http-handler.mjs';
import {inspectBuild,moduleIdentity,moduleAsset} from '../src/diagnostics.mjs';
async function fixture(action){
  const root=await fs.mkdtemp(path.join(os.tmpdir(),'openstage-build-'));
  try{
    const files={'package.json':'{"version":"0.1.0","private":"SECRET USER TOKEN"}','package-lock.json':'{}','scripts/project-server.mjs':'//server','scripts/build-snapshot.mjs':'//builder','scripts/http-handler.mjs':'//handler','scripts/security-headers.mjs':'//headers','public/taskpane.html':'<link href="/style.css"><script src="/vendor/jszip.min.js"></script><script type="module" src="/src/app.mjs"></script>','public/style.css':'body{}','public/help.html':'<html/>','public/picker.html':'<html/>','public/vendor/jszip.min.js':'//zip','public/vendor/pdf.min.js':'//pdf','public/vendor/pdf.worker.min.js':'//pdfworker','public/vendor/pdfjs-LICENSE.txt':'licence','public/assets/cue.png':'image','public/assets/private.docx':'PRIVATE DOCUMENT','public/private.json':'PRIVATE SETTINGS','src/app.mjs':"import './model.mjs'; import './word.mjs'; import './diagnostics.mjs';",'src/model.mjs':'export const model=1;','src/word.mjs':"import './document-store.mjs';",'src/document-store.mjs':'export const store=1;','src/diagnostics.mjs':'export const diagnostic=1;'};
    for(const [name,body] of Object.entries(files)){const full=path.join(root,name);await fs.mkdir(path.dirname(full),{recursive:true});await fs.writeFile(full,body);}
    await action(root);
  }finally{
    assert.equal(path.dirname(path.resolve(root)),path.resolve(os.tmpdir()));assert.ok(path.basename(root).startsWith('openstage-build-'));await fs.rm(root,{recursive:true,force:true});
  }
}
const url=(id,name='app.mjs')=>`https://localhost:3843/build/${id}/src/${name}`;
const registry=id=>['app','model','word','document-store','diagnostics'].map(n=>url(id,n+'.mjs'));
async function request(snapshot,url,method='GET'){
  const result={};await createHandler(snapshot)({url,method},{writeHead(code,headers){result.status=code;result.headers=headers;},end(body){result.body=body===undefined?'':String(body);}});return result;
}
const fetcher=snapshot=>async path=>{const result=await request(snapshot,path);return {ok:result.status===200,headers:{get:()=>null},text:async()=>result.body};};
test('snapshot pins HTML and executing module stamps, hashes served bytes and excludes document/settings files',async()=>fixture(async root=>{
  const s=await createBuildSnapshot(root),again=await createBuildSnapshot(root);assert.equal(s.buildId,again.buildId);assert.match(s.read('/taskpane.html').toString(),new RegExp(`/build/${s.buildId}/src/app.mjs`));assert.match(s.read('/src/model.mjs').toString(),/runtimeModules.*import.meta.url/);
  for(const file of s.record.served)assert.equal(createHash('sha256').update(s.read(file.path)).digest('hex'),file.sha256);
  for(const left of ['/picker.html','/vendor/pdf.min.js','/vendor/pdf.worker.min.js','/vendor/pdfjs-LICENSE.txt'])assert.equal(s.read(left),null,`${left} is the Word add-in's and is not served by the program`);
  assert.equal(s.read('/assets/private.docx'),null);assert.equal(s.read('/private.json'),null);assert.equal((await request(s,'/package.json')).status,404);assert.ok(!JSON.stringify(s.record).includes('PRIVATE DOCUMENT'));
  const copied=s.read('/src/model.mjs');copied.fill(0);assert.match(s.read('/src/model.mjs').toString(),/model=1/);
  const key=Symbol.for('openstage.guide.runtimeModules'),previous=globalThis[key];globalThis[key]=new Set();
  try{const name=path.join(root,'executed-model.mjs');await fs.writeFile(name,s.read('/src/model.mjs'));const moduleURL=pathToFileURL(name).href;await import(moduleURL);assert.ok(globalThis[key].has(moduleURL));}finally{if(previous===undefined)delete globalThis[key];else globalThis[key]=previous;}
}));
test('observed pinned core and matching validated record report matched without claiming native acceptance or exposing content',async()=>fixture(async root=>{
  const s=await createBuildSnapshot(root),result=await inspectBuild({entryURL:url(s.buildId),registry:registry(s.buildId).map(u=>u+'?PRIVATE QUERY'),fetcher:fetcher(s),runtime:{mode:'word',host:'Word',wordApi13:true,documentPath:'C:\\PRIVATE DOC.docx',project:{secret:'PRIVATE NOTES'},username:'PRIVATE USER',hostVersion:'C:\\PRIVATE USER'}});
  assert.equal(result.status,'matched');assert.equal(result.loadedRecordValidated,true);assert.equal(result.nativeManifestVerified,false);assert.equal(result.loadedBuildId,s.buildId);assert.doesNotMatch(JSON.stringify(result),/PRIVATE|credentials|username|"project"|documentPath/);assert.ok(result.buildFiles.some(f=>f.path==='package.json'));assert.ok(result.buildFiles.every(f=>!f.path.startsWith('public/assets')));
}));
test('server snapshot changes cannot relabel older executing code as current, even when its old metadata is unavailable',async()=>fixture(async root=>{
  const old=await createBuildSnapshot(root);await fs.writeFile(path.join(root,'src/model.mjs'),'export const model=2;');const current=await createBuildSnapshot(root);assert.notEqual(old.buildId,current.buildId);assert.match(old.read('/src/model.mjs').toString(),/model=1/);assert.equal((await request(current,`/build/${old.buildId}/src/model.mjs`)).status,404);
  const result=await inspectBuild({entryURL:url(old.buildId),registry:registry(old.buildId),fetcher:fetcher(current)});assert.equal(result.status,'stale');assert.equal(result.loadedBuildId,old.buildId);assert.equal(result.servedBuildId,current.buildId);assert.equal(result.loadedRecordValidated,false);
  const stillRunningOld=await inspectBuild({entryURL:url(old.buildId),registry:registry(old.buildId),fetcher:fetcher(old)});assert.equal(stillRunningOld.status,'matched');assert.equal(stillRunningOld.workspaceBuildVerified,false);assert.notEqual(stillRunningOld.loadedBuildId,current.buildId);assert.match(stillRunningOld.reason,/intended source build/);
}));
test('mixed executing versions, unpinned imports and incomplete core never become matched from fresh server metadata',async()=>fixture(async root=>{
  const s=await createBuildSnapshot(root);let calls=0;const fetch=async p=>{calls++;return fetcher(s)(p);};
  for(const modules of [[...registry(s.buildId),url('b'.repeat(64),'model.mjs')],[...registry(s.buildId),'https://localhost:3843/src/legacy.mjs?PRIVATE'],registry(s.buildId).slice(1)]){
    const r=await inspectBuild({entryURL:url(s.buildId),registry:modules,fetcher:fetch});assert.notEqual(r.status,'matched');assert.equal(r.loadedBuildId,s.buildId);assert.doesNotMatch(JSON.stringify(r),/PRIVATE/);
  }assert.equal(calls,0);const missing=await inspectBuild({entryURL:'file:///PRIVATE/path/app.mjs',registry:[],fetcher:fetch});assert.equal(missing.status,'unknown');assert.doesNotMatch(JSON.stringify(missing),/PRIVATE/);
}));
test('invalid/missing metadata stays unknown and mock HTTP rejects unsupported/private paths without live networking',async()=>fixture(async root=>{
  const s=await createBuildSnapshot(root),bad=async p=>({ok:true,headers:{get:()=>null},text:async()=>p==='/health'?JSON.stringify({app:'OpenStage Guide',buildId:s.buildId}):JSON.stringify({...s.record,inputs:[...s.record.inputs,{path:'PRIVATE.docx',sha256:'0'.repeat(64),bytes:4}]})});
  assert.equal((await inspectBuild({entryURL:url(s.buildId),registry:registry(s.buildId),fetcher:bad})).status,'unknown');
  assert.equal((await request(s,'/health','POST')).status,405);assert.equal((await request(s,'/src/%5cprivate.docx')).status,403);assert.equal((await request(s,'/assets/private.docx')).status,404);assert.equal((await request(s,'/health','HEAD')).body,'');assert.equal((await request(s,'/src/app.mjs')).headers['Cache-Control'],'no-store');
}));
test('snapshot refuses imports escaping the pinned graph, including nonliteral dynamic imports',async()=>fixture(async root=>{
  for(const code of ["import '/src/model.mjs';","import 'https://example.invalid/code.mjs';","const target='./model.mjs';import(target);","export {model} from '../outside.mjs';"]){await fs.writeFile(path.join(root,'src/app.mjs'),code);await assert.rejects(createBuildSnapshot(root),/pinned|literal/);}
}));
test('asset references inherit executing build identity and never use document/query parameters',()=>{
  const id='a'.repeat(64);assert.equal(moduleAsset('assets/icons/Cue.png',url(id,'word.mjs')),`/build/${id}/assets/icons/Cue.png`);assert.equal(moduleAsset('assets/icons/Cue.png','file:///test.mjs'),'/assets/icons/Cue.png');assert.equal(moduleIdentity(`https://example.invalid/build/${id}/src/app.mjs`),null);
});
