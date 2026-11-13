import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import vm from 'node:vm';
import {createRequire} from 'node:module';
import {course,addModule,addActivity,validate} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {agendaFiles,agendaBlob} from '../src/agenda.mjs';
import {planAgendaFiles,planAgendaBlob} from '../src/plan-agenda.mjs';
import {worksheetFiles,worksheetBlob} from '../src/worksheet.mjs';
import {materialsFiles,materialsBlob} from '../src/materials-output.mjs';
import {facilitatorFiles,facilitatorBlob} from '../src/facilitator-output.mjs';
import {planWorkbookFiles,planWorkbookBlob,planMaterialsFiles,planMaterialsBlob} from '../src/selected-plan-outputs.mjs';
import {planFacilitatorFiles,planFacilitatorBlob} from '../src/selected-plan-facilitator.mjs';
import {LEGACY_DOCX_PARTS,legacyDocxBlob} from '../src/legacy-docx-package.mjs';
const require=createRequire(import.meta.url),NodeZip=require('jszip');
const bundle=await fs.readFile(new URL('../public/vendor/jszip.min.js',import.meta.url),'utf8');
const sandbox={Uint8Array,Uint16Array,Uint32Array,Int32Array,ArrayBuffer,DataView,Promise,Blob,setTimeout,clearTimeout,setImmediate,clearImmediate};sandbox.window=sandbox.self=sandbox;
vm.runInContext(bundle,vm.createContext(sandbox),{filename:'shipped-browser-jszip.min.js',timeout:5000});
const BrowserZip=sandbox.JSZip;assert.equal(BrowserZip.support.nodebuffer,false);
const encoder=new TextEncoder(),decoder=new TextDecoder('utf-8',{fatal:true}),stamp='2026-10-28T18:00:00.000Z';
const directories=['_rels/','word/','word/_rels/','docProps/'];
function fixture(){
 let p=course('Public Unicode course');p=addModule(p,'Selected module','Selected learning objective.').next;
 p=addActivity(p,p.modules[0].id,{title:'BOUNDARY_SENTINEL',participant:'BOUNDARY_SENTINEL',facilitator:'BOUNDARY_SENTINEL',answers:'PRIVATE_ANSWER_SELECTED',minutes:5,lines:3,materials:[{item:'Public paper',quantity:2,unit:'sets',use:'consumable',notes:'PRIVATE_NOTE_SELECTED'}]}).next;
 p=addModule(p,'OMITTED_MODULE','OMITTED_OBJECTIVE').next;p=addActivity(p,p.modules[1].id,{title:'OMITTED_ACTIVITY',participant:'OMITTED_PARTICIPANT',facilitator:'PRIVATE_OMITTED_FACILITATOR',answers:'PRIVATE_OMITTED_ANSWER',minutes:7,lines:1,materials:[{item:'OMITTED_MATERIAL',quantity:3,unit:'sets',use:'consumable',notes:'PRIVATE_OMITTED_NOTE'}]}).next;
 const plan=editDeliveryPlan(p,{kind:'create',name:'Selected Unicode plan',targetMinutes:5,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]});return {p:plan.next,id:plan.id};
}
function routes(id,Zip,options={answers:false,participant:false}){return [
 ['agenda-full','title',p=>agendaFiles(p,'09:00',stamp),p=>agendaBlob(p,'09:00',Zip)],
 ['agenda-selected','title',p=>planAgendaFiles(p,id,'09:00',stamp),p=>planAgendaBlob(p,id,'09:00',stamp,Zip)],
 ['workbook-full','participant',p=>worksheetFiles(p,stamp),p=>worksheetBlob(p,Zip)],
 ['workbook-selected','participant',p=>planWorkbookFiles(p,id,stamp),p=>planWorkbookBlob(p,id,stamp,Zip)],
 ['materials-full','title',p=>materialsFiles(p,stamp),p=>materialsBlob(p,stamp,Zip)],
 ['materials-selected','title',p=>planMaterialsFiles(p,id,stamp),p=>planMaterialsBlob(p,id,stamp,Zip)],
 ['facilitator-full','facilitator',p=>facilitatorFiles(p,options,stamp),p=>facilitatorBlob(p,options,stamp,Zip)],
 ['facilitator-selected','facilitator',p=>planFacilitatorFiles(p,id,options,stamp),p=>planFacilitatorBlob(p,id,options,stamp,Zip)]
 ];}
async function fixedTime(fn){const Original=globalThis.Date;globalThis.Date=class extends Original{constructor(...args){super(...(args.length?args:[stamp]));}static now(){return Original.parse(stamp);}};try{return await fn();}finally{globalThis.Date=Original;}}
async function exact(blob,files){
 assert.ok(blob instanceof Blob);assert.equal(blob.type,'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
 const loaded=await NodeZip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true,createFolders:false});
 assert.deepEqual(Object.keys(loaded.files).sort(),[...LEGACY_DOCX_PARTS,...directories].sort());
 for(const name of LEGACY_DOCX_PARTS){const bytes=await loaded.file(name).async('uint8array');assert.deepEqual(bytes,encoder.encode(files[name]),name);assert.equal(decoder.decode(bytes),files[name],name);}
 for(const name of directories)assert.equal(loaded.files[name].dir,true);
 return loaded;
}
function wrapped(Zip,{file,generate,load}={}){
 function Wrapper(){const z=new Zip(),originalFile=z.file,originalGenerate=z.generateAsync;
  z.file=function(...args){return file?file.call(this,originalFile,args):originalFile.apply(this,args);};
  z.generateAsync=function(...args){return generate?generate.call(this,originalGenerate,args):originalGenerate.apply(this,args);};return z;}
 Wrapper.loadAsync=async(...args)=>load?load(Zip,args):Zip.loadAsync(...args);return Wrapper;
}
const base=()=>{const {p}=fixture();return worksheetFiles(p,stamp);};

test('all eight legacy APIs preserve canonical Unicode across16/32/48Ki boundaries in Node and the shipped browser bundle',async()=>fixedTime(async()=>{
 for(const Zip of [NodeZip,BrowserZip]){const {p:source,id}=fixture();for(const [name,field,filesFor,blobFor]of routes(id,Zip)){
  for(const [boundary,offset,character]of [[16384,0,'\u{1f642}'],[32768,0,'\u{1d11e}'],[49152,0,'\u{1f642}'],[16384,-1,'\u{1f642}'],[16384,1,'\u{1f642}'],[16384,0,'\u00e9']]){
   const p=structuredClone(source),a=p.modules[0].activities[0];for(const key of ['title','participant','facilitator'])a[key]='Short '+key;a[field]='BOUNDARY_SENTINEL';
   const prefix=filesFor(p)['word/document.xml'].indexOf('BOUNDARY_SENTINEL');assert.ok(prefix>=0&&prefix<boundary-1,name);
   a[field]='x'.repeat(boundary-1+offset-prefix)+character+' caf\u00e9 \u4e2d\u6587';assert.ok(a[field].length<=50000);validate(p);
   const before=JSON.stringify(p),files=filesFor(p);assert.equal(files['word/document.xml'].indexOf(character),boundary-1+offset,name);
   await exact(await blobFor(p),files);assert.equal(JSON.stringify(p),before,name);assert.equal(Object.isFrozen(p),false,name);
  }
 }}
}));

test('ordinary short output maps, selected omission and facilitator options remain exact across all eight routes',async()=>fixedTime(async()=>{
 for(const Zip of [NodeZip,BrowserZip])for(const answers of [false,true])for(const participant of [false,true]){
  const {p,id}=fixture(),before=JSON.stringify(p);p.modules[0].activities[0].participant='PUBLIC_PARTICIPANT_SELECTED';p.modules[0].activities[0].facilitator='PRIVATE_FACILITATOR_SELECTED';
  for(const [name,,filesFor,blobFor]of routes(id,Zip,{answers,participant})){
   const files=filesFor(p);await exact(await blobFor(p),files);const text=Object.values(files).join('');
   if(name.endsWith('selected'))assert.doesNotMatch(text,/OMITTED_/u,name);
   if(!name.startsWith('facilitator'))assert.doesNotMatch(text,/PRIVATE_/u,name);
   else {assert.match(text,/PRIVATE_FACILITATOR_SELECTED/u);assert.equal(text.includes('PRIVATE_ANSWER_SELECTED'),answers);assert.equal(text.includes('PUBLIC_PARTICIPANT_SELECTED'),participant);}
  }
  const compare=JSON.parse(before);compare.modules[0].activities[0].participant=p.modules[0].activities[0].participant;compare.modules[0].activities[0].facilitator=p.modules[0].activities[0].facilitator;assert.deepEqual(p,compare);
 }
}));

test('ordinary valid short legacy ZIPs keep complete bytes, six payloads, four directories, dates and compression behavior',async()=>fixedTime(async()=>{
 for(const Zip of [NodeZip,BrowserZip]){const {p,id}=fixture();for(const [name,,filesFor,blobFor]of routes(id,Zip)){
  const files=filesFor(p),old=new Zip();for(const [part,text]of Object.entries(files))old.file(part,text);
  const prior=await old.generateAsync({type:'uint8array',compression:'DEFLATE'}),now=new Uint8Array(await (await blobFor(p)).arrayBuffer());
  assert.deepEqual(now,prior,name);
 }}
}));

test('six-part input validation refuses null, extra, inherited, accessor and unsupported Unicode without invoking getters',async()=>{
 const files=base();let reads=0;
 const accessor={...files};Object.defineProperty(accessor,'word/document.xml',{enumerable:true,get(){reads++;return files['word/document.xml'];}});
 const inherited=Object.create(files),hidden={...files};Object.defineProperty(hidden,'word/document.xml',{value:files['word/document.xml'],enumerable:false});
 for(const input of [null,[],{},inherited,accessor,hidden,{...files,extra:''},{...files,[Symbol('x')]:''},{...files,'word/styles.xml':null},{...files,'word/document.xml':'<w:p>\ud800</w:p>'}])await assert.rejects(legacyDocxBlob(input,NodeZip),/Generated DOCX:/u);
 assert.equal(reads,0);await assert.rejects(legacyDocxBlob(files,function(){}),/unavailable/u);
});

test('private canonical bytes cannot alias encoder inputs and changed generation or ZIP members refuse before Blob return',async()=>{
 const files=base(),before=JSON.stringify(files);let retained=[];
 const mutatedInput=wrapped(NodeZip,{file(original,args){if(args[1] instanceof Uint8Array)retained.push(args[1]);return original.apply(this,args);},generate(original,args){retained[0].fill(0);return original.apply(this,args);}});
 await assert.rejects(legacyDocxBlob(files,mutatedInput),/payload .* changed/u);assert.equal(retained.length,6);assert.equal(JSON.stringify(files),before);
 for(const defect of ['xml','extra-directory','missing','trailing','wrong-type']){
  const broken=wrapped(NodeZip,{async generate(original,args){const bytes=await original.apply(this,args);if(defect==='wrong-type')return new ArrayBuffer(0);if(defect==='trailing'){const changed=new Uint8Array(bytes.length+1);changed.set(bytes);return changed;}
   const z=await NodeZip.loadAsync(bytes);if(defect==='xml')z.file('word/document.xml',encoder.encode('<changed/>'),{binary:true});if(defect==='extra-directory')z.folder('unreviewed');if(defect==='missing')z.remove('word/styles.xml');return z.generateAsync({type:'uint8array',compression:'DEFLATE'});}});
  await assert.rejects(legacyDocxBlob(base(),broken),/Generated DOCX:/u,defect);
 }
});

test('post-await source own-data checks refuse edits including a last-read accessor, without losing or freezing caller data',async()=>{
 for(const when of ['generate','first-read','last-read','last-accessor']){
  const files=base();let reads=0,getters=0;
  const changing=wrapped(NodeZip,{async generate(original,args){const bytes=await original.apply(this,args);if(when==='generate')files['word/document.xml']+='<!-- changed -->';return bytes;},async load(Zip,args){const z=await Zip.loadAsync(...args);for(const name of LEGACY_DOCX_PARTS){const entry=z.file(name),original=entry.async;entry.async=async function(...a){const bytes=await original.apply(this,a);if(++reads===(when==='first-read'?1:6)&&when!=='generate'){if(when==='last-accessor')Object.defineProperty(files,'word/document.xml',{enumerable:true,get(){getters++;return '<changed/>';}});else files['word/document.xml']+='<!-- changed -->';}return bytes;};}return z;}});
  await assert.rejects(legacyDocxBlob(files,changing),/source strings changed|visible own text data/u);assert.equal(getters,0);assert.equal(Object.isFrozen(files),false);
 }
});

test('raw envelope checks reject inconsistent size, version, flags, duplicate names, metadata and overlapping local records',async()=>{
 for(const defect of ['expanded-size','local-crc','version','flags','duplicate-name','attribute','local-name','overlap']){
  const changed=wrapped(NodeZip,{async generate(original,args){const result=await original.apply(this,args),bytes=Uint8Array.from(result),v=new DataView(bytes.buffer),end=bytes.length-22,entries=[];
   let at=v.getUint32(end+16,true);while(at<end){const n=v.getUint16(at+28,true),name=decoder.decode(bytes.subarray(at+46,at+46+n));entries.push({at,name,n,local:v.getUint32(at+42,true)});at+=46+n+v.getUint16(at+30,true)+v.getUint16(at+32,true);}
   const entry=entries.find(x=>x.name==='word/document.xml'),c=entry.at,l=entry.local;
   if(defect==='expanded-size')v.setUint32(c+24,0x7fffffff,true);
   if(defect==='local-crc')v.setUint32(l+14,0,true);
   if(defect==='version')v.setUint16(c+6,45,true);
   if(defect==='flags'){v.setUint16(c+8,1,true);v.setUint16(l+6,1,true);}
   if(defect==='duplicate-name'){const other=entries.find(x=>x.name==='docProps/core.xml');assert.equal(other.n,entry.n);bytes.set(encoder.encode(entry.name),other.at+46);bytes.set(encoder.encode(entry.name),other.local+30);}
   if(defect==='attribute')v.setUint32(c+38,16,true);
   if(defect==='local-name')bytes[l+30]=65;
   if(defect==='overlap')v.setUint32(c+42,0,true);
   return bytes;}});
  await assert.rejects(legacyDocxBlob(base(),changed),/Generated DOCX:/u,defect);
 }
});

test('generated byte snapshot is isolated before loader await and final offered bytes retain their validated content',async()=>{
 for(const Zip of [NodeZip,BrowserZip]){const files=base();let generated;
  const retaining=wrapped(Zip,{async generate(original,args){generated=await original.apply(this,args);return generated;},async load(Original,args){generated.fill(0);return Original.loadAsync(...args);}});
  await exact(await legacyDocxBlob(files,retaining),files);
 }
});
