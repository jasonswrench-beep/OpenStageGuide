import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import vm from 'node:vm';
import {createRequire} from 'node:module';
import {course,addModule,addActivity,validate} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {buildDeliveryPack,verifyDeliveryPack,deliveryPackName} from '../src/delivery-pack.mjs';
import {planAgendaFiles} from '../src/plan-agenda.mjs';
import {planWorkbookFiles,planMaterialsFiles} from '../src/selected-plan-outputs.mjs';
import {planFacilitatorFiles} from '../src/selected-plan-facilitator.mjs';
const require=createRequire(import.meta.url),NodeZip=require('jszip');
globalThis.DOMParser=require('@xmldom/xmldom').DOMParser;
const bundle=await fs.readFile(new URL('../public/vendor/jszip.min.js',import.meta.url),'utf8');
const sandbox={Uint8Array,Uint16Array,Uint32Array,Int32Array,ArrayBuffer,DataView,Promise,Blob,setTimeout,clearTimeout,setImmediate,clearImmediate};sandbox.window=sandbox.self=sandbox;
vm.runInContext(bundle,vm.createContext(sandbox),{filename:'shipped-browser-jszip.min.js',timeout:5000});
const BrowserZip=sandbox.JSZip;assert.equal(BrowserZip.support.nodebuffer,false);assert.equal('Buffer' in sandbox,false);
const stamp='2026-10-28T20:00:00.000Z',encoder=new TextEncoder(),decoder=new TextDecoder('utf-8',{fatal:true});
const publicNames=['agenda.docx','participant-workbook.docx','materials-checklist.docx'],privateName='PRIVATE-facilitator-guide.docx';
const opts=(route,answers=false,participant=false)=>route==='private'?{route,start:'09:00',stamp,answers,participant}:{route,start:'09:00',stamp};
function saved(){
 let p=addModule(course('Public café 中文'),'Selected module','Selected objective.').next;
 p=addActivity(p,p.modules[0].id,{title:'Selected task',participant:'PUBLIC_LEARNER',facilitator:'PRIVATE_DIRECTION',answers:'PRIVATE_ANSWER',minutes:5,lines:3,materials:[{item:'Cards',quantity:2,unit:'sets',use:'consumable',notes:'PRIVATE_PREP'}]}).next;
 p=addModule(p,'OMITTED_MODULE','OMITTED_OBJECTIVE').next;
 p=addActivity(p,p.modules[1].id,{title:'OMITTED_ACTIVITY',participant:'OMITTED_LEARNER',facilitator:'OMITTED_PRIVATE',answers:'OMITTED_ANSWER',minutes:7,lines:1}).next;
 p=editDeliveryPlan(p,{kind:'create',name:'Unicode plan',targetMinutes:5,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]}).next;
 return {p,id:p.deliveryPlans.plans[0].id};
}
function expected(p,id,options){
 const maps={'agenda.docx':planAgendaFiles(p,id,'09:00',stamp),'participant-workbook.docx':planWorkbookFiles(p,id,stamp,'09:00'),'materials-checklist.docx':planMaterialsFiles(p,id,stamp,'09:00')};
 if(options.route==='private')maps[privateName]=Object.fromEntries(Object.entries(planFacilitatorFiles(p,id,{answers:options.answers,participant:options.participant},stamp,'09:00')).map(([part,body])=>[part,part==='word/document.xml'?body.replace('Coordinated packs and existing-output regeneration are unavailable.','Existing-output regeneration remains unavailable.'):body]));
 return maps;
}
async function exact(result,p,id,options){
 const maps=expected(p,id,options),outer=await NodeZip.loadAsync(result.data,{checkCRC32:true,createFolders:false});
 const manifestBytes=await outer.file('manifest.json').async('uint8array'),manifestText=JSON.stringify(result.manifest,null,2)+'\n';
 assert.deepEqual(manifestBytes,encoder.encode(manifestText));assert.equal(decoder.decode(manifestBytes),manifestText);
 assert.deepEqual(Object.keys(outer.files),[...Object.keys(maps),'manifest.json']);
 for(const [name,files]of Object.entries(maps)){
  const inner=await NodeZip.loadAsync(await outer.file(name).async('uint8array'),{checkCRC32:true,createFolders:false});assert.deepEqual(Object.keys(inner.files),Object.keys(files));
  for(const [part,body]of Object.entries(files)){const raw=await inner.file(part).async('uint8array');assert.deepEqual(raw,encoder.encode(body),`${name}/${part}`);assert.equal(decoder.decode(raw),body);assert.doesNotMatch(body,/OMITTED_/);if(name!==privateName)assert.doesNotMatch(body,/PRIVATE_DIRECTION|PRIVATE_ANSWER|PRIVATE_PREP/);}
 }
 assert.deepEqual(await verifyDeliveryPack(result.data,NodeZip,{expectedManifestBytes:manifestText,expectedDocFiles:maps}),result.manifest);
 assert.equal(result.manifest.identity.minutes,5);assert.deepEqual(result.manifest.identity.scope,[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]);
 return outer;
}
function wrapped(Zip,{file,generate}={}){
 function Wrapper(){const z=new Zip(),originalFile=z.file,originalGenerate=z.generateAsync;z.file=function(...args){return file?file.call(this,originalFile,args):originalFile.apply(this,args);};z.generateAsync=function(...args){return generate?generate.call(this,originalGenerate,args):originalGenerate.apply(this,args);};return z;}
 Wrapper.loadAsync=(...args)=>Zip.loadAsync(...args);return Wrapper;
}

test('both ZIP engines preserve exact public/private source parts, options, identity, determinism and privacy',async()=>{
 for(const Zip of [NodeZip,BrowserZip])for(const options of [opts('shareable'),...['private'].flatMap(route=>[false,true].flatMap(a=>[false,true].map(p=>opts(route,a,p))))]){
  const {p,id}=saved(),before=JSON.stringify(p);const first=await buildDeliveryPack(p,id,options,Zip),second=await buildDeliveryPack(p,id,options,Zip);await exact(first,p,id,options);assert.deepEqual(first.data,second.data);assert.equal(JSON.stringify(p),before);assert.equal(Object.isFrozen(p),false);
  if(options.route==='private'){assert.deepEqual(first.manifest.facilitatorOptions,{answers:options.answers,participant:options.participant});const outer=await NodeZip.loadAsync(first.data),inner=await NodeZip.loadAsync(await outer.file(privateName).async('uint8array')),body=await inner.file('word/document.xml').async('string');assert.equal(body.includes('PRIVATE_ANSWER'),options.answers);assert.equal(body.includes('PUBLIC_LEARNER'),options.participant);}
  assert.notEqual(deliveryPackName(options.route),deliveryPackName(options.route));
 }
});

test('complete inner XML UTF-8 survives supplementary characters around16/32Ki boundaries in both ZIP engines',async()=>{
 for(const Zip of [NodeZip,BrowserZip])for(const route of ['shareable','private'])for(const boundary of [16384,32768])for(const offset of [-1,0,1]){
  const {p,id}=saved(),a=p.modules[0].activities[0],field=route==='private'?'facilitator':'participant',name=route==='private'?privateName:'participant-workbook.docx',options=opts(route);a[field]='BOUNDARY_SENTINEL';
  const prefix=expected(p,id,options)[name]['word/document.xml'].indexOf('BOUNDARY_SENTINEL');assert.ok(prefix>=0&&prefix<boundary-2);
  a[field]='x'.repeat(boundary-1+offset-prefix)+'\u{1f642} café 中文';validate(p);const maps=expected(p,id,options);assert.equal(maps[name]['word/document.xml'].indexOf('\u{1f642}'),boundary-1+offset);
  const before=JSON.stringify(p);await exact(await buildDeliveryPack(p,id,options,Zip),p,id,options);assert.equal(JSON.stringify(p),before);
 }
});

test('complete canonical manifest UTF-8 survives valid long source identity at16/32Ki boundaries',async()=>{
 for(const Zip of [NodeZip,BrowserZip])for(const route of ['shareable','private'])for(const boundary of [16384,32768])for(const offset of [-1,0,1]){
  const {p,id}=saved(),options=opts(route);p.id=p.deliveryPlans.projectId='IDENTITY_SENTINEL';validate(p);
  const seed=await buildDeliveryPack(p,id,options,Zip),prefix=(JSON.stringify(seed.manifest,null,2)+'\n').indexOf('IDENTITY_SENTINEL');assert.ok(prefix>0&&prefix<boundary-2);
  p.id=p.deliveryPlans.projectId='x'.repeat(boundary-1+offset-prefix)+'\u{1d11e} identity';validate(p);
  const before=JSON.stringify(p),result=await buildDeliveryPack(p,id,options,Zip);assert.equal((JSON.stringify(result.manifest,null,2)+'\n').indexOf('\u{1d11e}'),boundary-1+offset);await exact(result,p,id,options);assert.equal(JSON.stringify(p),before);
 }
});

test('every writer hands ZIP ordinary binary snapshots with unchanged fixed-date and folder options',async()=>{
 for(const Zip of [NodeZip,BrowserZip])for(const route of ['shareable','private']){
  const writes=[],Writer=wrapped(Zip,{file(original,args){if(args.length>1){const [name,body,options]=args;assert.equal(typeof body,'object',name);assert.equal(Object.getPrototypeOf(body),Uint8Array.prototype,name);assert.equal(options.createFolders,false);assert.equal(options.date.toISOString(),'2000-01-01T00:00:00.000Z');writes.push({name,body});}return original.apply(this,args);}});
  const {p,id}=saved(),options=opts(route),result=await buildDeliveryPack(p,id,options,Writer);await exact(result,p,id,options);assert.equal(writes.length,route==='private'?29:22);assert.equal(writes.filter(x=>x.name==='manifest.json').length,1);
 }
});

test('altered ZIP-input UTF-8 and manifest bytes refuse before any successful offer',async()=>{
 for(const Zip of [NodeZip,BrowserZip])for(const target of ['word/document.xml','manifest.json']){
  const Writer=wrapped(Zip,{file(original,args){if(args.length>1&&args[0]===target){const copy=new Uint8Array(args[1]);copy[0]^=1;args[1]=copy;}return original.apply(this,args);}}),{p,id}=saved();
  await assert.rejects(buildDeliveryPack(p,id,opts('shareable'),Writer),/XML|manifest|canonical|differs/i);
 }
});

test('every async freshness boundary refuses interruption while source and reviewed options remain unchanged',async()=>{
 for(const Zip of [NodeZip,BrowserZip])for(const route of ['shareable','private']){
  const {p,id}=saved(),options=opts(route,true,true),before=JSON.stringify({p,options}),healthy=await buildDeliveryPack(p,id,options,Zip);await exact(healthy,p,id,options);
  const expectedCalls=route==='private'?20:16;for(let stop=1;stop<=expectedCalls;stop++){let calls=0;await assert.rejects(buildDeliveryPack(p,id,options,Zip,()=>{if(++calls===stop)throw new Error('STALE_BOUNDARY');}),/STALE_BOUNDARY/);assert.equal(calls,stop);}
  assert.equal(JSON.stringify({p,options}),before);
 }
});
