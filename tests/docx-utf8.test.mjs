import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {prepareAssessmentExport,assessmentExportBlob} from '../src/assessment-output.mjs';
import {familyWorkbookBlob,prepareFamilyWorkbook,FAMILY_BRAND_FIELDS} from '../src/workbook-family-flow.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {FAMILY_WORKBOOK_PARTS} from '../src/workbook-family-output.mjs';
import {course,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {rawControl} from '../src/activity-draft.mjs';
import {unicodeAssessmentDefinition,unicodeAssessmentOptions} from './helpers/unicode-assessment-fixture.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip'),{DOMParser}=require('@xmldom/xmldom');
const encoder=new TextEncoder(),decoder=new TextDecoder('utf-8',{fatal:true});
const base=()=>prepareAssessmentExport(unicodeAssessmentDefinition(2),unicodeAssessmentOptions('respondent','fr-CA')).files;
const mutable=files=>Object.fromEntries(Object.entries(files));
function crossing(xml,boundary,offset=0,character='\u{1f642}'){
 const at=xml.indexOf('?>')+2,prefix=xml.slice(0,at)+'<!--',target=boundary-1+offset;
 return prefix+'x'.repeat(target-prefix.length)+character+' CAF\u00c9 \u4e2d\u6587 -->'+xml.slice(at);
}
async function exactParts(blob,files){
 const z=await Zip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true});assert.deepEqual(Object.keys(z.files).sort(),[...FAMILY_WORKBOOK_PARTS].sort());
 for(const name of FAMILY_WORKBOOK_PARTS){const bytes=await z.file(name).async('uint8array');assert.deepEqual(bytes,encoder.encode(files[name]),name);assert.equal(decoder.decode(bytes),files[name],name);}
 return z;
}

test('seven-part XML encoding preserves supplementary characters spanning16/32/48/64Ki UTF16 boundaries and nearby BMP controls',async()=>{
 const original=base(),before=JSON.stringify(original);
 for(const name of ['word/document.xml','word/styles.xml','docProps/core.xml','_rels/.rels'])for(const boundary of [16384,32768,49152,65536])for(const [offset,character] of [[0,'\u{1f642}'],[0,'\u{1d11e}'],[-1,'\u{1f642}'],[1,'\u{1f642}'],[0,'\u00e9']]){
  const files=mutable(original);files[name]=crossing(files[name],boundary,offset,character);assert.equal(files[name].slice(boundary-1+offset,boundary-1+offset+character.length),character);
  await exactParts(await familyWorkbookBlob(files,Zip),files);assert.equal(files[name].includes('\ufffd'),false);
 }
 assert.equal(JSON.stringify(original),before);
});

test('long valid24-item Likert and semantic differential respondent/PRIVATEguide2-and11-choice exports retain every multilingual byte',async()=>{
 for(const count of [2,11]){
  const d=unicodeAssessmentDefinition(count),before=JSON.stringify(d);
  for(const route of ['respondent','author'])for(const language of ['fr-CA','de-DE']){
   const options=unicodeAssessmentOptions(route,language),optionsBefore=JSON.stringify(options),prepared=prepareAssessmentExport(d,options),blob=await assessmentExportBlob(d,options,Zip);
   await exactParts(blob,prepared.files);const xml=prepared.files['word/document.xml'];assert.ok(xml.includes('\u{1f642}'));assert.ok(xml.includes('\u4e2d\u6587'));assert.ok(xml.includes('caf\u00e9'));
   const errors=[];new DOMParser({errorHandler:{warning:e=>errors.push(e),error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(xml,'application/xml');assert.deepEqual(errors,[]);
   if(route==='respondent')assert.doesNotMatch(xml,/KEYONLY_|Code:|Reverse keyed:/u);else assert.match(xml,/PRIVATE author scoring guide/u);
   assert.equal(JSON.stringify(options),optionsBefore);assert.equal(Object.isFrozen(options),false);
  }
  assert.equal(JSON.stringify(d),before);assert.equal(Object.isFrozen(d),false);
 }
});

test('all five full/selected seven-part family workbooks preserve long learner Unicode without private or omitted fields',async()=>{
 let p=course('Public Unicode \u{1f642} workbook');p=addModule(p,'Learning module \u4e2d\u6587','Learning objective \u{1f642} '+('Objective caf\u00e9 \u4e2d\u6587 '.repeat(300))).next;
 p=addActivity(p,p.modules[0].id,{title:'Long learner activity',minutes:5,participant:'Public learner \u{1f642} caf\u00e9 \u4e2d\u6587. '.repeat(1500),facilitator:'PRIVATE_FACILITATOR_CANARY',answers:'PRIVATE_ANSWERS_CANARY',lines:3}).next;
 p=addModule(p,'Full only module','Full only objective').next;p=addActivity(p,p.modules[1].id,{title:'Full only activity',minutes:7,participant:'FULL_ONLY_LEARNER_CANARY',facilitator:'PRIVATE_OTHER_TEACHING',answers:'PRIVATE_OTHER_ANSWERS',lines:1}).next;
 const plan=editDeliveryPlan(p,{kind:'create',name:'Selected \u{1f642}',targetMinutes:5,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]});p=plan.next;
 const before=JSON.stringify(p),rawBrand=Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,rawControl({value:''})]));
 for(const family of WORKBOOK_FAMILIES)for(const planId of ['',plan.id]){
  const prepared=prepareFamilyWorkbook(p,{planId,familyId:family.id,rawBrand,stamp:'2026-10-28T17:00:00.000Z'});assert.equal(prepared.output.profileVersion,3);await exactParts(await familyWorkbookBlob(prepared.output.files,Zip),prepared.output.files);
  const text=Object.values(prepared.output.files).join('');assert.doesNotMatch(text,/PRIVATE_/u);if(planId)assert.doesNotMatch(text,/FULL_ONLY_LEARNER_CANARY/u);else assert.match(text,/FULL_ONLY_LEARNER_CANARY/u);
 }
 assert.equal(JSON.stringify(p),before);assert.equal(Object.isFrozen(p),false);
});

test('canonical encoded expectations stay private from ZIP input arrays and asynchronous caller-map edits are refused',async()=>{
 const files=mutable(base()),before=JSON.stringify(files),file=Zip.prototype.file,generate=Zip.prototype.generateAsync,load=Zip.loadAsync;
 let retained=[];
 Zip.prototype.file=function(name,value,...args){if(value instanceof Uint8Array){retained.push(value);assert.equal(args[0].binary,true);}return file.call(this,name,value,...args);};
 Zip.prototype.generateAsync=async function(...args){retained[0].fill(0);return generate.apply(this,args);};
 try{await assert.rejects(familyWorkbookBlob(files,Zip),/changed/u);assert.equal(retained.length,7);assert.equal(JSON.stringify(files),before);}finally{Zip.prototype.file=file;Zip.prototype.generateAsync=generate;}
 for(const when of ['generate','first-read','last-read']){
  const input=mutable(base());let reads=0;
  Zip.prototype.generateAsync=async function(...args){const bytes=await generate.apply(this,args);if(when==='generate')input['word/document.xml']+='<!-- changed -->';return bytes;};
  Zip.loadAsync=async function(...args){const z=await load.apply(this,args);for(const entry of Object.values(z.files)){const read=entry.async;entry.async=async function(...a){const bytes=await read.apply(this,a);if(++reads===(when==='last-read'?7:1)&&when!=='generate')input['docProps/core.xml']+='<!-- changed -->';return bytes;};}return z;};
  try{await assert.rejects(familyWorkbookBlob(input,Zip),/source strings changed/u);}finally{Zip.prototype.generateAsync=generate;Zip.loadAsync=load;}
 }
});

test('post-read own-data validation refuses a changed accessor without invoking it; generated private-buffer protection remains',async()=>{
 const input=mutable(base()),load=Zip.loadAsync,generate=Zip.prototype.generateAsync;let reads=0,getterCalls=0,retained;
 Zip.loadAsync=async function(...args){const z=await load.apply(this,args);for(const entry of Object.values(z.files)){const read=entry.async;entry.async=async function(...a){const value=await read.apply(this,a);if(++reads===7)Object.defineProperty(input,'word/document.xml',{enumerable:true,get(){getterCalls++;return '<changed/>';}});return value;};}return z;};
 try{await assert.rejects(familyWorkbookBlob(input,Zip),/own data/u);assert.equal(getterCalls,0);}finally{Zip.loadAsync=load;}
 Zip.prototype.generateAsync=async function(...a){retained=await generate.apply(this,a);return retained;};Zip.loadAsync=async function(bytes,...a){retained.fill(0);return load.call(this,bytes,...a);};
 const healthy=base();try{await exactParts(await familyWorkbookBlob(healthy,Zip),healthy);}finally{Zip.prototype.generateAsync=generate;Zip.loadAsync=load;}
});
