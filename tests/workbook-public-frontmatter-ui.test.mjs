import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip');
let sequence=0;
async function app(){
 let p=fixture();p=editDeliveryPlan(p,{kind:'create',name:'Synthetic selected',targetMinutes:10,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]}).next;
 const h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);let fetches=0,storageReads=0,storageWrites=0;
 Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip:Zip,fetch:async()=>{fetches++;throw new Error('No frontmatter network');},localStorage:{getItem(){storageReads++;return null;},setItem(){storageWrites++;throw Error('No frontmatter storage writes');}}});
 await import(`../src/app.mjs?front-ui=${++sequence}`);await settle();
 const go=async name=>{await ui.routes.find(x=>x.dataset.view===name).onclick();await settle();};await go('output');
 return {p,h,ui,n,id:p.deliveryPlans.plans[0].id,go,calls:()=>({fetches,storageReads,storageWrites,saves:h.saves}),f:id=>n('family-front-'+id)};
}
const set=(a,id,value)=>{a.f(id).value=value;a.f(id).oninput?.();};
const include=(a,id,value=true)=>{a.f('include-'+id).checked=value;a.f('include-'+id).onchange();};
const click=async(a,id)=>{await a.n(id).onclick();await settle();};
function draft(a){set(a,'cover-title','Public copy title');set(a,'cover-subtitle','Authored subtitle');set(a,'cover-author','Public Author');set(a,'cover-edition','Edition 2');set(a,'cover-year','2026');set(a,'cover-audience','Public audience');set(a,'copyright-rightsHolder','Separate Rights Holder');set(a,'copyright-year','2024-2026');set(a,'copyright-noticeChoice','custom-text');set(a,'copyright-noticeText','Exact custom rights wording');set(a,'copyright-permissionsContact','Contact text');set(a,'copyright-credits','Third-party credits');include(a,'cover');include(a,'copyright');}
async function review(a){await click(a,'family-preview');assert.equal(a.n('error').hidden,true,a.n('error').textContent);assert.equal(a.n('family-reviewed').disabled,false,a.n('family-brand-status').textContent);a.n('family-reviewed').checked=true;a.n('family-reviewed').onchange();assert.equal(a.n('family-apply').disabled,false);}
async function offers(action){const blobs=[],names=[],oldUrl=URL.createObjectURL,oldCreate=document.createElement,oldTimer=globalThis.setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:synthetic';};document.createElement=tag=>{const e=oldCreate(tag);if(tag==='a'){e.click=()=>names.push(e.download);}return e;};globalThis.setTimeout=(fn,ms,...args)=>ms===60000?0:oldTimer(fn,ms,...args);try{await action();}finally{URL.createObjectURL=oldUrl;document.createElement=oldCreate;globalThis.setTimeout=oldTimer;}return {blobs,names};}

test('explicit off defaults and all25 family transitions preserve public text brand settings and source drafts',async()=>{
 const a=await app(),before=structuredClone(a.p);assert.equal(a.f('include-cover').checked,false);assert.equal(a.f('include-copyright').checked,false);assert.equal(a.f('copyright-noticeChoice').value,'');
 draft(a);a.n('family-brand-organizationName').value='Public Brand';a.n('family-brand-organizationName').oninput();a.n('family-brand-bodyPt').value='13';a.n('family-brand-bodyPt').oninput();a.n('activity-title').value='UNRELATED RAW ACTIVITY';
 for(const from of WORKBOOK_FAMILIES)for(const to of WORKBOOK_FAMILIES){a.n('family-template').value=from.id;a.n('family-template').onchange();await review(a);a.n('family-template').value=to.id;a.n('family-template').onchange();assert.equal(a.n('family-reviewed').checked,false);await review(a);assert.equal(a.f('cover-title').value,'Public copy title');assert.equal(a.f('copyright-noticeText').value,'Exact custom rights wording');assert.equal(a.n('family-brand-bodyPt').value,'13');const text=a.n('family-preview-content').textContent;assert.match(text,/Design:/u);assert.match(text,/Public copy title/u);assert.match(text,/Separate Rights Holder/u);assert.match(text,/Public Brand/u);assert.equal((text.match(/Rights holder: Separate Rights Holder/gu)||[]).length,1);}
 a.n('workbook-plan').value=a.id;a.n('workbook-plan').onchange();await review(a);await a.go('course');await a.go('output');assert.equal(a.f('cover-title').value,'Public copy title');assert.equal(a.f('copyright-noticeText').value,'Exact custom rights wording');assert.equal(a.n('activity-title').value,'UNRELATED RAW ACTIVITY');assert.deepEqual(a.p,before);assert.equal(a.h.saves,0);
});

test('Keep preserves raw and current review; reset and starter require matching context and clear only named draft fields',async()=>{
 const a=await app();draft(a);await review(a);const view=a.n('family-preview-content').textContent,prior=a.calls();a.n('family-brand-organizationName').value='Retained brand';a.n('family-brand-organizationName').oninput();await review(a);const sameReview=a.n('family-preview-content').textContent;
 a.f('reset').onclick();assert.equal(a.f('confirm').hidden,false);assert.equal(a.n('family-apply').disabled,true);a.f('keep').onclick();assert.equal(a.f('cover-title').value,'Public copy title');assert.equal(a.n('family-reviewed').checked,true);assert.equal(a.n('family-preview-content').textContent,sameReview);assert.equal(a.n('family-apply').disabled,false);
 a.f('reset').onclick();a.f('cover-title').value='EVENTLESS RETAIN';a.f('replace').onclick();assert.equal(a.f('cover-title').value,'EVENTLESS RETAIN');assert.equal(a.f('error').hidden,false);assert.equal(a.f('confirm').hidden,true);
 set(a,'copyright-noticeChoice','all-rights-reserved');a.f('starter').onclick();assert.equal(a.f('copyright-noticeText').value,'Exact custom rights wording');a.f('keep').onclick();assert.equal(a.f('copyright-noticeText').value,'Exact custom rights wording');a.f('starter').onclick();a.f('replace').onclick();assert.equal(a.f('copyright-noticeText').value,'All rights reserved.');assert.equal(a.f('copyright-rightsHolder').value,'Separate Rights Holder');assert.equal(a.f('cover-title').value,'EVENTLESS RETAIN');
 a.f('reset').onclick();a.f('replace').onclick();assert.equal(a.f('include-cover').checked,false);assert.equal(a.f('include-copyright').checked,false);assert.equal(a.f('copyright-noticeText').value,'');assert.equal(a.n('family-brand-organizationName').value,'Retained brand');assert.equal(a.n('family-brand-bodyPt').value,'');assert.deepEqual(a.calls(),prior);assert.equal(a.h.saves,0);
});

test('hidden detached tampered confirmation and busy cached actions refuse locally with no unrelated loss',async()=>{
 const a=await app();draft(a);await review(a);a.n('activity-title').value='KEEP RAW';a.n('agenda-reviewed').checked=true;const before=a.calls();
 a.f('reset').onclick();a.f('confirm-text').textContent='TAMPERED CONFIRMATION';a.f('replace').onclick();assert.equal(a.f('cover-title').value,'Public copy title');assert.equal(a.f('error').hidden,false);assert.equal(a.n('error').hidden,true);assert.equal(a.n('agenda-reviewed').checked,true);
 await a.go('course');a.f('reset').onclick();assert.equal(a.f('confirm').hidden,true);await a.go('output');a.f('controls').hidden=true;a.f('reset').onclick();assert.equal(a.f('confirm').hidden,true);a.f('controls').hidden=false;a.f('reset').isConnected=false;a.f('reset').onclick();assert.equal(a.f('confirm').hidden,true);a.f('reset').isConnected=true;
 let observed=false;a.h.hooks.beforeRefresh=()=>{if(observed)return;observed=true;const title=a.f('cover-title').value;a.f('reset').onclick();a.f('starter').onclick();assert.equal(a.f('cover-title').value,title);assert.equal(a.f('confirm').hidden,true);};await click(a,'refresh');assert.equal(observed,true);assert.equal(a.n('activity-title').value,'KEEP RAW');assert.equal(a.h.saves,0);assert.equal(a.n('agenda-reviewed').checked,true);assert.equal(a.calls().storageWrites,before.storageWrites);assert.equal(a.calls().fetches,before.fetches);
});

test('stale public raw fields validity checkbox choices and nested visibility block export while healthy repeated copies remain exact',async()=>{
 for(const change of [a=>{a.f('cover-title').value='EVENTLESS TITLE';},a=>{a.f('include-cover').checked=false;},a=>{a.f('copyright-noticeText').validity={customError:true};},a=>{a.f('controls').hidden=true;}]){
  const a=await app();draft(a);await review(a);change(a);const result=await offers(()=>click(a,'family-apply'));assert.equal(result.blobs.length,0);assert.equal(a.f('copyright-noticeText').value,'Exact custom rights wording');assert.equal(a.h.saves,0);
 }
 const a=await app();draft(a);await review(a);const result=await offers(async()=>{await click(a,'family-apply');await click(a,'family-apply');});assert.equal(result.blobs.length,2);assert.notEqual(result.names[0],result.names[1]);assert.match(result.names[0],/^Public copy title -/u);
 const zips=await Promise.all(result.blobs.map(async b=>Zip.loadAsync(await b.arrayBuffer(),{checkCRC32:true})));for(const name of Object.keys(zips[0].files))assert.equal(await zips[0].file(name).async('string'),await zips[1].file(name).async('string'));assert.equal(a.h.saves,0);
});

test('async public draft and acknowledgement roundtrips refuse download without altering raw fields',async()=>{
 for(const mutate of [a=>{set(a,'cover-title','During assembly');},a=>{a.n('family-reviewed').checked=false;a.n('family-reviewed').onchange();a.n('family-reviewed').checked=true;a.n('family-reviewed').onchange();}]){
  const a=await app();draft(a);await review(a);const original=Zip.prototype.generateAsync;Zip.prototype.generateAsync=async function(...args){const bytes=await original.apply(this,args);mutate(a);return bytes;};let result;try{result=await offers(()=>click(a,'family-apply'));}finally{Zip.prototype.generateAsync=original;}assert.equal(result.blobs.length,0);assert.equal(a.f('copyright-noticeText').value,'Exact custom rights wording');assert.equal(a.h.saves,0);
 }
});

test('copyright-only review includes existing public opening and inactive invalid license drafts remain retained',async()=>{
 const a=await app();draft(a);include(a,'cover',false);set(a,'copyright-licenseUrl','NOT A VALID URL');a.f('copyright-licenseUrl').validity={customError:true};await review(a);assert.match(a.n('family-preview-content').textContent,/Existing public opening:/u);assert.ok(a.n('family-preview-content').textContent.includes(a.p.title));assert.doesNotMatch(a.n('family-preview-content').textContent,/NOT A VALID URL/u);
 set(a,'copyright-noticeChoice','named-license');await click(a,'family-preview');assert.equal(a.n('family-apply').disabled,true);assert.equal(a.f('copyright-licenseUrl').value,'NOT A VALID URL');set(a,'copyright-noticeChoice','custom-text');await review(a);assert.equal(a.f('copyright-licenseUrl').value,'NOT A VALID URL');
});
