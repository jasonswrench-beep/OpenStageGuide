import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {WORKBOOK_RESEARCH_FIELDS} from '../src/workbook-research.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip');let sequence=0;
async function app(){
 let p=fixture();p=editDeliveryPlan(p,{kind:'create',name:'Selected',targetMinutes:10,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]}).next;
 const h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);let fetches=0,storageReads=0,storageWrites=0;
 Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip:Zip,fetch:async()=>{fetches++;throw Error('No research network');},localStorage:{getItem(){storageReads++;return null;},setItem(){storageWrites++;throw Error('No research storage writes');}}});
 await import(`../src/app.mjs?research-ui=${++sequence}`);await settle();const go=async name=>{await ui.routes.find(x=>x.dataset.view===name).onclick();await settle();};await go('output');
 return {p,h,ui,n,go,id:p.deliveryPlans.plans[0].id,r:id=>n('family-research-'+id),f:id=>n('family-front-'+id),calls:()=>({fetches,storageReads,storageWrites,saves:h.saves,runs:h.runs,refreshes:h.refreshes})};
}
const set=(a,k,v)=>{a.r(k).value=v;a.r(k).oninput?.();};
function draft(a){for(const k of WORKBOOK_RESEARCH_FIELDS)set(a,k,'Authored '+k+' text');a.r('include').checked=true;a.r('include').onchange();}
const click=async(a,id)=>{await a.n(id).onclick();await settle();};
async function review(a){await click(a,'family-preview');assert.equal(a.n('error').hidden,true,a.n('error').textContent);assert.equal(a.n('family-reviewed').disabled,false,a.n('family-brand-status').textContent);a.n('family-reviewed').checked=true;a.n('family-reviewed').onchange();assert.equal(a.n('family-apply').disabled,false);}
async function offers(fn){const blobs=[],names=[],oldURL=URL.createObjectURL,oldCreate=document.createElement,oldTimer=globalThis.setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:synthetic';};document.createElement=tag=>{const e=oldCreate(tag);if(tag==='a')e.click=()=>names.push(e.download);return e;};globalThis.setTimeout=(fn,ms,...args)=>ms===60000?0:oldTimer(fn,ms,...args);try{await fn();}finally{URL.createObjectURL=oldURL;document.createElement=oldCreate;globalThis.setTimeout=oldTimer;}return {blobs,names};}

test('research defaults off and all25 family transitions preserve authored scope branding publication and unrelated drafts',async()=>{
 const a=await app(),source=structuredClone(a.p);assert.equal(a.r('include').checked,false);for(const k of WORKBOOK_RESEARCH_FIELDS)assert.equal(a.r(k).value,'');
 draft(a);a.n('family-brand-organizationName').value='Public Brand';a.n('family-brand-organizationName').oninput();a.f('cover-title').value='Retained cover';a.f('cover-title').oninput();a.f('copyright-noticeText').value='Retained custom notice';a.f('copyright-noticeText').oninput();a.n('family-toc-include').checked=true;a.n('family-toc-include').onchange();a.n('activity-title').value='UNRELATED RAW';
 for(const from of WORKBOOK_FAMILIES)for(const to of WORKBOOK_FAMILIES){a.n('family-template').value=from.id;a.n('family-template').onchange();await review(a);a.n('family-template').value=to.id;a.n('family-template').onchange();assert.equal(a.n('family-reviewed').checked,false);await review(a);for(const k of WORKBOOK_RESEARCH_FIELDS)assert.equal(a.r(k).value,'Authored '+k+' text');assert.match(a.n('family-preview-content').textContent,/have not been verified/u);}
 a.n('workbook-plan').value=a.id;a.n('workbook-plan').onchange();await review(a);await a.go('course');await a.go('output');assert.equal(a.r('evidence').value,'Authored evidence text');assert.equal(a.f('cover-title').value,'Retained cover');assert.equal(a.f('copyright-noticeText').value,'Retained custom notice');assert.equal(a.n('family-brand-organizationName').value,'Public Brand');assert.equal(a.n('family-toc-include').checked,true);assert.equal(a.n('activity-title').value,'UNRELATED RAW');assert.deepEqual(a.p,source);assert.equal(a.h.saves,0);
});

test('research Keep/reset, stale context and confirmation protection affect only research draft',async()=>{
 const a=await app();draft(a);a.n('family-toc-include').checked=true;a.n('family-toc-include').onchange();a.f('cover-title').value='KEEP COVER';a.f('cover-title').oninput();await review(a);const text=a.n('family-preview-content').textContent,before=a.calls();
 a.r('reset').onclick();assert.equal(a.r('confirm').hidden,false);assert.equal(a.n('family-apply').disabled,true);a.r('keep').onclick();assert.equal(a.r('evidence').value,'Authored evidence text');assert.equal(a.n('family-reviewed').checked,true);assert.equal(a.n('family-preview-content').textContent,text);assert.equal(a.n('family-apply').disabled,false);
 a.r('reset').onclick();a.r('evidence').value='EVENTLESS RAW';a.r('replace').onclick();assert.equal(a.r('evidence').value,'EVENTLESS RAW');assert.equal(a.r('error').hidden,false);assert.equal(a.r('confirm').hidden,true);
 a.r('reset').onclick();a.r('confirm-text').textContent='TAMPER';a.r('replace').onclick();assert.equal(a.r('evidence').value,'EVENTLESS RAW');
 a.r('reset').onclick();a.r('confirm').onkeydown({key:'Escape',preventDefault(){}});assert.equal(a.r('confirm').hidden,true);assert.equal(a.r('include').checked,true);
 a.r('reset').onclick();a.r('replace').onclick();assert.equal(a.r('include').checked,false);for(const k of WORKBOOK_RESEARCH_FIELDS)assert.equal(a.r(k).value,'');assert.equal(a.f('cover-title').value,'KEEP COVER');assert.equal(a.n('family-toc-include').checked,true);assert.deepEqual(a.calls(),before);
});

test('publication and research replacements cannot overlap and cancellation preserves current review',async()=>{
 const a=await app();draft(a);await review(a);a.f('reset').onclick();assert.equal(a.f('confirm').hidden,false);a.r('reset').onclick();assert.equal(a.r('confirm').hidden,true);assert.equal(a.r('error').hidden,false);await click(a,'family-preview');assert.equal(a.f('confirm').hidden,false);a.f('keep').onclick();assert.equal(a.n('family-reviewed').checked,true);
 a.r('reset').onclick();a.f('reset').onclick();assert.equal(a.f('confirm').hidden,true);assert.equal(a.r('confirm').hidden,false);const result=await offers(()=>click(a,'family-apply'));assert.equal(result.blobs.length,0);a.r('keep').onclick();assert.equal(a.n('family-reviewed').checked,true);assert.equal(a.r('rationale').value,'Authored rationale text');
});

test('hidden detached busy controls and eventless malformed drafts refuse without Word calls or unrelated loss',async()=>{
 const a=await app();draft(a);await review(a);a.n('activity-title').value='KEEP RAW';a.n('agenda-reviewed').checked=true;const initial=a.calls();
 await a.go('course');const before=a.calls();a.r('reset').onclick();assert.equal(a.r('confirm').hidden,true);const hidden=await offers(()=>click(a,'family-apply'));assert.equal(hidden.blobs.length,0);assert.deepEqual(a.calls(),before);await a.go('output');
 a.r('controls').hidden=true;a.r('reset').onclick();assert.equal(a.r('confirm').hidden,true);a.r('controls').hidden=false;a.r('reset').isConnected=false;a.r('reset').onclick();assert.equal(a.r('confirm').hidden,true);a.r('reset').isConnected=true;
 const computed=globalThis.getComputedStyle;a.r('controls').nodeType=1;globalThis.getComputedStyle=node=>({display:node===a.r('controls')?'none':'block',visibility:'visible'});try{a.r('reset').onclick();assert.equal(a.r('confirm').hidden,true);}finally{globalThis.getComputedStyle=computed;}
 let observed=false;a.h.hooks.beforeRefresh=()=>{if(observed)return;observed=true;a.r('reset').onclick();assert.equal(a.r('confirm').hidden,true);assert.equal(a.r('rationale').value,'Authored rationale text');};await click(a,'refresh');assert.equal(observed,true);assert.equal(a.n('activity-title').value,'KEEP RAW');assert.equal(a.n('agenda-reviewed').checked,true);assert.equal(a.calls().storageWrites,initial.storageWrites);assert.equal(a.calls().fetches,initial.fetches);
 a.r('evidence').validity={customError:true};await click(a,'family-preview');assert.match(a.n('family-brand-status').textContent,/invalid|unfinished/u);assert.equal(a.r('evidence').value,'Authored evidence text');assert.equal(a.h.saves,0);
});

test('research async edits visibility and consent roundtrips offer no stale copy; healthy repeats retain exact content',async()=>{
 for(const mutate of [a=>{set(a,'evidence','Changed during packaging');},a=>{a.r('include').checked=false;a.r('include').onchange();a.r('include').checked=true;a.r('include').onchange();},a=>{a.r('controls').hidden=true;},a=>{a.n('family-reviewed').checked=false;a.n('family-reviewed').onchange();a.n('family-reviewed').checked=true;a.n('family-reviewed').onchange();}]){
  const a=await app();draft(a);await review(a);const original=Zip.prototype.generateAsync;Zip.prototype.generateAsync=async function(...args){const bytes=await original.apply(this,args);mutate(a);return bytes;};try{const r=await offers(()=>click(a,'family-apply'));assert.equal(r.blobs.length,0);assert.equal(a.r('references').value,'Authored references text');assert.equal(a.h.saves,0);}finally{Zip.prototype.generateAsync=original;}
 }
 const a=await app();draft(a);await review(a);const copies=await offers(async()=>{await click(a,'family-apply');await click(a,'family-apply');});assert.equal(copies.blobs.length,2);assert.notEqual(copies.names[0],copies.names[1]);const zips=await Promise.all(copies.blobs.map(async b=>Zip.loadAsync(await b.arrayBuffer(),{checkCRC32:true})));for(const k of Object.keys(zips[0].files))assert.equal(await zips[0].file(k).async('string'),await zips[1].file(k).async('string'));assert.equal(a.r('evidence').value,'Authored evidence text');
});
