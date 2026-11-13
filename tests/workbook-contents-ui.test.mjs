import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip');let serial=0;
async function app(){
 let p=fixture();p=editDeliveryPlan(p,{kind:'create',name:'Selected TOC test',targetMinutes:10,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]}).next;
 const h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);let fetches=0,reads=0,writes=0;
 Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip:Zip,fetch:async()=>{fetches++;throw Error('No contents network');},localStorage:{getItem(){reads++;return null;},setItem(){writes++;throw Error('No contents storage');}}});
 await import(`../src/app.mjs?toc-ui=${++serial}`);await settle();const go=async name=>{await ui.routes.find(x=>x.dataset.view===name).onclick();await settle();};await go('output');return {p,h,ui,n,go,id:p.deliveryPlans.plans[0].id,calls:()=>({fetches,reads,writes,saves:h.saves})};
}
const click=async(a,id)=>{await a.n(id).onclick();await settle();};
function include(a,value=true){a.n('family-toc-include').checked=value;a.n('family-toc-include').onchange();}
async function review(a){await click(a,'family-preview');assert.equal(a.n('error').hidden,true,a.n('error').textContent);assert.equal(a.n('family-reviewed').disabled,false,a.n('family-brand-status').textContent);a.n('family-reviewed').checked=true;a.n('family-reviewed').onchange();assert.equal(a.n('family-apply').disabled,false);}
async function offers(action){const blobs=[],names=[],oldUrl=URL.createObjectURL,oldCreate=document.createElement,oldTimer=globalThis.setTimeout;URL.createObjectURL=b=>{blobs.push(b);return 'blob:toc-test';};document.createElement=tag=>{const e=oldCreate(tag);if(tag==='a')e.click=()=>names.push(e.download);return e;};globalThis.setTimeout=(fn,ms,...args)=>ms===60000?0:oldTimer(fn,ms,...args);try{await action();}finally{URL.createObjectURL=oldUrl;document.createElement=oldCreate;globalThis.setTimeout=oldTimer;}return {blobs,names};}

test('contents default is off and all25 family transitions plus publication reset retain choice and unrelated drafts',async()=>{
 const a=await app(),before=structuredClone(a.p);assert.equal(a.n('family-toc-include').checked,false);include(a);a.n('family-front-cover-title').value='KEEP COVER';a.n('family-front-cover-title').oninput();a.n('family-front-copyright-noticeText').value='KEEP NOTICE';a.n('family-front-copyright-noticeText').oninput();a.n('family-brand-organizationName').value='KEEP BRAND';a.n('family-brand-organizationName').oninput();a.n('activity-title').value='KEEP ACTIVITY';
 for(const from of WORKBOOK_FAMILIES)for(const to of WORKBOOK_FAMILIES){a.n('family-template').value=from.id;a.n('family-template').onchange();await review(a);a.n('family-template').value=to.id;a.n('family-template').onchange();assert.equal(a.n('family-reviewed').checked,false);assert.equal(a.n('family-toc-include').checked,true);await review(a);assert.match(a.n('family-preview-content').textContent,/Contents navigation extension/u);assert.match(a.n('family-preview-content').textContent,/not automatically update/u);}
 a.n('workbook-plan').value=a.id;a.n('workbook-plan').onchange();await review(a);await a.go('course');await a.go('output');assert.equal(a.n('family-toc-include').checked,true);assert.equal(a.n('family-front-cover-title').value,'KEEP COVER');assert.equal(a.n('family-front-copyright-noticeText').value,'KEEP NOTICE');assert.equal(a.n('family-brand-organizationName').value,'KEEP BRAND');assert.equal(a.n('activity-title').value,'KEEP ACTIVITY');
 a.n('family-front-reset').onclick();assert.equal(a.n('family-apply').disabled,true);const result=await offers(()=>click(a,'family-apply'));assert.equal(result.blobs.length,0);a.n('family-front-keep').onclick();assert.equal(a.n('family-toc-include').checked,true);a.n('family-front-reset').onclick();a.n('family-front-replace').onclick();assert.equal(a.n('family-toc-include').checked,true);assert.equal(a.n('family-front-cover-title').value,'');assert.equal(a.n('family-brand-organizationName').value,'KEEP BRAND');assert.deepEqual(a.p,before);assert.equal(a.h.saves,0);
});

test('eventless contents choice validity and nested visibility tampering block file offers while repeated healthy downloads remain stable',async()=>{
 for(const mutate of [a=>{a.n('family-toc-include').checked=false;},a=>{a.n('family-toc-include').validity={customError:true};},a=>{a.n('family-toc-controls').hidden=true;},a=>{a.n('family-toc-include').hidden=true;}]){
  const a=await app();include(a);await review(a);a.n('activity-title').value='RETAIN';mutate(a);const result=await offers(()=>click(a,'family-apply'));assert.equal(result.blobs.length,0);assert.equal(a.n('activity-title').value,'RETAIN');assert.equal(a.h.saves,0);
 }
 const a=await app();include(a);await review(a);const result=await offers(async()=>{await click(a,'family-apply');await click(a,'family-apply');});assert.equal(result.blobs.length,2);assert.notEqual(result.names[0],result.names[1]);const zips=await Promise.all(result.blobs.map(async b=>Zip.loadAsync(await b.arrayBuffer(),{checkCRC32:true})));for(const name of Object.keys(zips[0].files))assert.equal(await zips[0].file(name).async('string'),await zips[1].file(name).async('string'));assert.match(await zips[0].file('word/document.xml').async('string'),/w:hyperlink w:anchor/u);assert.equal(a.n('family-toc-include').checked,true);assert.equal(a.h.saves,0);
});

test('async inclusion and acknowledgement roundtrips cancel offer without losing publication drafts',async()=>{
 for(const mutate of [a=>{include(a,false);include(a,true);},a=>{a.n('family-reviewed').checked=false;a.n('family-reviewed').onchange();a.n('family-reviewed').checked=true;a.n('family-reviewed').onchange();},a=>{a.n('family-toc-controls').hidden=true;}]){
  const a=await app();include(a);a.n('family-front-copyright-noticeText').value='KEEP PUBLIC WORDING';a.n('family-front-copyright-noticeText').oninput();await review(a);const old=Zip.prototype.generateAsync;Zip.prototype.generateAsync=async function(...args){const bytes=await old.apply(this,args);mutate(a);return bytes;};let result;try{result=await offers(()=>click(a,'family-apply'));}finally{Zip.prototype.generateAsync=old;}assert.equal(result.blobs.length,0);assert.equal(a.n('family-front-copyright-noticeText').value,'KEEP PUBLIC WORDING');assert.equal(a.h.saves,0);
 }
});

test('busy cached contents input preserves raw draft and rejects cached Apply without extra storage/network writes',async()=>{
 const a=await app();include(a);await review(a);const before=a.calls();let observed=false;
 a.h.hooks.beforeRefresh=()=>{if(observed)return;observed=true;assert.equal(a.n('family-toc-include').disabled,true);a.n('family-toc-include').onchange();a.n('family-apply').onclick();assert.equal(a.n('family-toc-include').checked,true);};
 const result=await offers(()=>click(a,'refresh'));assert.equal(observed,true);assert.equal(result.blobs.length,0);assert.equal(a.calls().writes,before.writes);assert.equal(a.calls().fetches,before.fetches);assert.equal(a.n('family-toc-include').checked,true);assert.equal(a.h.saves,0);
});

test('hidden workspace and ancestor actions refuse before Word calls; observed visibility roundtrips cancel async offers',async()=>{
 for(const hide of [a=>{a.n('view-output').hidden=true;},a=>{a.n('view-output').setAttribute('style','display:none');},a=>{const ancestor=a.ui.document.createElement('div');ancestor.hidden=true;a.n('view-output').parent=ancestor;},a=>{const ancestor=a.ui.document.createElement('div');ancestor.setAttribute('style','visibility:hidden');a.n('view-output').parent=ancestor;}]){
  const a=await app();include(a);a.n('family-front-cover-title').value='RETAIN COVER DRAFT';a.n('family-front-cover-title').oninput();a.n('family-front-copyright-noticeText').value='RETAIN NOTICE';a.n('family-front-copyright-noticeText').oninput();a.n('family-brand-organizationName').value='RETAIN BRAND';a.n('family-brand-organizationName').oninput();await review(a);const before={...a.calls(),runs:a.h.runs,refreshes:a.h.refreshes};hide(a);
  const result=await offers(async()=>{await click(a,'family-apply');await click(a,'family-preview');});assert.equal(result.blobs.length,0);assert.deepEqual({...a.calls(),runs:a.h.runs,refreshes:a.h.refreshes},before);assert.equal(a.n('error').hidden,true);assert.match(a.n('family-brand-status').textContent,/Open the Workbook workspace/u);assert.equal(a.n('family-toc-include').checked,true);assert.equal(a.n('family-front-cover-title').value,'RETAIN COVER DRAFT');assert.equal(a.n('family-front-copyright-noticeText').value,'RETAIN NOTICE');assert.equal(a.n('family-brand-organizationName').value,'RETAIN BRAND');
 }
 for(const mutate of [a=>{a.n('view-output').hidden=true;},a=>{a.n('view-output').hidden=true;a.n('family-reviewed').onchange();assert.equal(a.n('family-reviewed').checked,false);a.n('view-output').hidden=false;}]){
  const a=await app();include(a);a.n('family-front-cover-title').value='RETAIN ASYNC DRAFT';a.n('family-front-cover-title').oninput();await review(a);const original=Zip.prototype.generateAsync;Zip.prototype.generateAsync=async function(...args){const bytes=await original.apply(this,args);await mutate(a);return bytes;};let result;try{result=await offers(()=>click(a,'family-apply'));}finally{Zip.prototype.generateAsync=original;}assert.equal(result.blobs.length,0);assert.equal(a.n('family-front-cover-title').value,'RETAIN ASYNC DRAFT');assert.equal(a.n('family-toc-include').checked,true);assert.equal(a.h.saves,0);
 }
});
