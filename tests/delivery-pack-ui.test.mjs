import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {KEY} from '../src/document-store.mjs';

let sequence=0;
globalThis.DOMParser=DOMParser;
function saved(){let p=fixture();p=editDeliveryPlan(p,{kind:'create',name:'Reviewed short plan',targetMinutes:null,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]}).next;return p;}
async function app(){const p=saved(),h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});n('pack-start').value='09:00';n('pack-route').value='shareable';await import(`../src/app.mjs?delivery-pack=${++sequence}`);await settle();return {p,h,n,id:p.deliveryPlans.plans[0].id};}
async function click(n,id){await n(id).onclick();await settle();}
async function review(n,id,route='shareable'){n('pack-plan').value=id;n('pack-plan').onchange();n('pack-route').value=route;n('pack-route').onchange();await click(n,'pack-preview');n('pack-reviewed').checked=true;n('pack-reviewed').onchange();if(route==='private'){n('pack-private-consent').checked=true;n('pack-private-consent').onchange();}}

test('pack route has its own plan and consent; private choices begin off and raw drafts survive',async()=>{
 const {n,h,id}=await app();n('activity-facilitator').value='RAW AUTHOR DRAFT';n('plan-name').value='RAW PLAN DRAFT';n('workbook-plan').value='';n('facilitator-plan').value='';
 await review(n,id);assert.equal(n('pack-download').disabled,false);assert.match(n('pack-output').textContent,/SHAREABLE|shareable/);assert.doesNotMatch(n('pack-output').textContent,/SELECTED_PRIVATE|SECRET ANSWER/);
 assert.match(n('pack-output').textContent,/Learner task|Original activity/);assert.match(n('pack-output').textContent,/complete structured content/);
 assert.equal(n('facilitator-plan').value,'');assert.equal(n('workbook-plan').value,'');assert.equal(n('activity-facilitator').value,'RAW AUTHOR DRAFT');assert.equal(n('plan-name').value,'RAW PLAN DRAFT');assert.equal(h.saves,0);
 n('pack-route').value='private';n('pack-route').onchange();assert.equal(n('pack-download').disabled,true);assert.equal(n('pack-answers').checked,false);assert.equal(n('pack-participant').checked,false);
 await review(n,id,'private');assert.equal(n('pack-download').disabled,false);n('pack-private-consent').checked=false;n('pack-private-consent').onchange();assert.equal(n('pack-download').disabled,true);
});

test('pack download rechecks route, options, consent, display and saved source during asynchronous assembly',async()=>{
 for(const drift of ['route','plan','start','answers','consent','display','structure','source']){
  const {n,h,id}=await app();n('activity-facilitator').value='RAW KEEP';await review(n,id,'private');
  class ChangingZip extends JSZip{async generateAsync(o){
   if(drift==='route')n('pack-route').value='shareable';
   if(drift==='plan')n('pack-plan').value='';
   if(drift==='start')n('pack-start').value='10:00';
   if(drift==='answers')n('pack-answers').checked=true;
   if(drift==='consent')n('pack-private-consent').checked=false;
   if(drift==='display')n('pack-output').textContent='Altered review';
   if(drift==='structure'){const node=document.createElement('span');node.hidden=true;n('pack-output').append(node);}
   if(drift==='source')h.durable[KEY].deliveryPlans.plans[0].name='Changed plan';
   return super.generateAsync(o);
  }}
  globalThis.JSZip=ChangingZip;await click(n,'pack-download');assert.match(n('error').textContent,/changed|review/i,drift);assert.equal(n('pack-download').disabled,true);assert.equal(n('activity-facilitator').value,'RAW KEEP');assert.equal(h.saves,0);
 }
});

test('two healthy downloads retain the reviewed generation and get distinct safe names',async()=>{
 const {n,id}=await app();await review(n,id);
 const blobs=[],anchors=[],make=URL.createObjectURL,create=document.createElement,timer=globalThis.setTimeout;
 URL.createObjectURL=blob=>{blobs.push(blob);return `blob:delivery-pack-${blobs.length}`;};document.createElement=tag=>{const element=create();if(tag==='a')anchors.push(element);return element;};globalThis.setTimeout=(fn,ms)=>ms===60000?0:timer(fn,ms);
 try{await click(n,'pack-download');await click(n,'pack-download');}finally{URL.createObjectURL=make;document.createElement=create;globalThis.setTimeout=timer;}
 assert.equal(n('error').hidden,true);assert.equal(blobs.length,2);assert.notEqual(anchors[0].download,anchors[1].download);assert.match(anchors[0].download,/Shareable Learner Pack/);
 const first=await JSZip.loadAsync(await blobs[0].arrayBuffer()),second=await JSZip.loadAsync(await blobs[1].arrayBuffer()),a=JSON.parse(await first.file('manifest.json').async('string')),b=JSON.parse(await second.file('manifest.json').async('string'));
 assert.deepEqual(a.identity,b.identity);assert.deepEqual(a.members.map(m=>m.name),['agenda.docx','participant-workbook.docx','materials-checklist.docx']);
});

test('partial assembly never offers a ZIP and leaves drafts intact',async()=>{
 const {n,id}=await app();n('activity-facilitator').value='RAW KEEP';await review(n,id);
 let attempts=0,downloads=0;class FailingZip extends JSZip{async generateAsync(o){if(++attempts===2)throw new Error('Partial assembly stopped');return super.generateAsync(o);}}
 globalThis.JSZip=FailingZip;const old=URL.createObjectURL;URL.createObjectURL=()=>{downloads++;return 'blob:unexpected';};
 try{await click(n,'pack-download');}finally{URL.createObjectURL=old;}
 assert.match(n('error').textContent,/Partial assembly stopped/);assert.equal(downloads,0);assert.equal(n('activity-facilitator').value,'RAW KEEP');
});

test('hiding or styling the reviewed pack root during assembly withholds both routes',async()=>{
 for(const route of ['shareable','private'])for(const drift of ['hidden','style']){
  const {n,id}=await app();n('activity-facilitator').value='RAW KEEP';await review(n,id,route);
  let offered=0;const old=URL.createObjectURL;URL.createObjectURL=()=>{offered++;return 'blob:unexpected';};
  class ChangingZip extends JSZip{async generateAsync(o){if(drift==='hidden')n('pack-output').hidden=true;else n('pack-output').style={cssText:'display:none'};return super.generateAsync(o);}}
  globalThis.JSZip=ChangingZip;try{await click(n,'pack-download');}finally{URL.createObjectURL=old;}
  assert.match(n('error').textContent,/review changed/i);assert.equal(offered,0);assert.equal(n('activity-facilitator').value,'RAW KEEP');
 }
});
