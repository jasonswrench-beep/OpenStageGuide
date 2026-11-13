import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {KEY} from '../src/document-store.mjs';

let sequence=0;
function saved(){
 let p=fixture();p=addActivity(p,p.modules[0].id,{title:'Selected private session',icon:'Discussion',minutes:12,facilitator:'PRIVATE_SELECTED_TEACH',participant:'SELECTED_LEARNER',answers:'PRIVATE_SELECTED_ANSWER',lines:3,materials:[{item:'Cards',quantity:2,unit:'sets',use:'consumable',notes:'PRIVATE_SELECTED_PREP'}]}).next;
 p=editDeliveryPlan(p,{kind:'create',name:'Short private plan',targetMinutes:null,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}]}).next;
 return p;
}
async function app(p=saved()){
 const h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});
 await import(`../src/app.mjs?selected-private=${++sequence}`);await settle();return {h,n,p,id:p.deliveryPlans.plans[0].id};
}
async function click(n,id){await n(id).onclick();await settle();}

test('selected private scope is independent, opt-ins stay off, and raw authoring drafts survive review',async()=>{
 const {h,n,id}=await app();assert.equal(n('facilitator-plan').value,'');assert.equal(!!n('facilitator-answers').checked,false);assert.equal(!!n('facilitator-participant').checked,false);
 n('activity-facilitator').value='RAW PRIVATE DRAFT';n('plan-name').value='RAW PLAN DRAFT';n('plan-select').value=id;
 n('facilitator-plan').value=id;n('facilitator-plan').onchange();await click(n,'facilitator-preview');
 const view=n('facilitator-output').textContent;assert.match(view,/Short private plan/);assert.match(view,/PRIVATE_SELECTED_TEACH/);assert.match(view,/PRIVATE_SELECTED_PREP/);assert.match(view,/Course details and front matter are omitted/);assert.doesNotMatch(view,/PRIVATE DIRECTIONS|SECRET ANSWER|SELECTED_LEARNER|PRIVATE_SELECTED_ANSWER/);
 assert.equal(n('workbook-plan').value,'');assert.equal(n('materials-checklist-plan').value,'');assert.equal(n('plan-select').value,id);assert.equal(n('facilitator-download').disabled,true);
 n('facilitator-reviewed').checked=true;n('facilitator-reviewed').onchange();assert.equal(n('facilitator-download').disabled,false);
 assert.equal(n('activity-facilitator').value,'RAW PRIVATE DRAFT');assert.equal(n('plan-name').value,'RAW PLAN DRAFT');assert.equal(h.saves,0);
});

test('selected private async changes to scope, options, consent, preview or saved source withhold downloads',async()=>{
 for(const drift of ['scope','answers','participant','consent','display','structure','source']){
  const {n,h,id}=await app();n('activity-facilitator').value='RAW KEEP';n('facilitator-plan').value=id;n('facilitator-plan').onchange();await click(n,'facilitator-preview');n('facilitator-reviewed').checked=true;n('facilitator-reviewed').onchange();
  class ChangingZip extends JSZip{async generateAsync(options){
   if(drift==='scope')n('facilitator-plan').value='';
   if(drift==='answers')n('facilitator-answers').checked=true;
   if(drift==='participant')n('facilitator-participant').checked=true;
   if(drift==='consent')n('facilitator-reviewed').checked=false;
   if(drift==='display')n('facilitator-output').textContent='Changed reviewed text';
   if(drift==='structure'){const hidden=document.createElement('span');hidden.hidden=true;n('facilitator-output').append(hidden);}
   if(drift==='source')h.durable[KEY].deliveryPlans.plans[0].name='Changed saved plan';
   return super.generateAsync(options);
  }}
  globalThis.JSZip=ChangingZip;await click(n,'facilitator-download');assert.match(n('error').textContent,/changed/i,drift);assert.equal(n('facilitator-download').disabled,true);assert.equal(n('activity-facilitator').value,'RAW KEEP');assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);
 }
});

test('repeated healthy selected private downloads retain reviewed stamp and create distinct private names',async()=>{
 const {n,id}=await app();n('facilitator-plan').value=id;n('facilitator-plan').onchange();n('facilitator-answers').checked=true;n('facilitator-answers').onchange();n('facilitator-participant').checked=true;n('facilitator-participant').onchange();await click(n,'facilitator-preview');n('facilitator-reviewed').checked=true;n('facilitator-reviewed').onchange();
 const blobs=[],anchors=[],make=URL.createObjectURL,create=document.createElement,timer=globalThis.setTimeout;
 URL.createObjectURL=blob=>{blobs.push(blob);return `blob:private-test-${blobs.length}`;};document.createElement=tag=>{const element=create();if(tag==='a')anchors.push(element);return element;};globalThis.setTimeout=(fn,ms)=>ms===60000?0:timer(fn,ms);
 try{await click(n,'facilitator-download');await click(n,'facilitator-download');}finally{URL.createObjectURL=make;document.createElement=create;globalThis.setTimeout=timer;}
 assert.equal(n('error').hidden,true);assert.equal(blobs.length,2);assert.notEqual(anchors[0].download,anchors[1].download);assert.match(anchors[0].download,/Private Facilitator Guide/);
 const first=await JSZip.loadAsync(await blobs[0].arrayBuffer()),second=await JSZip.loadAsync(await blobs[1].arrayBuffer());assert.equal(await first.file('docProps/core.xml').async('string'),await second.file('docProps/core.xml').async('string'));
 const body=await first.file('word/document.xml').async('string');assert.match(body,/PRIVATE_SELECTED_ANSWER/);assert.match(body,/SELECTED_LEARNER/);assert.doesNotMatch(body,/SECRET ANSWER|PRIVATE DIRECTIONS/);
});

test('malformed selected plan refuses while full-course private preview remains available',async()=>{
 const {n,h,id}=await app();h.durable[KEY].deliveryPlans.plans[0].groups=[{moduleId:h.durable[KEY].modules[0].id,activityIds:['missing']}];n('facilitator-plan').value=id;n('facilitator-plan').onchange();await click(n,'facilitator-preview');assert.match(n('error').textContent,/repair/i);assert.equal(n('facilitator-download').disabled,true);
 n('facilitator-plan').value='';n('facilitator-plan').onchange();await click(n,'facilitator-preview');assert.equal(n('error').hidden,true);assert.match(n('facilitator-output').textContent,/Full course/);
});
