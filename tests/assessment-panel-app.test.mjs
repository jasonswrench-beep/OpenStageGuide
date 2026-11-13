import test from 'node:test';
import assert from 'node:assert/strict';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';

let sequence=0;
async function app(project=fixture(),{office=true}={}){
  const h=makeHost(project),ui=dom(),n=id=>ui.nodes.get(id);let fetches=0,storageReads=0,storageWrites=0;
  Object.assign(globalThis,{Office:office?h.Office:undefined,Word:h.Word,document:ui.document,location:{search:''},
    fetch:async()=>{fetches++;throw new Error('No assessment network path is expected.');},
    localStorage:{getItem(){storageReads++;return null;},setItem(){storageWrites++;throw new Error('No assessment storage write is expected.');}}});
  await import(`../src/app.mjs?assessmentPanel=${++sequence}`);await settle();
  const go=async name=>{const route=ui.routes.find(b=>b.dataset.view===name);assert.ok(route);await route.onclick();await settle();};
  return {h,ui,n,go,calls:()=>({fetches,storageReads,storageWrites,runs:h.runs,saves:h.saves,refreshes:h.refreshes})};
}
const set=(a,id,value)=>{a.n('measure-'+id).value=value;a.n('measure-'+id).oninput?.();};

test('real app exposes session assessment workspace without a course and preserves it across section navigation',async()=>{
  const a=await app(null);await a.go('assessments');assert.equal(a.n('view-assessments').hidden,false);assert.doesNotMatch(a.n('status').textContent,/Create your course/u);
  const baseline=a.calls();set(a,'title','Session draft without course');set(a,'item-prompt','Kept across navigation');a.n('measure-item-save').onclick();assert.equal(a.n('measure-error').hidden,true,a.n('measure-error').textContent);
  await a.go('course');assert.equal(a.n('view-assessments').hidden,true);await a.go('assessments');assert.equal(a.n('measure-title').value,'Session draft without course');assert.equal(a.n('measure-item-prompt').value,'Kept across navigation');assert.deepEqual(a.calls(),baseline);assert.equal(a.h.saves,0);
});

test('assessment validation stays local and retains unrelated raw drafts and reviewed output controls without Word calls',async()=>{
  const a=await app();await a.go('assessments');const rawIds=['activity-title','activity-participant','activity-facilitator','course-detail-audience','agenda-start','plan-name','family-brand-organizationName','family-preset-name','library-edit-name','materials-item'];
  for(const id of rawIds)if(a.n(id))a.n(id).value=`KEEP ${id}`;
  a.n('activity-minutes').value='';a.n('activity-minutes').validity={badInput:true};a.n('family-reviewed').checked=true;a.n('family-preview-content').textContent='KEEP REVIEWED FAMILY';a.n('agenda-reviewed').checked=true;a.n('agenda-summary').textContent='KEEP AGENDA';
  const before=Object.fromEntries(rawIds.filter(id=>a.n(id)).map(id=>[id,a.n(id).value])),globalError=a.n('error').textContent,body=structuredClone(a.h.records),calls=a.calls();
  set(a,'title','Invalid assessment');a.n('measure-item-save').onclick();assert.equal(a.n('measure-error').hidden,false);assert.equal(a.n('error').textContent,globalError);assert.deepEqual(Object.fromEntries(Object.keys(before).map(id=>[id,a.n(id).value])),before);assert.equal(a.n('activity-minutes').validity.badInput,true);assert.equal(a.n('family-reviewed').checked,true);assert.equal(a.n('family-preview-content').textContent,'KEEP REVIEWED FAMILY');assert.equal(a.n('agenda-reviewed').checked,true);assert.equal(a.n('agenda-summary').textContent,'KEEP AGENDA');assert.deepEqual(a.h.records,body);assert.deepEqual(a.calls(),calls);
});

test('real app busy gating rejects cached assessment actions and retains raw work through Word refresh success or failure',async()=>{
  for(const fail of [false,true]){
    const a=await app();await a.go('assessments');set(a,'title','KEEP WHILE WORD BUSY');set(a,'item-prompt','Unfinished item');a.n('measure-anchor-count').value='2.';a.n('measure-anchor-count').validity={badInput:true};a.n('measure-anchor-count').oninput();let observed=false;
    a.h.hooks.beforeRefresh=()=>{if(observed)return;observed=true;assert.equal(a.ui.document.querySelector('main').attributes['aria-busy'],'true');assert.equal(a.n('measure-item-save').disabled,true);assert.equal(a.n('measure-title').disabled,true);const calls=a.calls();a.n('measure-new').onclick();a.n('measure-item-save').onclick();assert.deepEqual(a.calls(),calls);assert.equal(a.n('measure-confirm').hidden,true);assert.equal(a.n('measure-title').value,'KEEP WHILE WORD BUSY');};
    if(fail)a.h.refreshFailures.add(a.h.refreshes+1);
    await a.n('refresh').onclick();await settle();assert.equal(observed,true);assert.equal(a.ui.document.querySelector('main').attributes['aria-busy'],'false');assert.equal(a.n('measure-title').value,'KEEP WHILE WORD BUSY');assert.equal(a.n('measure-anchor-count').value,'2.');assert.equal(a.n('measure-anchor-count').validity.badInput,true);assert.equal(a.n('measure-item-save').disabled,false);
  }
});

test('session assessment remains usable without Word readiness while global host errors stay untouched',async()=>{
  const a=await app(null,{office:false});const globalError=a.n('error').textContent;assert.match(globalError,/Open this panel/u);await a.go('assessments');const calls=a.calls();set(a,'title','Local synthetic draft');set(a,'item-prompt','One local item');a.n('measure-item-save').onclick();assert.equal(a.n('measure-error').hidden,true,a.n('measure-error').textContent);a.n('measure-review').onclick();assert.equal(a.n('measure-error').hidden,true,a.n('measure-error').textContent);assert.equal(a.n('error').textContent,globalError);assert.deepEqual(a.calls(),calls);
});
