import test from 'node:test';
import assert from 'node:assert/strict';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {createRequire} from 'node:module';
import {dom,settle} from './helpers/app-dom.mjs';

const require=createRequire(import.meta.url),Zip=require('jszip');
let sequence=0;
async function app(project=fixture(),{office=true}={}){
  const h=makeHost(project),ui=dom(),n=id=>ui.nodes.get(id),downloads=[];let fetches=0,storageReads=0,storageWrites=0;
  const create=ui.document.createElement;ui.document.createElement=tag=>{const el=create(tag);if(tag==='a')el.click=()=>downloads.push({name:el.download,url:el.href});return el;};
  const blobs=new Map();URL.createObjectURL=blob=>{const key='blob:synthetic-'+blobs.size;blobs.set(key,blob);return key;};URL.revokeObjectURL=()=>{};
  const timer=globalThis.setTimeout;globalThis.setTimeout=(fn,ms,...args)=>ms===60000?0:timer(fn,ms,...args);globalThis.JSZip=Zip;
  Object.assign(globalThis,{Office:office?h.Office:undefined,Word:h.Word,document:ui.document,location:{search:''},
    fetch:async()=>{fetches++;throw new Error('No assessment network path is expected.');},
    localStorage:{getItem(){storageReads++;return null;},setItem(){storageWrites++;throw new Error('No assessment storage write is expected.');}}});
  await import(`../src/app.mjs?assessmentPanel=${++sequence}`);await settle();
  const go=async name=>{const route=ui.routes.find(b=>b.dataset.view===name);assert.ok(route);await route.onclick();await settle();};
  return {h,ui,n,go,downloads,blobs,calls:()=>({fetches,storageReads,storageWrites,runs:h.runs,saves:h.saves,refreshes:h.refreshes})};
}
const set=(a,id,value)=>{a.n('measure-'+id).value=value;a.n('measure-'+id).oninput?.();};

function capture(a){set(a,'title','Actual app synthetic export');set(a,'item-prompt','Select a public anchor');a.n('measure-item-save').onclick();a.n('measure-review').onclick();a.n('measure-reviewed').checked=true;a.n('measure-reviewed').onchange();a.n('measure-capture').onclick();assert.equal(a.n('measure-error').hidden,true,a.n('measure-error').textContent);}
function review(a){a.n('measure-export-preview').onclick();assert.equal(a.n('measure-export-error').hidden,true,a.n('measure-export-error').textContent);a.n('measure-export-reviewed').checked=true;a.n('measure-export-reviewed').onchange();}

test('actual app downloadBlob and JSZip create respondent bytes without Word storage network calls or lost drafts',async()=>{
  const a=await app();await a.go('assessments');capture(a);set(a,'export-instructions','Keep these authored instructions');a.n('activity-title').value='KEEP OTHER DRAFT';a.n('activity-minutes').value='1.';a.n('activity-minutes').validity={badInput:true};a.n('family-preview-content').textContent='KEEP FAMILY REVIEW';a.n('family-reviewed').checked=true;a.n('agenda-summary').textContent='KEEP AGENDA';a.n('agenda-reviewed').checked=true;
  const before=a.calls(),records=structuredClone(a.h.records);review(a);await a.n('measure-export-download').onclick();assert.equal(a.downloads.length,1);assert.match(a.downloads[0].name,/Blank-Respondent/u);const blob=a.blobs.get(a.downloads[0].url),zip=await Zip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true});assert.equal(Object.keys(zip.files).length,7);const body=await zip.file('word/document.xml').async('string');assert.match(body,/Actual app synthetic export/u);assert.match(body,/Keep these authored instructions/u);assert.doesNotMatch(body,/raw code|Reverse keyed|PRIVATE/u);assert.deepEqual(a.calls(),before);assert.deepEqual(a.h.records,records);assert.equal(a.n('activity-title').value,'KEEP OTHER DRAFT');assert.equal(a.n('activity-minutes').value,'1.');assert.equal(a.n('activity-minutes').validity.badInput,true);assert.equal(a.n('family-preview-content').textContent,'KEEP FAMILY REVIEW');assert.equal(a.n('family-reviewed').checked,true);assert.equal(a.n('agenda-reviewed').checked,true);
  a.n('measure-export-instructions').value='eventless changed';await a.n('measure-export-download').onclick();assert.equal(a.downloads.length,1);assert.equal(a.n('measure-export-error').hidden,false);assert.deepEqual(a.calls(),before);assert.equal(a.n('family-reviewed').checked,true);assert.equal(a.n('agenda-summary').textContent,'KEEP AGENDA');
});

test('actual app Word busy rejects cached export actions and section navigation retains drafts but invalidates review',async()=>{
  const a=await app();await a.go('assessments');capture(a);review(a);let observed=false;
  a.h.hooks.beforeRefresh=()=>{if(observed)return;observed=true;assert.equal(a.n('measure-export-download').disabled,true);const before=a.calls();a.n('measure-export-preview').onclick();a.n('measure-export-download').onclick();assert.deepEqual(a.calls(),before);assert.equal(a.downloads.length,0);};
  await a.n('refresh').onclick();await settle();assert.equal(observed,true);assert.equal(a.downloads.length,0);assert.equal(a.n('measure-title').value,'Actual app synthetic export');
  review(a);await a.go('course');await a.go('assessments');assert.equal(a.n('measure-export-reviewed').checked,false);assert.equal(a.n('measure-title').value,'Actual app synthetic export');assert.equal(a.downloads.length,0);
});

