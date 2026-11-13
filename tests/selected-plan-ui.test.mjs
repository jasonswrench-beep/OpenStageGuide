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
 let p=fixture();p=addActivity(p,p.modules[0].id,{title:'Second selected activity',minutes:9,participant:'Learner second',lines:2,materials:[{item:'Pads',quantity:2,unit:'each',use:'consumable',notes:'PRIVATE PREP'}]}).next;
 const groups=[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}];
 p=editDeliveryPlan(p,{kind:'create',name:'Short session',targetMinutes:null,groups}).next;
 return p;
}
async function app(p=saved()){
 const h=makeHost(p),ui=dom(),n=id=>ui.nodes.get(id);
 Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip});
 await import(`../src/app.mjs?selected-ui=${++sequence}`);await settle();return {h,n,p,id:p.deliveryPlans.plans[0].id};
}
async function click(n,id){await n(id).onclick();await settle();}
async function previewWorkbook(n){await click(n,'preview');}
async function previewMaterials(n){await click(n,'materials-checklist-preview');}

test('Workbook and Materials keep independent full-course defaults, selected order and raw authoring drafts',async()=>{
 const {n,p,id,h}=await app();assert.equal(n('workbook-plan').value,'');assert.equal(n('materials-checklist-plan').value,'');
 n('activity-title').value='RAW ACTIVITY DRAFT';n('plan-name').value='RAW PLAN DRAFT';n('plan-select').value=id;
 await previewWorkbook(n);assert.match(n('workbook-preview').textContent,/Full course/);assert.equal(n('workbook-plan-review-label').hidden,true);
 n('workbook-plan').value=id;n('workbook-plan').onchange();await previewWorkbook(n);
 assert.match(n('workbook-preview').textContent,/Short session/);assert.match(n('workbook-preview').textContent,/Second selected activity/);assert.doesNotMatch(n('workbook-preview').textContent,/PRIVATE PREP/);assert.equal(n('download').disabled,true);
 assert.equal(n('materials-checklist-plan').value,'');assert.equal(n('plan-select').value,id);
 n('workbook-plan-reviewed').checked=true;n('workbook-plan-reviewed').onchange();assert.equal(n('download').disabled,false);
 n('materials-checklist-plan').value=id;n('materials-checklist-plan').onchange();await previewMaterials(n);
 assert.match(n('materials-checklist-output').textContent,/Short session/);assert.match(n('materials-checklist-output').textContent,/Pads/);assert.doesNotMatch(n('materials-checklist-output').textContent,/PRIVATE PREP/);assert.equal(n('materials-checklist-download').disabled,true);
 n('materials-checklist-reviewed').checked=true;n('materials-checklist-reviewed').onchange();assert.equal(n('materials-checklist-download').disabled,false);
 assert.equal(n('activity-title').value,'RAW ACTIVITY DRAFT');assert.equal(n('plan-name').value,'RAW PLAN DRAFT');assert.equal(h.saves,0);assert.equal(p.deliveryPlans.plans[0].id,id);
});

test('selected output downloads require their own exact review and refuse async display or source drift',async()=>{
 for(const kind of ['workbook','materials'])for(const drift of ['display','source']){
  const {n,h,p,id}=await app();const isWorkbook=kind==='workbook',select=isWorkbook?'workbook-plan':'materials-checklist-plan',preview=isWorkbook?'preview':'materials-checklist-preview',review=isWorkbook?'workbook-plan-reviewed':'materials-checklist-reviewed',download=isWorkbook?'download':'materials-checklist-download',display=isWorkbook?'workbook-preview':'materials-checklist-output';
  n(select).value=id;n(select).onchange();await click(n,preview);n(review).checked=true;n(review).onchange();assert.equal(n(download).disabled,false);
  class ChangingZip extends JSZip{async generateAsync(options){if(drift==='display')n(display).textContent='Changed displayed review';else h.durable[KEY].deliveryPlans.plans[0].name='Changed after review';return super.generateAsync(options);}}
  globalThis.JSZip=ChangingZip;await click(n,download);assert.match(n('error').textContent,/changed/i,`${kind}/${drift}`);assert.equal(n(download).disabled,true);assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);assert.equal(p.modules.length,1);
 }
});

test('switching checklist scope clears consent; broken selected plans leave full-course preview usable',async()=>{
 const {n,id,h}=await app();n('materials-checklist-plan').value=id;n('materials-checklist-plan').onchange();await previewMaterials(n);n('materials-checklist-reviewed').checked=true;n('materials-checklist-reviewed').onchange();assert.equal(n('materials-checklist-download').disabled,false);
 n('materials-checklist-plan').value='';n('materials-checklist-plan').onchange();assert.equal(n('materials-checklist-reviewed').checked,false);assert.equal(n('materials-checklist-download').disabled,true);
 await previewMaterials(n);assert.match(n('materials-checklist-output').textContent,/Full course/);
 h.durable[KEY].deliveryPlans.plans[0].groups=[{moduleId:h.durable[KEY].modules[0].id,activityIds:['missing']}];
 n('workbook-plan').value=id;n('workbook-plan').onchange();await previewWorkbook(n);assert.match(n('error').textContent,/needs repair/i);assert.equal(n('download').disabled,true);
 n('workbook-plan').value='';n('workbook-plan').onchange();await previewWorkbook(n);assert.match(n('workbook-preview').textContent,/Full course/);
});

test('a missing-ID plan option cannot alias a valid sibling ID in either output selector',async()=>{
 const p=saved();p.deliveryPlans.plans[0].id='invalid:0';p.deliveryPlans.plans.unshift({name:'Broken missing ID',targetMinutes:null,groups:[]});
 const {n}=await app(p);
 for(const [select,preview,download] of [['workbook-plan','preview','download'],['materials-checklist-plan','materials-checklist-preview','materials-checklist-download']]){
  const options=n(select).children;assert.equal(options[1].value,'!invalid:0');assert.equal(options[2].value,'invalid:0');
  n(select).value=options[1].value;n(select).onchange();await click(n,preview);assert.match(n('error').textContent,/unique saved delivery plan/i);assert.equal(n(download).disabled,true);
  await click(n,'refresh');assert.equal(n(select).value,'!invalid:0');
  n(select).value='invalid:0';n(select).onchange();await click(n,preview);assert.equal(n('error').hidden,true);assert.equal(n(select).value,'invalid:0');
 }
});

test('selected workbook review binds visible writing-line count, placement and visibility before and after packaging',async()=>{
 for(const change of ['remove','move','hide','add','async-remove']){
  const {n,h,id}=await app();n('activity-title').value='RAW WRITING SPACE DRAFT';n('workbook-plan').value=id;n('workbook-plan').onchange();await previewWorkbook(n);
  n('workbook-plan-reviewed').checked=true;n('workbook-plan-reviewed').onchange();assert.equal(n('download').disabled,false);
  const root=n('workbook-preview'),before=root.textContent,lines=root.children.filter(child=>child.className==='writing-line');assert.equal(lines.length,2);
  if(change==='remove')lines[0].remove();
  if(change==='move'){lines[0].remove();const index=root.children.findIndex(child=>child.textContent==='Second selected activity');root.children.splice(index,0,lines[0]);lines[0].parent=root;}
  if(change==='hide')lines[0].hidden=true;
  if(change==='add'){const added=document.createElement('div');added.className='writing-line';root.append(added);}
  if(change==='async-remove'){
   class ChangingZip extends JSZip{async generateAsync(options){lines[0].remove();return super.generateAsync(options);}}
   globalThis.JSZip=ChangingZip;
  }
  if(change!=='async-remove')assert.equal(root.textContent,before,'text-only comparison would miss writing-space change');
  await click(n,'download');assert.match(n('error').textContent,/review changed/i,change);assert.equal(n('download').disabled,true);assert.equal(n('activity-title').value,'RAW WRITING SPACE DRAFT');assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);
 }
});

test('healthy repeated selected downloads retain each reviewed stamp and use distinct new filenames',async()=>{
 for(const kind of ['workbook','materials']){
  const {n,id}=await app(),select=kind==='workbook'?'workbook-plan':'materials-checklist-plan',preview=kind==='workbook'?'preview':'materials-checklist-preview',review=kind==='workbook'?'workbook-plan-reviewed':'materials-checklist-reviewed',download=kind==='workbook'?'download':'materials-checklist-download';
  n(select).value=id;n(select).onchange();await click(n,preview);n(review).checked=true;n(review).onchange();
  const generated=[],anchors=[],make=URL.createObjectURL,create=document.createElement,timer=globalThis.setTimeout;
  URL.createObjectURL=blob=>{generated.push(blob);return `blob:openstage-test-${generated.length}`;};
  document.createElement=tag=>{const element=create();if(tag==='a')anchors.push(element);return element;};
  globalThis.setTimeout=(fn,ms)=>ms===60000?0:timer(fn,ms);
  try{await click(n,download);await click(n,download);}finally{URL.createObjectURL=make;document.createElement=create;globalThis.setTimeout=timer;}
  assert.equal(n('error').hidden,true);assert.equal(generated.length,2);assert.equal(anchors.length,2);assert.notEqual(anchors[0].download,anchors[1].download);
  const first=await JSZip.loadAsync(await generated[0].arrayBuffer()),second=await JSZip.loadAsync(await generated[1].arrayBuffer());
  const firstCore=await first.file('docProps/core.xml').async('string'),secondCore=await second.file('docProps/core.xml').async('string');assert.equal(firstCore,secondCore);
  assert.match(firstCore,/2026-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z/);
 }
});
