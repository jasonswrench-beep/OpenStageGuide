import test from 'node:test';
import assert from 'node:assert/strict';
import {course,addModule,addActivity} from '../src/model.mjs';
import {searchOutline} from '../src/outline.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {agendaFiles} from '../src/agenda.mjs';
import {materialsFiles} from '../src/materials-output.mjs';
import {facilitatorFiles} from '../src/facilitator-output.mjs';
const stamp='2026-10-26T08:30:00.000Z';
const row=(item='  Markers  ',unit='  packs  ')=>({item,quantity:2,unit,use:'consumable',notes:'PRIVATE_MATERIAL_NOTE'});
function fixture(){let p=course('Outline material check');p=addModule(p,'Opening','Explain the topic').next;p=addModule(p,'Later','Practice the topic').next;p=addModule(p,'Empty module','Wrap up').next;
 const activity=(module,title,icon,materials)=>{p=addActivity(p,p.modules[module].id,{title,icon,minutes:5,participant:'Try '+title,facilitator:'Teach '+title,answers:'Discuss '+title,lines:1,...(materials===undefined?{}:{materials})}).next;};
 activity(0,'No list','Discussion');activity(0,'Explicit empty','Learning activity',[]);activity(0,'Saved supplies','Discussion',[row()]);activity(1,'Later no list','Discussion');return p;}

test('no recorded materials selects only absent or empty activity lists with stable source context',()=>{
 const p=fixture(),before=structuredClone(p),out=searchOutline(p,{status:'no-recorded-materials'}),ids=out.rows.filter(r=>r.type==='activity').map(r=>r.id);
 assert.deepEqual(ids,[p.modules[0].activities[0].id,p.modules[0].activities[1].id,p.modules[1].activities[0].id]);assert.equal(out.activityCount,3);assert.equal(out.moduleCount,2);assert.equal(out.rows[0].context,true);assert.deepEqual(out.rows.filter(r=>r.type==='activity').map(r=>[r.moduleIndex,r.activityIndex]),[[0,0],[0,1],[1,0]]);assert.ok(!out.rows.some(r=>r.id===p.modules[2].id));assert.ok(!out.rows.some(r=>r.id===p.modules[0].activities[2].id));assert.equal(JSON.stringify(out).includes('PRIVATE_MATERIAL_NOTE'),false);assert.deepEqual(p,before);
 const narrowed=searchOutline(p,{status:'no-recorded-materials',query:'Later topic',icon:'Discussion'});assert.equal(narrowed.activityCount,1);assert.equal(narrowed.rows[1].id,p.modules[1].activities[0].id);assert.equal(searchOutline(p,{status:'no-recorded-materials',query:'PRIVATE_MATERIAL_NOTE'}).rows.length,0);
 assert.equal(searchOutline(p,{status:'missing-objective'}).activityCount,0);assert.equal(searchOutline(p,{status:'all'}).activityCount,4);
});

test('supported nonempty material rows, including trim-compatible whitespace, never count as missing or change on filter',()=>{
 const p=fixture(),a=p.modules[0].activities[2],before=structuredClone(p);assert.equal(a.materials[0].item,'Markers');assert.equal(a.materials[0].unit,'packs');
 a.materials=[row()];const raw=structuredClone(p);const out=searchOutline(p,{status:'no-recorded-materials'});assert.ok(!out.rows.some(r=>r.id===a.id));assert.deepEqual(p,raw);
 const outputs=()=>[worksheetFiles(p,stamp),agendaFiles(p,'09:00',stamp),materialsFiles(p,stamp),facilitatorFiles(p,{answers:false,participant:false},stamp)];const beforeOutputs=outputs();searchOutline(p,{status:'no-recorded-materials',query:'topic',icon:'Discussion'});assert.deepEqual(outputs(),beforeOutputs);assert.deepEqual(p,raw);assert.notDeepEqual(raw,before);
});

test('unsupported or blank saved material values refuse without being recast as no recording',()=>{
 for(const bad of [null,'',{},0,false,[row('  ')],[row('Pencils',' \t ')]]){
  const p=fixture();p.modules[0].activities[0].materials=bad;const before=structuredClone(p);
  for(const status of ['no-recorded-materials','all','unfinished'])assert.throws(()=>searchOutline(p,{status}),e=>e instanceof Error&&!(e instanceof TypeError)&&/materials list|material name and unit|material entry/i.test(e.message));
  assert.deepEqual(p,before);
 }
});
