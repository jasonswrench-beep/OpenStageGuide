import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {activityMaterials,editActivityMaterials,addActivity,validate} from '../src/model.mjs';
import {normalizeMaterials} from '../src/materials.mjs';
import {planMaterials} from '../src/planning.mjs';
import {fixture,makeHost,adapter,journal} from './helpers/mock-office.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {masterFixture,installParser,assets} from './helpers/checkpoint-fixture.mjs';
import {createCheckpoint,readCheckpoint} from '../src/checkpoint.mjs';
import {emptyLibrary,saveLibraryCopy,serializeLibrary,readLibrary,removeLibraryCopy,undoLibraryRemoval,libraryPlaceholders} from '../src/library.mjs';
const material=(quantity=2,use='consumable')=>({item:'Handout',quantity,unit:'copies',use,notes:'PRIVATE MATERIAL PREPARATION'});
test('older records need no migration; material edits are immutable metadata-only and semantic noops retain revision',()=>{
  const p=fixture(),id=p.modules[0].activities[0].id;p.unknown={keep:true};p.modules[0].activities[0].future={keep:true};const before=JSON.stringify(p);
  assert.deepEqual(activityMaterials(p,id),[]);assert.deepEqual(editActivityMaterials(p,id,[]),p);const next=editActivityMaterials(p,id,[material()]);assert.equal(next.revision,p.revision+1);assert.deepEqual(next.modules[0].activities[0].future,{keep:true});assert.equal(next.modules[0].activities[0].id,id);assert.equal(JSON.stringify(p),before);assert.deepEqual(editActivityMaterials(next,id,[material('2.000')]),next);
});
test('material boundaries refuse unknown/duplicate entries and unsupported quantities without changing input',()=>{
  for(const quantity of ['',true,0,-1,Infinity,100001,0.0001,1.2345])assert.throws(()=>normalizeMaterials([material(quantity)]));
  for(const entry of [{...material(),item:''},{...material(),unit:''},{...material(),use:'unknown'},{...material(),notes:'x'.repeat(2001)},{...material(),sourceProject:{secret:true}}])assert.throws(()=>normalizeMaterials([entry]));
  assert.throws(()=>normalizeMaterials([material(),{...material(),item:' HANDOUT ',unit:'copies'}]));assert.throws(()=>normalizeMaterials(Array.from({length:101},()=>material())));assert.throws(()=>validate({...fixture(),modules:[{id:'m',activities:[{id:'a',materials:null}]}]}));
});
test('preparation totals add consumables, maximize sequential reusable needs and separate units/types exactly',()=>{
  let p=fixture(),id=p.modules[0].activities[0].id;p=editActivityMaterials(p,id,[material(0.1),{...material(2,'reusable'),item:'Board',unit:'pieces'}]);const added=addActivity(p,p.modules[0].id,{title:'Later task',participant:'Learner task',minutes:5,lines:0});p=editActivityMaterials(added.next,added.item.id,[{...material(0.2),item:' handout ',unit:'copies'},{...material(3,'reusable'),item:'BOARD',unit:'pieces'},{...material(1),unit:'sets'}]);
  const before=JSON.stringify(p),plan=planMaterials(p);assert.equal(plan.rows.length,3);assert.equal(plan.rows[0].quantity,0.3);assert.equal(plan.rows[1].quantity,3);assert.equal(plan.rows[2].quantity,1);assert.equal(plan.configuredActivities,2);assert.equal(plan.rows[0].notes.length,1);assert.match(plan.assumption,/sequential/);assert.equal(JSON.stringify(p),before);
});
test('materials persist through settings-only save/reopen and preserve timing evidence and every Word field',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=p.modules[0].activities[0].id;const timed=await w.updateActivitySettings(p,id,{minutes:7.5,lines:3}),body=structuredClone(h.records),cueWrites=h.cueWrites;
  const saved=await w.updateActivityMaterials(timed,id,[material()]);assert.deepEqual(h.records,body);assert.equal(h.cueWrites,cueWrites);assert.equal(h.paragraphs.length,0);assert.equal(journal(h).kind,'metadata');assert.ok(journal(h).cueChecks.includes(id));const saves=h.saves;await w.updateActivityMaterials(saved,id,[material()]);assert.equal(h.saves,saves);assert.deepEqual(await (await adapter(h.reopen())).loadProject(),saved);
});
test('invalid input stays healthy and stale material drafts cannot overwrite direct Word or another metadata edit',async()=>{
  const p=fixture(),id=p.modules[0].activities[0].id,h=makeHost(p),w=await adapter(h);await assert.rejects(w.updateActivityMaterials(p,id,[material(-1)]),e=>e.code==='OSG_INPUT');assert.equal(w.isBlocked(),false);assert.equal(h.saves,0);
  h.records.find(r=>r.tag===`osg:participant:${id}`).text='New direct learner text';await assert.rejects(w.updateActivityMaterials(p,id,[material()]),e=>e.code==='OSG_STALE');assert.equal(h.saves,0);assert.equal(w.isBlocked(),false);
  const current=await w.persist(await w.readProject()),saved=await w.updateActivityMaterials(current,id,[material(3)]);await assert.rejects(w.updateActivityMaterials(current,id,[material(2)]),e=>e.code==='OSG_STALE');assert.deepEqual(await w.loadProject(),saved);
});
for(const save of [1,2])for(const persisted of [false,true])test(`interrupted materials metadata save ${save}/${persisted} recovers without body writes or repeated revisions`,async()=>{
  const p=fixture(),id=p.modules[0].activities[0].id,h=makeHost(p),w=await adapter(h),body=structuredClone(h.records);h.saveFailures.set(save,{persist:persisted});await assert.rejects(w.updateActivityMaterials(p,id,[material()]),/paused/);assert.equal(w.isBlocked(),true);const recovered=await w.recover();assert.deepEqual(recovered,editActivityMaterials(p,id,[material()]));assert.deepEqual(h.records,body);assert.equal(h.cueWrites,0);assert.equal(h.paragraphs.length,0);const saves=h.saves;assert.deepEqual(await w.recover(),recovered);assert.equal(h.saves,saves);
});
test('materials survive full-master backup while every learner part excludes preparation and unknown material metadata',async()=>{
  const p=fixture(),id=p.modules[0].activities[0].id,next=editActivityMaterials(p,id,[material()]);for(const text of Object.values(worksheetFiles(next)))assert.doesNotMatch(text,/PRIVATE MATERIAL PREPARATION|Handout|copies/);
  installParser();const {master}=await masterFixture(next),backup=await createCheckpoint(master,next,JSZip,assets),checked=await readCheckpoint(backup.data,JSZip);assert.deepEqual(checked.project.modules[0].activities[0].materials,[material()]);
});
test('unit case remains significant and reusable totals explicitly assume a shared available pool',()=>{
  let p=fixture(),id=p.modules[0].activities[0].id;p=editActivityMaterials(p,id,[{...material(2),unit:'mg'},{...material(3),unit:'Mg'}]);const plan=planMaterials(p);assert.equal(plan.rows.length,2);assert.deepEqual(plan.rows.map(r=>r.quantity),[2,3]);assert.match(plan.assumption,/shared pool available again/);assert.match(plan.assumption,/retained supplies add/);
});
test('materials survive portable library copy, import/export, insertion and remove/undo without source-ID transplantation',()=>{
  const p=fixture(),{id,...fields}=p.modules[0].activities[0];fields.materials=[material()];const library=saveLibraryCopy(emptyLibrary(),'Prepared activity',[],fields),portable=readLibrary(serializeLibrary(library)),entry=portable.entries[0],copy=addActivity(p,p.modules[0].id,entry.fields);
  assert.deepEqual(copy.item.materials,[material()]);assert.notEqual(copy.item.id,id);assert.notEqual(copy.item.id,entry.id);copy.item.materials[0].notes='Independent edit';assert.equal(entry.fields.materials[0].notes,'PRIVATE MATERIAL PREPARATION');const {next,removed}=removeLibraryCopy(portable,entry.id);assert.deepEqual(undoLibraryRemoval(next,removed),portable);for(const part of Object.values(worksheetFiles(copy.next)))assert.doesNotMatch(part,/Independent edit|Handout|copies/);
  entry.fields.materials[0].notes='TODO preparation';assert.deepEqual(libraryPlaceholders(entry.fields),['material 1 notes']);
});
