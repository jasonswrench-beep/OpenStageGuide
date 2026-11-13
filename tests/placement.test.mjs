import test from 'node:test';
import assert from 'node:assert/strict';
import {addModule,addActivity} from '../src/model.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {buildAgenda} from '../src/planning.mjs';
import {fixture,records,makeHost,adapter,journal} from './helpers/mock-office.mjs';
function sample(){let p=fixture();p=addModule(p,'Middle module','').next;return addModule(p,'Final module','').next;}
const fields={title:'Earlier module practice',icon:'Discussion',minutes:7.5,participant:'EARLIER_LEARNER',facilitator:'EARLIER_PRIVATE',answers:'EARLIER_ANSWER',lines:3};
const textRows=h=>h.records.map(({tag,text})=>({tag,text}));
test('an unflagged guide with moved modules cannot append to the model-final module at body End',async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules.at(-1).id,fields);
  h.physicalOrder=h.records.map(r=>r.tag).reverse();
  await assert.rejects(w.insert(p,next,'activity',item),/module order/);
  assert.equal(h.paragraphs.length,0);assert.equal(h.pictures.length,0);assert.deepEqual(h.records,records(p));
});
for(const defect of ['physical-order','nested','table','unrelated','shared-paragraph','duplicate','missing'])test(`ambiguous earlier-module boundary (${defect}) stops without content mutation`,async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[0].id,fields);
  const anchor=`osg:module:${p.modules[1].id}`;
  if(defect==='physical-order'){h.physicalOrder=h.records.map(r=>r.tag);const at=h.physicalOrder.indexOf(anchor);h.physicalOrder.unshift(...h.physicalOrder.splice(at,2));}
  if(defect==='nested')h.nestedTags=new Set([anchor]);
  if(defect==='table')h.tableTags=new Set([anchor]);
  if(defect==='unrelated')h.unrelatedTags=new Set([anchor]);
  if(defect==='shared-paragraph')h.paragraphText={[anchor]:'Manual prefix '+p.modules[1].title};
  if(defect==='duplicate')h.records.push({...h.records.find(r=>r.tag===anchor)});
  if(defect==='missing')h.records=h.records.filter(r=>r.tag!==anchor);
  const original=structuredClone(h.records);
  await assert.rejects(w.insert(p,next,'activity',item));
  assert.equal(h.paragraphs.length,0);assert.equal(h.pictures.length,0);assert.deepEqual(h.records,original);
});

test('complete fields in the wrong physical module block finish and recovery, even when collection order is unchanged',async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[0].id,fields);
  h.hooks.beforeSync=n=>{if(n===10){const tags=h.records.map(r=>r.tag);h.physicalOrder=[...tags.filter(t=>!t.endsWith(item.id)),...tags.filter(t=>t.endsWith(item.id))];}};
  await assert.rejects(w.insert(p,next,'activity',item),/module order/);
  const content=structuredClone(h.records),counts={paragraphs:h.paragraphs.length,pictures:h.pictures.length,saves:h.saves};
  await assert.rejects(w.recover(),/module order/);await assert.rejects(w.readProject(),/paused/);
  assert.deepEqual(h.records,content);assert.deepEqual({paragraphs:h.paragraphs.length,pictures:h.pictures.length,saves:h.saves},counts);
  assert.equal(journal(h).phase,'prepared');
});

test('physical order is checked again on committed reopen and retained through later metadata saves',async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[0].id,fields);
  await w.insert(p,next,'activity',item);
  h.records.find(r=>r.tag.startsWith('osg:participant:')).text='Manual edit after placement';
  const fresh=await w.readProject();await w.persist(fresh);assert.equal(journal(h).placement,true);
  h.physicalOrder=h.records.map(r=>r.tag).reverse();
  const reopened=await adapter(h.reopen());await assert.rejects(reopened.loadProject(),/module order/);await assert.rejects(reopened.recover(),/module order/);
});
test('an anchor moved after validation pauses completion and recovery without deleting partial work',async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[0].id,fields);
  h.hooks.beforeSync=n=>{if(n===5)h.physicalOrder=h.records.map(r=>r.tag).reverse();};
  await assert.rejects(w.insert(p,next,'activity',item),/module order/);
  const content=structuredClone(h.records),count=h.paragraphs.length;
  await assert.rejects(w.recover(),/module order/);assert.deepEqual(h.records,content);assert.equal(h.paragraphs.length,count);
});

test('duplicate module titles still use the selected stable ID',async()=>{
  const p=sample();p.modules.forEach(m=>m.title='Same module name');
  const h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[1].id,fields);
  await w.insert(p,next,'activity',item);assert.deepEqual(textRows(h),records(next));
  assert.equal(next.modules[0].activities.length,1);assert.equal(next.modules[1].activities[0].id,item.id);
  assert.ok(h.paragraphs.every(r=>r.beforeTag===`osg:module:${p.modules[2].id}`));
});

for(const target of [0,1,2])test(`activity placement in module ${target+1} keeps other modules, text, IDs and output order`,async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[target].id,fields);
  const result=await w.insert(p,next,'activity',item);
  assert.deepEqual(result,next);assert.deepEqual(textRows(h),records(next));
  assert.deepEqual(h.records.filter(r=>!r.tag.endsWith(item.id)),records(p));
  assert.equal(h.pictures.length,1);assert.equal(h.pictures[0].width,72);
  const anchor=p.modules[target+1];
  assert.ok(h.paragraphs.every(para=>para.beforeTag===(anchor?`osg:module:${anchor.id}`:null)));
  assert.equal(h.paragraphs.length,9);
  assert.equal(journal(h).phase,'committed');
  const before={paragraphs:h.paragraphs.length,saves:h.saves};
  assert.deepEqual(await w.recover(),next);assert.deepEqual(await w.recover(),next);
  assert.deepEqual({paragraphs:h.paragraphs.length,saves:h.saves},before);
  const reopened=await adapter(h.reopen());assert.deepEqual(await reopened.loadProject(),next);
  assert.equal(buildAgenda(result,'09:00').rows.at(-1).moduleId,p.modules[target].id);
  const parts=worksheetFiles(result);for(const text of Object.values(parts)){assert.ok(!text.includes('EARLIER_PRIVATE'));assert.ok(!text.includes('EARLIER_ANSWER'));}
  assert.ok(parts['word/document.xml'].includes('EARLIER_LEARNER'));
});

for(const stage of [3,4,5,6,7,8,9,10,11,12,13])test(`earlier-module synchronization failure at stage ${stage} preserves other modules and never replays images`,async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[0].id,fields);
  h.syncFailures.set(stage,{apply:true});await assert.rejects(w.insert(p,next,'activity',item),/paused/);
  const savedRows=structuredClone(h.records),counts={paragraphs:h.paragraphs.length,pictures:h.pictures.length};
  assert.deepEqual(h.records.filter(r=>!r.tag.endsWith(item.id)),records(p));
  if(stage<9)await assert.rejects(w.recover(),/deleted/);else assert.deepEqual(await w.recover(),next);
  assert.deepEqual(h.records,savedRows);assert.deepEqual({paragraphs:h.paragraphs.length,pictures:h.pictures.length},counts);
  if(stage>=9){assert.deepEqual(await w.recover(),next);assert.equal(h.pictures.length,1);}
});

test('earlier-module destination cannot smuggle changes to another module or reuse a stale baseline',async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[0].id,fields);
  next.modules[1].title='Unexpected replacement';
  await assert.rejects(w.insert(p,next,'activity',item),/changed/);
  assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);
  const valid=addActivity(p,p.modules[0].id,fields);
  h.records.find(r=>r.tag.startsWith('osg:participant:')).text='Manual learner edit';
  await assert.rejects(w.insert(p,valid.next,'activity',valid.item),/changed/);
  assert.equal(h.saves,0);assert.equal(h.paragraphs.length,0);
  assert.equal(h.records.find(r=>r.tag.startsWith('osg:participant:')).text,'Manual learner edit');
});

for(const persist of [false,true])test(`earlier-module uncertain commit (persisted=${persist}) recovers at the selected module after reopen`,async()=>{
  const p=sample(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[1].id,fields);
  h.saveFailures.set(2,{persist});await assert.rejects(w.insert(p,next,'activity',item),/paused/);
  const counts={paragraphs:h.paragraphs.length,pictures:h.pictures.length};
  const reopened=await adapter(h.reopen());assert.deepEqual(await reopened.recover(),next);
  assert.deepEqual(textRows(h),records(next));assert.deepEqual({paragraphs:h.paragraphs.length,pictures:h.pictures.length},counts);
});
