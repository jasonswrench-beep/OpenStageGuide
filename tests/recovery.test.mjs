import test from 'node:test';
import assert from 'node:assert/strict';
import {KEY,PENDING} from '../src/document-store.mjs';
import {addModule,addActivity,course,reconcile} from '../src/model.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {fixture,records,makeHost,adapter,journal} from './helpers/mock-office.mjs';

test('successful insertion retains a confirmed receipt and recovery never inserts it again',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','An objective.');
  const result=await w.insert(p,next,'module',item),count=h.paragraphs.length,saves=h.saves;
  assert.deepEqual(result,next);assert.equal(journal(h).phase,'committed');assert.equal(w.isBlocked(),false);
  assert.deepEqual(await w.recover(),result);assert.deepEqual(await w.recover(),result);
  assert.equal(h.paragraphs.length,count);assert.equal(h.saves,saves);
  const reopened=await adapter(h.reopen());assert.deepEqual(await reopened.loadProject(),result);
});
for(const persist of [false,true])test(`failed initial journal save (persisted=${persist}) cannot write Word content`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');h.saveFailures.set(1,{persist});
  await assert.rejects(w.insert(p,next,'module',item),/paused/);
  assert.equal(h.paragraphs.length,0);assert.deepEqual(h.records,records(p));assert.equal(w.isBlocked(),true);
  await assert.rejects(w.readProject(p),/paused/);await assert.rejects(w.persist(p),/paused/);await assert.rejects(w.recover(),/deleted/);
  const reopened=await adapter(h.reopen());
  if(persist)await assert.rejects(reopened.loadProject(),/paused/);else assert.deepEqual(await reopened.loadProject(),p);
});
for(const persist of [false,true])test(`failed final settings save (persisted=${persist}) blocks stale pane operations and recovers once`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');h.saveFailures.set(2,{persist});
  await assert.rejects(w.insert(p,next,'module',item),/paused/);assert.equal(w.isBlocked(),true);assert.equal(h.memory[PENDING].phase,'prepared');
  const count=h.paragraphs.length,saves=h.saves;
  await assert.rejects(w.readProject(p),/paused/);await assert.rejects(w.persist(p),/paused/);await assert.rejects(w.insert(p,next,'module',item),/paused/);await assert.rejects(w.jump(p.modules[0].id),/paused/);
  assert.equal(h.saves,saves);assert.equal(h.paragraphs.length,count);
  const recovered=await w.recover();assert.deepEqual(recovered,next);assert.equal(w.isBlocked(),false);
  assert.deepEqual(h.durable[KEY],next);assert.equal(journal(h).phase,'committed');
  assert.deepEqual(await w.recover(),next);assert.equal(h.paragraphs.length,count);
});
for(const persist of [false,true])test(`reopening after an uncertain final save (persisted=${persist}) uses durable evidence`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');h.saveFailures.set(2,{persist});
  await assert.rejects(w.insert(p,next,'module',item));const count=h.paragraphs.length;
  const reopened=await adapter(h.reopen());
  if(persist)assert.deepEqual(await reopened.loadProject(),next);else await assert.rejects(reopened.loadProject(),/paused/);
  assert.deepEqual(await reopened.recover(),next);assert.equal(h.paragraphs.length,count);
});
for(const persist of [false,true])test(`failed recovery settings save (persisted=${persist}) remains blocked and retry is idempotent`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');
  h.saveFailures.set(2,{persist:false});h.saveFailures.set(3,{persist});await assert.rejects(w.insert(p,next,'module',item));
  await assert.rejects(w.recover(),/paused/);assert.equal(w.isBlocked(),true);await assert.rejects(w.persist(p),/paused/);
  const count=h.paragraphs.length;assert.deepEqual(await w.recover(),next);assert.deepEqual(await w.recover(),next);assert.equal(h.paragraphs.length,count);
});
for(const stage of [1,2,3,4,5,6])test(`Word synchronization failure at module transaction stage ${stage} preserves content and gates recovery`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','Objective.');h.syncFailures.set(stage,{apply:true});
  await assert.rejects(w.insert(p,next,'module',item),/paused/);assert.equal(w.isBlocked(),true);
  const content=structuredClone(h.records),count=h.paragraphs.length;
  if(stage===1)assert.deepEqual(await w.recover(),p);
  else if(stage<4)await assert.rejects(w.recover(),/deleted/);
  else assert.deepEqual(await w.recover(),next);
  assert.deepEqual(h.records,content);assert.equal(h.paragraphs.length,count);
});
test('a dropped first write batch leaves the durable prepared journal and cannot recover incomplete fields',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');h.syncFailures.set(3,{apply:false});
  await assert.rejects(w.insert(p,next,'module',item));assert.equal(h.paragraphs.length,0);assert.equal(journal(h).phase,'prepared');
  const reopened=await adapter(h.reopen());await assert.rejects(reopened.loadProject(),/paused/);await assert.rejects(reopened.recover(),/deleted/);assert.deepEqual(h.durable[KEY],p);
});
for(const stage of [3,4,5,6,7,8,9,10,11,12])test(`activity synchronization failure at stage ${stage} never duplicates its image or fields`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addActivity(p,p.modules[0].id,{title:'New activity',minutes:5,participant:'Learner task',facilitator:'',answers:'',lines:1});
  h.syncFailures.set(stage,{apply:true});await assert.rejects(w.insert(p,next,'activity',item));const count=h.records.length,images=h.pictures.length;
  if(stage<8)await assert.rejects(w.recover(),/deleted/);else assert.deepEqual(await w.recover(),next);
  assert.equal(h.records.length,count);assert.equal(h.pictures.length,images);
});
test('recovery preserves direct manual edits without leaking instructor fields into the workbook',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');h.saveFailures.set(2,{persist:false});await assert.rejects(w.insert(p,next,'module',item));
  h.records.find(x=>x.tag.startsWith('osg:participant:')).text='Manually revised learner task';
  h.records.find(x=>x.tag===`osg:objective:${item.id}`).text='Manually added objective';
  const body=structuredClone(h.records),recovered=await w.recover();
  assert.equal(recovered.revision,next.revision+1);assert.equal(recovered.modules[0].activities[0].participant,'Manually revised learner task');assert.equal(recovered.modules[1].objective,'Manually added objective');
  assert.deepEqual(h.records,body);assert.deepEqual(await w.recover(),recovered);
  const files=JSON.stringify(worksheetFiles(recovered));assert.match(files,/Manually revised learner task/);assert.doesNotMatch(files,/PRIVATE DIRECTIONS|SECRET ANSWER/);
});
test('edits during insertion are reconciled into the returned committed record, never overwritten',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');
  h.hooks.beforeSync=(n,host)=>{if(n===4)host.records.find(x=>x.tag.startsWith('osg:participant:')).text='Edited during insertion';};
  const result=await w.insert(p,next,'module',item);assert.equal(result.modules[0].activities[0].participant,'Edited during insertion');assert.deepEqual(h.durable[KEY],result);
});
test('manual edits during preparation stop content mutation and remain in Word',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');
  h.hooks.beforeSave=(n,host)=>{if(n===1)host.records[0].text='Manual course title';};
  await assert.rejects(w.insert(p,next,'module',item),/changed during preparation/);assert.equal(h.paragraphs.length,0);assert.equal(h.records[0].text,'Manual course title');assert.equal(journal(h).phase,'prepared');
});
for(const mode of ['missing','duplicate','orphan','unknown-type'])test(`${mode} managed fields block reads, persistence, insertion and incomplete recovery`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h);
  if(mode==='missing')h.records.pop();if(mode==='duplicate')h.records.push({...h.records[0]});if(mode==='orphan')h.records.push({tag:'osg:module:unrecorded',text:'Preserve me'});if(mode==='unknown-type')h.records.push({tag:'osg:other:unrecorded',text:'Preserve me'});
  const content=structuredClone(h.records);await assert.rejects(w.loadProject());assert.equal(w.isBlocked(),true);await assert.rejects(w.readProject(p));await assert.rejects(w.persist(p));await assert.rejects(w.recover());
  assert.deepEqual(h.records,content);assert.equal(h.saves,0);
});
test('a fresh document with orphaned managed controls cannot silently start a new course',async()=>{
  const h=makeHost(null),w=await adapter(h);h.records.push({tag:'osg:course:orphan',text:'Existing course'});
  await assert.rejects(w.loadProject(),/unexpected/);await assert.rejects(w.insert(null,course('Another course'),'course'),/paused/);assert.equal(h.saves,0);
});
test('stale pane refresh adopts current committed metadata and stale persist cannot replace it',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');await w.insert(p,next,'module',item);
  assert.deepEqual(await w.readProject(p),next);const saves=h.saves;
  await assert.rejects(w.persist(p),/document changed/);assert.equal(h.saves,saves);assert.deepEqual(h.durable[KEY],next);
  assert.deepEqual(await w.persist(await w.readProject(p)),next);
});
for(const persist of [false,true])test(`metadata refresh save failure (persisted=${persist}) cannot write stale data`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h);h.records[0].text='Manually changed title';
  const fresh=await w.readProject(p);h.saveFailures.set(2,{persist});await assert.rejects(w.persist(fresh));await assert.rejects(w.persist(p),/paused/);
  const result=await w.recover();assert.equal(result.title,'Manually changed title');assert.equal(result.revision,fresh.revision);assert.equal(h.paragraphs.length,0);
});
test('settings refresh failure after the final save stays blocked until durable completion is checked',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');h.refreshFailures.add(4);
  await assert.rejects(w.insert(p,next,'module',item));assert.equal(w.isBlocked(),true);assert.equal(journal(h).phase,'committed');
  assert.deepEqual(await w.recover(),next);assert.equal(h.saves,2);
});
test('a mismatched committed receipt is not treated as a recoverable older snapshot',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h);h.durable[PENDING]={schema:2,phase:'committed',txId:'mismatch',project:{...p,revision:p.revision+10}};
  await assert.rejects(w.loadProject(),/paused/);await assert.rejects(w.recover(),/does not match/);assert.deepEqual(h.durable[KEY],p);assert.equal(h.saves,0);
});
test('legacy pending insertion recovers only a complete compatible append and never replays Word writes',async()=>{
  const p=fixture(),h=makeHost(p),{next,item}=addModule(p,'Legacy module','');h.records=records(next);h.durable[PENDING]={project:next,kind:'module',id:item.id};h.reopen();
  const w=await adapter(h);await assert.rejects(w.loadProject());assert.deepEqual(await w.recover(),next);assert.equal(h.paragraphs.length,0);assert.equal(journal(h).phase,'committed');
});
test('model reconciliation rejects unique unknown fields even when all expected controls exist',()=>{
  const p=fixture();assert.throws(()=>reconcile(p,[...records(p),{tag:'osg:activity:new-id',text:'Keep this'}]),/unexpected/);
});
test('a stale settings cache cannot hide another pending operation from reads or persistence',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next}=addModule(p,'External module','');
  h.durable[PENDING]={schema:2,phase:'prepared',txId:'external',kind:'module',base:p,project:next};h.records=records(next);
  assert.equal(h.memory[PENDING],undefined);await assert.rejects(w.readProject(p),/paused/);await assert.rejects(w.persist(p),/paused/);
  const other=await adapter(h.reopen());await assert.rejects(other.persist(p),/paused/);assert.equal(h.saves,0);
});
test('a stale settings cache refreshes a newer committed course before any metadata write',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next}=addModule(p,'External module','');
  h.durable[KEY]=structuredClone(next);h.durable[PENDING]={schema:2,phase:'committed',txId:'external',kind:'module',project:next};h.records=records(next);
  assert.deepEqual(h.memory[KEY],p);await assert.rejects(w.persist(p),/document changed/);assert.equal(h.saves,0);assert.deepEqual(await w.readProject(p),next);
});
test('overlapping runtime operations cannot interleave a structural write with a read',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'New module','');let release;
  const refresh=h.settings.refreshAsync;h.settings.refreshAsync=cb=>{release=()=>refresh(cb);};
  const insertion=w.insert(p,next,'module',item);
  await assert.rejects(w.readProject(p),/already running/);await assert.rejects(w.persist(p),/already running/);
  h.settings.refreshAsync=refresh;release();assert.deepEqual(await insertion,next);
});
for(const persist of [false,true])test(`new-course final save failure (persisted=${persist}) cannot create duplicate course controls`,async()=>{
  const h=makeHost(null),w=await adapter(h),p=course('New course');h.saveFailures.set(2,{persist});
  await assert.rejects(w.insert(null,p,'course'));assert.equal(h.records.length,1);
  assert.deepEqual(await w.recover(),p);assert.deepEqual(await w.recover(),p);assert.equal(h.records.length,1);
});
for(const mode of ['missing','duplicate','orphan'])test(`pending recovery with ${mode} fields preserves all content and leaves metadata unadvanced`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'Pending module','');h.saveFailures.set(2,{persist:false});
  await assert.rejects(w.insert(p,next,'module',item));
  if(mode==='missing')h.records.pop();if(mode==='duplicate')h.records.push({...h.records[0]});if(mode==='orphan')h.records.push({tag:'osg:module:orphan',text:'Keep this text'});
  const snapshot=structuredClone(h.records),saves=h.saves;await assert.rejects(w.recover());
  assert.deepEqual(h.records,snapshot);assert.deepEqual(h.durable[KEY],p);assert.equal(h.saves,saves);assert.equal(journal(h).phase,'prepared');
});
test('legacy recovery preparation failure keeps the legacy journal and the complete Word fields',async()=>{
  const p=fixture(),h=makeHost(p),{next,item}=addModule(p,'Legacy module','');h.records=records(next);h.durable[PENDING]={project:next,kind:'module',id:item.id};h.reopen();h.saveFailures.set(1,{persist:false});
  const w=await adapter(h);await assert.rejects(w.recover());assert.equal(w.isBlocked(),true);assert.deepEqual(h.durable[KEY],p);assert.equal(h.durable[PENDING].schema,undefined);
  assert.deepEqual(await w.recover(),next);assert.deepEqual(h.records,records(next));assert.equal(h.paragraphs.length,0);
});
test('a foreign committed record arriving during the Word write is not overwritten',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),{next,item}=addModule(p,'Our module','');const foreign={...structuredClone(p),revision:p.revision+10};
  h.hooks.beforeSync=(n,host)=>{if(n===4){host.durable[KEY]=foreign;host.durable[PENDING]={schema:2,phase:'committed',txId:'foreign',kind:'metadata',project:foreign};}};
  await assert.rejects(w.insert(p,next,'module',item),/changed while/);assert.deepEqual(h.durable[KEY],foreign);assert.equal(h.saves,1);
  await assert.rejects(w.recover(),/unexpected/);assert.deepEqual(h.durable[KEY],foreign);
});
