import test from 'node:test';
import assert from 'node:assert/strict';
import {editActivitySettings,addActivity} from '../src/model.mjs';
import {buildAgenda} from '../src/planning.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {fixture,makeHost,adapter,journal} from './helpers/mock-office.mjs';
import {KEY,PENDING} from '../src/document-store.mjs';
const settings={minutes:7.5,lines:8};
const first=p=>p.modules[0].activities[0];

test('settings change duration and writing space only, preserving IDs, teaching text and other activities',async()=>{
  let p=fixture();p=addActivity(p,p.modules[0].id,{title:first(p).title,participant:'Other learner',minutes:2,lines:1}).next;
  const h=makeHost(p),w=await adapter(h),id=first(p).id;
  h.records.find(r=>r.tag===`osg:participant:${id}`).text='A direct Word edit';
  const loaded=await w.persist(await w.readProject());
  const result=await w.updateActivitySettings(loaded,id,settings);
  assert.equal(result.revision,loaded.revision+1);assert.equal(first(result).id,id);
  assert.deepEqual({...first(result),...{minutes:first(loaded).minutes,lines:first(loaded).lines}},first(loaded));
  assert.deepEqual(result.modules[0].activities[1],loaded.modules[0].activities[1]);
  assert.equal(h.cues[id],'Learning activity · 7.5 min');assert.equal(h.cueWrites,1);
  assert.equal(h.paragraphs.length,0);assert.equal(h.pictures.length,0);
  assert.equal(buildAgenda(result,'09:00').rows[0].end,'09:07:30');
  const xml=worksheetFiles(result)['word/document.xml'];assert.ok(xml.includes('A direct Word edit'));
  for(const text of Object.values(worksheetFiles(result))){assert.ok(!text.includes('PRIVATE DIRECTIONS'));assert.ok(!text.includes('SECRET ANSWER'));}
  const saves=h.saves;
  assert.deepEqual(await w.recover(),result);assert.equal(h.cueWrites,1);assert.equal(h.saves,saves);
  assert.deepEqual(await (await adapter(h.reopen())).loadProject(),result);
});

test('line-only settings commit without a Word text write; identical settings do not save a revision',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;
  const result=await w.updateActivitySettings(p,id,{minutes:10,lines:0});assert.equal(first(result).lines,0);assert.equal(h.cueWrites,0);
  const saves=h.saves;assert.deepEqual(await w.updateActivitySettings(result,id,{minutes:10,lines:0}),result);assert.equal(h.saves,saves);
});

test('direct title and learner edits after a settings commit reconcile normally; a changed timing cue pauses reads',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;
  await w.updateActivitySettings(p,id,settings);
  h.records.find(r=>r.tag===`osg:activity:${id}`).text='A revised title in Word';
  h.records.find(r=>r.tag===`osg:participant:${id}`).text='A revised learner task in Word';
  const read=await w.readProject();assert.equal(first(read).title,'A revised title in Word');assert.equal(first(read).participant,'A revised learner task in Word');
  h.cues[id]='Custom timing cue';await assert.rejects(w.readProject(),/cue/);
});

test('backup cannot erase timing evidence by persisting a simultaneous direct learner edit',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;
  await w.updateActivitySettings(p,id,settings);
  h.cues[id]='Learning activity · 99 min';h.records.find(r=>r.tag===`osg:participant:${id}`).text='A newer direct edit';
  const durable=structuredClone(h.durable),body=structuredClone(h.records),saves=h.saves;let captured=false;
  // Run through the exported capture path, supplying a harmless Office stub.
  h.Office.context.document.getFileAsync=()=>{captured=true;throw new Error('Should not capture');};
  await assert.rejects(w.captureMaster(),/cue/);
  assert.equal(captured,false);assert.equal(h.saves,saves);assert.deepEqual(h.durable,durable);assert.deepEqual(h.records,body);
  assert.equal(h.durable[PENDING].schema,3);
});

test('timing evidence survives valid direct-edit persistence and later module insertion',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;
  await w.updateActivitySettings(p,id,settings);
  h.records.find(r=>r.tag===`osg:participant:${id}`).text='Reconciled learner text';
  const result=await w.persist(await w.readProject());assert.equal(journal(h).schema,3);assert.deepEqual(journal(h).cueChecks,[id]);
  const {addModule}=await import('../src/model.mjs'),{next,item}=addModule(result,'Later module','');
  await w.insert(result,next,'module',item);assert.equal(journal(h).schema,3);assert.deepEqual(journal(h).cueChecks,[id]);
  h.cues[id]='Learning activity · 99 min';await assert.rejects(w.readProject(),/cue/);
});

test('invalid settings cannot silently coerce missing numbers or fractional lines',()=>{
  const p=fixture(),id=first(p).id;
  for(const bad of [{minutes:'',lines:2},{minutes:NaN,lines:2},{minutes:-1,lines:2},{minutes:1441,lines:2},{minutes:2,lines:2.5},{minutes:2,lines:''},{minutes:2,lines:21}])assert.throws(()=>editActivitySettings(p,id,bad));
  assert.throws(()=>editActivitySettings(p,'missing',settings));
});

for(const defect of ['cue-text','cue-missing','cue-nested','cue-table','cue-child','title-paragraph'])test(`ambiguous settings boundary ${defect} never writes content or a journal`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;
  if(defect==='cue-text')h.cues[id]='Manual timing note';
  if(defect==='cue-missing')delete h.cues[id];
  if(defect==='cue-nested')h.cueNested=true;
  if(defect==='cue-table')h.cueTable=true;
  if(defect==='cue-child')h.cueChildren=true;
  if(defect==='title-paragraph')h.paragraphText={[`osg:activity:${id}`]:'Manual prefix '+first(p).title};
  const before=structuredClone(h.durable),body=structuredClone(h.records),cues=structuredClone(h.cues);
  await assert.rejects(w.updateActivitySettings(p,id,settings));
  assert.deepEqual(h.durable,before);assert.deepEqual(h.records,body);assert.deepEqual(h.cues,cues);assert.equal(h.cueWrites,0);assert.equal(h.saves,0);
});

test('legacy inline icon in the timing paragraph survives numeric duration replacement',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;h.cuePictures=true;
  await w.updateActivitySettings(p,id,settings);assert.equal(h.cueWrites,1);assert.equal(h.cuePictures,true);assert.equal(h.pictures.length,0);
});

test('direct Word text changed after loading settings rejects stale form without overwriting it',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;
  h.records.find(r=>r.tag===`osg:participant:${id}`).text='Newer Word instructions';
  await assert.rejects(w.updateActivitySettings(p,id,settings),e=>e.code==='OSG_STALE');
  assert.equal(w.isBlocked(),false);assert.equal(h.saves,0);assert.equal(h.cueWrites,0);
});

test('text edited during preparation survives a stopped settings update and cannot recover incomplete duration',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;
  h.hooks.beforeSave=n=>{if(n===1)h.records.find(r=>r.tag===`osg:participant:${id}`).text='Edit during prepare';};
  await assert.rejects(w.updateActivitySettings(p,id,settings),/changed during preparation/);
  await assert.rejects(w.recover(),/cue/);assert.equal(h.cueWrites,0);
  assert.equal(h.records.find(r=>r.tag===`osg:participant:${id}`).text,'Edit during prepare');
});

for(const save of [1,2])for(const persist of [false,true])test(`uncertain settings save ${save}, persisted=${persist} never replays the duration edit`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=first(p).id;h.saveFailures.set(save,{persist});
  await assert.rejects(w.updateActivitySettings(p,id,settings),/paused/);
  const writes=h.cueWrites,body=structuredClone(h.records),cues=structuredClone(h.cues);
  if(save===1)await assert.rejects(w.recover(),/cue/);
  else {const recovered=await w.recover();assert.equal(first(recovered).minutes,7.5);assert.equal(first(recovered).lines,8);assert.equal(journal(h).phase,'committed');}
  assert.equal(h.cueWrites,writes);assert.deepEqual(h.records,body);assert.deepEqual(h.cues,cues);
});

test('every Word synchronization failure before and after a duration replacement preserves all other text and never replays',async()=>{
  const p=fixture(),id=first(p).id,probe=makeHost(p);await (await adapter(probe)).updateActivitySettings(p,id,settings);
  for(let stage=1;stage<=probe.syncs;stage++)for(const apply of [false,true]){
    const h=makeHost(p),w=await adapter(h);h.syncFailures.set(stage,{apply});
    await assert.rejects(w.updateActivitySettings(p,id,settings),undefined,`sync ${stage}/${apply}`);
    const writes=h.cueWrites,cues=structuredClone(h.cues),body=structuredClone(h.records);
    try{await w.recover();assert.equal(h.durable[KEY].modules[0].activities[0].minutes,7.5);}catch{assert.equal(h.durable[KEY].modules[0].activities[0].minutes,10);}
    assert.equal(h.cueWrites,writes);assert.deepEqual(h.cues,cues);assert.deepEqual(h.records,body);
  }
});
