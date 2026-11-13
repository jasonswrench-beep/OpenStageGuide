import test from 'node:test';
import assert from 'node:assert/strict';
import {courseDetails,editCourseDetails,validate} from '../src/model.mjs';
import {KEY,PENDING} from '../src/document-store.mjs';
import {fixture,makeHost,adapter,journal} from './helpers/mock-office.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {masterFixture,installParser,assets} from './helpers/checkpoint-fixture.mjs';
import {createCheckpoint,readCheckpoint} from '../src/checkpoint.mjs';
import JSZip from 'jszip';
const fields={audience:'COURSE_PRIVATE_AUDIENCE',purpose:'COURSE_PRIVATE_PURPOSE',authors:'COURSE_PRIVATE_AUTHOR',attribution:'COURSE_PRIVATE_RIGHTS'};

test('older courses load blank details; editing only the four fields preserves IDs and unknown metadata',()=>{
  const p=fixture();p.unrelated={keep:['unknown']};
  assert.deepEqual(courseDetails(p),{audience:'',purpose:'',authors:'',attribution:''});
  assert.deepEqual(editCourseDetails(p,{audience:'',purpose:'',authors:'',attribution:''}),p);
  const next=editCourseDetails(p,{...fields,title:'Attempted title change',modules:[]});
  assert.deepEqual(next,{...p,details:fields,revision:p.revision+1});assert.equal(p.details,undefined);
  next.details.futureSetting={keep:true};
  const edited=editCourseDetails(next,{...fields,purpose:'  Revised\r\npurpose  '});
  assert.equal(edited.details.purpose,'Revised\npurpose');assert.deepEqual(edited.details.futureSetting,{keep:true});
  assert.deepEqual(editCourseDetails(edited,courseDetails(edited)),edited);
});

test('invalid detail shapes and overlong or nontext input are refused',()=>{
  const p=fixture();
  for(const details of [null,[],5,'text',{audience:5},{purpose:'x'.repeat(50001)}])assert.throws(()=>validate({...p,details}));
  for(const input of [null,{}, {...fields,authors:{}},{...fields,attribution:'x'.repeat(50001)}])assert.throws(()=>editCourseDetails(p,input));
});

test('course detail save uses settings only, preserves document text/images and round-trips on reopen',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),body=structuredClone(h.records);h.durable['unrelated.extension']={keep:true};
  const result=await w.updateCourseDetails(p,fields);
  assert.deepEqual(result,editCourseDetails(p,fields));assert.deepEqual(h.records,body);assert.equal(h.cueWrites,0);
  assert.equal(h.paragraphs.length,0);assert.equal(h.pictures.length,0);assert.deepEqual(h.durable['unrelated.extension'],{keep:true});
  assert.equal(journal(h).kind,'metadata');assert.equal(journal(h).phase,'committed');
  const saves=h.saves;await w.updateCourseDetails(result,fields);assert.equal(h.saves,saves);
  const reopened=await adapter(h.reopen());assert.deepEqual(await reopened.loadProject(),result);
});

test('invalid form input does not quarantine a healthy course',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h);
  await assert.rejects(w.updateCourseDetails(p,{...fields,purpose:{}}),e=>e.code==='OSG_INPUT');
  assert.equal(w.isBlocked(),false);assert.equal(h.saves,0);assert.equal(h.cueWrites,0);
});

for(const change of ['word-text','course-details'])test(`a stale course-details draft cannot overwrite newer ${change}`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h);
  if(change==='word-text')h.records.find(r=>r.tag.startsWith('osg:participant:')).text='NEWER DIRECT WORD EDIT';
  else h.durable[KEY]=editCourseDetails(p,{...fields,purpose:'Newer settings from another pane'});
  const before=structuredClone(h.durable),body=structuredClone(h.records);
  await assert.rejects(w.updateCourseDetails(p,fields),e=>e.code==='OSG_STALE');
  assert.deepEqual(h.durable,before);assert.deepEqual(h.records,body);assert.equal(h.saves,0);assert.equal(w.isBlocked(),false);
});

for(const save of [1,2])for(const persist of [false,true])test(`uncertain course-details save ${save}/persist=${persist} recovers metadata without body writes or extra revisions`,async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),body=structuredClone(h.records);h.saveFailures.set(save,{persist});
  await assert.rejects(w.updateCourseDetails(p,fields),/paused/);assert.equal(w.isBlocked(),true);
  const result=await w.recover();assert.deepEqual(result,editCourseDetails(p,fields));
  const saves=h.saves;assert.deepEqual(await w.recover(),result);assert.equal(h.saves,saves);
  assert.deepEqual(h.records,body);assert.equal(h.paragraphs.length,0);assert.equal(h.pictures.length,0);assert.equal(h.cueWrites,0);
  assert.deepEqual(await (await adapter(h.reopen())).loadProject(),result);
});

test('direct Word edits during a details save are retained, not replaced by the form snapshot',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h);
  h.hooks.beforeSave=n=>{if(n===1)h.records.find(r=>r.tag.startsWith('osg:participant:')).text='EDIT DURING DETAILS SAVE';};
  const result=await w.updateCourseDetails(p,fields);
  assert.equal(result.modules[0].activities[0].participant,'EDIT DURING DETAILS SAVE');assert.deepEqual(result.details,fields);
  assert.equal(h.cueWrites,0);
});

test('course details preserve activity timing evidence and never turn a mismatched cue into a healthy record',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h),id=p.modules[0].activities[0].id;
  const timed=await w.updateActivitySettings(p,id,{minutes:7.5,lines:4});
  const result=await w.updateCourseDetails(timed,fields);
  assert.equal(journal(h).schema,3);assert.deepEqual(journal(h).cueChecks,[id]);assert.deepEqual(result.details,fields);
  h.cues[id]='Learning activity · 99 min';const before=structuredClone(h.durable),saves=h.saves;
  await assert.rejects(w.updateCourseDetails(result,{...fields,purpose:'Should not save'}),/cue/);
  assert.deepEqual(h.durable,before);assert.equal(h.saves,saves);
});

test('incomplete managed content blocks course-detail recovery without advancing stored data',async()=>{
  const p=fixture(),h=makeHost(p),w=await adapter(h);h.saveFailures.set(2,{persist:false});
  await assert.rejects(w.updateCourseDetails(p,fields));h.records.pop();const before=structuredClone(h.durable),saves=h.saves;
  await assert.rejects(w.recover(),/deleted/);assert.deepEqual(h.durable,before);assert.equal(h.saves,saves);assert.equal(h.durable[PENDING].phase,'prepared');
});

test('backups retain course details and unknown fields while learner packages exclude all master-only course notes',async()=>{
  installParser();const p=editCourseDetails(fixture(),fields);p.details.futureSetting='KEEP_UNKNOWN_DETAIL';
  const {master}=await masterFixture(p),backup=await createCheckpoint(master,p,JSZip,assets),restored=await readCheckpoint(backup.data,JSZip);
  assert.deepEqual(restored.project,p);assert.deepEqual(restored.master,master);
  for(const part of Object.values(worksheetFiles(p)))for(const text of [...Object.values(fields),'KEEP_UNKNOWN_DETAIL'])assert.ok(!part.includes(text));
});
