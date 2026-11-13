import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {course,courseDetails,normalizeCourseDetails,editCourseDetails,addModule,addActivity} from '../src/model.mjs';
import {KEY,PENDING} from '../src/document-store.mjs';
import {fixture,makeHost,adapter,records} from './helpers/mock-office.mjs';
import {masterFixture,installParser,assets} from './helpers/checkpoint-fixture.mjs';
import {createCheckpoint,readCheckpoint} from '../src/checkpoint.mjs';
import {worksheetBlob} from '../src/worksheet.mjs';
import {agendaBlob} from '../src/agenda.mjs';
import {materialsBlob} from '../src/materials-output.mjs';
import {auditExport} from './helpers/export-contract.mjs';
const details={audience:'SETUP_PRIVATE_AUDIENCE',purpose:'SETUP_PRIVATE_PURPOSE',authors:'SETUP_PRIVATE_AUTHORS',attribution:'SETUP_PRIVATE_ATTRIBUTION'};
const blanks={audience:'',purpose:'',authors:'',attribution:''};
test('minimal setup retains title-only legacy shape and optional explicit blank details at revision zero with independent inputs',()=>{
  const old=course('  Title only  ');assert.deepEqual(Object.keys(old).sort(),['id','modules','revision','schema','title']);assert.equal(old.title,'Title only');assert.equal(old.revision,0);assert.deepEqual(old.modules,[]);assert.deepEqual(courseDetails(old),blanks);
  const empty=course('Blank details',blanks),input={...details},next=course('Guided title',input);assert.deepEqual(empty.details,blanks);assert.equal(next.revision,0);assert.deepEqual(next.modules,[]);input.audience='later input edit';assert.equal(next.details.audience,details.audience);next.details.purpose='later course edit';assert.equal(input.purpose,details.purpose);assert.notEqual(next.id,old.id);assert.deepEqual(old.modules,[]);
});
test('creation and editing share detail normalization text limits and malformed-input refusals',()=>{
  const input={...details,purpose:'  First\r\nSecond\rThird  '},normalized=normalizeCourseDetails(input),p=course('Guided',input),edited=editCourseDetails(course('Existing'),input);assert.equal(p.revision,0);assert.equal(edited.revision,1);assert.deepEqual(p.details,normalized);assert.deepEqual(edited.details,normalized);assert.equal(p.details.purpose,'First\nSecond\nThird');assert.equal(input.purpose,'  First\r\nSecond\rThird  ');
  for(const bad of [null,[],{}, {...details,authors:5},{...details,attribution:{}},{...details,purpose:'x'.repeat(50001)}]){assert.throws(()=>course('Guided',bad));assert.throws(()=>editCourseDetails(p,bad));}assert.throws(()=>course(' ',details));assert.equal(course('Limit',{...details,authors:'x'.repeat(50000)}).details.authors.length,50000);
});
test('guided creation carries all private details in exactly one prepared and committed course transaction with one title field',async()=>{
  const p=course('Guided course',details),h=makeHost(null),w=await adapter(h),saves=[];h.durable['other.extension']={keep:'UNCHANGED'};h.hooks.beforeSave=()=>saves.push(structuredClone({journal:h.memory[PENDING],project:h.memory[KEY]}));
  assert.deepEqual(await w.insert(null,p,'course'),p);assert.equal(h.saves,2);assert.equal(saves.length,2);assert.equal(saves[0].journal.kind,'course');assert.equal(saves[0].journal.phase,'prepared');assert.equal(saves[0].journal.base,null);assert.equal(saves[0].project,undefined);assert.deepEqual(saves[0].journal.project,p);assert.equal(saves[1].journal.phase,'committed');assert.equal(saves[1].journal.kind,'course');assert.equal(saves[1].journal.txId,saves[0].journal.txId);assert.deepEqual(saves[1].project,p);assert.deepEqual(h.durable[KEY],p);assert.equal(p.revision,0);assert.deepEqual(p.modules,[]);assert.deepEqual(h.records.map(({tag,text})=>({tag,text})),[{tag:`osg:course:${p.id}`,text:p.title}]);assert.equal(h.records[0].cannotDelete,true);assert.equal(h.paragraphs.length,1);assert.equal(h.pictures.length,0);assert.equal(h.cueWrites,0);assert.deepEqual(h.durable['other.extension'],{keep:'UNCHANGED'});
  const count=h.saves;assert.deepEqual(await w.recover(),p);assert.equal(h.saves,count);assert.deepEqual(await (await adapter(h.reopen())).loadProject(),p);assert.equal(h.saves,count);
});
test('failed or uncertain initial setup persistence preserves full intent but cannot replay a missing course title',async()=>{
  for(const persist of [false,true]){const p=course('Prepared setup',details),h=makeHost(null),w=await adapter(h);h.saveFailures.set(1,{persist});await assert.rejects(w.insert(null,p,'course'));assert.equal(h.records.length,0);assert.equal(h.paragraphs.length,0);assert.deepEqual(h.memory[PENDING].project,p);const durable=structuredClone(h.durable),saves=h.saves;await assert.rejects(w.recover());assert.equal(h.saves,saves);assert.deepEqual(h.durable,durable);assert.equal(h.records.length,0);const reopened=await adapter(h.reopen());if(persist){await assert.rejects(reopened.loadProject());await assert.rejects(reopened.recover());}else assert.equal(await reopened.loadProject(),null);assert.equal(h.records.length,0);}
});
test('uncertain final setup saves recover and reopen without losing details adding revisions or replaying title insertion',async()=>{
  for(const persist of [false,true]){const p=course('Final setup',details),h=makeHost(null),w=await adapter(h);h.saveFailures.set(2,{persist});await assert.rejects(w.insert(null,p,'course'));assert.equal(h.records.length,1);assert.equal(h.paragraphs.length,1);const reopened=await adapter(h.reopen());if(persist)assert.deepEqual(await reopened.loadProject(),p);else await assert.rejects(reopened.loadProject());assert.deepEqual(await reopened.recover(),p);const saves=h.saves;assert.deepEqual(await reopened.recover(),p);assert.deepEqual(await w.recover(),p);assert.equal(h.saves,saves);assert.deepEqual(h.durable[KEY],p);assert.deepEqual(h.durable[PENDING].project.details,details);assert.equal(h.durable[KEY].revision,0);assert.equal(h.records.length,1);assert.equal(h.paragraphs.length,1);}
});
test('uncertain setup recovery saves remain idempotent and preserve full metadata',async()=>{
  for(const persist of [false,true]){const p=course('Recover setup',details),h=makeHost(null),w=await adapter(h);h.saveFailures.set(2,{persist:false});h.saveFailures.set(3,{persist});await assert.rejects(w.insert(null,p,'course'));await assert.rejects(w.recover());assert.deepEqual(await w.recover(),p);const saves=h.saves;assert.deepEqual(await w.recover(),p);assert.equal(h.saves,saves);assert.equal(h.records.length,1);assert.equal(h.paragraphs.length,1);assert.deepEqual(h.durable[KEY].details,details);}
});
test('existing course refuses a second setup without changing content settings or unknown legacy details',async()=>{
  const p=fixture();p.details={...details,futureSetting:{keep:true}};const h=makeHost(p),w=await adapter(h),before=structuredClone(h.durable),body=structuredClone(h.records);await assert.rejects(w.insert(null,course('Other setup',details),'course'),e=>e.code==='OSG_STALE');assert.equal(h.saves,0);assert.deepEqual(h.records,body);assert.deepEqual(h.durable,before);assert.deepEqual(await w.loadProject(),p);assert.deepEqual(h.records,records(p));
});
test('initial details round-trip in full master backup and remain excluded from all public output packages after lessons are added',async()=>{
  installParser();let p=course('Guided privacy',details);p.details.futureSetting='SETUP_PRIVATE_UNKNOWN';const {master}=await masterFixture(p),saved=await createCheckpoint(master,p,JSZip,assets),restored=await readCheckpoint(saved.data,JSZip);assert.deepEqual(restored.master,master);assert.deepEqual(restored.project,p);assert.equal(restored.project.revision,0);assert.deepEqual(restored.project.modules,[]);
  p=addModule(p,'Learning module','Practice a skill.').next;p=addActivity(p,p.modules[0].id,{title:'Practice',icon:'Discussion',minutes:5,participant:'Write a response.',facilitator:'SETUP_PRIVATE_TEACHER',answers:'SETUP_PRIVATE_ANSWER',lines:1,materials:[{item:'Sheets',quantity:1,unit:'sets',use:'consumable',notes:'SETUP_PRIVATE_NOTE'}]}).next;const privateCanaries=[...Object.values(details),'SETUP_PRIVATE_UNKNOWN','SETUP_PRIVATE_TEACHER','SETUP_PRIVATE_ANSWER','SETUP_PRIVATE_NOTE'];for(const kind of ['workbook','agenda','materials']){const blob=kind==='workbook'?await worksheetBlob(p,JSZip):kind==='agenda'?await agendaBlob(p,'09:00',JSZip):await materialsBlob(p,undefined,JSZip);await auditExport(blob,{project:p,privateCanaries,kind});}assert.deepEqual(p.details,{...details,futureSetting:'SETUP_PRIVATE_UNKNOWN'});
});
