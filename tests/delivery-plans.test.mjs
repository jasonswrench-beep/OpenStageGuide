import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {deliveryPlans,editDeliveryPlan,planSchedule,selectedObjectiveLinks,selectedPlanObjectiveReport} from '../src/delivery-plans.mjs';
import {planAgendaPreview,planAgendaFiles,planAgendaBlob,planAgendaName} from '../src/plan-agenda.mjs';
import {agendaFiles} from '../src/agenda.mjs';
import {worksheetBlob} from '../src/worksheet.mjs';
import {materialsBlob} from '../src/materials-output.mjs';
import {facilitatorBlob} from '../src/facilitator-output.mjs';
import {createCheckpoint,readCheckpoint} from '../src/checkpoint.mjs';
import {masterFixture,assets,installParser} from './helpers/checkpoint-fixture.mjs';
import {makeHost,adapter} from './helpers/mock-office.mjs';
import {frontHost,paragraph,control} from './helpers/frontmatter-fixture.mjs';
import {KEY,PENDING} from '../src/document-store.mjs';
import {emptyLibrary,saveLibraryCopy,serializeLibrary,libraryFields} from '../src/library.mjs';

function fixture(wrench=false){let p=wrench?wrenchCourse('Plan course'):course('Plan course');for(let i=0;i<2;i++){p=addModule(p,'Module '+i,'Objective '+i).next;for(let j=0;j<2;j++)p=addActivity(p,p.modules[i].id,{title:`Task ${i}.${j}`,minutes:10+i+j,participant:`Learner ${i}.${j}`,facilitator:`PRIVATE TEACH ${i}.${j}`,answers:`PRIVATE KEY ${i}.${j}`,lines:1}).next;}return p;}
const groups=p=>[{moduleId:p.modules[1].id,activityIds:[p.modules[1].activities[1].id,p.modules[1].activities[0].id]},{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}];
const create=(p,name='One hour',targetMinutes=60)=>({kind:'create',name,targetMinutes,groups:groups(p)});
const filesText=files=>Object.values(files).join('\n');
function packageCues(h){h.packageTransform=async z=>{let doc=await z.file('word/document.xml').async('string');for(const m of h.memory[KEY].modules)for(const a of m.activities){const marker=control(h.records.find(r=>r.tag===`osg:activity:${a.id}`));doc=doc.replace(marker,paragraph(`${a.icon} · ${a.minutes} min`)+marker);}z.file('word/document.xml',doc);};}

test('named plan keeps canonical source order and minutes while its agenda follows explicit cross-module group order and target',()=>{
 const p=fixture(),before=structuredClone(p),{next:q,id}=editDeliveryPlan(p,create(p));assert.deepEqual(p,before);assert.equal(q.revision,p.revision+1);assert.equal(q.modules[0].activities[0].minutes,10);assert.deepEqual(q.modules,p.modules);
 const a=planSchedule(q,id,'23:50');assert.deepEqual(a.rows.map(r=>r.activityId),groups(p).flatMap(g=>g.activityIds));assert.equal(a.targetMinutes,60);assert.equal(a.varianceMinutes,60-a.minutes);assert.match(a.end,/Day 2/);
 const doc=planAgendaFiles(q,id,'09:00','2026-10-26T10:00:00.000Z')['word/document.xml'];assert.ok(doc.includes(id));assert.ok(doc.includes(p.id));assert.ok(doc.includes('Generation revision'));assert.ok(doc.includes('2026-10-26T10:00:00.000Z'));assert.ok(doc.indexOf('Task 1.1')<doc.indexOf('Task 1.0'));assert.ok(doc.indexOf('Task 1.0')<doc.indexOf('Task 0.1'));assert.ok(!doc.includes('Task 0.0'));
 assert.match(planAgendaName('Plan course','One hour'),/^Plan course - One hour - Agenda - [a-f0-9-]+\.docx$/);assert.equal(deliveryPlans(q).plans[0].reasons.length,0);
});

test('valid unrelated plans remain usable when another entry is stale, malformed or has duplicate plan ID',()=>{
 let p=fixture();const first=editDeliveryPlan(p,create(p));p=first.next;const known=p.deliveryPlans.plans[0];p.deliveryPlans.plans.push({id:'broken',name:'Broken',groups:[{moduleId:p.modules[0].id,activityIds:['deleted-activity']}],future:{keep:'OPAQUE'}});p.deliveryPlans.plans.push({id:'same',name:'Ambiguous A',groups:[]},{id:'same',name:'Ambiguous B',groups:[]});
 const view=deliveryPlans(p);assert.equal(view.plans[0].reasons.length,0);assert.ok(view.plans[1].reasons.length);assert.match(view.plans[2].reasons.join(' '),/duplicated/);assert.throws(()=>planSchedule(p,'same'),/unique/);assert.equal(planSchedule(p,known.id).rows.length,3);
 const opaque=structuredClone(p.deliveryPlans.plans.slice(1));const result=editDeliveryPlan(p,{kind:'update',id:known.id,name:'Revised',targetMinutes:null,groups:groups(p)}).next;assert.deepEqual(result.deliveryPlans.plans.slice(1),opaque);assert.equal(result.deliveryPlans.plans[0].name,'Revised');
 p.deliveryPlans.version=99;assert.equal(deliveryPlans(p).supported,false);assert.doesNotThrow(()=>agendaFiles(p));
});

test('empty plan is editable but never downloads a misleading agenda; missing and wrong-parent selected IDs refuse only that plan',()=>{
 let p=fixture();let added=editDeliveryPlan(p,{kind:'create',name:'Empty',targetMinutes:null,groups:[]});p=added.next;assert.equal(deliveryPlans(p).plans[0].reasons.length,0);assert.throws(()=>planSchedule(p,added.id),/No activities selected/);
 const active=editDeliveryPlan(p,create(p,'Active'));p=active.next;assert.equal(planSchedule(p,active.id).rows.length,3);
 p.deliveryPlans.plans[0].groups=[{moduleId:p.modules[0].id,activityIds:[p.modules[1].activities[0].id]}];assert.match(deliveryPlans(p).plans[0].reasons.join(' '),/belongs to another module/);assert.throws(()=>planSchedule(p,added.id),/needs repair/);assert.equal(planSchedule(p,active.id).rows.length,3);assert.doesNotThrow(()=>agendaFiles(p));
});

test('target accepts real three-decimal values, refuses stored blank strings and invalid precision, and bounds complete selected references',()=>{
 const p=fixture();let q=editDeliveryPlan(p,create(p,'Precise',1.001)).next;assert.equal(q.deliveryPlans.plans[0].targetMinutes,1.001);assert.equal(planSchedule(q,q.deliveryPlans.plans[0].id).targetMinutes,1.001);q.deliveryPlans.plans[0].targetMinutes='';assert.match(deliveryPlans(q).plans[0].reasons.join(' '),/target/);assert.throws(()=>planSchedule(q,q.deliveryPlans.plans[0].id),/needs repair/);
 for(const bad of [0,10080.001,1.0001,NaN,'2'])assert.throws(()=>editDeliveryPlan(p,create(p,'Bad',bad)));
 const large=Array.from({length:2001},(_,i)=>`missing-${i}`);q=structuredClone(p);q.deliveryPlans={version:1,projectId:q.id,plans:[{id:'oversize',name:'Over',groups:[{moduleId:q.modules[0].id,activityIds:large}]}]};assert.match(deliveryPlans(q).plans[0].reasons.join(' '),/too many/);assert.doesNotThrow(()=>agendaFiles(q));
});

test('saved plan IDs requiring output transformation stay unchanged, visible and unusable while valid plans and full-course outputs work',()=>{
 for(const bad of ['plan\u0001A','plan\rA']){let p=fixture();p=editDeliveryPlan(p,create(p)).next;const good=p.deliveryPlans.plans[0].id;p.deliveryPlans.plans.push({...structuredClone(p.deliveryPlans.plans[0]),id:bad,name:'Unsupported ID'});const before=structuredClone(p),state=deliveryPlans(p);assert.equal(state.plans[0].reasons.length,0);assert.match(state.plans[1].reasons.join(' '),/unsupported characters/);assert.throws(()=>planSchedule(p,bad),/needs repair/);assert.throws(()=>planAgendaPreview(p,bad),/needs repair/);assert.throws(()=>planAgendaFiles(p,bad),/needs repair/);assert.equal(planAgendaPreview(p,good).planId,good);assert.doesNotThrow(()=>agendaFiles(p));assert.deepEqual(p,before);}
});

test('reserve and overrun use bounded decimal display while fractional seconds and day rollover remain canonical',()=>{
 for(const [minutes,target,status] of [[1.1,1.2,'0.1 minutes reserved'],[1.1,1,'0.1 minutes over'],[1.1,1.1,'0 minutes reserved'],[1,1.001,'0.001 minutes reserved']]){let p=fixture();p.modules[0].activities[0].minutes=minutes;p=editDeliveryPlan(p,{kind:'create',name:'Decimal plan',targetMinutes:target,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}]}).next;const id=p.deliveryPlans.plans[0].id,a=planSchedule(p,id,'23:59'),preview=planAgendaPreview(p,id,'23:59'),doc=planAgendaFiles(p,id,'23:59','2026-10-26T12:00:00.000Z')['word/document.xml'];assert.match(preview.targetStatus,new RegExp(status));assert.ok(doc.includes(status));assert.ok(!doc.includes('999999999'));assert.equal(a.rows[0].end,minutes===1.1?'Day 2 · 00:00:06':'Day 2 · 00:00');assert.equal(a.end,a.rows[0].end);assert.equal(preview.varianceMinutes,a.varianceMinutes);}
});

test('ambiguous saved group inheritance refuses update without losing unknown fields or mutating source',async()=>{
 let p=fixture();p=editDeliveryPlan(p,create(p)).next;const id=p.deliveryPlans.plans[0].id,m=p.modules[0].id;p.deliveryPlans.plans[0].groups=[{moduleId:m,activityIds:[p.modules[0].activities[0].id],future:'FIRST'},{moduleId:m,activityIds:[p.modules[0].activities[1].id],future:'SECOND'}];const before=structuredClone(p),change={kind:'update',id,name:'Edited',targetMinutes:null,groups:[{moduleId:m,activityIds:[p.modules[0].activities[0].id]}]};assert.throws(()=>editDeliveryPlan(p,change),/ambiguous or oversized/);assert.deepEqual(p,before);
 const h=makeHost(p),w=await adapter(h);await assert.rejects(w.updateDeliveryPlan(p,change),e=>e.code==='OSG_INPUT');assert.equal(h.saves,0);assert.equal(w.isBlocked(),false);assert.deepEqual(h.durable[KEY],before);
});

test('oversized saved group inheritance refuses before cloning old references or saving, preserving other plans',async()=>{
 let p=fixture();p=editDeliveryPlan(p,create(p)).next;const good=p.deliveryPlans.plans[0].id,m=p.modules[0].id;p.deliveryPlans.plans.push({id:'oversized',name:'Oversized legacy entry',groups:[{moduleId:m,activityIds:Array.from({length:2001},(_,i)=>`old-${i}`),future:'KEEP'}]});const before=structuredClone(p),change={kind:'update',id:'oversized',name:'Attempted repair',targetMinutes:null,groups:[{moduleId:m,activityIds:[p.modules[0].activities[0].id]}]};assert.throws(()=>editDeliveryPlan(p,change),/oversized/);assert.deepEqual(p,before);assert.equal(planSchedule(p,good).rows.length,3);const h=makeHost(p),w=await adapter(h);await assert.rejects(w.updateDeliveryPlan(p,change),e=>e.code==='OSG_INPUT');assert.equal(h.saves,0);assert.deepEqual(h.durable[KEY],before);
});

test('metadata-only save, no-op, freshness and uncertain repeated recovery preserve selected plan without Word writes',async()=>{
 for(const persist of [false,true]){const p=fixture(),h=makeHost(p),w=await adapter(h),change=create(p);h.saveFailures.set(2,{persist});await assert.rejects(w.updateDeliveryPlan(p,change));assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);h.saveFailures.clear();const reopened=await adapter(h.reopen()),q=await reopened.recover();assert.equal(q.deliveryPlans.plans.length,1);const count=h.saves;await reopened.recover();assert.equal(h.saves,count);assert.equal(h.durable[PENDING].phase,'committed');}
 const p=fixture(),h=makeHost(p),w=await adapter(h),q=await w.updateDeliveryPlan(p,create(p));assert.equal(h.saves,2);const same=await w.updateDeliveryPlan(q,{kind:'update',id:q.deliveryPlans.plans[0].id,name:'One hour',targetMinutes:60,groups:groups(p)});assert.deepEqual(same,q);assert.equal(h.saves,2);h.records.find(r=>r.tag===`osg:activity:${p.modules[0].activities[0].id}`).text='Edited in Word';await assert.rejects(w.updateDeliveryPlan(q,{kind:'update',id:q.deliveryPlans.plans[0].id,name:'Renamed',targetMinutes:60,groups:groups(p)}),e=>e.code==='OSG_STALE');assert.equal(h.saves,2);
});

test('Wrench profile5 keeps physical evidence and a second final-save Word edit refuses without an unbounded correction loop',async()=>{
 for(const governed of [false,true]){const p=fixture(governed),h=governed?frontHost(makeHost(p)):makeHost(p);if(governed)packageCues(h);const w=await adapter(h),q=await w.updateDeliveryPlan(p,create(p));assert.equal(h.durable[PENDING].schema,governed?5:2);assert.equal(q.deliveryPlans.plans.length,1);assert.equal(h.paragraphs.length,0);assert.equal(h.cueWrites,0);if(governed)assert.ok(h.fileReads>0);}
 const p=fixture(),h=makeHost(p),w=await adapter(h);h.hooks.beforeSave=n=>{if(n===2)h.records.find(r=>r.tag===`osg:activity:${p.modules[0].activities[0].id}`).text='First concurrent title';if(n===4)h.records.find(r=>r.tag===`osg:activity:${p.modules[1].activities[0].id}`).text='Second concurrent title';};await assert.rejects(w.updateDeliveryPlan(p,create(p)),e=>e.code==='OSG_STALE');assert.equal(h.saves,4);assert.equal(w.isBlocked(),false);assert.equal(h.durable[PENDING].phase,'committed');assert.equal(h.durable[KEY].deliveryPlans.plans.length,1);h.hooks.beforeSave=undefined;const fresh=await w.readProject();assert.equal(fresh.modules[0].activities[0].title,'First concurrent title');assert.equal(fresh.modules[1].activities[0].title,'Second concurrent title');assert.equal(h.saves,4);assert.equal(h.paragraphs.length,0);
});

test('allowlisted plan agenda excludes private metadata and unselected content; full master backup retains exact plans and library copy stays unlinked',async()=>{
 installParser();let p=fixture();p.deliveryPlans={version:1,projectId:p.id,plans:[],future:{secret:'PRIVATE ROOT'}};p=editDeliveryPlan(p,create(p,'Plan PRIVATE NAME')).next;p.deliveryPlans.plans[0].rationale='PRIVATE PLAN RATIONALE';p.deliveryPlans.plans[0].groups[0].future='PRIVATE GROUP FIELD';p.objectiveCoverage={version:99,private:'PRIVATE COVERAGE'};
 const id=p.deliveryPlans.plans[0].id,stamp='2026-10-26T10:00:00.000Z',doc=filesText(planAgendaFiles(p,id,'09:00',stamp));for(const secret of ['PRIVATE ROOT','PRIVATE PLAN RATIONALE','PRIVATE GROUP FIELD','PRIVATE COVERAGE','PRIVATE TEACH','PRIVATE KEY','Task 0.0'])assert.ok(!doc.includes(secret),secret);
 assert.ok(doc.includes(id));assert.ok(doc.includes(p.modules[1].activities[1].id));assert.ok(doc.includes(stamp));const blob=await planAgendaBlob(p,id,'09:00',stamp,JSZip),zip=await JSZip.loadAsync(await blob.arrayBuffer());assert.equal(Object.values(zip.files).filter(f=>!f.dir).length,6);assert.ok((await zip.file('word/document.xml').async('string')).includes(stamp));
 for(const other of [agendaFiles(p),await (async()=>{const z=await JSZip.loadAsync(await (await worksheetBlob(p,JSZip)).arrayBuffer());return {'word/document.xml':await z.file('word/document.xml').async('string')};})(),await (async()=>{const z=await JSZip.loadAsync(await (await materialsBlob(p,undefined,JSZip)).arrayBuffer());return {'word/document.xml':await z.file('word/document.xml').async('string')};})()])assert.ok(!filesText(other).includes(id));
 const facilitator=await JSZip.loadAsync(await (await facilitatorBlob(p,{answers:true,participant:true},stamp,JSZip)).arrayBuffer());assert.ok(!(await facilitator.file('word/document.xml').async('string')).includes(id));
 const master=(await masterFixture(p)).master,backup=await createCheckpoint(master,p,JSZip,assets),restored=await readCheckpoint(backup.data,JSZip);assert.deepEqual(restored.master,master);assert.deepEqual(restored.project.deliveryPlans,p.deliveryPlans);
 const a=p.modules[0].activities[0],library=saveLibraryCopy(emptyLibrary(),'Copy',[],libraryFields(Object.fromEntries(['title','icon','minutes','facilitator','participant','answers','lines'].map(k=>[k,a[k]]))));assert.ok(!serializeLibrary(library).includes(id));assert.ok(!serializeLibrary(library).includes('PRIVATE PLAN RATIONALE'));
});

test('selected-plan objective comparison uses every current course objective, exact selected cross-module links and honest role gaps',async()=>{
 let p=fixture();const [m0,m1]=p.modules,[a00]=m0.activities,[a10,a11]=m1.activities;
 p=editDeliveryPlan(p,{kind:'create',name:'Short',targetMinutes:30,groups:[{moduleId:m1.id,activityIds:[a11.id,a10.id]}]}).next;
 p.objectiveCoverage={version:1,projectId:p.id,privateFuture:'PRIVATE UNKNOWN COVERAGE FIELD',links:[
  {activityId:a00.id,objectiveModuleId:m0.id,roles:['practice','assessment'],reviewedObjective:m0.objective,rationale:'PRIVATE UNSELECTED RATIONALE'},
  {activityId:a11.id,objectiveModuleId:m0.id,roles:['practice'],reviewedObjective:m0.objective,rationale:'PRIVATE CROSS-MODULE RATIONALE'},
  {activityId:a10.id,objectiveModuleId:m0.id,roles:['practice'],reviewedObjective:'STALE PRIVATE OBJECTIVE',rationale:'PRIVATE STALE RATIONALE'},
  {activityId:a11.id,objectiveModuleId:m1.id,roles:['assessment'],reviewedObjective:m1.objective,rationale:'PRIVATE ASSESSMENT RATIONALE'}
 ],exceptions:[{objectiveModuleId:m0.id,role:'assessment',reviewedObjective:m0.objective,note:'PRIVATE COURSE GAP NOTE'}]};
 const before=structuredClone(p),id=p.deliveryPlans.plans[0].id,stamp='2026-10-26T12:00:00.000Z',fullBefore=agendaFiles(p,'09:00',stamp),schedule=planSchedule(p,id),privateReport=selectedPlanObjectiveReport(p,id),preview=planAgendaPreview(p,id),publicJson=JSON.stringify(preview),doc=filesText(planAgendaFiles(p,id,'09:00',stamp));
 assert.equal(schedule.coverage.state,'ready');assert.equal(schedule.coverage.rows,undefined);assert.equal(privateReport.rows.length,2);assert.deepEqual(privateReport.summary,{courseObjectives:2,nonblankObjectives:2,blankObjectives:0,selectedActivities:2,lostPractice:0,lostAssessment:1,preexistingPracticeGaps:1,preexistingAssessmentGaps:0,staleSelectedLinks:1});
 assert.equal(privateReport.rows[0].roles.practice.courseCurrent,2);assert.equal(privateReport.rows[0].roles.practice.selectedCurrent,1);assert.equal(privateReport.rows[0].roles.assessment.lostBySelection,true);assert.equal(privateReport.rows[0].roles.practice.selectedStale,1);assert.equal(privateReport.rows[0].roles.assessment.noteConflict,true);assert.equal(privateReport.rows[1].roles.practice.preexistingGap,true);
 assert.match(publicJson,/Objectives losing all current declared coverage in this selection: 0 for practice, 1 for assessment/);assert.ok(doc.includes('0 for practice, 1 for assessment'));
 for(const secret of ['PRIVATE','STALE',m0.id,m0.title,m0.objective,a00.id]){assert.ok(!publicJson.includes(secret),`preview leaked ${secret}`);assert.ok(!doc.includes(secret),`agenda leaked ${secret}`);}
 const zip=await JSZip.loadAsync(await (await planAgendaBlob(p,id,'09:00',stamp,JSZip)).arrayBuffer());for(const name of Object.keys(zip.files)){const body=await zip.file(name)?.async('string');if(body)for(const secret of ['PRIVATE','STALE',m0.id,m0.title,m0.objective,a00.id])assert.ok(!body.includes(secret),`${name} leaked ${secret}`);}
 const reversed=selectedObjectiveLinks(p,[{moduleId:m1.id,activityIds:[a10.id,a11.id]}]);assert.deepEqual(reversed.summary,privateReport.summary);
 assert.ok(JSON.stringify(privateReport).includes('PRIVATE COURSE GAP NOTE')===false);assert.deepEqual(agendaFiles(p,'09:00',stamp),fullBefore);assert.deepEqual(p,before);
});

test('objective comparison distinguishes no objectives, blank text, absent records and unsupported saved annotations',()=>{
 const empty=course('Empty'),zero=selectedObjectiveLinks(empty,[]);assert.equal(zero.state,'no-objectives');assert.match(zero.notice,/no module objectives/);assert.equal(zero.summary.courseObjectives,0);
 const p=fixture(),scope=[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[0].id]}],absent=selectedObjectiveLinks(p,scope);assert.equal(absent.state,'no-annotations');assert.equal(absent.summary.preexistingPracticeGaps,2);assert.match(absent.notice,/No saved objective annotations/);
 const blank=structuredClone(p);for(const m of blank.modules)m.objective='';const noText=selectedObjectiveLinks(blank,scope);assert.equal(noText.state,'no-objective-text');assert.equal(noText.summary.blankObjectives,2);assert.match(noText.notice,/all 2 course module objectives are blank/);
 const mixed=structuredClone(p);mixed.modules[0].objective='';mixed.objectiveCoverage={version:1,projectId:mixed.id,links:[],exceptions:[]};const partial=selectedObjectiveLinks(mixed,scope);assert.equal(partial.state,'ready');assert.equal(partial.summary.blankObjectives,1);
 for(const data of [{version:99,projectId:p.id,links:[],exceptions:[]},{version:1,projectId:'foreign',links:[],exceptions:[]}]){const q=structuredClone(p);q.objectiveCoverage=data;const before=structuredClone(q),view=selectedObjectiveLinks(q,scope);assert.equal(view.state,'unsupported');assert.equal(view.summary,null);assert.equal(view.rows.length,0);assert.match(view.notice,/unavailable/);assert.deepEqual(q,before);}
});

test('nontext objective in an omitted module makes only derived coverage unavailable and preserves the valid agenda',()=>{
 let p=fixture();p=editDeliveryPlan(p,{kind:'create',name:'Second module only',targetMinutes:null,groups:[{moduleId:p.modules[1].id,activityIds:[p.modules[1].activities[0].id]}]}).next;
 const omitted=p.modules[0],id=p.deliveryPlans.plans[0].id;omitted.objective={future:'PRIVATE NONSTRING OBJECTIVE'};const before=structuredClone(p),preview=planAgendaPreview(p,id),doc=filesText(planAgendaFiles(p,id,'09:00','2026-10-26T12:00:00.000Z'));
 assert.equal(preview.coverage.state,'unsupported-objective-field');assert.equal(preview.coverage.summary,null);assert.match(preview.coverageNotice,/unsupported text field/);assert.equal(preview.rows.length,1);assert.ok(!JSON.stringify(preview).includes('PRIVATE NONSTRING OBJECTIVE'));assert.ok(!doc.includes('PRIVATE NONSTRING OBJECTIVE'));assert.ok(!doc.includes(omitted.id));assert.deepEqual(p,before);
});
