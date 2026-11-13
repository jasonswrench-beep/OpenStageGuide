import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {course,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan,planSchedule} from '../src/delivery-plans.mjs';
import {planAgendaPreview,planAgendaFiles} from '../src/plan-agenda.mjs';
import {resolveSelectedPlanScope} from '../src/selected-plan-scope.mjs';
import {planWorkbookPreview,planWorkbookFiles,planWorkbookBlob,planWorkbookName,planMaterialsPreview,planMaterialsFiles,planMaterialsBlob,planMaterialsName} from '../src/selected-plan-outputs.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {materialsFiles} from '../src/materials-output.mjs';
import {agendaFiles} from '../src/agenda.mjs';
import {facilitatorFiles} from '../src/facilitator-output.mjs';
const stamp='2026-10-26T10:00:00.000Z';
const item=(name,quantity,use='consumable')=>({item:name,quantity,unit:use==='reusable'?'each':'boxes',use,notes:'PRIVATE_PREP_NOTES'});
function fixture(){let p=course('COURSE_TITLE_CANARY');for(let i=0;i<3;i++)p=addModule(p,`Module ${i}`,i===2?'EMPTY_GROUP_OBJECTIVE':`Objective ${i}`).next;
 const add=(mi,title,minutes,materials)=>{p=addActivity(p,p.modules[mi].id,{title,minutes,participant:`LEARN_${title}`,facilitator:`PRIVATE_TEACH_${title}`,answers:`PRIVATE_ANSWER_${title}`,lines:1,...(materials?{materials}:{})}).next;};
 add(0,'OMITTED_TASK',7,[item('Markers',100)]);add(0,'SELECTED_0',8,[item('Markers',0.75)]);
 add(1,'SELECTED_1A',11,[item('Markers',2),item('Projector',5,'reusable')]);add(1,'SELECTED_1B',13,[item('Markers',1.25),item('Projector',2,'reusable')]);
 const groups=[{moduleId:p.modules[2].id,activityIds:[]},{moduleId:p.modules[1].id,activityIds:[p.modules[1].activities[1].id,p.modules[1].activities[0].id]},{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}];
 const created=editDeliveryPlan(p,{kind:'create',name:'Reviewed plan',targetMinutes:50,groups});p=created.next;p.opaque={secret:'PRIVATE_UNKNOWN_CANARY'};return {p,id:created.id,groups};
}
const contents=files=>Object.values(files).join('\n');

test('one immutable selected identity and saved order align workbook, materials and agenda without empty-group content',()=>{
 const {p,id,groups}=fixture(),before=structuredClone(p),scope=resolveSelectedPlanScope(p,id),agenda=planAgendaPreview(p,id),work=planWorkbookPreview(p,id,stamp),materials=planMaterialsPreview(p,id,stamp);
 assert.equal(Object.isFrozen(scope),true);assert.equal(Object.isFrozen(scope.scope),true);assert.deepEqual(scope.scope,groups);assert.equal(scope.groups.length,2);assert.equal(scope.minutes,agenda.minutes);assert.equal(work.minutes,agenda.minutes);assert.equal(materials.minutes,agenda.minutes);
 assert.deepEqual(scope.scope.flatMap(g=>g.activityIds),agenda.rows.map(r=>r.activityId));assert.deepEqual(work.scope,agenda.scope);assert.deepEqual(materials.scope,agenda.scope);assert.deepEqual(work.modules.map(m=>m.title),['Module 1','Module 0']);assert.deepEqual(work.modules[0].activities.map(a=>a.title),['SELECTED_1B','SELECTED_1A']);assert.equal(work.modules[1].activities[0].title,'SELECTED_0');
 const workbook=contents(planWorkbookFiles(p,id,stamp)),list=contents(planMaterialsFiles(p,id,stamp));for(const forbidden of ['OMITTED_TASK','EMPTY_GROUP_OBJECTIVE','PRIVATE_TEACH','PRIVATE_ANSWER','PRIVATE_PREP_NOTES','PRIVATE_UNKNOWN_CANARY']){assert.ok(!workbook.includes(forbidden),`workbook leaked ${forbidden}`);assert.ok(!list.includes(forbidden),`materials leaked ${forbidden}`);}assert.ok(workbook.includes('Objective 1'));assert.ok(workbook.includes('Objective 0'));assert.ok(workbook.includes('selecting activities does not confirm objective coverage'));assert.ok(workbook.includes(id));assert.ok(list.includes(id));assert.ok(!list.includes('COURSE_TITLE_CANARY'));
 assert.equal(materials.configuredActivities,3);assert.equal(materials.totalActivities,3);assert.equal(materials.rows.find(r=>r.item==='Markers').quantity,4);assert.equal(materials.rows.find(r=>r.item==='Projector').quantity,5);assert.deepEqual(materials.rows.find(r=>r.item==='Markers').activities,['1.1 SELECTED_1B','1.2 SELECTED_1A','2.1 SELECTED_0']);assert.match(materials.assumption,/Missing entries do not prove/);assert.match(planMaterialsName('Reviewed plan'),/^OpenStage - Reviewed plan - Materials Checklist - /);assert.doesNotMatch(planMaterialsName('Reviewed plan'),/COURSE_TITLE_CANARY/);assert.match(planWorkbookName(p.title,'Reviewed plan'),/Participant Workbook/);assert.deepEqual(p,before);
});

test('selected materials ignore unsupported learner/objective fields while selected workbook refuses without source freezing',()=>{
 const {p,id}=fixture(),selected=p.modules[1].activities[0];p.modules[1].objective={opaque:'PRIVATE_OBJECTIVE'};selected.participant={toString:null,opaque:'PRIVATE_LEARNER'};selected.lines={opaque:'PRIVATE_LINES'};const before=structuredClone(p),participant=selected.participant,lines=selected.lines,objective=p.modules[1].objective;
 const view=planMaterialsPreview(p,id,stamp);assert.equal(view.rows.find(r=>r.item==='Markers').quantity,4);assert.equal(view.title,undefined);assert.throws(()=>planWorkbookPreview(p,id,stamp),/selected learning objective is unsupported/i);assert.equal(Object.isFrozen(p),false);assert.equal(Object.isFrozen(participant),false);assert.equal(Object.isFrozen(lines),false);assert.equal(Object.isFrozen(objective),false);assert.deepEqual(p,before);
 p.modules[1].objective='Objective 1';assert.throws(()=>planWorkbookPreview(p,id,stamp),/selected participant instructions is unsupported/i);assert.equal(planMaterialsPreview(p,id,stamp).totalActivities,3);p.modules[1].activities[0].participant='Learner instructions';assert.throws(()=>planWorkbookPreview(p,id,stamp),/writing-line setting/i);assert.equal(Object.isFrozen(lines),false);
});

test('omitted malformed learner/objective fields are not selected requirements; invalid selected materials retain strict global refusal',()=>{
 const {p,id}=fixture();p.modules[2].objective={opaque:'EMPTY_GROUP_OBJECTIVE'};p.modules[0].activities[0].participant={opaque:'OMITTED_LEARNER'};p.modules[0].activities[0].title={opaque:'OMITTED_TITLE'};const before=structuredClone(p);
 assert.doesNotThrow(()=>planWorkbookPreview(p,id,stamp));assert.doesNotThrow(()=>planMaterialsPreview(p,id,stamp));assert.deepEqual(p,before);
 p.modules[0].activities[0].materials=[item('  ',1)];const invalid=structuredClone(p);assert.throws(()=>planMaterialsPreview(p,id,stamp),/material name and unit/i);assert.deepEqual(p,invalid);
});

test('broken or foreign selected plans refuse without changing valid full-course output contracts',()=>{
 const {p,id}=fixture(),full=[worksheetFiles(p,stamp),materialsFiles(p,stamp),agendaFiles(p,'09:00',stamp),facilitatorFiles(p,{answers:false,participant:false},stamp)];
 for(const bad of ['missing','duplicate','empty','wrong-parent']){
  const q=structuredClone(p);if(bad==='duplicate')q.deliveryPlans.plans.push(structuredClone(q.deliveryPlans.plans[0]));if(bad==='empty')q.deliveryPlans.plans[0].groups=[{moduleId:q.modules[0].id,activityIds:[]}];if(bad==='wrong-parent')q.deliveryPlans.plans[0].groups=[{moduleId:q.modules[0].id,activityIds:[q.modules[1].activities[0].id]}];const before=structuredClone(q);
  assert.throws(()=>planWorkbookPreview(q,bad==='missing'?'missing':id,stamp),/unique|repair|No activities/i);assert.throws(()=>planMaterialsPreview(q,bad==='missing'?'missing':id,stamp),/unique|repair|No activities/i);assert.deepEqual(q,before);
  assert.deepEqual([worksheetFiles(q,stamp),materialsFiles(q,stamp),agendaFiles(q,'09:00',stamp),facilitatorFiles(q,{answers:false,participant:false},stamp)],full);
 }
 const foreign=structuredClone(p);foreign.deliveryPlans.projectId='other';assert.throws(()=>planWorkbookPreview(foreign,id,stamp),/foreign project/i);assert.doesNotThrow(()=>worksheetFiles(foreign,stamp));
});

test('selected output identity refuses transformed IDs and selected fields without coercion',()=>{
 for(const kind of ['project','module','activity']){
  const {p,id}=fixture();if(kind==='project'){p.id='bad\rproject';p.deliveryPlans.projectId=p.id;}else if(kind==='module'){const m=p.modules[1],old=m.id;m.id='bad\rmodule';p.deliveryPlans.plans[0].groups.find(g=>g.moduleId===old).moduleId=m.id;}else{const a=p.modules[1].activities[0],old=a.id;a.id='bad\ractivity';p.deliveryPlans.plans[0].groups.find(g=>g.moduleId===p.modules[1].id).activityIds=p.deliveryPlans.plans[0].groups.find(g=>g.moduleId===p.modules[1].id).activityIds.map(x=>x===old?a.id:x);}const before=structuredClone(p);
  assert.throws(()=>planWorkbookPreview(p,id,stamp),/cannot be shown exactly/i);assert.throws(()=>planMaterialsPreview(p,id,stamp),/cannot be shown exactly/i);assert.deepEqual(p,before);
 }
 const {p,id}=fixture();p.modules[1].activities[0].participant='  ';assert.throws(()=>planWorkbookPreview(p,id,stamp),/participant instructions text/i);assert.doesNotThrow(()=>planMaterialsPreview(p,id,stamp));assert.throws(()=>planWorkbookPreview(p,id,'bad stamp'),/generation stamp/i);
 for(const badTitle of [{opaque:'PRIVATE_OBJECT_TITLE'},17,false]){const q=fixture();q.p.title=badTitle;const before=structuredClone(q.p);assert.throws(()=>planWorkbookPreview(q.p,q.id,stamp),/unsupported|supported course title/i);assert.throws(()=>planMaterialsPreview(q.p,q.id,stamp),/unsupported|supported course title/i);assert.deepEqual(q.p,before);}
});

test('selected publication refuses required titles emptied by public text processing',()=>{
 const first=fixture();first.p.deliveryPlans.plans[0].name='\u0001';const before=structuredClone(first.p);
 assert.throws(()=>planWorkbookPreview(first.p,first.id,stamp),/plan name is empty/i);assert.throws(()=>planMaterialsPreview(first.p,first.id,stamp),/plan name is empty/i);assert.deepEqual(first.p,before);
 const second=fixture();second.p.title='\u0001';const saved=structuredClone(second.p);
 assert.throws(()=>planWorkbookPreview(second.p,second.id,stamp),/selected course title text/i);assert.doesNotThrow(()=>planMaterialsPreview(second.p,second.id,stamp));assert.deepEqual(second.p,saved);
 const third=fixture();third.p.title='合作 🙂';third.p.deliveryPlans.plans[0].name='Café \t plan';assert.equal(planWorkbookPreview(third.p,third.id,stamp).title,'合作 🙂');assert.match(planMaterialsPreview(third.p,third.id,stamp).planName,/Café/);
});

test('all generated selected-plan package parts stay within declared public projection',async()=>{
 const {p,id}=fixture();for(const [kind,make,preview] of [['workbook',planWorkbookBlob,planWorkbookPreview],['materials',planMaterialsBlob,planMaterialsPreview]]){
  const body=await make(p,id,stamp,JSZip),zip=await JSZip.loadAsync(await body.arrayBuffer()),before=structuredClone(p);for(const name of Object.keys(zip.files)){const part=await zip.file(name)?.async('string');if(!part)continue;for(const marker of ['OMITTED_TASK','EMPTY_GROUP_OBJECTIVE','PRIVATE_TEACH','PRIVATE_ANSWER','PRIVATE_PREP_NOTES','PRIVATE_UNKNOWN_CANARY'])assert.ok(!part.includes(marker),`${kind}/${name} leaked ${marker}`);if(kind==='materials')assert.ok(!part.includes('COURSE_TITLE_CANARY'),`${kind}/${name} leaked course title`);}assert.equal(preview(p,id,stamp).planId,id);assert.deepEqual(p,before);
 }
});
