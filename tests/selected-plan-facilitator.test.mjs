import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {facilitatorFiles,facilitatorParagraphs,facilitatorOptions,FACILITATOR_PRIVATE} from '../src/facilitator-output.mjs';
import {planFacilitatorPreview,planFacilitatorFiles,planFacilitatorBlob,planFacilitatorName,SELECTED_FACILITATOR_LIMITS} from '../src/selected-plan-facilitator.mjs';
import {auditFacilitator} from './helpers/facilitator-contract.mjs';
const stamp='2026-10-26T16:00:00.000Z';
const supply=(name,notes)=>({item:name,quantity:1.125,unit:'sets',use:'consumable',notes});
function fixture(wrench=false){
 let p=wrench?wrenchCourse('Selected private course',{audience:'OMITTED_AUDIENCE',purpose:'OMITTED_PURPOSE',authors:'OMITTED_AUTHORS',attribution:'OMITTED_ATTRIBUTION'},['background']):course('Selected private course');
 if(wrench){p.frontMatter.sections[0].body='OMITTED_OPENING';p.frontMatter.sections[1].body='OMITTED_BACKGROUND';}
 for(const title of ['One','Two','Empty'])p=addModule(p,title,title==='Empty'?'OMITTED_EMPTY_OBJECTIVE':`Objective ${title}`).next;
 const add=(mi,title)=>{p=addActivity(p,p.modules[mi].id,{title,icon:'Discussion',minutes:7.5,participant:`LEARN_${title}`,facilitator:`PRIVATE_TEACH_${title}`,answers:`PRIVATE_ANSWER_${title}`,lines:2,materials:[supply(`Material_${title}`,`PRIVATE_PREP_${title}`)]}).next;};
 add(0,'OMITTED');add(0,'A');add(1,'B');add(1,'C');
 const groups=[{moduleId:p.modules[2].id,activityIds:[]},{moduleId:p.modules[1].id,activityIds:[p.modules[1].activities[1].id,p.modules[1].activities[0].id]},{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}];
 p=editDeliveryPlan(p,{kind:'create',name:'Private session',targetMinutes:null,groups}).next;p.unknown='OMITTED_OPAQUE';p.modules[0].activities[1].librarySource={secret:'OMITTED_LIBRARY'};return {p,id:p.deliveryPlans.plans[0].id,groups};
}
const allParts=files=>Object.values(files).join('\n');

test('Minimal and Wrench selected private copies keep authored order, exact identity and only selected private fields for all opt-ins',async()=>{
 for(const wrench of [false,true])for(const answers of [false,true])for(const participant of [false,true]){
  const {p,id,groups}=fixture(wrench),before=structuredClone(p),options={answers,participant},v=planFacilitatorPreview(p,id,options,stamp),body=allParts(planFacilitatorFiles(p,id,options,stamp));
  assert.deepEqual(v.selectedScope,groups);assert.equal(v.minutes,22.5);assert.deepEqual(v.modules.map(m=>m.title),['Two','One']);assert.deepEqual(v.modules[0].activities.map(a=>a.title),['C','B']);assert.deepEqual(v.details,[]);assert.deepEqual(v.frontMatter,[]);assert.equal(v.options.answers,answers);assert.equal(v.options.participant,participant);
  assert.match(body,/Private session/);assert.match(body,/PRIVATE_TEACH_C/);assert.match(body,/PRIVATE_PREP_C/);assert.match(body,/Course details and front matter are omitted/);assert.match(body,/selecting activities does not confirm objective coverage/i);
  for(const secret of ['OMITTED_AUDIENCE','OMITTED_PURPOSE','OMITTED_AUTHORS','OMITTED_ATTRIBUTION','OMITTED_OPENING','OMITTED_BACKGROUND','OMITTED_EMPTY_OBJECTIVE','PRIVATE_TEACH_OMITTED','PRIVATE_PREP_OMITTED','OMITTED_OPAQUE','OMITTED_LIBRARY'])assert.ok(!body.includes(secret),`${wrench}/${answers}/${participant} leaked ${secret}`);
  assert.equal(body.includes('PRIVATE_ANSWER_C'),answers);assert.equal(body.includes('LEARN_C'),participant);assert.equal(body.includes('Participant writing-line setting: 2'),participant);
  assert.equal(facilitatorParagraphs(v).some(([,text])=>text===FACILITATOR_PRIVATE),true);assert.equal(facilitatorParagraphs(v).some(([,text])=>text===SELECTED_FACILITATOR_LIMITS),true);
  await auditFacilitator(await planFacilitatorBlob(p,id,options,stamp,JSZip),{expectedRows:facilitatorParagraphs(v),stamp,forbidden:['OMITTED_OPAQUE','OMITTED_OPENING','PRIVATE_TEACH_OMITTED',...answers?[]:['PRIVATE_ANSWER_C'],...participant?[]:['LEARN_C']]});assert.deepEqual(p,before);
 }
});

test('selected private publication refuses malformed selected fields but ignores omitted optional text and off fields',()=>{
 const {p,id}=fixture(true),activity=p.modules[1].activities[0];p.modules[0].activities[0].facilitator={private:'OMITTED_BAD_TEACH'};p.modules[0].activities[0].answers={private:'OMITTED_BAD_ANSWER'};p.modules[2].objective={private:'OMITTED_BAD_OBJECTIVE'};p.frontMatter.sections[0].body='\ufffe';
 const before=structuredClone(p);assert.doesNotThrow(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},stamp));assert.deepEqual(p,before);
 activity.answers={private:'SELECTED_BAD_ANSWER'};activity.participant={private:'SELECTED_BAD_LEARNER'};activity.lines={private:'SELECTED_BAD_LINES'};
 assert.doesNotThrow(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},stamp));assert.throws(()=>planFacilitatorPreview(p,id,{answers:true,participant:false},stamp),/selected answers\/debrief/i);assert.throws(()=>planFacilitatorPreview(p,id,{answers:false,participant:true},stamp),/writing-line settings/i);
 activity.lines=2;assert.throws(()=>planFacilitatorPreview(p,id,{answers:false,participant:true},stamp),/selected participant instructions/i);
 activity.participant='Learner';activity.answers='Answer';activity.facilitator={private:'SELECTED_BAD_TEACH'};assert.throws(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},stamp),/selected facilitator instructions/i);
});

test('broken, ambiguous, foreign and transformed scope identities refuse without changing full-course copy',()=>{
 const {p,id}=fixture(),full=facilitatorFiles(p,{answers:false,participant:false},stamp);
 for(const kind of ['missing','duplicate','foreign','wrong-parent','empty','transformed']){
  const q=structuredClone(p);if(kind==='duplicate')q.deliveryPlans.plans.push(structuredClone(q.deliveryPlans.plans[0]));
  if(kind==='foreign')q.deliveryPlans.projectId='foreign';
  if(kind==='wrong-parent')q.deliveryPlans.plans[0].groups[1].activityIds=[q.modules[0].activities[1].id];
  if(kind==='empty')q.deliveryPlans.plans[0].groups=[{moduleId:q.modules[2].id,activityIds:[]}];
  if(kind==='transformed'){const old=q.modules[1].id;q.modules[1].id='bad\rmodule';q.deliveryPlans.plans[0].groups.find(g=>g.moduleId===old).moduleId=q.modules[1].id;}
  const before=structuredClone(q);assert.throws(()=>planFacilitatorPreview(q,kind==='missing'?'missing':id,{answers:false,participant:false},stamp));assert.deepEqual(q,before);
  if(kind!=='transformed')assert.deepEqual(facilitatorFiles(q,{answers:false,participant:false},stamp),full);
 }
});

test('required public title/name and selected IDs cannot silently collapse while legitimate Unicode survives',()=>{
 for(const field of ['title','planName','projectId']){
  const {p,id}=fixture();if(field==='title')p.title='\u0001';if(field==='planName')p.deliveryPlans.plans[0].name='\u0001';if(field==='projectId'){p.id='bad\rproject';p.deliveryPlans.projectId=p.id;}
  const before=structuredClone(p);assert.throws(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},stamp));assert.deepEqual(p,before);
 }
 const {p,id}=fixture();p.title='合作 🙂';p.deliveryPlans.plans[0].name='Café plan';assert.equal(planFacilitatorPreview(p,id,{answers:false,participant:false},stamp).title,'合作 🙂');assert.match(planFacilitatorName('Café plan'),/^OpenStage Private Facilitator Guide - Café plan - [a-f0-9-]{36}\.docx$/);assert.notEqual(planFacilitatorName('Café plan'),planFacilitatorName('Café plan'));
 assert.throws(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},'bad'),/stamp/i);
});

test('private inclusion options require own data booleans and never invoke unsupported getters',()=>{
 let calls=0;const inherited=Object.assign(Object.create({answers:true,participant:true}),{x:0,y:0}),accessor={participant:false};Object.defineProperty(accessor,'answers',{enumerable:true,get(){calls++;return true;}});
 for(const bad of [inherited,accessor,{answers:false,participant:false,[Symbol('hidden')]:true}])assert.throws(()=>facilitatorOptions(bad),/supported facilitator-copy inclusion options/i);
 assert.equal(calls,0);assert.deepEqual(facilitatorOptions({answers:true,participant:false}),{answers:true,participant:false});
});

test('selected private fields refuse malformed cue, objective and preparation text while omitted records stay excluded',()=>{
 const {p,id}=fixture(),selected=p.modules[1].activities[0],omitted=p.modules[0].activities[0];
 omitted.icon='OMITTED_BAD_CUE';omitted.minutes='OMITTED_BAD_DURATION';omitted.materials[0].notes='OMITTED_BAD_PREP\ufffe';
 assert.doesNotThrow(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},stamp));
 for(const [field,bad,pattern] of [['icon','bad cue',/supported selected instructional cue/i],['minutes','ten',/selected activity duration/i]]){
  const old=selected[field];selected[field]=bad;const before=structuredClone(p);assert.throws(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},stamp),pattern);assert.deepEqual(p,before);selected[field]=old;
 }
 p.modules[1].objective={secret:'BAD_SELECTED_OBJECTIVE'};assert.throws(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},stamp),/selected learning objective/i);p.modules[1].objective='Objective Two';
 selected.materials[0].notes='SELECTED_BAD_PREP\ufffe';assert.throws(()=>planFacilitatorPreview(p,id,{answers:false,participant:false},stamp),/selected preparation notes contains an unsupported character/i);
});
