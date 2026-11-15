// The generated pages name the forms the guide's blocks cite: a "Forms used" column on Training at a Glance, and "Participant handouts" lines under a
// module's lessons table. A guide that cites no form gets the pages it always had.
import test from 'node:test';
import assert from 'node:assert/strict';
import {pane,submit,addModule,addLesson,click,moveTo,startsWith} from './helpers/guide-harness.mjs';
import {formTitleBlock} from '../src/guide-forms-ooxml.mjs';

const index=(a,text)=>a.adapter.snapshot().findIndex(p=>p.text.includes(text));
async function guideWithCitedForms(){
  const a=pane();
  await a.p.refresh();
  a.type('guide-title','Forms pages');await submit(a,'guide-start-form');
  await addModule(a,'One');await addLesson(a,'First');
  await addModule(a,'Two');await addLesson(a,'Second');
  await a.adapter.insert(()=>formTitleBlock({id:'f1',label:'Handout',number:'',title:'Goals'})+formTitleBlock({id:'f2',label:'Worksheet',number:'',title:'Answers',facilitator:true}),{where:'after',after:index(a,'Lesson 1.1')+1,spare:false});
  return a;
}
const cite=async(a,lessonText,formId,words)=>{
  a.p.useTool('instructionalCues');await a.p.refresh();
  await moveTo(a,startsWith(`OSG Lesson: ${lessonText}`));
  a.pick('handout');a.type('guide-block-which',formId);a.type('guide-block-text',words);
  await submit(a,'guide-block-form');assert.equal(a.err(),'');
};
const tableRows=a=>{
  const ps=a.adapter.snapshot(),at=ps.findIndex(p=>p.style==='OSG Table Head'&&p.text.trim()==='Module or lesson');
  const cells=[];for(let i=at;i<ps.length&&/^OSG Table (Head|Cell)$/.test(ps[i].style);i++)cells.push(ps[i].text.trim());
  return cells;
};

test('a guide that cites no form keeps the three-column Training at a Glance',async()=>{
  const a=await guideWithCitedForms();
  await click(a,'guide-update');
  await click(a,'guide-build-glance');assert.equal(a.err(),'');
  assert.deepEqual(tableRows(a).slice(0,3),['Module or lesson','Minutes','Starts at minute']);
  assert.ok(!tableRows(a).includes('Forms used'));
});

test('once blocks cite forms, Training at a Glance has a Forms used column: each lesson lists what its blocks cite, a module what its lessons cite',async()=>{
  const a=await guideWithCitedForms();
  await click(a,'guide-update');
  await cite(a,'Lesson 1.1','f1','Give out');
  await cite(a,'Lesson 2.1','f2','Check');
  await click(a,'guide-update');
  await click(a,'guide-build-glance');assert.equal(a.err(),'');
  const cells=tableRows(a);
  assert.deepEqual(cells.slice(0,4),['Module or lesson','Minutes','Starts at minute','Forms used']);
  const body=cells.slice(4),rows=[];for(let i=0;i<body.length;i+=4)rows.push(body.slice(i,i+4));
  const by=name=>rows.find(r=>r[0].includes(name));
  assert.match(by('Module 1')[3],/Handout 1\.1/,'the module lists the form its lesson cites');
  assert.match(by('Lesson 1.1')[3],/^Handout 1\.1$/);
  assert.match(by('Lesson 2.1')[3],/^Worksheet 2\.1$|^Worksheet 1\.2$|^Worksheet/,'a form for the facilitator is listed too');
  assert.equal(by('Lesson 1.2')?.[3]??'','','a lesson that cites nothing has an empty cell');
  assert.match(a.n('guide-glance-note').textContent,/matches your guide/);
});

test('the page falls out of date when a block starts to cite a form, and building it again brings the column in',async()=>{
  const a=await guideWithCitedForms();
  await click(a,'guide-update');
  await click(a,'guide-build-glance');
  assert.match(a.n('guide-glance-note').textContent,/matches your guide/);
  await cite(a,'Lesson 1.1','f1','Give out');
  await a.p.refresh();
  assert.match(a.n('guide-glance-note').textContent,/out of date/,'the new citation is not on the page yet');
  await click(a,'guide-build-glance');
  assert.ok(tableRows(a).includes('Forms used'));
  assert.match(a.n('guide-glance-note').textContent,/matches your guide/);
});

test('a module overview names the participant handouts and the facilitator forms its blocks cite, and says so again when they change',async()=>{
  const a=await guideWithCitedForms();
  await click(a,'guide-update');
  await cite(a,'Lesson 1.1','f1','Give out');
  await cite(a,'Lesson 1.1','f2','Check');
  await click(a,'guide-update');
  await moveTo(a,startsWith('OSG Module: Module 1'));
  await click(a,'guide-build-overview');assert.equal(a.err(),'');
  const texts=a.adapter.snapshot().map(p=>p.text.trim());
  assert.ok(texts.some(t=>/^Participant handouts: Handout 1\.1\.$/.test(t)),texts.filter(t=>/handouts|forms/i.test(t)).join(' | '));
  assert.ok(texts.some(t=>/^Facilitator forms: Worksheet 1\.2\.$/.test(t)));
  assert.match(a.n('guide-overview-note').textContent,/matches your guide/);
  // Another module's overview has none of these lines.
  await moveTo(a,startsWith('OSG Module: Module 2'));
  await click(a,'guide-build-overview');
  const second=a.adapter.snapshot().map(p=>p.text.trim());
  assert.equal(second.filter(t=>/^Participant handouts:/.test(t)).length,1,'only the first module cites a handout');
  // Building the first one again does not stack the lines.
  await moveTo(a,startsWith('OSG Module: Module 1'));
  await click(a,'guide-build-overview');
  assert.equal(a.adapter.snapshot().filter(p=>/^Participant handouts:/.test(p.text.trim())).length,1,'the lines are replaced, not added again');
});
