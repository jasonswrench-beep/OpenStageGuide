// The five workbook layouts (src/workbook-layouts.mjs): what the participant workbook adds around the guide's own learner content, held to the workbook's own leak check and
// structure check, with nothing meant for the facilitator in it.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readSource,buildWorkbook} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {workbookStructure} from '../src/workbook-checks.mjs';
import {readStyles,textOf} from '../src/guide-check.mjs';
import {WORKBOOK_LAYOUTS,layoutById,layoutPages,participantPlan} from '../src/workbook-layouts.mjs';
import {started,addModule,addLesson,addBlock,submit} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml={DOMParserClass:DOMParser,SerializerClass:XMLSerializer};
const parse=t=>new DOMParser().parseFromString(t,'application/xml');
const text=(files,name)=>files.find(f=>f.name===name)?.text;
const rowsOf=files=>{
  const styles=readStyles(parse(text(files,'word/styles.xml')));
  return Array.from(parse(text(files,'word/document.xml')).getElementsByTagNameNS(W,'p')).map(p=>{
    const id=Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val')||'Normal';
    return `${styles.get(id)?.name||id}: ${textOf(p)}`.trim();
  });
};
const strict=built=>leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
const sourceOf=async a=>readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser});
const build=async(a,options={})=>buildWorkbook(await sourceOf(a),options,xml);

// A guide made from the Skill Practice Workshop template through the tool, plus an objective, and one activity for the facilitator alone.
async function guide(){
  const a=await started('Feedback skills');
  await addModule(a,'Practice');await addLesson(a,'Rehearse','10');
  await addBlock(a,'objective','Give feedback that names an example.');
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  a.n('guide-tpl-list').children[0].children[2].children[0].onclick();
  a.type('guide-tpl-title','Skill practice');
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  return a;
}

test('five layouts and the guide as it is are offered, each with who it fits and what is different about it',()=>{
  assert.deepEqual(WORKBOOK_LAYOUTS.map(l=>l.label),['The guide as it is','Facilitated Workshop','Skills Practice','Case and Discussion','Reflection and Transfer','Assessment and Evidence']);
  assert.ok(WORKBOOK_LAYOUTS.every(l=>l.fit&&l.distinct&&l.note));
  assert.equal(layoutById('skills').label,'Skills Practice');
  assert.equal(layoutById(undefined).id,'');
  assert.equal(layoutById('nope'),null);
  assert.throws(()=>layoutPages('nope',{}),/Choose one of the five workbook layouts/);
});

test('what learners meet is read from the activities\' learner parts only: title, grouping and output, never the facilitator\'s',async()=>{
  const a=await guide();
  const plan=participantPlan(await a.adapter.paragraphs());
  assert.equal(plan.length,6);
  assert.deepEqual(plan[2],{lesson:'Skill practice',title:'Scenario rehearsal',output:'A rehearsed response and one revision',grouping:'Triads'});
  const text=JSON.stringify(plan);
  assert.doesNotMatch(text,/Offer a written simulation|Assign performer|Observers describe feelings/);
});

test('each layout makes a workbook that passes the leak check, has no facilitator words, and adds its own pages',async()=>{
  const a=await guide();
  const base=await build(a);
  const baseRows=rowsOf(base.files);
  const facilitatorOnly=['Offer a written simulation','Assign performer, counterpart, and observer','Observers describe feelings, not behavior','What changed when you repeated the moment','Timed steps','Likely challenges','Access options','Debrief prompts'];
  const expected={
    facilitated:['heading 1: Agenda','Table Cell: Scenario rehearsal'],
    skills:['heading 1: Observation checklist','heading 2: Revision plan','Form Text: What I will change:'],
    case:['heading 1: Evidence table','heading 2: Decision prompts','Form Text: The decision to make:'],
    reflection:['heading 1: Reflection','heading 2: Practice log','heading 2: Review checkpoints'],
    assessment:['heading 1: Rating items','heading 2: How I see it','heading 2: Evidence summary']
  };
  for(const layout of Object.keys(expected)){
    const built=await build(a,{layout});
    const rows=rowsOf(built.files);
    assert.equal(strict(built).passed,true,`${layout}: ${strict(built).summary}`);
    for(const want of expected[layout])assert.ok(rows.some(r=>r.includes(want)),`${layout} has "${want}"`);
    const all=rows.join('\n');
    for(const word of facilitatorOnly)assert.ok(!all.includes(word),`${layout} leaks "${word}"`);
    assert.ok(rows.length>baseRows.length,`${layout} adds pages`);
    const structure=workbookStructure(built.files,xml);
    assert.deepEqual(structure.fix,[],`${layout}: ${JSON.stringify(structure.fix)}`);
  }
});

test('the guide as it is adds nothing, and an unknown layout is refused',async()=>{
  const a=await guide();
  const plain=await build(a),same=await build(a,{layout:''});
  assert.deepEqual(rowsOf(plain.files),rowsOf(same.files));
  await assert.rejects(()=>build(a,{layout:'nope'}),/Choose one of the five workbook layouts/);
});

test('the Facilitated Workshop agenda lists each activity with how people work and what they make, and its notes are labelled for it',async()=>{
  const a=await guide();
  const built=await build(a,{layout:'facilitated'});
  const rows=rowsOf(built.files);
  const at=rows.findIndex(r=>r==='heading 1: Agenda');
  assert.ok(at>=0);
  const table=rows.slice(at,at+40).join('\n');
  assert.match(table,/Table Head: Activity/);
  assert.match(table,/Table Cell: Triads/);
  assert.match(table,/Table Cell: A rehearsed response and one revision/);
  assert.ok(rows.some(r=>r.includes('Your notes:')));
});

test('each layout labels the lines to write on after an activity in its own words',async()=>{
  const a=await guide();
  const labels={skills:'What I tried:',case:'What the evidence says:',reflection:'What I noticed:',assessment:'My evidence:'};
  for(const [layout,label] of Object.entries(labels)){
    const rows=rowsOf((await build(a,{layout})).files);
    assert.ok(rows.some(r=>r.endsWith(label)),`${layout} uses "${label}"`);
    assert.ok(!rows.some(r=>r.endsWith('Your notes:')),`${layout} does not use the plain label`);
  }
});

test('the rating items name their ends, and the scoring words are left out unless the author asks for them',async()=>{
  const a=await guide();
  const without=rowsOf((await build(a,{layout:'assessment'})).files).join('\n');
  assert.match(without,/1 Strongly disagree/);assert.match(without,/5 Strongly agree/);
  assert.match(without,/Left end/);assert.match(without,/Right end/);
  assert.doesNotMatch(without,/How to score/);
  const withKey=await build(a,{layout:'assessment',scoreKey:true});
  assert.match(rowsOf(withKey.files).join('\n'),/How to score the rating items/);
  assert.equal(strict(withKey).passed,true);
  // the evidence summary lists the guide's own objectives
  assert.match(without,/Table Cell: Give feedback that names an example\./);
});

test('the workbook screen offers the layouts, says what each adds, remembers the choice, and the workbook it makes has the layout\'s pages',async()=>{
  const storage=(()=>{const data={};return {data,getItem:k=>k in data?data[k]:null,setItem:(k,v)=>{data[k]=String(v);}};})();
  const a=await started('Feedback skills',{storage});
  await addModule(a,'Practice');await addLesson(a,'Rehearse','10');
  await addBlock(a,'objective','Give feedback that names an example.');
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  a.n('guide-tpl-list').children[0].children[2].children[0].onclick();
  await submit(a,'guide-tpl-apply-form');
  assert.deepEqual(a.n('guide-wb-layout').children.map(o=>o.textContent),['The guide as it is','Facilitated Workshop (In-person or virtual cohorts)','Skills Practice (Practice and coached feedback)','Case and Discussion (Case analysis and classroom discussion)','Reflection and Transfer (Spaced follow-up and field practice)','Assessment and Evidence (Surveys, assessments and training evaluation)']);
  assert.equal(a.n('guide-wb-layout').value,'');
  assert.equal(a.n('guide-wb-score-key').disabled,true,'scoring words belong to one layout only');
  a.type('guide-wb-layout','skills');a.n('guide-wb-layout').onchange();
  assert.equal(a.n('guide-wb-layout-note').textContent,'An observation checklist and a revision plan.');
  assert.equal(JSON.parse(storage.data['openstage-guide.workbook-look']).layout,'skills');
  await a.n('guide-tab-workbook').onclick();
  await a.n('guide-workbook-go').onclick();
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  const {readDocxFiles}=await import('../scripts/workbook-leak-check.mjs');
  const files=await readDocxFiles(Buffer.from(opened.base64,'base64'));
  const rows=rowsOf(Array.isArray(files)?files:[...files].map(([name,v])=>({name,...v})));
  assert.ok(rows.includes('heading 1: Observation checklist'));
  assert.match(a.n('guide-workbook-ok').textContent,/Leak check passed/);
  a.type('guide-wb-layout','assessment');a.n('guide-wb-layout').onchange();
  assert.equal(a.n('guide-wb-score-key').disabled,false);
  // the next pane remembers the layout
  const b=await started('Another',{storage});
  assert.equal(b.n('guide-wb-layout').value,'assessment');
});

test('a guide with no activities adds no agenda, and a facilitator-only activity is not in it',async()=>{
  const a=await started('Plain');
  await addModule(a,'M');await addLesson(a,'L','10');
  await addBlock(a,'say','Hello.');await addBlock(a,'ask','What do you think?');
  const rows=rowsOf((await build(a,{layout:'facilitated'})).files);
  assert.ok(!rows.some(r=>r==='heading 1: Agenda'));
  const b=await started('Mixed');
  await addModule(b,'M');await addLesson(b,'L','10');
  b.pick('activity');b.n('guide-activity-parts').checked=true;b.n('guide-activity-parts').onchange();
  b.type('guide-block-text','Secret rehearsal');b.type('guide-act-instructions','Do it.');
  b.n('guide-audience-facilitator').checked=true;b.n('guide-audience-everyone').checked=false;
  await submit(b,'guide-block-form');
  assert.equal(participantPlan(await b.adapter.paragraphs()).length,0);
});
