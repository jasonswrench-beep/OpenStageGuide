// Session templates: browse, preview, apply as a copy and save as a template (src/session-templates.mjs), proven by writing a template into a guide the way the tool does and
// reading the guide back with the guide's own readers.
import test from 'node:test';
import assert from 'node:assert/strict';
import {makeContext,buildPackage} from '../src/guide-ooxml.mjs';
import {readGuide,readBlock,updatePlan,guideIssues,kindOf,cueFromText} from '../src/guide-model.mjs';
import {readFacts} from '../src/guide-facts.mjs';
import {ACTIVITY_PARTS} from '../src/guide-cues.mjs';
import {TEMPLATES,ACTIVITIES} from '../src/session-catalog.mjs';
import {BUNDLED,TIME_BANDS,allTemplates,templateById,templateCard,matchTemplates,templatePreview,checkStart,sessionXml,moduleXml,templateFromLesson,addMine,removeMine,cleanMine,serializeMine,groupText} from '../src/session-templates.mjs';
import {readSteps,readSource} from '../src/session-text.mjs';
import {paragraphsOf} from './helpers/guide-package.mjs';

const pack=build=>{const ctx=makeContext(),xml=build(ctx);return buildPackage(xml,ctx,Object.fromEntries([...ctx.used].map(i=>[i,'AAAA'])));};
const guideOf=(t,opts={})=>paragraphsOf(pack(ctx=>moduleXml(t,{title:'Practice day',m:1,before:0,...opts},ctx)));
const t01=templateById('T01');

test('every bundled template has its rows, written out from the catalog, and is read-only',()=>{
  assert.equal(BUNDLED.length,16);
  for(const t of BUNDLED){
    assert.equal(t.rows.length,TEMPLATES.find(x=>x.id===t.id).rows.length);
    assert.equal(t.rows.reduce((n,r)=>n+r.minutes,0),t.minutes);
    assert.ok(Object.isFrozen(t)&&Object.isFrozen(t.rows)&&Object.isFrozen(t.rows[0]));
  }
  assert.equal(allTemplates().length,16);
  assert.equal(templateById('T99'),null);
  assert.deepEqual(templateCard(t01),{id:'T01',title:'Skill Practice Workshop',pack:'Launch',family:'Teach and practice',minutes:60,audience:'Trainers and instructors',group:'6 to 24 people',purpose:'Teach one skill and give everyone a try with feedback',output:'Revised performance and one transfer action',mine:false});
  assert.equal(groupText({min:4,max:20}),'4 to 20 people');
});

test('search looks across the title, goal, output, audience and tags; filters by family, pack and time',()=>{
  const ids=(o)=>matchTemplates(allTemplates(),o).map(t=>t.id);
  assert.deepEqual(ids({text:'retrospective'}),['T03']);
  assert.deepEqual(ids({text:'ranked shortlist'}),['T04']);              // an output
  assert.deepEqual(ids({text:'onboarding teams'}),['T07']);              // an audience
  assert.ok(ids({text:'silent idea exchange'}).includes('T04'));          // an activity title is a tag
  assert.ok(ids({text:'observation checklist'}).includes('T01'));         // a material is a tag
  assert.deepEqual(ids({family:'Improve instruction'}),['T08','T10']);
  assert.equal(ids({pack:'Expansion'}).length,6);
  assert.equal(ids({band:'short'}).length+ids({band:'medium'}).length+ids({band:'long'}).length,16);
  assert.deepEqual(ids({band:'long'}).sort(),['T08','T11','T12','T14']);
  assert.deepEqual(ids({text:'feedback',family:'Teach and practice',pack:'Launch'}),['T01','T06']);
  assert.deepEqual(ids({text:'zzz'}),[]);
  assert.equal(TIME_BANDS.length,3);
});

test('a preview shows the outcome, the agenda with its times, the materials and the ways to adapt, and changes nothing',()=>{
  const before=JSON.stringify(t01);
  const p=templatePreview(t01,{start:'9:00'});
  assert.deepEqual(p.agenda.map(r=>r.from),['09:00','09:05','09:17','09:35','09:47','09:52']);
  assert.equal(p.agenda.at(-1).to,'10:00');
  assert.equal(p.total,60);
  assert.deepEqual(p.outcomes,['Revised performance and one transfer action']);
  assert.ok(p.rationale.length>20&&p.prerequisites.length>10);
  assert.deepEqual(p.materials,['Scenario cards','Observation checklist','Facilitator guide']);
  assert.ok(p.adjustments.includes('Offer a written simulation.'));
  assert.deepEqual(p.agenda.map(r=>r.phase),['Open','Practice','Practice','Practice','Close','Close']);
  // with no start time the agenda shows minutes from the start
  const plain=templatePreview(t01);
  assert.deepEqual(plain.agenda.map(r=>r.from),['0 min','5 min','17 min','35 min','47 min','52 min']);
  assert.equal(JSON.stringify(t01),before);
  assert.throws(()=>templatePreview(t01,{start:'25:99'}),/clock time such as 9:00/);
  assert.equal(checkStart(' 09:30 '),'9:30');assert.equal(checkStart(''),'');
});

test('applied to a guide, a template is a module with a lesson, its time, its facts and one Activity block for each row, in the right order and minutes',()=>{
  const paras=guideOf(t01,{audience:'sales trainers',start:'9:00'});
  const guide=readGuide(paras),lesson=guide.modules[0].lessons[0];
  assert.equal(guide.modules[0].text,'Module 1: Practice day');
  assert.equal(lesson.text,'Lesson 1.1: Practice day');
  assert.equal(guide.modules[0].hasTransition,true);
  assert.equal(lesson.timeMinutes,60);
  assert.equal(lesson.blockMinutes,60);
  const {facts}=readFacts(paras,lesson);
  assert.equal(facts.purpose,'Teach one skill and give everyone a try with feedback. For sales trainers.');
  assert.equal(facts.materials,'Scenario cards, Observation checklist, Facilitator guide');
  assert.equal(facts.group,'6 to 24 people');
  assert.equal(facts.newDay,'9:00');
  assert.deepEqual(facts.extra,['Template: Skill Practice Workshop (T01, version 1)']);
  const blocks=[];
  for(let i=lesson.index+1;i<=lesson.end;i++)if(kindOf(paras[i])==='cue'){const c=cueFromText(paras[i].text);if(c.id==='activity')blocks.push(readBlock(paras,i));}
  assert.deepEqual(blocks.map(b=>b.text),t01.rows.map(r=>r.title));
  assert.deepEqual(blocks.map(b=>Number(b.minutes)),[5,12,18,12,5,8]);
  // every part came through, and the steps still add up to the block's minutes when read back
  for(const b of blocks){
    assert.ok(b.parts.instructions&&b.parts.output&&b.parts.steps&&b.parts.source&&b.parts.phase);
    assert.equal(readSteps(b.parts.steps).minutes,Number(b.minutes));
    assert.deepEqual(readSource(b.parts.source)?.version,1);
  }
  assert.equal(readSource(blocks[2].parts.source).id,'A13');
  // the objective block says what the group leaves with
  assert.ok(paras.some(p=>/Objective\s+Participants leave with: Revised performance and one transfer action\./.test(p.text)));
});

test('the guide made from a template is a sound guide: its times agree and the Review has nothing to say about its steps',()=>{
  const paras=guideOf(t01);
  const plan=updatePlan(paras);
  assert.deepEqual(plan.edits,[]);
  const issues=guideIssues(plan).filter(i=>!/TRANSITION|GUIDE-TITLE/.test(i.rule));
  assert.deepEqual(issues,[]);
  assert.equal(plan.totalMinutes,60);
});

test('every bundled template can be applied, alone or in a module, and reads back with its rows and total',()=>{
  for(const t of BUNDLED){
    const alone=paragraphsOf(pack(ctx=>moduleXml(t,{title:t.title,m:1},ctx)));
    const lesson=readGuide(alone).modules[0].lessons[0];
    assert.equal(lesson.blockMinutes,t.minutes,t.id);
    assert.equal(lesson.cues,t.rows.length+1,`${t.id}: its rows and its objective`);
    const lonely=paragraphsOf(pack(ctx=>sessionXml(t,{title:t.title,m:2,l:3,before:10},ctx)));
    assert.ok(lonely.some(p=>p.text===`Lesson 2.3: ${t.title}`),t.id);
  }
});

test('the catalog and another copy are untouched when one copy is edited',()=>{
  const before=JSON.stringify([BUNDLED,ACTIVITIES]);
  const a=guideOf(t01),b=guideOf(t01);
  const edited=a.map(p=>/Scenario rehearsal/.test(p.text)?{...p,text:p.text.replace('Scenario rehearsal','My rehearsal')}:p);
  assert.ok(edited.some(p=>/My rehearsal/.test(p.text)));
  assert.ok(b.some(p=>/Scenario rehearsal/.test(p.text))&&!b.some(p=>/My rehearsal/.test(p.text)));
  assert.equal(JSON.stringify([BUNDLED,ACTIVITIES]),before);
  assert.throws(()=>{'use strict';t01.rows[0].title='x';});
});

test('applying needs a title and a sensible start time, and says so in plain words',()=>{
  assert.throws(()=>pack(ctx=>sessionXml(t01,{title:'  ',m:1,l:1},ctx)),/Type a title for the session first/);
  assert.throws(()=>pack(ctx=>sessionXml(t01,{title:'X',start:'nine',m:1,l:1},ctx)),/clock time such as 9:00/);
});

test('a lesson saved as a template keeps its instructions, steps and prompts, and drops examples and notes unless asked to keep them',()=>{
  // a lesson with one catalog activity, one activity with a case study and notes, one plain activity, and a Say block that is not an activity
  const paras=paragraphsOf(pack(ctx=>moduleXml(BUNDLED.find(t=>t.id==='T05'),{title:'Case day',m:1},ctx)));
  const withExamples=[...paras];
  const custom=pack(ctx=>
    moduleXml({...t01,rows:[...t01.rows.slice(0,1),{title:'Dana case',minutes:6,parts:{instructions:'Read it.\nTalk it over.',caseStudy:'Dana and Sam met on Monday.',objective:'Weigh the facts.'},notes:'Walk the room. Dana is shy.'}]},{title:'Mixed',m:1},ctx));
  const mixed=paragraphsOf(custom);
  const at=mixed.findIndex(p=>/^Lesson 1\.1/.test(p.text))+3;
  const {template,skipped}=templateFromLesson(mixed,at,{name:'My mixed day'});
  assert.equal(skipped,0);
  assert.equal(template.title,'My mixed day');
  assert.deepEqual(template.rows.map(r=>r.title),['Outcome and agenda briefing','Dana case']);
  assert.equal(template.rows[1].parts.caseStudy,'');
  assert.equal(template.rows[1].notes,'');
  assert.equal(template.rows[1].parts.instructions,'Read it.\nTalk it over.');
  assert.equal(template.rows[0].parts.steps.split('\n').length,3);
  assert.equal(template.materials.length,3);
  assert.deepEqual(template.group,{min:6,max:24});
  const kept=templateFromLesson(mixed,at,{name:'Kept',keepExamples:true,keepNotes:true}).template;
  assert.equal(kept.rows[1].parts.caseStudy,'Dana and Sam met on Monday.');
  assert.equal(kept.rows[1].notes,'Walk the room. Dana is shy.');
  assert.ok(withExamples.length>10);
  // outside a lesson there is nothing to save, and a lesson with no activities says so
  assert.throws(()=>templateFromLesson(mixed,0,{name:'x'}),/inside a lesson/);
  const empty=paragraphsOf(pack(ctx=>`${'<w:p><w:pPr><w:pStyle w:val="OSGModule"/></w:pPr><w:r><w:t>Module 1: A</w:t></w:r></w:p>'}`));
  assert.throws(()=>templateFromLesson(empty,0,{name:'x'}),/inside a lesson/);
});

test('your own templates are kept, applied like the bundled ones, replaced when saved again, listed with the others and removed',()=>{
  const mixed=paragraphsOf(pack(ctx=>moduleXml(t01,{title:'Mine',m:1},ctx)));
  const {template}=templateFromLesson(mixed,mixed.findIndex(p=>/^Lesson 1\.1/.test(p.text))+2,{name:'  My workshop  '});
  const first=addMine([],template,{now:()=>1000});
  assert.equal(first.replaced,false);
  assert.equal(first.list.length,1);
  const mine=first.list[0];
  assert.equal(mine.title,'My workshop');
  assert.equal(mine.minutes,60);
  assert.equal(mine.pack,'Yours');
  assert.equal(allTemplates(first.list).length,17);
  assert.equal(templateById(mine.id,first.list).title,'My workshop');
  // it round trips through the text the settings file keeps
  const again=cleanMine(JSON.parse(serializeMine(first.list)));
  assert.deepEqual(again,first.list);
  // saved again with the same name: replaced, same id, not duplicated
  const second=addMine(first.list,{...template,purpose:'Changed'},{now:()=>2000});
  assert.equal(second.replaced,true);
  assert.equal(second.list.length,1);
  assert.equal(second.list[0].id,mine.id);
  assert.equal(second.list[0].purpose,'Changed');
  // it applies as a copy
  const paras=paragraphsOf(pack(ctx=>moduleXml(mine,{title:'From mine',m:1},ctx)));
  assert.equal(readGuide(paras).modules[0].lessons[0].blockMinutes,60);
  assert.ok(readFacts(paras,readGuide(paras).modules[0].lessons[0]).facts.extra.includes('Template: My workshop (your template)'));
  assert.deepEqual(removeMine(second.list,mine.id),[]);
  // the most that are kept, and what is not whole is left out rather than half used
  assert.deepEqual(cleanMine('nope'),[]);
  assert.deepEqual(cleanMine([{id:'x',title:'No rows',rows:[]},{title:'No id',rows:[{title:'a',parts:{instructions:'b'}}]},null,{id:'y',title:'Ok',rows:[{title:'a',minutes:5,parts:{instructions:'b'}}]}]).map(t=>t.id),['y']);
  assert.throws(()=>addMine(Array.from({length:30},(_,i)=>({...mine,id:`m${i}`,title:`T${i}`})),template),/most/);
});

test('every part an Activity can hold is named once, and the parts the catalog writes are all known to the guide',()=>{
  const ids=ACTIVITY_PARTS.map(p=>p.id);
  assert.equal(new Set(ids).size,ids.length);
  assert.equal(new Set(ACTIVITY_PARTS.map(p=>p.label.toLowerCase())).size,ids.length);
  for(const t of BUNDLED)for(const r of t.rows)for(const k of Object.keys(r.parts))assert.ok(ids.includes(k),k);
});
