// The Session templates tool on its own panel (src/guide-pane-templates.mjs): browse, preview without change, apply a copy as a lesson or a module, save a lesson as a template.
// Drives the real pane module and the real taskpane.html ids in memory. It is not native Word.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {readGuide,readBlock,updatePlan} from '../src/guide-model.mjs';
import {readFacts} from '../src/guide-facts.mjs';
import {TEMPLATES_KEY} from '../src/session-text.mjs';
import {started,click,addModule,addLesson,submit} from './helpers/guide-harness.mjs';

const memoryStorage=(start={})=>{const data={...start};return {data,getItem:k=>k in data?data[k]:null,setItem:(k,v)=>{data[k]=String(v);}};};
const withTool=async(storage=memoryStorage())=>{
  const a=await started('Practice day',{storage});
  await addModule(a,'Module one');await addLesson(a,'Lesson one','10');
  a.p.useTool('sessionTemplates');
  await a.p.openTool('sessionTemplates');
  return {a,storage};
};
const names=(a,id='guide-tpl-list')=>a.n(id).children.map(li=>li.children[0].textContent.replace(/ \(\d+ min\).*$/,''));
const cardOf=(a,title,id='guide-tpl-list')=>a.n(id).children.find(li=>li.children[0].textContent.startsWith(`${title} (`));
const button=(li,label)=>li.children[2].children.find(b=>b.textContent===label);
const guideOf=async a=>readGuide(await a.adapter.paragraphs());

test('the tool is in the Start menu, opens on its own panel, and lists the 16 templates with what each is for',async()=>{
  const {a}=await withTool();
  assert.equal(a.n('guide-sec-templates').hidden,false);
  assert.equal(a.n('guide-sec-library').hidden,true);assert.equal(a.n('guide-sec-module').hidden,true);
  assert.equal(names(a).length,16);
  assert.equal(names(a)[0],'Skill Practice Workshop');
  assert.equal(a.n('guide-tpl-count').textContent,'16 templates.');
  const first=a.n('guide-tpl-list').children[0];
  assert.match(first.children[0].textContent,/Skill Practice Workshop \(60 min\)Teach and practice/);
  assert.match(first.children[1].textContent,/Trainers and instructors, 6 to 24 people\. Teach one skill and give everyone a try with feedback\. Makes: Revised performance and one transfer action\./);
  assert.equal(a.n('guide-tpl-preview').hidden,true,'no preview until one is chosen');
  assert.equal(a.err(),'');
});

test('the page says plainly that no AI helper is set up, so nothing is sent anywhere',()=>{
  const html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');
  assert.match(html,/id="guide-tpl-ai-note"[^>]*>No AI helper is set up in this program, so nothing here is sent anywhere\./);
});

test('the list narrows by words, by purpose, by length and by set, and says how many were found',async()=>{
  const {a}=await withTool();
  a.type('guide-tpl-find','retrospective');a.n('guide-tpl-find').oninput();
  assert.deepEqual(names(a),['Team Retrospective']);
  assert.equal(a.n('guide-tpl-count').textContent,'1 of 16 templates found.');
  a.type('guide-tpl-find','');a.type('guide-tpl-family','Improve instruction');a.n('guide-tpl-family').onchange();
  assert.deepEqual(names(a),['Training Transfer Planning','Accessible Lesson Redesign']);
  a.type('guide-tpl-family','');a.type('guide-tpl-band','long');a.n('guide-tpl-band').onchange();
  assert.deepEqual(names(a),['Training Transfer Planning','Difficult Conversation Rehearsal','Problem Diagnosis','Prototype Sprint']);
  a.type('guide-tpl-band','');a.type('guide-tpl-pack','Expansion');a.n('guide-tpl-pack').onchange();
  assert.equal(names(a).length,6);
  a.type('guide-tpl-pack','Yours');a.n('guide-tpl-pack').onchange();
  assert.equal(a.n('guide-tpl-list').children.length,0);
  assert.equal(a.n('guide-tpl-count').textContent,'0 of 16 templates found.');
  assert.deepEqual(a.n('guide-tpl-pack').children.map(o=>o.value),['','Launch','Expansion','Yours']);
});

test('Preview shows the outcome, the agenda with its times, the materials and the ways to adapt it, and changes nothing in the guide',async()=>{
  const {a}=await withTool();
  const before=a.rows();
  await button(cardOf(a,'Skill Practice Workshop'),'Preview').onclick();
  assert.equal(a.n('guide-tpl-preview').hidden,false);
  assert.equal(a.n('guide-tpl-preview-title').textContent,'Skill Practice Workshop (60 min)');
  const facts=a.n('guide-tpl-preview-facts').children.map(p=>p.textContent);
  assert.ok(facts.some(f=>/^What the group leaves with: Revised performance and one transfer action/.test(f)));
  assert.ok(facts.some(f=>/^Before you start: A clear skill/.test(f)));
  assert.ok(facts.some(f=>/^Why the activities come in this order: /.test(f)));
  assert.ok(facts.some(f=>/^Materials: Scenario cards, Observation checklist, Facilitator guide/.test(f)));
  assert.ok(facts.some(f=>/^Ways to adapt it: .*Offer a written simulation\./.test(f)));
  const rows=a.n('guide-tpl-preview-body').children.map(tr=>tr.children.map(c=>c.textContent));
  assert.equal(rows.length,6);
  assert.deepEqual(rows[1],['5 min to 17 min','Model and notice','Practice','12','Observation notes tied to one criterion']);
  // a start time turns the minutes into clock times
  a.type('guide-tpl-start','9:00');a.n('guide-tpl-start').onchange();
  assert.deepEqual(a.n('guide-tpl-preview-body').children.map(tr=>tr.children[0].textContent),['09:00 to 09:05','09:05 to 09:17','09:17 to 09:35','09:35 to 09:47','09:47 to 09:52','09:52 to 10:00']);
  assert.equal(a.n('guide-tpl-title').value,'Skill Practice Workshop');
  assert.match(a.n('guide-tpl-where-note').textContent,/It goes after Lesson 1\.1: Lesson one, as Lesson 1\.2\./);
  assert.match(a.n('guide-status').textContent,/Nothing in the guide has changed/);
  assert.deepEqual(a.rows(),before,'the guide is exactly as it was');
  await click(a,'guide-tpl-close');
  assert.equal(a.n('guide-tpl-preview').hidden,true);
  assert.deepEqual(a.rows(),before);
});

test('Apply puts a copy in the guide as a new lesson in the module, with its time, facts, objective and one Activity block for each row',async()=>{
  const {a}=await withTool();
  await button(cardOf(a,'Skill Practice Workshop'),'Preview').onclick();
  a.type('guide-tpl-title','Feedback skills');a.type('guide-tpl-audience','new managers');a.type('guide-tpl-start','9:00');
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/A copy of "Skill Practice Workshop" is in the guide as Lesson 1\.2: Feedback skills\. The template itself is unchanged\./);
  const paras=await a.adapter.paragraphs(),guide=readGuide(paras),lesson=guide.modules[0].lessons[1];
  assert.equal(guide.modules[0].lessons.length,2);
  assert.equal(lesson.text,'Lesson 1.2: Feedback skills');
  assert.equal(lesson.blockMinutes,60);
  const {facts}=readFacts(paras,lesson);
  assert.equal(facts.purpose,'Teach one skill and give everyone a try with feedback. For new managers.');
  assert.equal(facts.newDay,'9:00');
  assert.deepEqual(facts.extra,['Template: Skill Practice Workshop (T01, version 1)']);
  const activity=paras.findIndex(p=>/Scenario rehearsal \(18 min\)/.test(p.text));
  const b=readBlock(paras,activity);
  assert.equal(b.parts.phase,'Practice');
  assert.match(b.parts.steps,/3 min: Assign performer, counterpart, and observer/);
  assert.match(b.parts.source,/\(A13, version 1\)$/);
  // the lesson's time line is right once Update guide has run, and the Review has nothing to say about its steps
  await click(a,'guide-update');
  assert.deepEqual(updatePlan(await a.adapter.paragraphs()).edits,[]);
  // applied again, it is another copy, and the first is untouched
  await submit(a,'guide-tpl-apply-form');
  assert.equal((await guideOf(a)).modules[0].lessons.length,3);
  assert.equal(a.n('guide-tpl-list').children.length,16,'the catalog is as it was');
});

test('Apply as a new module writes a module with its own Transition block, after the module the cursor is in',async()=>{
  const {a}=await withTool();
  await button(cardOf(a,'Team Retrospective'),'Preview').onclick();
  a.n('guide-tpl-where-new').checked=true;a.n('guide-tpl-where-new').onchange();
  assert.match(a.n('guide-tpl-where-note').textContent,/as Module 2, with its own Transition block/);
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  const guide=await guideOf(a);
  assert.equal(guide.modules.length,2);
  assert.equal(guide.modules[1].text,'Module 2: Team Retrospective');
  assert.equal(guide.modules[1].hasTransition,true);
  assert.equal(guide.modules[1].lessons[0].blockMinutes,60);
  assert.match(a.n('guide-status').textContent,/as Module 2: Team Retrospective/);
});

test('Apply needs a preview first, a title and a usable start time, says so, and writes nothing until they are right',async()=>{
  const {a}=await withTool();
  const before=a.rows();
  await submit(a,'guide-tpl-apply-form');
  assert.match(a.err(),/Choose a template and Preview it first/);
  await button(cardOf(a,'Skill Practice Workshop'),'Preview').onclick();
  a.type('guide-tpl-title','   ');
  await submit(a,'guide-tpl-apply-form');
  assert.match(a.err(),/Type a title for the session first/);
  a.type('guide-tpl-title','Feedback skills');a.type('guide-tpl-start','nine');
  await submit(a,'guide-tpl-apply-form');
  assert.match(a.err(),/clock time such as 9:00/);
  assert.deepEqual(a.rows(),before,'nothing was written');
  assert.equal(a.n('guide-tpl-title').value,'Feedback skills','and what was typed is still there');
});

test('a lesson can be saved as a template, kept on this computer, found in the next pane, applied and removed',async()=>{
  const storage=memoryStorage();
  const {a}=await withTool(storage);
  await button(cardOf(a,'Case Discussion'),'Preview').onclick();
  a.type('guide-tpl-title','Case morning');
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  // the cursor is now in the lesson that was written
  const paras=await a.adapter.paragraphs();
  a.adapter.place(paras.findIndex(p=>/^Lesson 1\.2/.test(p.text))+1);await a.p.quietRefresh();
  a.type('guide-tpl-name','My case morning');
  await click(a,'guide-tpl-save');
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/My case morning was saved to your templates, with 5 activities\./);
  assert.deepEqual(names(a,'guide-tpl-mine'),['My case morning']);
  assert.equal(a.n('guide-tpl-count').textContent,'17 templates, 1 of them yours.');
  assert.equal(JSON.parse(storage.data[TEMPLATES_KEY]).length,1);
  assert.equal(a.n('guide-tpl-name').value,'');
  // the next pane, a new document, finds it
  const b=await started('Another guide',{storage});
  await addModule(b,'M');await addLesson(b,'L','10');
  b.p.useTool('sessionTemplates');await b.p.openTool('sessionTemplates');
  assert.deepEqual(names(b,'guide-tpl-mine'),['My case morning']);
  await button(cardOf(b,'My case morning','guide-tpl-mine'),'Preview').onclick();
  await submit(b,'guide-tpl-apply-form');
  assert.equal(b.err(),'');
  assert.equal((await readGuide(await b.adapter.paragraphs())).modules[0].lessons[1].blockMinutes,58);
  // removing it leaves the copy in the guide
  await button(cardOf(b,'My case morning','guide-tpl-mine'),'Remove').onclick();
  assert.deepEqual(names(b,'guide-tpl-mine'),[]);
  assert.equal(b.n('guide-tpl-preview').hidden,true);
  assert.equal(JSON.parse(storage.data[TEMPLATES_KEY]).length,0);
  assert.equal((await readGuide(await b.adapter.paragraphs())).modules[0].lessons.length,2);
});

test('saving says what it cannot do: no lesson with activities, and storage that refuses',async()=>{
  const {a}=await withTool();
  a.type('guide-tpl-name','Empty');
  await click(a,'guide-tpl-save');
  assert.match(a.err(),/no activities to save/);
  const refusing={getItem:()=>null,setItem:()=>{throw new Error('full');}};
  const {a:b}=await withTool(refusing);
  await button(cardOf(b,'Team Retrospective'),'Preview').onclick();
  await submit(b,'guide-tpl-apply-form');
  const paras=await b.adapter.paragraphs();
  b.adapter.place(paras.findIndex(p=>/^Lesson 1\.2/.test(p.text))+1);await b.p.quietRefresh();
  b.type('guide-tpl-name','Will not stay');
  await click(b,'guide-tpl-save');
  assert.match(b.err(),/would not keep your template/);
  assert.deepEqual(names(b,'guide-tpl-mine'),[]);
});

test('the keep-examples and keep-notes choices are read when a template is saved',async()=>{
  const storage=memoryStorage();
  const {a}=await withTool(storage);
  await button(cardOf(a,'Case Discussion'),'Preview').onclick();
  await submit(a,'guide-tpl-apply-form');
  const paras=await a.adapter.paragraphs();
  a.adapter.place(paras.findIndex(p=>/^Lesson 1\.2/.test(p.text))+1);await a.p.quietRefresh();
  a.type('guide-tpl-name','Plain');
  await click(a,'guide-tpl-save');
  const plain=JSON.parse(storage.data[TEMPLATES_KEY])[0];
  assert.ok(plain.rows.every(r=>r.parts.caseStudy===''&&r.notes===''));
  a.n('guide-tpl-keep-examples').checked=true;a.n('guide-tpl-keep-notes').checked=true;
  a.type('guide-tpl-name','Kept');
  await click(a,'guide-tpl-save');
  assert.equal(JSON.parse(storage.data[TEMPLATES_KEY]).length,2);
});
