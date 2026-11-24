// The Session materials tool on its own panel (src/guide-pane-materials.mjs): the run sheet, the preparation checklist, and the list of what has been made with its status.
// Drives the real pane module and the real taskpane.html ids in memory. It is not native Word.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {readSessions} from '../src/session-record.mjs';
import {started,click,addModule,addLesson,submit} from './helpers/guide-harness.mjs';

const withTool=async()=>{
  const a=await started('Skill day');
  await addModule(a,'Practice');await addLesson(a,'Warm up','5');
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  a.n('guide-tpl-list').children[0].children[2].children[0].onclick();
  a.type('guide-tpl-title','Skill practice');
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  a.p.useTool('sessionMaterials');await a.p.openTool('sessionMaterials');
  return a;
};
const items=a=>a.n('guide-mat-outputs').children;
const buttonOf=(li,label)=>li.children[2].children.find(b=>b.textContent===label);
const recorded=async a=>readSessions(await a.adapter.sessions());
const docText=async d=>{const z=await JSZip.loadAsync(Buffer.from(d.base64,'base64'));return z.file('word/document.xml').async('string');};

test('the tool is in the Make menu, opens on its own panel, and says how it counts people and what has been made',async()=>{
  const a=await withTool();
  assert.equal(a.n('guide-sec-materials').hidden,false);
  assert.equal(a.n('guide-sec-decks').hidden,true);assert.equal(a.n('guide-sec-built').hidden,true);
  assert.equal(a.n('guide-mat-scope-note').textContent,'It will be made for every lesson in the guide.');
  assert.match(a.n('guide-mat-counts').textContent,/does not say how many people and teams/);
  assert.match(a.n('guide-mat-outputs-note').textContent,/Nothing has been made yet/);
  assert.deepEqual(a.n('guide-mat-paper').children.map(o=>o.value),['Letter','A4']);
  assert.equal(items(a).length,0);
  assert.equal(a.err(),'');
});

test('Make the run sheet opens a Word file with the timeline and the steps, and lists it as Ready',async()=>{
  const a=await withTool();
  await click(a,'guide-mat-runsheet');
  assert.equal(a.err(),'');
  const opened=a.adapter.openedDocuments();
  assert.equal(opened.length,1);
  assert.equal(opened[0].fileName??opened[0].name,'Skill day - run sheet.docx');
  const xml=await docText(opened[0]);
  assert.match(xml,/Run sheet: Skill day/);assert.match(xml,/Scenario rehearsal \(18 min\)/);
  assert.match(a.n('guide-status').textContent,/The run sheet is made\. It is open in a new Word window/);
  assert.equal(items(a).length,1);
  assert.match(items(a)[0].children[0].textContent,/^Run sheet: Run sheet$/);
  assert.match(items(a)[0].children[1].textContent,/The whole guide\. Made .*Ready: the guide has not changed since it was made\./);
  const rec=await recorded(a);
  assert.equal(rec.outputs.length,1);assert.equal(rec.outputs[0].kind,'runsheet');assert.ok(rec.outputs[0].sources.length>=6);
});

test('after the guide changes the run sheet is Review needed, what changed is named, and making just that one again makes it Ready',async()=>{
  const a=await withTool();
  await click(a,'guide-mat-runsheet');
  await click(a,'guide-mat-checklist');
  assert.equal(items(a).length,2);
  // add another session to the guide: both outputs were made from the guide as it was
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  a.n('guide-tpl-list').children[4].children[2].children[0].onclick();
  a.type('guide-tpl-title','Case morning');
  await submit(a,'guide-tpl-apply-form');
  a.p.useTool('sessionMaterials');await a.p.openTool('sessionMaterials');
  const run=items(a).find(li=>/Run sheet/.test(li.children[0].textContent));
  assert.match(run.children[1].textContent,/Review needed: the guide has changed since it was made\./);
  await buttonOf(run,'What changed').onclick();
  assert.equal(a.n('guide-mat-changes-box').hidden,false);
  assert.match(a.n('guide-mat-changes-title').textContent,/Run sheet: Review needed/);
  const changes=a.n('guide-mat-changes').children.map(li=>li.textContent);
  assert.ok(changes.some(c=>/Case analysis is new in the guide\./.test(c)),changes.join(' | '));
  // the checklist reads supplies, which the new session also names, so it is flagged too; make only the run sheet again
  await buttonOf(run,'Make this one again').onclick();
  assert.equal(a.err(),'');
  assert.equal(a.adapter.openedDocuments().length,3,'one new file, not a replaced one');
  const now=items(a);
  assert.match(now.find(li=>/Run sheet/.test(li.children[0].textContent)).children[1].textContent,/Ready:/);
  assert.match(now.find(li=>/Preparation checklist/.test(li.children[0].textContent)).children[1].textContent,/Review needed:/,'the other output keeps its own status');
  assert.equal((await recorded(a)).outputs.length,2,'the same record was refreshed, not a second one');
});

test('an output can be left out on purpose, included again, and forgotten, and the file already made is not touched',async()=>{
  const a=await withTool();
  await click(a,'guide-mat-checklist');
  const li=()=>items(a)[0];
  await buttonOf(li(),'Leave it out').onclick();
  assert.match(li().children[1].textContent,/Excluded: left out on purpose\./);
  assert.equal((await recorded(a)).outputs[0].excluded,true);
  await buttonOf(li(),'Include it').onclick();
  assert.match(li().children[1].textContent,/Ready:/);
  await buttonOf(li(),'Forget it').onclick();
  assert.equal(items(a).length,0);
  assert.equal((await recorded(a)).outputs.length,0);
  assert.equal(a.adapter.openedDocuments().length,1);
});

test('one lesson can be made on its own, needs the cursor in a lesson, and is watched against that lesson only',async()=>{
  const a=await withTool();
  a.n('guide-mat-scope-lesson').checked=true;a.n('guide-mat-scope-guide').checked=false;a.n('guide-mat-scope-lesson').onchange();
  const paras=await a.adapter.paragraphs();
  a.adapter.place(paras.findIndex(p=>/^Module 1/.test(p.text)));await a.p.quietRefresh();
  await click(a,'guide-mat-runsheet');
  assert.match(a.err(),/Put your cursor inside a lesson first/);
  a.adapter.place(paras.findIndex(p=>/^Lesson 1\.2/.test(p.text))+2);await a.p.quietRefresh();
  assert.match(a.n('guide-mat-scope-note').textContent,/^It will be made for Lesson 1\.2: Skill practice\./);
  await click(a,'guide-mat-runsheet');
  assert.equal(a.err(),'');
  const xml=await docText(a.adapter.openedDocuments()[0]);
  assert.match(xml,/Lesson 1\.2: Skill practice/);assert.doesNotMatch(xml,/Lesson 1\.1: Warm up/);
  assert.match(items(a)[0].children[1].textContent,/^Lesson 1\.2: Skill practice\./);
  assert.equal((await recorded(a)).outputs[0].lessonTitle,'Lesson 1.2: Skill practice');
});

test('private notes are asked for with a tick, a bad start time is refused, and the workbook button opens the workbook tool',async()=>{
  const a=await withTool();
  a.type('guide-mat-start','nine');
  await click(a,'guide-mat-runsheet');
  assert.match(a.err(),/clock time such as 9:00/);
  assert.equal(a.adapter.openedDocuments().length,0);
  a.type('guide-mat-start','8:30');a.n('guide-mat-notes').checked=true;
  await click(a,'guide-mat-runsheet');
  assert.equal(a.err(),'');
  assert.match(items(a)[0].children[0].textContent,/Run sheet with private notes/);
  assert.match(await docText(a.adapter.openedDocuments()[0]),/08:30 to 08:35|09:30|08:3\d/);
  await click(a,'guide-mat-workbook');
  assert.equal(a.n('guide-workbook-go').hidden,false);
});

test('a participant workbook made from the workbook screen is listed here with its layout, and is flagged when the guide changes',async()=>{
  const a=await withTool();
  a.p.useTool('participantWorkbook');await a.p.openTool('participantWorkbook');
  a.type('guide-wb-layout','skills');a.n('guide-wb-layout').onchange();
  await click(a,'guide-workbook-go');
  assert.ok(a.adapter.openedDocuments().length>=1,a.n('guide-workbook-problem').textContent);
  const rec=await recorded(a);
  assert.equal(rec.outputs[0].kind,'workbook');assert.equal(rec.outputs[0].name,'Participant workbook (Skills Practice)');
  a.p.useTool('sessionMaterials');await a.p.openTool('sessionMaterials');
  assert.match(items(a)[0].children[0].textContent,/Participant workbook: Participant workbook \(Skills Practice\)/);
  assert.match(items(a)[0].children[1].textContent,/Ready:/);
  assert.equal(buttonOf(items(a)[0],'Open its tool').textContent,'Open its tool');
});
