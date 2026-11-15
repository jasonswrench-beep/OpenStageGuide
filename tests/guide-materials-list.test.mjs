// The Materials list page (kit slice K5): what to have ready, grouped for the facilitator, each person, each team, each table and the room, with quantities.
// It is built from each lesson's Materials fact and the print lines of the forms, when the author asks, and it is kept out of the participant workbook.
// It drives the real pane and the real taskpane.html ids in memory.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {pane,submit,addModule,addLesson,click,moveTo,startsWith} from './helpers/guide-harness.mjs';
import {formTitleBlock,printLine,formText} from '../src/guide-forms-ooxml.mjs';
import {materialsRows,materialsBody,MATERIALS_HEADING} from '../src/guide-pages.mjs';
import {readGuide} from '../src/guide-model.mjs';
import {readSettings,writeSettings} from '../src/guide-forms.mjs';
import {readSource,buildWorkbook} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';

const index=(a,text)=>a.adapter.snapshot().findIndex(p=>p.text.includes(text));
async function guide(){
  const a=pane();
  await a.p.refresh();
  a.type('guide-title','Materials guide');await submit(a,'guide-start-form');
  await addModule(a,'One');await addLesson(a,'First','10');
  await addModule(a,'Two');await addLesson(a,'Second','5');
  return a;
}
const addForms=async a=>{
  await a.adapter.insert(()=>
    formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})+printLine({copies:{n:1,per:'person'},paper:'Letter',orientation:'portrait'})+formText('Read this first.')
    +formTitleBlock({id:'f2',label:'Checklist',number:'1.2',title:'Room'})+printLine({copies:{n:2,per:'team'}})+formText('Tick them.')
    +formTitleBlock({id:'f3',label:'Poster',number:'1.3',title:'Rules'})+printLine({copies:{n:1,per:'room'}})+formText('Hang it.'),{where:'after',after:index(a,'Lesson 1.1'),spare:false});
};
const tableCells=a=>{
  const ps=a.adapter.snapshot(),at=ps.findIndex(p=>p.style==='OSG Table Head'&&p.text.trim()==='For');
  const cells=[];for(let i=at;i<ps.length&&/^OSG Table (Head|Cell)$/.test(ps[i].style);i++)cells.push(ps[i].text.trim());
  return cells;
};

test('a guide with nothing to list says what to give it, and builds nothing',async()=>{
  const a=await guide();
  assert.match(a.n('guide-materials-note').textContent,/needs a lesson with materials, or a form with a print line/);
  await click(a,'guide-build-materials');
  assert.match(a.err(),/There is nothing to list yet/);
  assert.equal(index(a,MATERIALS_HEADING),-1,'no page was started');
});

test('the Materials list gathers each lesson\'s materials and the copies to make, grouped and with quantities',async()=>{
  const a=await guide();
  await addForms(a);
  a.p.useTool('forms');await a.p.refresh();
  a.type('guide-forms-people','24');await a.n('guide-forms-people').onchange();
  a.type('guide-forms-teams','6');await a.n('guide-forms-teams').onchange();
  // a lesson's materials are the Materials fact under its time line
  await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));
  a.p.useTool('lessonDetails');await a.p.refresh();
  a.type('guide-details-materials','Flip chart and markers');a.n('guide-details-method').value='Lecture';
  await submit(a,'guide-details-form');assert.equal(a.err(),'');
  const paras=await a.adapter.paragraphs();
  const settings=readSettings(await a.adapter.formSettings());
  const rows=materialsRows(paras,readGuide(paras),settings);
  assert.deepEqual(rows.map(r=>[r.group,r.item.replace(/^.*: /,''),r.quantity]),[
    ['The facilitator','Flip chart and markers',''],
    ['Each person','Goals','1 each (24 in all)'],
    ['Each team','Room','2 each (12 in all)'],
    ['The room','Rules','1']
  ].map(r=>r),'the forms come out grouped, in the order the groups are listed');
  await click(a,'guide-build-materials');assert.equal(a.err(),'');
  assert.deepEqual(tableCells(a).slice(0,3),['For','Item','How many']);
  assert.match(a.n('guide-materials-note').textContent,/matches your guide/);
  assert.match(a.n('guide-status').textContent,/The materials list is built with 4 lines./);
  assert.equal(a.adapter.snapshot().filter(p=>p.text.trim()===MATERIALS_HEADING).length,1,'one heading');
});

test('the page falls out of date when a number changes, and building it again replaces it rather than adding another',async()=>{
  const a=await guide();
  await addForms(a);
  a.p.useTool('forms');await a.p.refresh();
  a.type('guide-forms-people','24');await a.n('guide-forms-people').onchange();
  a.type('guide-forms-teams','6');await a.n('guide-forms-teams').onchange();
  await click(a,'guide-build-materials');assert.equal(a.err(),'');
  assert.match(a.n('guide-materials-note').textContent,/matches your guide/);
  a.p.useTool('forms');await a.p.refresh();
  a.type('guide-forms-people','30');await a.n('guide-forms-people').onchange();
  assert.match(a.n('guide-materials-note').textContent,/out of date/);
  await click(a,'guide-build-materials');assert.equal(a.err(),'');
  assert.equal(a.adapter.snapshot().filter(p=>p.text.trim()===MATERIALS_HEADING).length,1,'replaced, not added again');
  assert.ok(tableCells(a).some(c=>/30 in all/.test(c)),'the new number is in the table');
  assert.match(a.n('guide-status').textContent,/The earlier one was replaced/);
});

test('the pure builder says what to do when there is nothing to list, and a form whose units are unknown is listed with that said',()=>{
  assert.throws(()=>materialsBody([]),/There is nothing to list yet/);
  const paras=[{style:'Normal',text:'x'}];
  assert.deepEqual(materialsRows(paras,{modules:[]},{}),[]);
  const body=materialsBody([{group:'Each person',item:'Goals',quantity:'1 each; how many in all is not known yet'}],{withHeading:true});
  assert.ok(body.includes(MATERIALS_HEADING)&&body.includes('how many in all is not known yet'));
});

test('the Materials list is for the person making the guide: the participant workbook leaves it out and the leak check passes',async()=>{
  const a=await guide();
  await addForms(a);
  a.p.useTool('forms');await a.p.refresh();
  a.type('guide-forms-people','24');await a.n('guide-forms-people').onchange();
  await click(a,'guide-build-materials');assert.equal(a.err(),'');
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},{DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const doc=built.files.find(f=>f.name==='word/document.xml').text;
  assert.ok(!doc.includes('What to have ready'),'the page is not in the workbook');
  assert.ok(!doc.includes('Materials list'),'nor its heading');
  const leak=leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  assert.equal(leak.passed,true,JSON.stringify(leak.findings.filter(f=>f.level!=='pass')));
});
