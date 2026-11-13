// Forms, slices F6 and F7: the print line and the copies to make, and the Review checks for what is wrong inside a form.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {updatePlan,guideIssues,kindOf} from '../src/guide-model.mjs';
import {parsePrint,printText,readSettings,writeSettings,readForms,formPlan,copiesToMake} from '../src/guide-forms.mjs';
import {formTitleBlock,formText,printLine,checkList} from '../src/guide-forms-ooxml.mjs';
import {classify,readSource,buildWorkbook} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {STYLE_NAMES,STYLE_IDS,stylesPart} from '../src/guide-styles.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const P=(style,text,marks)=>({style,text,...(marks?{marks}:{})});
const mod=(n,t)=>P('OSG Module',`Module ${n}: ${t}`);
const form=(id,text,{facilitator=false}={})=>P(facilitator?'OSG Form Title Facilitator':'OSG Form Title',text,{bookmarks:id?[`osg_form_${id}`]:[],refs:[]});
const body=text=>P('OSG Form Text',text);
const print=text=>P('OSG Form Print',text);
const cite=(id,text,style='OSG Cue Body')=>P(style,`Give out ${text}`,{bookmarks:[],refs:[{id,text}]});
const titleOf=()=>P('OSG Guide Title','A guide');
const rules=(paras,settings)=>guideIssues(updatePlan(paras,{forms:settings})).filter(i=>/^FORM-/.test(i.rule)).map(i=>i.rule);

// ---- F6: the print line
test('a print line reads copies, paper, page, sides and "editable", and keeps words it does not know as notes',()=>{
  assert.deepEqual(parsePrint('Print: 1 per person, Letter, landscape.'),{copies:{n:1,per:'person',size:null},paper:'Letter',orientation:'landscape',sides:null,editable:false,notes:[]});
  assert.deepEqual(parsePrint('print: 2 PER Team of 4, a4, two-sided, editable download, on card stock'),{copies:{n:2,per:'team',size:4},paper:'A4',orientation:null,sides:'two-sided',editable:true,notes:['on card stock']});
  assert.deepEqual(parsePrint('Print: 30 copies').copies,{n:30,per:'fixed',size:null});
  assert.deepEqual(parsePrint('Print: 1 copy.').copies,{n:1,per:'fixed',size:null});
  assert.deepEqual(parsePrint('Print: 3 per table of 6').copies,{n:3,per:'table',size:6});
  assert.equal(parsePrint('Copies: 3'),null);
  assert.equal(parsePrint(''),null);
});

test('a print line is written back in plain words and reads the same', ()=>{
  for(const spec of [
    {copies:{n:1,per:'person',size:null},paper:'Letter',orientation:'landscape'},
    {copies:{n:2,per:'team',size:4},paper:'A4',sides:'two-sided',editable:true,notes:['on card stock']},
    {copies:{n:30,per:'fixed'}},{copies:{n:1,per:'fixed'}},{copies:{n:1,per:'room'}},{paper:'Letter'},{editable:true}
  ]){
    const words=printText(spec);
    assert.match(words,/^Print: .*\.$/);
    assert.equal(printText(parsePrint(words)),words,words);
  }
  assert.equal(printText({copies:{n:1,per:'fixed'}}),'Print: 1 copy.');
  assert.equal(printText({copies:{n:2,per:'person'},orientation:'portrait'}),'Print: 2 per person, portrait.');
  for(const bad of [{},{copies:{n:0,per:'person'}},{copies:{n:1.5,per:'person'}},{copies:{n:1,per:'galaxy'}},{paper:'Folio'},{orientation:'diagonal'},{sides:'three-sided'}])
    assert.throws(()=>printText(bad),Error,JSON.stringify(bad));
});

test('the print line is a paragraph of its own style under the title, and a form reads it back',()=>{
  const x=printLine({copies:{n:1,per:'person'},paper:'Letter',orientation:'landscape'});
  assert.match(x,new RegExp(`<w:pStyle w:val="${STYLE_IDS.formPrint}"/>`));
  assert.match(x,/Print: 1 per person, Letter, landscape\./);
  assert.equal(kindOf({style:STYLE_NAMES.formPrint}),'form-print');
  assert.ok(stylesPart().includes(`w:styleId="${STYLE_IDS.formPrint}"`));
  const read=readForms([form('f1','Handout 1.1: A'),print('Print: 2 per team of 4, A4.'),body('x'),form('f2','Handout 1.2: B'),body('y'),print('Print: 1 per person.')]);
  assert.deepEqual(read.forms[0].print.copies,{n:2,per:'team',size:4});
  assert.equal(read.forms[0].printIndex,1);
  assert.equal(read.forms[1].print,null,'a print line that is not directly under a title belongs to no form');
});

test('the settings keep how many people, teams and tables there are, and drop what is not a count',()=>{
  assert.deepEqual(readSettings(writeSettings({scheme:'letter',group:{people:24,teams:6,tables:4}})),{scheme:'letter',group:{people:24,teams:6,tables:4}});
  assert.deepEqual(readSettings(writeSettings({group:{people:12}})),{scheme:'auto',group:{people:12,teams:null,tables:null}});
  assert.deepEqual(readSettings('{"group":{"people":0,"teams":-1,"tables":"4","x":9}}'),{scheme:'auto'},'nothing usable, no group');
  assert.deepEqual(readSettings('{"scheme":"unit","group":{"people":10000}}'),{scheme:'unit'});
  assert.deepEqual(readSettings(null),{scheme:'auto'});
});

test('copies to make: each print line is multiplied by the people, teams or tables the guide records, and what cannot be added up is listed',()=>{
  const paras=[titleOf(),mod(1,'One'),
    form('f1','Handout 1.1: A'),print('Print: 1 per person.'),body('x'),
    form('f2','Handout 1.2: B'),print('Print: 2 per team.'),body('x'),
    mod(2,'Two'),
    form('f3','Handout 2.1: C'),print('Print: 1 per table of 6.'),body('x'),
    form('f4','Handout 2.2: D'),print('Print: 25 copies.'),body('x'),
    form('f5','Handout 2.3: E'),print('Print: 1 per room.'),body('x'),
    form('f6','Handout 2.4: F'),body('x'),
    form('f7','Handout 2.5: G'),print('Print: 1 per team of 4.'),body('x')];
  const full=copiesToMake(paras,{group:{people:24,teams:6,tables:4}});
  assert.deepEqual(full.rows.map(r=>[r.id,r.module,r.per,r.units,r.total]),[['f1',1,'person',24,24],['f2',1,'team',6,12],['f3',2,'table',4,4],['f4',2,'fixed',1,25],['f5',2,'room',1,1],['f7',2,'team',6,6]]);
  assert.equal(full.total,24+12+4+25+1+6);
  assert.deepEqual(full.missing,[]);
  // With only the number of people, a team of four is worked out and the rest are listed.
  const some=copiesToMake(paras,{group:{people:23}});
  assert.equal(some.rows.find(r=>r.id==='f7').units,6,'23 people in teams of four is six teams');
  assert.equal(some.rows.find(r=>r.id==='f7').total,6);
  assert.deepEqual(some.missing.map(m=>m.text),['Handout 1.2: B','Handout 2.1: C']);
  assert.match(some.missing[0].reason,/how many teams/);
  assert.equal(some.total,23+25+1+6);
  const none=copiesToMake(paras,{});
  assert.equal(none.rows.find(r=>r.id==='f1').total,null);
  assert.equal(none.total,25+1);
});

test('the Review says when a print line cannot be added up, cannot be read, or is missing from a form that is given out',()=>{
  const base=[titleOf(),mod(1,'One'),form('f1','Handout 1.1: A'),print('Print: 1 per person.'),body('x'),form('f2','Handout 1.2: B'),print('Whatever'),body('x'),form('f3','Handout 1.3: C'),body('x'),
    cite('f1','Handout 1.1: A'),cite('f2','Handout 1.2: B'),cite('f3','Handout 1.3: C')];
  const text=paras=>guideIssues(updatePlan(paras,{forms:{}})).filter(i=>/^FORM-PRINT/.test(i.rule)).map(i=>`${i.rule} ${i.message}`);
  const got=text(base);
  assert.ok(got.some(g=>/^FORM-PRINT-GROUP .*Handout 1\.1: A is printed for each person/.test(g)));
  assert.ok(got.some(g=>/^FORM-PRINT-NAME .*Handout 1\.2: B/.test(g)));
  assert.ok(got.some(g=>/^FORM-PRINT-NONE .*Handout 1\.3: C is given out/.test(g)));
  assert.equal(got.filter(g=>/NONE/.test(g)).length,1,'a form with an unreadable print line is not also told it has none');
  const withGroup=guideIssues(updatePlan(base,{forms:{group:{people:20}}})).filter(i=>i.rule==='FORM-PRINT-GROUP');
  assert.deepEqual(withGroup,[],'once the guide says how many people, the line can be added up');
  assert.deepEqual(rules([titleOf(),mod(1,'One'),form('f1','Handout 1.1: A'),body('x')],{}).filter(r=>/PRINT/.test(r)),[],'a form that is not cited is not told to have a print line');
});

// ---- F7: what is wrong inside a form
test('a form with no title, nothing in it, or a field with no label is told to the author',()=>{
  const got=rules([titleOf(),mod(1,'One'),P('OSG Form Title',''),body('x'),form('f1','Handout 1.1: Empty'),form('f2','Handout 1.2: Fields'),P('OSG Form Field',''),P('OSG Form Field','Name:')],{});
  assert.ok(got.includes('FORM-EMPTY-TITLE'));
  assert.ok(got.includes('FORM-EMPTY'));
  assert.ok(got.includes('FORM-FIELD-LABEL'));
  const issues=guideIssues(updatePlan([titleOf(),mod(1,'One'),P('OSG Form Title','')]));
  assert.equal(issues.find(i=>i.rule==='FORM-EMPTY-TITLE').level,'fail');
  assert.equal(issues.filter(i=>i.rule==='FORM-NAME').length,0,'an empty title is told once, as empty, not as badly named');
});

test('a table in a form with no header row is a failure, and each table is judged on its own',()=>{
  const ok=[titleOf(),mod(1,'One'),form('f1','Handout 1.1: A'),P('OSG Table Head','Goal'),P('OSG Table Head','Done'),P('OSG Table Cell','Plan'),P('OSG Rating Box','x')];
  assert.ok(!rules(ok,{}).includes('FORM-TABLE-HEADER'));
  const bad=[...ok,body('Between the tables.'),P('OSG Table Cell','a'),P('OSG Table Cell','b')];
  const issue=guideIssues(updatePlan(bad)).find(i=>i.rule==='FORM-TABLE-HEADER');
  assert.equal(issue.level,'fail');
  assert.equal(issue.paragraph,bad.length-2,'it points at the first cell of the table without a header');
  const atEnd=[titleOf(),mod(1,'One'),form('f1','Handout 1.1: A'),P('OSG Table Cell','a')];
  assert.ok(rules(atEnd,{}).includes('FORM-TABLE-HEADER'),'a headerless table at the very end of the guide is found too');
  const outside=[titleOf(),mod(1,'One'),P('OSG Table Cell','a')];
  assert.ok(!rules(outside,{}).includes('FORM-TABLE-HEADER'),'tables outside a form are not judged here');
});

test('a form for the facilitator that a block for everyone cites is told to the author; a facilitator block may cite it',()=>{
  const forFacilitator=form('f1','Checklist 1.1: Set-up',{facilitator:true});
  const make=style=>[titleOf(),mod(1,'One'),forFacilitator,body('x'),cite('f1','Checklist 1.1: Set-up',style)];
  const told=guideIssues(updatePlan(make('OSG Cue'))).filter(i=>i.rule==='FORM-AUDIENCE');
  assert.equal(told.length,1);assert.match(told[0].message,/for the facilitator only/);
  assert.equal(guideIssues(updatePlan(make('OSG Cue Body'))).filter(i=>i.rule==='FORM-AUDIENCE').length,0,'a line with no block above it is not a block for everyone');
  const after=[titleOf(),mod(1,'One'),forFacilitator,body('x'),P('OSG Cue','Ask  Why?'),cite('f1','Checklist 1.1: Set-up','OSG Cue Body')];
  assert.equal(guideIssues(updatePlan(after)).filter(i=>i.rule==='FORM-AUDIENCE').length,1,'a line under a block for everyone is part of that block');
  const facilitatorBlock=[titleOf(),mod(1,'One'),forFacilitator,body('x'),P('OSG Cue Facilitator','Say  Hello'),cite('f1','Checklist 1.1: Set-up','OSG Cue Body')];
  assert.equal(guideIssues(updatePlan(facilitatorBlock)).filter(i=>i.rule==='FORM-AUDIENCE').length,0);
});

// ---- the print line and the workbook
test('the print line is for the facilitator: the workbook leaves it out and watches for its words',()=>{
  const classified=classify([P('OSG Guide Title','G'),mod(1,'One'),P('OSG Lesson','Lesson 1.1: First'),form('f1','Handout 1.1: A'),print('Print: 1 per person, Letter, landscape.'),body('Read.')]);
  assert.deepEqual(classified.decisions.slice(3).map(d=>[d.keep,d.reason]),[[true,'form'],[false,'form-print'],[true,'form']]);
});

test('through the whole pipeline: the file made for learners has the form but not its print line, and the leak check passes',async()=>{
  const a=await started('Print guide');
  await addModule(a,'One');await addLesson(a,'First');await addBlock(a,'ask','What makes it hard?');
  const at=a.adapter.snapshot().findIndex(p=>p.text.includes('What makes it hard?'));
  await a.adapter.insert(ctx=>formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})+printLine({copies:{n:1,per:'person'},paper:'Letter',notes:['bring a pen']})+formText('Read this first.')+checkList({items:['Read it'],ctx}),{where:'after',after:at,spare:false});
  const list=await a.adapter.paragraphs();
  assert.equal(readForms(list).forms[0].print.copies.per,'person');
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},{DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const doc=built.files.find(f=>f.name==='word/document.xml').text,styles=built.files.find(f=>f.name==='word/styles.xml').text;
  assert.ok(doc.includes('Handout 1.1: Goals')&&doc.includes('Read this first.'));
  assert.ok(!doc.includes('Print: 1 per person')&&!doc.includes('bring a pen')&&!doc.includes('OSGFormPrint'));
  assert.ok(!styles.includes('OSGFormPrint'),'the workbook does not even define the print style');
  assert.ok(built.facilitatorTexts.some(t=>t.startsWith('Print: 1 per person')));
  const result=leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  assert.equal(result.passed,true,JSON.stringify(result.findings.filter(f=>f.level!=='pass')));
  const leaked=built.files.map(f=>f.name==='word/document.xml'?{...f,text:f.text.replace('</w:body>','<w:p><w:r><w:t>Print: 1 per person, Letter, bring a pen.</w:t></w:r></w:p></w:body>')}:f);
  assert.equal(leakCheck({files:leaked,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser}).passed,false,'a print line that reaches the workbook is caught');
});

test('Update guide in the project and the plan know nothing new for a guide with no forms',()=>{
  const plan=formPlan([titleOf(),mod(1,'One')],{});
  assert.deepEqual([plan.edits,plan.issues],[[],[]]);
  assert.deepEqual(copiesToMake([titleOf()],{}),{rows:[],total:0,missing:[]});
});
