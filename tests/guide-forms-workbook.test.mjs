// Forms in the learner workbook (spec slice F2): a form for everyone is kept, whole, with its tables and check boxes; a form for the
// facilitator is left out and its words are watched for by the leak check. Runs against stand-ins for Word, like the rest of the workbook tests.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {classify,summarize,readSource,buildWorkbook} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {readStyles,textOf} from '../src/guide-check.mjs';
import {para,run} from '../src/guide-ooxml.mjs';
import {STYLE_NAMES,STYLE_IDS} from '../src/guide-styles.mjs';
import {formTitleBlock,formText,writingLines,fieldRow,checkList,formTable,ratingGrid,sourceLine,citationControl} from '../src/guide-forms-ooxml.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml={DOMParserClass:DOMParser,SerializerClass:XMLSerializer};
const text=(files,name)=>files.find(f=>f.name===name)?.text;
const parseXml=t=>new DOMParser().parseFromString(t,'application/xml');
const outParas=files=>{
  const styles=readStyles(parseXml(text(files,'word/styles.xml')));
  return Array.from(parseXml(text(files,'word/document.xml')).getElementsByTagNameNS(W,'p')).map(p=>{
    const id=Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val')||'Normal';
    return {style:styles.get(id)?.name||id,text:textOf(p).trim()};
  });
};

// ---- who a form is for, decided from the paragraphs alone
const P=(style,text='x')=>({style,text});
const guide=[P('OSG Guide Title','G'),P('OSG Module','Module 1: One'),P('OSG Lesson','Lesson 1.1: First')];

test('a form for everyone is kept, with its text, lines, fields, source line and tables; its audience comes from its title',()=>{
  const {decisions}=classify([...guide,P('OSG Form Title','Handout 1.1: Goals'),P('OSG Form Text','Read.'),P('OSG Form Field','Name:'),P('OSG Form Line',''),P('OSG Form Box',''),
    P('OSG Table Head','Done'),P('OSG Table Cell','Plan'),P('OSG Rating Box','x'),P('OSG Form Source','Credit.')]);
  assert.ok(decisions.slice(3).every(d=>d.keep&&d.reason==='form'),'everything in the form is kept, an empty line to write on included');
});

test('a form for the facilitator is left out whole, and what follows it that is not part of it is judged on its own',()=>{
  const {decisions}=classify([...guide,P('OSG Form Title Facilitator','Checklist 1.2: Set-up'),P('OSG Form Text','Secret.'),P('OSG Table Head','Done'),P('OSG Table Cell','Move'),P('OSG Rating Box','x'),
    P('OSG Cue','Ask  Why?'),P('OSG Form Text','Strays after a block')]);
  assert.deepEqual(decisions.slice(3,8).map(d=>[d.keep,d.reason]),[[false,'facilitator-form'],[false,'facilitator-form'],[false,'facilitator-form'],[false,'facilitator-form'],[false,'facilitator-form']]);
  assert.deepEqual([decisions[8].keep,decisions[8].reason],[true,'block'],'a block for everyone after it is kept as ever');
  assert.deepEqual([decisions[9].keep,decisions[9].reason],[false,'own'],'form text with no form above it is the author\'s own text, left out and listed');
});

test('a table whose first heading reads Lesson is a form table, not a module overview, when it sits in a form',()=>{
  const {decisions}=classify([...guide,P('OSG Form Title','Handout 1.1: Plan'),P('OSG Table Head','Lesson'),P('OSG Table Cell','One')]);
  assert.deepEqual(decisions.slice(4).map(d=>d.reason),['form','form']);
});

test('a module or lesson ends a form; so does an assessment',()=>{
  const {decisions}=classify([...guide,P('OSG Form Title','Handout 1.1: A'),P('OSG Lesson','Lesson 1.2: Second'),P('OSG Table Cell','Not in a form')]);
  assert.equal(decisions[4].reason,'heading');
  assert.deepEqual([decisions[5].keep,decisions[5].reason],[false,'own'],'a table cell outside any form and any section is left out');
});

test('the summary counts forms for everyone and forms left for the facilitator',()=>{
  const s=summarize([...guide,P('OSG Form Title','Handout 1.1: A'),P('OSG Form Text','x'),P('OSG Form Title Facilitator','Checklist 1.2: B'),P('OSG Form Text','y'),P('OSG Cue','Ask  Why?')]);
  assert.equal(s.forms,1);
  assert.equal(s.left.facilitatorForms,1);
});

// ---- the file that is made
async function guideWithForms(){
  const a=await started('Forms guide');
  await addModule(a,'One');await addLesson(a,'First');
  await addBlock(a,'ask','What makes it hard?');
  const at=a.adapter.snapshot().findIndex(p=>p.text.includes('What makes it hard?'));
  // A line under the Ask block that cites the handout, as the Handout cue will write it.
  await a.adapter.insert(ctx=>para(STYLE_IDS.body,run('Give out ')+citationControl({id:'f1',text:'Handout 1.1: Goals'},ctx)+run(' now.')),{where:'after',after:at,spare:false});
  await a.adapter.insert(ctx=>
    formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})+formText('Read this first.')+fieldRow([{label:'Name',width:'full'},{label:'Date',width:'short'}])
    +writingLines({label:'Your notes',count:2})+checkList({items:['Read it','Try it'],columns:['When'],ctx})
    +formTable({columns:[{name:'Goal',weight:3},{name:'Done',kind:'check'}],rows:[['Plan ahead']],ctx,caption:'Goals'})
    +ratingGrid({statements:['I plan ahead.'],evidence:true,ctx})+sourceLine('Adapted with permission.')
    +formTitleBlock({id:'f2',label:'Checklist',number:'1.2',title:'Room set-up',facilitator:true})+formText('SECRETPLANNING words for the facilitator.')+checkList({items:['Move the SECRETTABLES'],ctx,caption:'Set-up'})
    ,{where:'after',after:at+1,spare:false});
  return a;
}

test('the workbook carries a form for everyone whole, flattens a citation to plain words, and leaves the facilitator\'s form out',async()=>{
  const a=await guideWithForms();
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},xml);
  const doc=text(built.files,'word/document.xml'),rows=outParas(built.files).map(p=>`${p.style}: ${p.text}`);
  assert.ok(rows.includes('OSG Form Title: Handout 1.1: Goals'));
  assert.ok(rows.includes('OSG Form Text: Read this first.'));
  assert.ok(rows.some(r=>r.startsWith('OSG Form Field: Name')));
  assert.equal(rows.filter(r=>r.startsWith('OSG Form Line')).length,2);
  assert.ok(rows.includes('OSG Form Source: Adapted with permission.'));
  assert.ok(rows.includes('OSG Table Cell: Plan ahead'));
  assert.equal((doc.match(/<w14:checkbox/g)||[]).length,2+1+5,'checklist, the table\'s check column and the rating grid keep their real check boxes');
  assert.equal((doc.match(/<w:tblHeader/g)||[]).length,3,'each learner table keeps the header row Word knows about');
  assert.ok(rows.some(r=>/Give out Handout 1\.1: Goals now\./.test(r)),'a citation keeps its words');
  assert.ok(!/osg-ref|osg_form_/.test(doc),'the tag and the bookmark are not in the workbook');
  for(const gone of ['SECRETPLANNING','SECRETTABLES','Room set-up','OSGFormTitleFacilitator'])assert.ok(!doc.includes(gone),`${gone} is not in the workbook`);
  assert.ok(built.facilitatorTexts.some(t=>t.includes('SECRETPLANNING')),'the facilitator form\'s words are watched for');
  assert.ok(built.facilitatorTexts.some(t=>t.includes('Room set-up')));
  const styles=text(built.files,'word/styles.xml');
  assert.ok(styles.includes('OSGFormLine')&&styles.includes('OSGFormTitle'),'the workbook defines the form styles it uses');
  assert.ok(!styles.includes('OSGFormTitleFacilitator'),'and not the facilitator one');
  const result=leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  assert.equal(result.passed,true,JSON.stringify(result.findings.filter(f=>f.level!=='pass')));
});

test('the leak check catches a facilitator form that got into the workbook',async()=>{
  const a=await guideWithForms();
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},xml);
  const leaked=built.files.map(f=>f.name==='word/document.xml'?{...f,text:f.text.replace('</w:body>',`<w:p><w:pPr><w:pStyle w:val="OSGFormTitleFacilitator"/></w:pPr><w:r><w:t>Checklist 1.2: Room set-up</w:t></w:r></w:p></w:body>`)}:f);
  const result=leakCheck({files:leaked,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  assert.equal(result.passed,false,'a paragraph in the facilitator form title style, or its words, fails the check');
});

test('a guide whose only learner content is a form can make a workbook',async()=>{
  const a=await started('Only forms');
  await addModule(a,'One');await addLesson(a,'First');
  const at=a.adapter.snapshot().findIndex(p=>p.text.includes('Lesson 1.1'));
  await a.adapter.insert(ctx=>formTitleBlock({id:'f1',label:'Worksheet',number:'1.1',title:'Plans'})+writingLines({count:3}),{where:'after',after:at+1,spare:false});
  const s=summarize(await a.adapter.paragraphs());
  assert.equal(s.forms,1);
  const pkg=await a.adapter.sourcePackage();
  assert.doesNotThrow(()=>buildWorkbook(readSource(pkg,{DOMParserClass:DOMParser}),{},xml));
});
