// A form printed landscape (its print line says so) is a section of its own in a Word file: the portrait pages end before it, and are portrait again after it.
import test from 'node:test';
import assert from 'node:assert/strict';
import {landscapeFlags,formSpans} from '../src/guide-forms.mjs';
import {bodyWithContents,withSection} from '../src/guide-docx.mjs';
import {STYLE_NAMES} from '../src/guide-styles.mjs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readSource,buildWorkbook} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {formTitleBlock,formText,printLine} from '../src/guide-forms-ooxml.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const P=(style,text)=>({style,text,xml:`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`});
const guide=(...parts)=>parts.flat();
const form=(title,print,...more)=>[P(STYLE_NAMES.formTitle,title),...(print?[P(STYLE_NAMES.formPrint,print)]:[]),P(STYLE_NAMES.formText,'Write here.'),...more];
const front=[P(STYLE_NAMES.module,'Module 1: One'),P(STYLE_NAMES.lesson,'Lesson 1.1: First'),P(STYLE_NAMES.cue,'Say Hello')];
const count=(s,re)=>(s.match(re)||[]).length;

test('the extent of a form runs from its title to what ends it, and its print line gives its orientation',()=>{
  const paras=guide(front,form('Handout 1.1: Goals','Print: 1 per person, Letter, landscape.'),P(STYLE_NAMES.cue,'Say Next'),form('Handout 1.2: Plans','Print: 2 per team, A4.'));
  const spans=formSpans(paras);
  assert.deepEqual(spans.map(s=>[s.index,s.end,s.orientation]),[[3,5,'landscape'],[7,9,null]]);
  assert.deepEqual(landscapeFlags(paras),[false,false,false,true,true,true,false,false,false,false]);
});

test('a landscape form in the middle of a guide is a section of its own: a portrait section ends before it and a landscape one after it',()=>{
  const paras=guide(front,form('Handout 1.1: Goals','Print: 1 per person, Letter, landscape.'),P(STYLE_NAMES.cue,'Say Next'));
  const body=withSection(bodyWithContents(paras));
  assert.equal(count(body,/<w:sectPr>/g),3,'two section breaks and the page setup of the end');
  assert.equal(count(body,/w:orient="landscape"/g),1);
  const at=s=>body.indexOf(s);
  // portrait break, then the form, then the landscape break, then the block that follows, then the portrait setup of the end
  assert.ok(at('Say Hello')<at('<w:pgSz w:w="12240" w:h="15840"/>'),'the portrait break comes after the blocks before the form');
  assert.ok(at('<w:pgSz w:w="12240" w:h="15840"/>')<at('Handout 1.1: Goals'),'and before the form');
  assert.ok(at('Write here.')<at('w:orient="landscape"'),'the landscape break follows the form');
  assert.ok(at('w:orient="landscape"')<at('Say Next'),'and comes before the block after it');
  assert.match(body,/<\/w:sectPr>$/,'the page setup of the end is last');
  assert.doesNotMatch(body.slice(body.lastIndexOf('<w:sectPr>')),/orient/,'and it is portrait');
});

test('a guide that ends in a landscape form ends in a landscape section, with no empty portrait page after it',()=>{
  const paras=guide(front,form('Handout 1.1: Goals','Print: 1 per person, Letter, landscape.'));
  const body=withSection(bodyWithContents(paras));
  assert.equal(count(body,/<w:sectPr>/g),2);
  assert.match(body,/w:orient="landscape"\/>[^]*<\/w:sectPr>$/,'the last page setup is landscape');
});

test('a guide with no landscape form has no section break at all, as before',()=>{
  const paras=guide(front,form('Handout 1.1: Goals','Print: 1 per person, Letter, portrait.'),P(STYLE_NAMES.cue,'Say Next'));
  const body=bodyWithContents(paras);
  assert.doesNotMatch(body,/sectPr/);
  assert.equal(count(withSection(body),/<w:sectPr>/g),1);
});

test('two landscape forms in a row share one landscape section; a pack may say it with flags of its own and start with a landscape form',()=>{
  const two=guide(front,form('Handout 1.1: A','Print: 1 per person, landscape.'),form('Handout 1.2: B','Print: 1 per person, landscape.'),P(STYLE_NAMES.cue,'Say Next'));
  assert.equal(count(withSection(bodyWithContents(two)),/<w:sectPr>/g),3,'one portrait break, one landscape break, the end');
  const flagged=[{...P(STYLE_NAMES.formTitle,'Handout A: One'),landscape:true},{...P(STYLE_NAMES.formText,'Text'),landscape:true},P(STYLE_NAMES.formTitle,'Handout B: Two')];
  const body=withSection(bodyWithContents(flagged));
  assert.equal(count(body,/<w:sectPr>/g),2,'no portrait break before a first landscape form, one landscape break after it');
  assert.ok(body.indexOf('w:orient="landscape"')>body.indexOf('Text')&&body.indexOf('w:orient="landscape"')<body.indexOf('Handout B'));
});

// ---- the participant workbook
async function workbookWith(...forms){
  const a=await started('Orientation guide');
  await addModule(a,'One');await addLesson(a,'First');await addBlock(a,'ask','What makes it hard?');
  let at=a.adapter.snapshot().findIndex(p=>p.text.includes('What makes it hard?'));
  for(const [id,title,orientation] of forms){
    await a.adapter.insert(()=>formTitleBlock({id,label:'Handout',number:'',title})+printLine({copies:{n:1,per:'person'},orientation})+formText(`Body of ${title}`),{where:'after',after:at,spare:false});
    at=a.adapter.snapshot().findIndex(p=>p.text.includes(`Body of ${title}`));
  }
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},{DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const file=name=>built.files.find(f=>f.name===name)?.text;
  return {built,file,doc:file('word/document.xml')};
}

test('a landscape form in the workbook is a section of its own with a footer made for the wider page, and the cover keeps its own first page',async()=>{
  const {built,file,doc}=await workbookWith(['f1','Goals','portrait'],['f2','Wide plan','landscape']);
  assert.equal((doc.match(/<w:sectPr>/g)||[]).length,2,'a break before the landscape form, and the page setup of the end, which is landscape because the guide ends in that form');
  assert.equal((doc.match(/w:orient="landscape"/g)||[]).length,1);
  assert.ok(file('word/footer2.xml'),'a footer for the landscape section');
  assert.match(file('word/_rels/document.xml.rels'),/Id="rIdFooterWide"[^>]*Target="footer2\.xml"/);
  assert.match(file('[Content_Types].xml'),/PartName="\/word\/footer2\.xml"/);
  const wide=/w:tab w:val="right" w:pos="(\d+)"/.exec(file('word/footer2.xml'))[1],narrow=/w:tab w:val="right" w:pos="(\d+)"/.exec(file('word/footer1.xml'))[1];
  assert.ok(Number(wide)>Number(narrow),'the page number of the landscape footer sits further right');
  const first=doc.indexOf('<w:sectPr>');
  assert.match(doc.slice(first,doc.indexOf('</w:sectPr>',first)),/<w:titlePg\/>/,'the first section still has no footer on the cover');
  assert.doesNotMatch(doc.slice(doc.lastIndexOf('<w:sectPr>')),/<w:titlePg\/>/,'and the last one has a footer on its first page');
  const leak=leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  assert.equal(leak.passed,true,JSON.stringify(leak.findings.filter(f=>f.level!=='pass')));
});

test('portrait pages return after a landscape form, and a workbook with no landscape form is the one it always was',async()=>{
  const {file,doc}=await workbookWith(['f1','Wide plan','landscape'],['f2','Goals','portrait']);
  assert.equal((doc.match(/<w:sectPr>/g)||[]).length,3);
  assert.equal((doc.match(/w:orient="landscape"/g)||[]).length,1,'only the own section of the form');
  assert.doesNotMatch(doc.slice(doc.lastIndexOf('<w:sectPr>')),/orient/,'the end of the guide is portrait');
  const plain=await workbookWith(['f1','Goals','portrait']);
  assert.equal((plain.doc.match(/<w:sectPr>/g)||[]).length,1);
  assert.equal(plain.file('word/footer2.xml'),undefined);
  assert.doesNotMatch(plain.file('word/_rels/document.xml.rels'),/rIdFooterWide/);
  void file;
});
