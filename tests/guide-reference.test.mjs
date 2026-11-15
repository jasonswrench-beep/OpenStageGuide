// Reference pages (kit slice K4): a glossary and a list of references kept in alphabetical order, figure captions numbered by Update guide, pull quotes and the licence page.
// It drives the real pane and the real taskpane.html ids in memory, and checks what reaches the participant workbook.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {started,submit,addModule,addLesson,addBlock,moveTo,startsWith,click} from './helpers/guide-harness.mjs';
import {planTerm,planReference,captionLine,quoteBlock,licenceBody,findLicence,captionEdits,readGlossary,LICENCES,GLOSSARY_HEADING} from '../src/guide-reference.mjs';
import {readSource,buildWorkbook} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {STYLE_NAMES} from '../src/guide-styles.mjs';

const texts=async(a,style)=>(await a.adapter.paragraphs()).filter(p=>p.style===style).map(p=>p.text.trim());
async function guide(){
  const a=await started('Reference guide');
  await addModule(a,'One');await addLesson(a,'First','10');await addBlock(a,'ask','What makes it hard?');
  return a;
}
const open=async a=>{a.p.useTool('referencePages');await a.p.refresh();};
const term=async(a,t,m)=>{a.type('guide-ref-term',t);a.type('guide-ref-meaning',m);await submit(a,'guide-ref-term-form');};

test('the glossary is started at the end of the guide and its terms are kept in alphabetical order, whatever order they are added in',async()=>{
  const a=await guide();await open(a);
  await term(a,'Facilitator','Someone who guides a group.');assert.equal(a.err(),'');
  await term(a,'Agenda','The plan for the time.');
  await term(a,'debrief','Talking about what happened.');
  await term(a,'Zone','A place.');
  assert.deepEqual(await texts(a,STYLE_NAMES.term),['Agenda: The plan for the time.','debrief: Talking about what happens.'.replace('happens','happened'),'Facilitator: Someone who guides a group.','Zone: A place.'],'sorted without regard to case');
  const rows=(await a.adapter.paragraphs()).map(p=>`${p.style}: ${p.text.trim()}`);
  assert.equal(rows.filter(r=>r==='heading 1: Glossary').length,1,'one Glossary heading');
  assert.ok(rows.findIndex(r=>r==='heading 1: Glossary')<rows.findIndex(r=>r.startsWith('OSG Term: Agenda')),'terms follow the heading');
  assert.match(a.n('guide-status').textContent,/"Zone" is in the glossary\./);
  assert.equal(a.n('guide-ref-term').value,'','the boxes are empty for the next term');
});

test('a term that is already there, an empty one, and one with a colon are refused with their reason',async()=>{
  const a=await guide();await open(a);
  await term(a,'Agenda','The plan.');
  await term(a,'AGENDA','Another plan.');assert.match(a.err(),/already has "Agenda"/);
  await term(a,'','Words.');assert.match(a.err(),/Type the term first/);
  await term(a,'Plan','');assert.match(a.err(),/Type what it means first/);
  await term(a,'Time: of day','Words.');assert.match(a.err(),/cannot hold a colon/);
  assert.equal((await texts(a,STYLE_NAMES.term)).length,1);
});

test('references go into their own list in alphabetical order, and the same reference twice is refused',async()=>{
  const a=await guide();await open(a);
  const ref=async t=>{a.type('guide-ref-text',t);await submit(a,'guide-ref-form');};
  await ref('Wrench, J. (2020). Communication. Press.');await ref('Adams, B. (2019). Listening. House.');await ref('Mills, T. (2021). Groups. Press.');
  assert.deepEqual(await texts(a,STYLE_NAMES.reference),['Adams, B. (2019). Listening. House.','Mills, T. (2021). Groups. Press.','Wrench, J. (2020). Communication. Press.']);
  await ref('adams, b. (2019). listening. house.');assert.match(a.err(),/already hold that one/);
});

test('a caption goes below the cursor, and Update guide numbers captions by module and keeps them numbered after a move',async()=>{
  const a=await guide();await open(a);
  await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));await a.p.refresh();
  a.type('guide-cap-text','A group at a flip chart');a.type('guide-cap-source','Photo by the author');
  await submit(a,'guide-cap-form');assert.equal(a.err(),'');
  assert.deepEqual(await texts(a,STYLE_NAMES.caption),['Figure: A group at a flip chart. Source: Photo by the author.']);
  await addModule(a,'Two');await addLesson(a,'Second','5');
  await moveTo(a,startsWith('OSG Lesson: Lesson 2.1'));
  a.type('guide-cap-text','A second figure');await submit(a,'guide-cap-form');
  const edits=captionEdits(await a.adapter.paragraphs());
  assert.deepEqual(edits.map(e=>e.text),['Figure 1.1: A group at a flip chart. Source: Photo by the author.','Figure 2.1: A second figure.']);
  await click(a,'guide-update');
  assert.deepEqual(await texts(a,STYLE_NAMES.caption),['Figure 1.1: A group at a flip chart. Source: Photo by the author.','Figure 2.1: A second figure.']);
  assert.deepEqual(captionEdits(await a.adapter.paragraphs()),[],'a second Update has nothing left to do');
  assert.throws(()=>captionLine({caption:''}),/Type the caption first/);
});

test('a pull quote is the words in quotation marks and who said them, in two styles',async()=>{
  const a=await guide();await open(a);
  await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));await a.p.refresh();
  a.type('guide-quote-text','"Listening is the first skill."');a.type('guide-quote-by','A teacher');
  await submit(a,'guide-quote-form');assert.equal(a.err(),'');
  assert.deepEqual(await texts(a,STYLE_NAMES.quote),['“Listening is the first skill.”']);
  assert.deepEqual(await texts(a,STYLE_NAMES.quoteBy),['A teacher']);
  a.type('guide-quote-text','Some words');a.type('guide-quote-by','');await submit(a,'guide-quote-form');assert.match(a.err(),/Type who said it first/);
});

test('the licence page is built at the end, built again over itself, and says what the author said',async()=>{
  const a=await guide();await open(a);
  assert.match(a.n('guide-licence-note').textContent,/not built yet/);
  assert.deepEqual(a.n('guide-licence-kind').children.map(o=>o.textContent),LICENCES.map(l=>l.label));
  a.type('guide-licence-kind','cc-by-nc-sa-4.0');a.type('guide-licence-holder','Jason Wrench');a.type('guide-licence-year','2026');
  a.type('guide-licence-reviewed','15 November 2026');a.type('guide-licence-place','the state of New York, USA');
  await submit(a,'guide-licence-form');assert.equal(a.err(),'');
  const rows=(await a.adapter.paragraphs()).map(p=>p.text.trim());
  const at=rows.indexOf('Licence and permissions');
  assert.ok(at>0);
  assert.match(rows[at+1],/^© 2026 Jason Wrench\. This guide is licensed under CC BY-NC-SA 4\.0 \(https:\/\/creativecommons\.org\/licenses\/by-nc-sa\/4\.0\/\)\./);
  assert.equal(rows[at+2],'Content that dates (figures, laws, guidance) was last reviewed on 15 November 2026.');
  assert.equal(rows[at+3],'Where the law applies: the state of New York, USA.');
  assert.match(a.n('guide-licence-note').textContent,/in the guide\. Building it again replaces it/);
  a.type('guide-licence-kind','all-rights');a.type('guide-licence-reviewed','');a.type('guide-licence-place','');
  await submit(a,'guide-licence-form');
  const after=(await a.adapter.paragraphs()).map(p=>p.text.trim());
  assert.equal(after.filter(t=>t==='Licence and permissions').length,1,'replaced, not added again');
  assert.match(after[after.indexOf('Licence and permissions')+1],/All rights reserved\./);
  assert.ok(!after.some(t=>/last reviewed on/.test(t)),'the lines that are no longer wanted are gone');
  a.type('guide-licence-holder','');await submit(a,'guide-licence-form');assert.match(a.err(),/Type who holds the rights first/);
  a.type('guide-licence-kind','custom');a.type('guide-licence-words','');await submit(a,'guide-licence-form');assert.match(a.err(),/Type the words of the licence or permission first/);
});

test('the pure builders say what to do for a bad year, an unknown licence and a section with no heading',()=>{
  assert.throws(()=>licenceBody({licence:'cc-by-4.0',holder:'A',year:'26'}),/four digits/);
  assert.throws(()=>licenceBody({licence:'nope',holder:'A',year:'2026'}),/Choose a licence/);
  assert.equal(findLicence([{style:'Normal',text:'x'}]),null);
  assert.equal(readGlossary([{style:'Normal',text:'x'}]),null);
  assert.equal(planTerm([{style:'Normal',text:'x'}],'A','B').created,true);
  assert.equal(planReference([{style:'Normal',text:'x'}],'A. (2020). B.').after,0);
  assert.match(quoteBlock({quote:'Words',by:'Me'}),/OSGQuote/);
  assert.equal(GLOSSARY_HEADING,'Glossary');
});

test('the reference pages reach the participant workbook, in styles the workbook defines, and the leak check still passes',async()=>{
  const a=await guide();await open(a);
  await term(a,'Agenda','The plan for the time.');
  a.type('guide-ref-text','Adams, B. (2019). Listening. House.');await submit(a,'guide-ref-form');
  a.type('guide-licence-kind','cc-by-4.0');a.type('guide-licence-holder','The author');a.type('guide-licence-year','2026');await submit(a,'guide-licence-form');
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},{DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const doc=built.files.find(f=>f.name==='word/document.xml').text,styles=built.files.find(f=>f.name==='word/styles.xml').text;
  assert.ok(doc.includes('The plan for the time.')&&doc.includes('Adams, B. (2019). Listening. House.'),'glossary and references are in');
  assert.ok(styles.includes('OSGTerm')&&styles.includes('OSGReference'),'and their styles are defined in the workbook');
  const leak=leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  assert.equal(leak.passed,true,JSON.stringify(leak.findings.filter(f=>f.level!=='pass')));
});
