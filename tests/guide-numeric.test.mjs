// The numeric item: a question answered with a number, with how close an answer can be and its units in the key.
// Proven against stand-ins for Word, not Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {makeContext,buildPackage} from '../src/guide-ooxml.mjs';
import {itemBlock} from '../src/guide-assess-ooxml.mjs';
import {numericKey,readNumericKey,ITEM_TYPES} from '../src/guide-items.mjs';
import {readItem,readAssessments} from '../src/guide-model.mjs';
import {textOf} from '../src/guide-check.mjs';
import {parse,paragraphsOf} from './helpers/guide-package.mjs';
import {started,WORD_QUIRKS,click,submit,moveTo,addModule,addLesson,startsWith} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const make=item=>{const ctx=makeContext();return buildPackage(itemBlock(item,ctx),ctx,{});};
const rowsOf=pkg=>paragraphsOf(pkg).map(p=>`${p.style}: ${p.text.trim()}`.trim());
const BASE={type:'numeric',number:3,stem:'A car goes 120 km in 2 hours. What is its average speed?',answer:'60'};

test('numeric is one of the types, in the quiz group, with a key',()=>{
  const t=ITEM_TYPES.find(x=>x.id==='numeric');
  assert.deepEqual([t.group,t.keyLabel,!!t.noKey],['check','numeric',false]);
});

test('a numeric item is the question, a line to write the number on, and a key with the number',()=>{
  assert.deepEqual(rowsOf(make(BASE)),[
    'OSG Question: Question 3. A car goes 120 km in 2 hours. What is its average speed?',
    'OSG Answer Line: Your answer:',
    'OSG Answer Key: Answer key (numeric): 60'
  ]);
});

test('how close an answer can be and the units go in the key, and the units are told to the learner on the answer line',()=>{
  assert.deepEqual(rowsOf(make({...BASE,answer:'60.5',tolerance:'0.5',units:'km per hour',why:'120 divided by 2.',points:2})),[
    'OSG Question: Question 3. A car goes 120 km in 2 hours. What is its average speed?',
    'OSG Answer Line: Your answer in km per hour:',
    'OSG Answer Key: Answer key (numeric): 60.5 (plus or minus 0.5) km per hour',
    'OSG Answer Key: Why: 120 divided by 2.',
    'OSG Answer Key: Points: 2'
  ]);
  assert.equal(numericKey({answer:'-3',units:'  degrees   C '}),'-3 degrees C');
});

test('the key is read back into its number, tolerance and units, whatever is left out',()=>{
  assert.deepEqual(readNumericKey('60.5 (plus or minus 0.5) km per hour'),{answer:'60.5',tolerance:'0.5',units:'km per hour'});
  assert.deepEqual(readNumericKey('60'),{answer:'60',tolerance:'',units:''});
  assert.deepEqual(readNumericKey('-3 degrees C'),{answer:'-3',tolerance:'',units:'degrees C'});
  assert.deepEqual(readNumericKey('42 (plus or minus 1)'),{answer:'42',tolerance:'1',units:''});
  assert.deepEqual(readNumericKey('about forty'),{answer:'about forty',tolerance:'',units:''},'a key typed by hand is kept as it is');
  for(const k of ['60','60.5 (plus or minus 0.5) km per hour','-3 degrees C','42 (plus or minus 1)'])assert.equal(numericKey(readNumericKey(k)),k);
});

test('the builder says what is wrong with the number, the tolerance or the units',()=>{
  const bad=(patch,pattern)=>assert.throws(()=>make({...BASE,...patch}),pattern);
  bad({answer:''},/Type the correct number for the key/);
  bad({answer:'1,200'},/without commas or spaces, like 42 or 3\.5/);
  bad({answer:'sixty'},/must be a number/);
  bad({tolerance:'-1'},/not negative/);
  bad({tolerance:'a bit'},/must be a number that is not negative/);
  bad({units:'x'.repeat(31)},/up to 30 characters/);
  assert.throws(()=>make({...BASE,stem:' '}),/Type the question first/);
});

test('a numeric item reads back for editing, and is told from short answer by its key',()=>{
  const ctx=makeContext(),xml=itemBlock({...BASE,answer:'60.5',tolerance:'0.5',units:'km per hour',points:2},ctx);
  const rows=paragraphsOf(buildPackage(xml,ctx,{}));
  const it=readItem(rows,0);
  assert.deepEqual([it.type,it.answer,it.tolerance,it.units,it.points,it.number],['numeric','60.5','0.5','km per hour','2',3]);
  assert.equal(it.stem,BASE.stem);
});

// ---- the form
const typeNumeric=(a,fields)=>{
  a.itemRadio('numeric').checked=true;a.itemRadio('numeric').onchange();
  for(const [id,v] of Object.entries(fields))a.n(id).value=v;
};
async function inSection(){
  const a=await started('Numeric',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await click(a,'guide-tab-assess');await click(a,'guide-sec-level-two');
  assert.equal(a.err(),'');return a;
}

test('choosing numeric shows the number, how close and the units, and nothing else it does not need',async()=>{
  const a=await started('Fields',{quirks:WORD_QUIRKS});
  const ids=['guide-item-options','guide-item-truefalse','guide-item-answer-box','guide-item-numeric','guide-item-lines-box','guide-item-matching','guide-item-key-box'];
  const shown=()=>ids.filter(id=>!a.n(id).hidden);
  a.itemRadio('numeric').onchange();
  assert.deepEqual(shown(),['guide-item-answer-box','guide-item-numeric','guide-item-key-box']);
  assert.equal(a.n('guide-item-answer-label').textContent,'Correct number');
  a.itemRadio('short-answer').onchange();
  assert.equal(a.n('guide-item-numeric').hidden,true);
});

test('Insert item writes the numeric item and clears the form, and Edit this item opens it again with its parts',async()=>{
  const a=await inSection();
  typeNumeric(a,{'guide-item-stem':BASE.stem,'guide-item-answer':'60.5','guide-item-tolerance':'0.5','guide-item-units':'km per hour','guide-item-points':'2'});
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  const at=a.rows().findIndex(r=>r.startsWith('OSG Question'));
  assert.deepEqual(a.rows().slice(at,at+4),[
    'OSG Question: Question 1. A car goes 120 km in 2 hours. What is its average speed?',
    'OSG Answer Line: Your answer in km per hour:',
    'OSG Answer Key: Answer key (numeric): 60.5 (plus or minus 0.5) km per hour',
    'OSG Answer Key: Points: 2'
  ]);
  for(const id of ['guide-item-answer','guide-item-tolerance','guide-item-units'])assert.equal(a.n(id).value,'');
  await moveTo(a,startsWith('OSG Question'));
  await click(a,'guide-assess-edit');
  assert.deepEqual(['guide-item-answer','guide-item-tolerance','guide-item-units'].map(id=>a.n(id).value),['60.5','0.5','km per hour']);
  a.n('guide-item-answer').value='61';a.n('guide-item-tolerance').value='';
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Question')).length,1,'replaced, not added');
  assert.ok(a.rows().includes('OSG Answer Key: Answer key (numeric): 61 km per hour'));
});

test('a bad number is refused at the form with the reason, and nothing is inserted',async()=>{
  const a=await inSection();
  const before=a.rows();
  typeNumeric(a,{'guide-item-stem':'How many?','guide-item-answer':'1,200'});
  await submit(a,'guide-item-form');
  assert.match(a.err(),/without commas or spaces/);
  assert.deepEqual(a.rows(),before);
});

test('the numeric key never reaches the learner workbook, and the line to write the number on does',async()=>{
  const a=await inSection();
  typeNumeric(a,{'guide-item-stem':BASE.stem,'guide-item-answer':'60.5','guide-item-tolerance':'0.5','guide-item-units':'km per hour'});
  await submit(a,'guide-item-form');
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true,'the leak check passes');
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64')),doc=parse(await zip.file('word/document.xml').async('string'));
  const text=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>textOf(p)).join('\n');
  assert.ok(text.includes('What is its average speed?')&&text.includes('Your answer in km per hour:'));
  assert.ok(!text.includes('Answer key')&&!text.includes('plus or minus'));
});

test('a numeric item counts in its section like any other, and the review has nothing to fix',async()=>{
  const a=await inSection();
  typeNumeric(a,{'guide-item-stem':'First?','guide-item-answer':'1'});await submit(a,'guide-item-form');
  typeNumeric(a,{'guide-item-stem':'Second?','guide-item-answer':'2'});await submit(a,'guide-item-form');
  const items=readAssessments(await a.adapter.paragraphs()).sections[0].items;
  assert.deepEqual(items.map(i=>[i.number,i.type,i.hasKey]),[[1,'numeric',true],[2,'numeric',true]]);
  assert.equal(items[0].label,'Question 1 (numeric)');
  await click(a,'guide-open-review');
  assert.deepEqual(a.n('guide-review-fix').children.map(li=>li.textContent),['Nothing to fix.']);
});
