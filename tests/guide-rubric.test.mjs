// The performance task with a rubric: a table in the key, a row for each criterion and a column for each level.
// Proven against stand-ins for Word, not Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {makeContext,buildPackage} from '../src/guide-ooxml.mjs';
import {itemBlock} from '../src/guide-assess-ooxml.mjs';
import {DEFAULT_RUBRIC_LEVELS,parseRubricRows,ITEM_TYPES,LIMITS} from '../src/guide-items.mjs';
import {readItem,readAssessments,kindOf} from '../src/guide-model.mjs';
import {summarize} from '../src/guide-workbook.mjs';
import {textOf} from '../src/guide-check.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {parse,paragraphsOf,packageToDocx} from './helpers/guide-package.mjs';
import {started,WORD_QUIRKS,click,submit,moveTo,addModule,addLesson,startsWith} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const LEVELS=[{text:'Exceeds',points:'3'},{text:'Meets',points:'2'},{text:'Approaching',points:'1'}];
const ROWS=[
  ['Listens','Never interrupts','Interrupts once or twice','Often interrupts'],
  ['Restates the point','Restates it accurately','Restates part of it','Does not restate it']
];
const RUBRIC={type:'rubric',number:2,stem:'Role-play a reframe with a partner for two minutes.',levels:LEVELS,rubricRows:ROWS,points:6};
const make=item=>{const ctx=makeContext();return {ctx,xml:itemBlock(item,ctx),pkg:buildPackage(itemBlock(item,makeContext()),ctx,{})};};
const rowsOf=pkg=>paragraphsOf(pkg).map(p=>`${p.style}: ${p.text.trim()}`.trim());

test('the rubric type is in the list, in the quiz group, and has a key',()=>{
  const t=ITEM_TYPES.find(x=>x.id==='rubric');
  assert.deepEqual([t.group,t.keyLabel,!!t.noKey,t.lines],['check','performance rubric',false,6]);
  assert.equal(DEFAULT_RUBRIC_LEVELS.split('\n').length,3);
});

test('the learner gets the task and lines to write on; the key holds the whole rubric as a table',()=>{
  const rows=rowsOf(make(RUBRIC).pkg);
  assert.deepEqual(rows,[
    'OSG Question: Question 2. Role-play a reframe with a partner for two minutes.',
    'OSG Answer Line: Your response or notes:',...Array(5).fill('OSG Answer Line:'),
    'OSG Answer Key: Answer key (performance rubric): score each criterion by the level the learner reaches.',
    'OSG Key Table Head: Criterion','OSG Key Table Head: Exceeds (3 points)','OSG Key Table Head: Meets (2 points)','OSG Key Table Head: Approaching (1 point)',
    'OSG Key Table Cell: Listens','OSG Key Table Cell: Never interrupts','OSG Key Table Cell: Interrupts once or twice','OSG Key Table Cell: Often interrupts',
    'OSG Key Table Cell: Restates the point','OSG Key Table Cell: Restates it accurately','OSG Key Table Cell: Restates part of it','OSG Key Table Cell: Does not restate it',
    'OSG Answer Key: Points: 6'
  ]);
});

test('the table has a header row, a caption and description, no merged cells and no empty cell',()=>{
  const {xml}=make(RUBRIC);
  assert.match(xml,/<w:tblHeader\/>/);
  assert.match(xml,/<w:tblCaption w:val="Rubric: Role-play a reframe/);
  assert.match(xml,/<w:tblDescription w:val="[^"]*Exceeds, Meets, Approaching/);
  assert.doesNotMatch(xml,/gridSpan|vMerge/);
  const doc=parse(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`);
  const cells=Array.from(doc.getElementsByTagNameNS(W,'tc'));
  assert.equal(cells.length,4*3);
  assert.ok(cells.every(c=>textOf(c).trim()));
});

test('levels can have no points, and two to five levels are allowed',()=>{
  const two=rowsOf(make({...RUBRIC,levels:[{text:'Yes',points:''},{text:'No',points:''}],rubricRows:[['Listens','Yes it does','No it does not']]}).pkg);
  assert.ok(two.includes('OSG Key Table Head: Yes')&&two.includes('OSG Key Table Head: No'));
  const five=Array.from({length:5},(_,i)=>({text:`Level ${i+1}`,points:''}));
  assert.doesNotThrow(()=>make({...RUBRIC,levels:five,rubricRows:[['Listens','a','b','c','d','e']]}));
});

test('the builder says what to fix: the levels, the points, the criteria and each line',()=>{
  const bad=(patch,pattern)=>assert.throws(()=>make({...RUBRIC,...patch}),pattern);
  bad({levels:[{text:'Only one',points:''}],rubricRows:[['A','x']]},/2 to 5 levels/);
  bad({levels:Array.from({length:6},(_,i)=>({text:`L${i}`,points:''})),rubricRows:[['A','1','2','3','4','5','6']]},/2 to 5 levels/);
  bad({levels:[{text:'Same',points:''},{text:'same',points:''}],rubricRows:[['A','x','y']]},/different name/);
  bad({levels:[{text:'Meets',points:'lots'},{text:'Not yet',points:''}],rubricRows:[['A','x','y']]},/points for "Meets" must be a number/);
  bad({rubricRows:[]},/Type at least one criterion/);
  bad({rubricRows:Array.from({length:9},(_,i)=>[`C${i}`,'a','b','c'])},/up to 8 criteria/);
  bad({rubricRows:[['Listens','Never interrupts','Sometimes interrupts']]},/Line 1 has 2 descriptions and the rubric has 3 levels/);
  bad({rubricRows:[['Listens','Never','','Often']]},/Line 1 has an empty description for "Meets"/);
  bad({rubricRows:[['','Never','Sometimes','Often']]},/Line 1 has an empty criterion/);
  bad({stem:' '},/Type the question first/);
  assert.equal(LIMITS.rubricRows,8);
});

test('the rubric reads back for editing, and is told from the checklist by its key',()=>{
  const {pkg}=make(RUBRIC);
  const it=readItem(paragraphsOf(pkg),0);
  assert.deepEqual([it.type,it.number,it.points,it.lines],['rubric',2,'6',6]);
  assert.deepEqual(it.levels,LEVELS);
  assert.deepEqual(it.rubricRows,ROWS);
  assert.equal(it.label,'Question 2 (performance rubric)');
  assert.deepEqual(parseRubricRows('A | b | c\n\nD|e|f'),[['A','b','c'],['D','e','f']]);
  assert.equal(kindOf({style:'OSG Key Table Cell'}),'key-table-cell');
});

// ---- the form
const typeRubric=(a,fields)=>{a.itemRadio('rubric').checked=true;a.itemRadio('rubric').onchange();for(const [id,v] of Object.entries(fields))a.n(id).value=v;};
async function inSection(){
  const a=await started('Rubric',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await click(a,'guide-tab-assess');await click(a,'guide-sec-level-two');
  assert.equal(a.err(),'');return a;
}
const ROWS_TEXT=ROWS.map(r=>r.join(' | ')).join('\n');

test('choosing the rubric shows the levels with a starting set, the criteria and the lines, and nothing it does not need',async()=>{
  const a=await started('Fields',{quirks:WORD_QUIRKS});
  const ids=['guide-item-options','guide-item-answer-box','guide-item-lines-box','guide-item-performance','guide-item-rubric','guide-item-key-box'];
  a.itemRadio('rubric').onchange();
  assert.deepEqual(ids.filter(id=>!a.n(id).hidden),['guide-item-lines-box','guide-item-rubric','guide-item-key-box']);
  assert.equal(a.n('guide-item-rubric-levels').value,DEFAULT_RUBRIC_LEVELS);
  a.itemRadio('performance').onchange();
  assert.equal(a.n('guide-item-rubric').hidden,true);
});

test('Insert item writes the rubric, and Edit this item opens it again; Update item replaces the whole table once',async()=>{
  const a=await inSection();
  typeRubric(a,{'guide-item-stem':RUBRIC.stem,'guide-item-rubric-rows':ROWS_TEXT,'guide-item-points':'6'});
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Key Table Head')).length,4);
  assert.ok(a.rows().includes('OSG Key Table Cell: Never interrupts'));
  assert.equal(a.n('guide-item-rubric-rows').value,'');assert.equal(a.n('guide-item-rubric-levels').value,DEFAULT_RUBRIC_LEVELS,'ready for the next one');
  await moveTo(a,startsWith('OSG Question'));
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-rubric-levels').value,'Exceeds | 3\nMeets | 2\nApproaching | 1');
  assert.equal(a.n('guide-item-rubric-rows').value,ROWS_TEXT);
  a.n('guide-item-rubric-rows').value='Listens | Never interrupts | Sometimes | Often';
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Key Table Head: Criterion')).length,1,'replaced, not added');
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Key Table Cell')).length,4);
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Question')).length,1);
});

test('a line with the wrong number of descriptions is refused at the form, and the form keeps what was typed',async()=>{
  const a=await inSection();
  const before=a.rows();
  typeRubric(a,{'guide-item-stem':'Do it.','guide-item-rubric-rows':'Listens | Never | Often'});
  await submit(a,'guide-item-form');
  assert.match(a.err(),/Line 1 has 2 descriptions and the rubric has 3 levels/);
  assert.deepEqual(a.rows(),before);
  assert.equal(a.n('guide-item-rubric-rows').value,'Listens | Never | Often');
});

// ---- the workbook, the review, the structure check
test('the workbook has the task and the lines, and never the rubric: every cell of it is key',async()=>{
  const a=await inSection();
  typeRubric(a,{'guide-item-stem':RUBRIC.stem,'guide-item-rubric-rows':ROWS_TEXT});
  await submit(a,'guide-item-form');
  const s=summarize(await a.adapter.paragraphs());
  assert.equal(s.left.keys,1,'the key, its table and its lines are one run');assert.deepEqual(s.leftOut,[]);
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true,'the leak check passes');
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64')),doc=parse(await zip.file('word/document.xml').async('string'));
  const text=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>textOf(p)).join('\n');
  assert.ok(text.includes('Role-play a reframe')&&text.includes('Your response or notes:'));
  for(const gone of ['Never interrupts','Exceeds','Approaching','Restates it accurately','performance rubric'])assert.ok(!text.includes(gone),`${gone} is not in the workbook`);
});

test('a section with a rubric passes the review and the structure check, the table headed',async()=>{
  const a=await inSection();
  typeRubric(a,{'guide-item-stem':RUBRIC.stem,'guide-item-rubric-rows':ROWS_TEXT});
  await submit(a,'guide-item-form');
  await click(a,'guide-open-review');
  assert.deepEqual(a.n('guide-review-fix').children.map(li=>li.textContent),['Nothing to fix.']);
  const result=await checkDocx(await packageToDocx(await a.adapter.sourcePackage(),'Rubric'));
  assert.deepEqual(result.findings.filter(f=>f.level==='fail'),[]);
  assert.deepEqual(result.findings.filter(f=>['TABLE-BLANK','TABLE-MERGED','EMPTY-PARAGRAPHS','TYPED-LIST'].includes(f.rule)&&f.level!=='pass'),[]);
  assert.equal(result.findings.filter(f=>f.rule==='TABLE-HEADER'&&f.level==='pass').length,1);
  const items=readAssessments(await a.adapter.paragraphs()).sections[0].items;
  assert.deepEqual(items.map(i=>[i.number,i.type,i.hasKey]),[[1,'rubric',true]]);
});

// ---- the rubric learners can see
test('with "Show the rubric to learners" the table goes under the task in the ordinary table styles, and the key keeps only its line',()=>{
  const rows=rowsOf(make({...RUBRIC,learnerRubric:true}).pkg);
  assert.deepEqual(rows,[
    'OSG Question: Question 2. Role-play a reframe with a partner for two minutes.',
    'OSG Table Head: Criterion','OSG Table Head: Exceeds (3 points)','OSG Table Head: Meets (2 points)','OSG Table Head: Approaching (1 point)',
    'OSG Table Cell: Listens','OSG Table Cell: Never interrupts','OSG Table Cell: Interrupts once or twice','OSG Table Cell: Often interrupts',
    'OSG Table Cell: Restates the point','OSG Table Cell: Restates it accurately','OSG Table Cell: Restates part of it','OSG Table Cell: Does not restate it',
    'OSG Answer Line: Your response or notes:',...Array(5).fill('OSG Answer Line:'),
    'OSG Answer Key: Answer key (performance rubric): score each criterion by the level the learner reaches.',
    'OSG Answer Key: Points: 6'
  ]);
  const {xml}=make({...RUBRIC,learnerRubric:true});
  assert.match(xml,/<w:tblHeader\/>/);
  assert.doesNotMatch(xml,/facilitator's key/);
});

test('either version reads back with its own setting, and the same levels and rows',()=>{
  for(const learnerRubric of [false,true]){
    const it=readItem(paragraphsOf(make({...RUBRIC,learnerRubric}).pkg),0);
    assert.deepEqual([it.type,it.learnerRubric,it.levels,it.rubricRows],['rubric',learnerRubric,LEVELS,ROWS],`learnerRubric ${learnerRubric}`);
  }
});

test('ticking the box writes the rubric where learners see it; Edit this item opens it ticked, and unticking moves it back to the key',async()=>{
  const a=await inSection();
  typeRubric(a,{'guide-item-stem':RUBRIC.stem,'guide-item-rubric-rows':ROWS_TEXT});
  a.n('guide-item-rubric-learners').checked=true;
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.ok(a.rows().includes('OSG Table Head: Exceeds (3 points)')&&!a.rows().some(r=>r.startsWith('OSG Key Table')));
  assert.equal(a.n('guide-item-rubric-learners').checked,false,'the form is ready for the next item');
  await moveTo(a,startsWith('OSG Question'));
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-rubric-learners').checked,true);
  a.n('guide-item-rubric-learners').checked=false;
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.ok(a.rows().includes('OSG Key Table Head: Exceeds (3 points)'));
  assert.ok(!a.rows().includes('OSG Table Head: Exceeds (3 points)'),'moved, not copied');
  assert.equal(a.rows().filter(r=>r.endsWith('Criterion')).length,1);
});

test('a rubric learners can see is in the workbook, the key line is not, and the leak check passes',async()=>{
  const a=await inSection();
  typeRubric(a,{'guide-item-stem':RUBRIC.stem,'guide-item-rubric-rows':ROWS_TEXT});
  a.n('guide-item-rubric-learners').checked=true;
  await submit(a,'guide-item-form');
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true);
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64')),doc=parse(await zip.file('word/document.xml').async('string'));
  const text=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>textOf(p)).join('\n');
  for(const kept of ['Role-play a reframe','Exceeds (3 points)','Never interrupts','Does not restate it','Your response or notes:'])assert.ok(text.includes(kept),`${kept} is in the workbook`);
  assert.ok(!text.includes('performance rubric')&&!text.includes('Answer key'));
  await click(a,'guide-tab-guide');await click(a,'guide-open-review');
  assert.deepEqual(a.n('guide-review-fix').children.map(li=>li.textContent),['Nothing to fix.']);
});
