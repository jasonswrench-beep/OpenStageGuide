// Ordering, performance task, semantic differential and action plan: the four item types added after the first nine.
// Proven against stand-ins for Word, not Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {makeContext,buildPackage} from '../src/guide-ooxml.mjs';
import {sectionBlock,itemBlock} from '../src/guide-assess-ooxml.mjs';
import {ITEM_TYPES,LIMITS,scrambleOrder,parseCriteria,ACTION_PROMPTS} from '../src/guide-items.mjs';
import {readItem,readAssessments} from '../src/guide-model.mjs';
import {readSource,buildWorkbook,zipWorkbook} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {readStyles,textOf} from '../src/guide-check.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {parse,partXml,paragraphsOf,packageToDocx} from './helpers/guide-package.mjs';
import {started,WORD_QUIRKS,click,submit,moveTo,addModule,addLesson} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const make=body=>{const ctx=makeContext(),xml=body(ctx);return {ctx,xml,pkg:buildPackage(xml,ctx,{})};};
const rowsOf=pkg=>paragraphsOf(pkg).map(p=>`${p.style}: ${p.text.trim()}`.trim());

const ORDER={type:'ordering',number:1,stem:'Put the steps of a good reframe in order.',orderItems:['Name the feeling','Describe the facts','Ask for a change'],why:'Feelings first, then facts.',points:3};
const PERF={type:'performance',number:2,stem:'Role-play a reframe with a partner for two minutes.',criteria:[{text:'Listens without interrupting',points:'2'},{text:'Restates the other person\'s point',points:'1'},{text:'Asks one neutral question',points:''}],points:3};
const SEMANTIC={type:'semantic-differential',stem:'Mark the point closest to how the session felt.',sdPairs:[{item:'Boring',match:'Interesting'},{item:'Rushed',match:'Well paced'}],sdPoints:5};
const ACTION={type:'action-plan',stem:'What will you do differently after this training?',actions:2};

test('the four new types are in the list, in the right groups',()=>{
  const by=id=>ITEM_TYPES.find(t=>t.id===id);
  assert.deepEqual(['ordering','performance'].map(id=>by(id).group),['check','check']);
  assert.deepEqual(['semantic-differential','action-plan'].map(id=>by(id).group),['survey','survey']);
  assert.ok(by('semantic-differential').noKey&&by('action-plan').noKey&&!by('ordering').noKey&&!by('performance').noKey);
  assert.equal(ITEM_TYPES.length,15,'the thirteen, numeric and the rubric');
});

// ---- ordering
test('ordering lists the items mixed up with letters, and the key gives the right order by letter',()=>{
  const {ctx,pkg}=make(c=>itemBlock(ORDER,c));
  assert.deepEqual(rowsOf(pkg),[
    'OSG Question: Question 1. Put the steps of a good reframe in order. Put the items in order by writing their letters, first to last, on the line.',
    'OSG Option: Ask for a change','OSG Option: Describe the facts','OSG Option: Name the feeling',
    'OSG Answer Line: Your order, by letter:',
    'OSG Answer Key: Answer key (ordering): c, b, a.','OSG Answer Key: Why: Feelings first, then facts.','OSG Answer Key: Points: 3'
  ]);
  assert.equal(ctx.lists,1,'a real lettered list');
});

test('ordering is never shown in the right order, and the mixed-up order is the same every time',()=>{
  assert.deepEqual(scrambleOrder(['b','a','c']),['a','b','c']);
  assert.deepEqual(scrambleOrder(['a','b','c']),['c','b','a'],'alphabetical is already the right order, so it is shown backwards');
  assert.deepEqual(scrambleOrder(['c','b','a']),['a','b','c']);
  const rows=items=>rowsOf(make(c=>itemBlock({type:'ordering',number:1,stem:'S',orderItems:items},c)).pkg);
  assert.ok(rows(['Alpha','Beta','Gamma']).includes('OSG Answer Key: Answer key (ordering): c, b, a.'));
  assert.deepEqual(rows(['Gamma','Alpha','Beta']),rows(['Gamma','Alpha','Beta']));
});

test('ordering says what is missing',()=>{
  const bad=(orderItems,pattern)=>assert.throws(()=>make(c=>itemBlock({...ORDER,orderItems},c)),pattern);
  bad(['One','Two'],/at least 3 items/);
  bad(Array.from({length:9},(_,i)=>`Item ${i}`),/up to 8 items/);
  bad(['Same','same','Other'],/Each item must be different/);
  assert.throws(()=>make(c=>itemBlock({...ORDER,stem:' '},c)),/Type the question first/);
});

// ---- performance task
test('a performance task gives the learner the task and lines, and the facilitator a checklist in the key',()=>{
  const {pkg}=make(c=>itemBlock(PERF,c));
  assert.deepEqual(rowsOf(pkg),[
    'OSG Question: Question 2. Role-play a reframe with a partner for two minutes.',
    'OSG Answer Line: Your response or notes:','OSG Answer Line:','OSG Answer Line:','OSG Answer Line:','OSG Answer Line:','OSG Answer Line:',
    'OSG Answer Key: Answer key (performance checklist): score each criterion the learner meets.',
    'OSG Answer Key: Criterion 1: Listens without interrupting (2 points)',
    'OSG Answer Key: Criterion 2: Restates the other person\'s point (1 point)',
    'OSG Answer Key: Criterion 3: Asks one neutral question',
    'OSG Answer Key: Points: 3'
  ]);
  assert.equal(rowsOf(make(c=>itemBlock({...PERF,lines:2},c)).pkg).filter(r=>r.startsWith('OSG Answer Line')).length,2);
});

test('a checklist line is the criterion and, after a bar, optional points; a bar that is not points is part of the words',()=>{
  assert.deepEqual(parseCriteria('Listens | 2\nSummarises\nTone | warm\n\nOne | 1.5'),[{text:'Listens',points:'2'},{text:'Summarises',points:''},{text:'Tone | warm',points:''},{text:'One',points:'1.5'}]);
  const bad=(criteria,pattern)=>assert.throws(()=>make(c=>itemBlock({...PERF,criteria},c)),pattern);
  bad([],/at least one criterion/);
  bad(Array.from({length:13},(_,i)=>({text:`c${i}`,points:''})),/up to 12 criteria/);
});

// ---- semantic differential
test('a semantic differential is a table: the left word, a check box for each point, the right word',()=>{
  const {pkg}=make(c=>itemBlock(SEMANTIC,c));
  assert.deepEqual(rowsOf(pkg).slice(0,13),[
    'OSG Question: Mark the point closest to how the session felt.',
    'OSG Table Head: Left word','OSG Table Head: 1','OSG Table Head: 2','OSG Table Head: 3','OSG Table Head: 4','OSG Table Head: 5','OSG Table Head: Right word',
    'OSG Table Cell: Boring','OSG Rating Box: ☐','OSG Rating Box: ☐','OSG Rating Box: ☐','OSG Rating Box: ☐'
  ]);
  const doc=parse(partXml(pkg,'/word/document.xml')),tbl=doc.getElementsByTagNameNS(W,'tbl')[0];
  assert.equal(tbl.getElementsByTagNameNS(W,'tblHeader').length,1);
  assert.equal(Array.from(tbl.getElementsByTagNameNS(W,'gridCol')).reduce((s,g)=>s+Number(g.getAttributeNS(W,'w')),0)<=9360,true);
  const aliases=Array.from(doc.getElementsByTagNameNS(W,'alias')).map(a=>a.getAttributeNS(W,'val'));
  assert.equal(aliases.length,10);assert.equal(aliases[0],'Boring to Interesting: 1 of 5');assert.equal(aliases[9],'Rushed to Well paced: 5 of 5');
  assert.equal(rowsOf(make(c=>itemBlock({...SEMANTIC,sdPoints:7},c)).pkg).filter(r=>r.startsWith('OSG Table Head')).length,9,'seven points: seven numbers and the two words');
});

test('a semantic differential says what is missing',()=>{
  const bad=(extra,pattern)=>assert.throws(()=>make(c=>itemBlock({...SEMANTIC,...extra},c)),pattern);
  bad({sdPairs:[]},/at least one pair of opposite words/);
  bad({sdPairs:[{item:'Boring',match:''}]},/Line 1 is missing the second word/);
  bad({sdPairs:[{item:'Fine',match:'Fine'},{item:'',match:'Warm'}]},/Line 2 is missing the first word/);
  bad({sdPairs:Array.from({length:13},(_,i)=>({item:`a${i}`,match:`b${i}`}))},/up to 12 pairs/);
  bad({sdPoints:6},/Choose a scale of 5 or 7 points/);
});

// ---- action plan
test('an action plan is labelled lines to write on, four for each action, and has no key',()=>{
  const {pkg}=make(c=>itemBlock(ACTION,c));
  assert.deepEqual(rowsOf(pkg),[
    'OSG Question: What will you do differently after this training?',
    'OSG Item Note: Action 1',...ACTION_PROMPTS.map(p=>`OSG Answer Line: ${p}`),
    'OSG Item Note: Action 2',...ACTION_PROMPTS.map(p=>`OSG Answer Line: ${p}`)
  ]);
  for(const actions of [0,6,1.5,NaN,'x'])assert.throws(()=>make(c=>itemBlock({...ACTION,actions},c)),/Choose how many actions to ask for: 1 to 5/);
});

// ---- reading back
const section=(...items)=>make(c=>sectionBlock('level-two')+items.map(i=>itemBlock(i,c)).join(''));
const questionAt=(paras,n=0)=>paras.map((p,i)=>p.style==='OSG Question'?i:-1).filter(i=>i>=0)[n];

test('each of the four reads back into the fields it was built from, and is told apart from the types that look like it',()=>{
  const paras=paragraphsOf(section(ORDER,PERF,SEMANTIC,ACTION,{type:'rating-statements',stem:'Rate.',statements:['One.'],scale:'agree'},{type:'open-comment',stem:'Anything else?'}).pkg);
  const types=readAssessments(paras).sections[0].items.map(i=>i.type);
  assert.deepEqual(types,['ordering','performance','semantic-differential','action-plan','rating-statements','open-comment']);
  const o=readItem(paras,questionAt(paras,0));
  assert.deepEqual([o.type,o.stem,o.orderItems,o.why,o.points,o.label],['ordering',ORDER.stem,ORDER.orderItems,ORDER.why,'3','Question 1 (ordering)']);
  const p=readItem(paras,questionAt(paras,1));
  assert.deepEqual([p.type,p.stem,p.criteria,p.points,p.lines],['performance',PERF.stem,PERF.criteria,'3',6]);
  const s=readItem(paras,questionAt(paras,2));
  assert.deepEqual([s.type,s.sdPoints,s.sdPairs,s.label],['semantic-differential',5,SEMANTIC.sdPairs,'Semantic differential']);
  const a=readItem(paras,questionAt(paras,3));
  assert.deepEqual([a.type,a.actions,a.stem],['action-plan',2,ACTION.stem]);
});

test('an item built from what was read back is the same item: nothing is lost by editing',()=>{
  const paras=paragraphsOf(section(ORDER,PERF,{...SEMANTIC,sdPoints:7},ACTION).pkg);
  [ORDER,PERF,{...SEMANTIC,sdPoints:7},ACTION].forEach((original,n)=>{
    const it=readItem(paras,questionAt(paras,n));
    assert.deepEqual(rowsOf(make(c=>itemBlock({...it,number:it.number},c)).pkg),rowsOf(make(c=>itemBlock(original,c)).pkg),original.type);
  });
});

// ---- the pane
const typeItem=(a,type,fields)=>{a.itemRadio(type).checked=true;a.itemRadio(type).onchange();for(const [id,v] of Object.entries(fields))a.n(id).value=v;};
async function inSection(level){
  const a=await started('More types',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await click(a,'guide-tab-assess');await click(a,`guide-sec-${level}`);
  assert.equal(a.err(),'');return a;
}
const sectionRows=a=>a.rows().slice(a.rows().findIndex(r=>r.startsWith('OSG Assessment')));

test('choosing each new type shows only the fields it needs',async()=>{
  const a=await started('Fields',{quirks:WORD_QUIRKS});
  const ids=['guide-item-options','guide-item-answer-box','guide-item-lines-box','guide-item-matching','guide-item-rating','guide-item-ordering','guide-item-performance','guide-item-semantic','guide-item-action','guide-item-key-box'];
  const shown=()=>ids.filter(id=>!a.n(id).hidden);
  a.itemRadio('ordering').onchange();assert.deepEqual(shown(),['guide-item-ordering','guide-item-key-box']);
  a.itemRadio('performance').onchange();assert.deepEqual(shown(),['guide-item-lines-box','guide-item-performance','guide-item-key-box']);
  assert.equal(a.n('guide-item-lines').value,'6');
  a.itemRadio('semantic-differential').onchange();assert.deepEqual(shown(),['guide-item-semantic']);
  a.itemRadio('action-plan').onchange();assert.deepEqual(shown(),['guide-item-action']);
});

test('ordering and a performance task go into a Level Two section from their forms, and edit in place',async()=>{
  const a=await inSection('level-two');
  typeItem(a,'ordering',{'guide-item-stem':ORDER.stem,'guide-item-order-items':ORDER.orderItems.join('\n')});
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  typeItem(a,'performance',{'guide-item-stem':PERF.stem,'guide-item-criteria':'Listens without interrupting | 2\nAsks one neutral question'});
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  const rows=sectionRows(a);
  assert.ok(rows.includes('OSG Answer Key: Answer key (ordering): c, b, a.'));
  assert.ok(rows.includes('OSG Answer Key: Criterion 1: Listens without interrupting (2 points)'));
  assert.equal(rows.filter(r=>r.startsWith('OSG Answer Line')).length,7,'one line for the order and six for the task');
  await moveTo(a,r=>r==='OSG Option: Name the feeling');
  assert.equal(a.n('guide-assess-selected-text').textContent,'Your cursor is in Question 1 (ordering).');
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-order-items').value,ORDER.orderItems.join('\n'),'back in the right order, not the mixed-up one');
  a.n('guide-item-order-items').value='Name the feeling\nDescribe the facts\nAsk for a change\nAgree a next step';
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  const after=sectionRows(a);
  assert.equal(after.filter(r=>r.startsWith('OSG Question: Question 1.')).length,1);
  assert.ok(after.includes('OSG Answer Key: Answer key (ordering): d, c, b, a.'),'four items: the first is now d, as the mixed-up list has four');
  await moveTo(a,r=>r==='OSG Answer Key: Criterion 2: Asks one neutral question');
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-criteria').value,'Listens without interrupting | 2\nAsks one neutral question');
});

test('a semantic differential and an action plan go into a Level One survey, and edit in place',async()=>{
  const a=await inSection('level-one');
  typeItem(a,'semantic-differential',{'guide-item-stem':SEMANTIC.stem,'guide-item-sd-pairs':'Boring | Interesting\nRushed\tWell paced'});
  a.n('guide-item-sd-7').checked=true;a.n('guide-item-sd-5').checked=false;
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  typeItem(a,'action-plan',{'guide-item-stem':ACTION.stem,'guide-item-actions':'2'});
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  const rows=sectionRows(a);
  assert.ok(rows.includes('OSG Table Head: 7'));assert.ok(rows.includes('OSG Item Note: Action 2'));
  assert.equal(rows.filter(r=>r==='OSG Rating Box: ☐').length,14,'two rows of seven boxes');
  await moveTo(a,r=>r==='OSG Table Cell: Boring');
  assert.equal(a.n('guide-assess-selected-text').textContent,'Your cursor is in Semantic differential.');
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-sd-pairs').value,'Boring | Interesting\nRushed | Well paced');assert.equal(a.n('guide-item-sd-7').checked,true);
  a.n('guide-item-sd-5').checked=true;a.n('guide-item-sd-7').checked=false;
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  assert.equal(sectionRows(a).filter(r=>r==='OSG Rating Box: ☐').length,10,'now five points, in one table');
  await moveTo(a,r=>r==='OSG Item Note: Action 1');
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-actions').value,'2');
  a.n('guide-item-actions').value='4';
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  assert.equal(sectionRows(a).filter(r=>r.startsWith('OSG Item Note: Action')).length,4);
});

test('a form that is not finished says what is missing, above Insert item',async()=>{
  const a=await inSection('level-two');
  const before=a.rows();
  typeItem(a,'ordering',{'guide-item-stem':'Order them.','guide-item-order-items':'Only\nTwo'});
  await submit(a,'guide-item-form');assert.match(a.n('guide-item-problem').textContent,/at least 3 items/);
  typeItem(a,'performance',{'guide-item-stem':'Do it.','guide-item-criteria':''});
  await submit(a,'guide-item-form');assert.match(a.n('guide-item-problem').textContent,/at least one criterion/);
  typeItem(a,'semantic-differential',{'guide-item-stem':'Mark.','guide-item-sd-pairs':'Boring'});
  await submit(a,'guide-item-form');assert.match(a.n('guide-item-problem').textContent,/Line 1 is missing the second word/);
  typeItem(a,'action-plan',{'guide-item-stem':'Plan.','guide-item-actions':'9'});
  await submit(a,'guide-item-form');assert.match(a.n('guide-item-problem').textContent,/Choose how many actions to ask for: 1 to 5/);
  assert.deepEqual(a.rows(),before);
});

// ---- the whole section, the structure check and the learner workbook
test('a section with all four passes the structure check, with the new table headed',async()=>{
  const {pkg}=make(c=>sectionBlock('level-two')+itemBlock(ORDER,c)+itemBlock(PERF,c)+itemBlock(SEMANTIC,c)+itemBlock(ACTION,c));
  const result=await checkDocx(await packageToDocx(pkg,'More types'));
  assert.deepEqual(result.findings.filter(f=>f.level==='fail'),[]);
  assert.deepEqual(result.findings.filter(f=>['TABLE-BLANK','TABLE-MERGED','EMPTY-PARAGRAPHS','TYPED-LIST'].includes(f.rule)),[]);
  assert.equal(result.findings.filter(f=>f.rule==='TABLE-HEADER'&&f.level==='pass').length,1);
});

test('in the learner workbook the order and the criteria stay in the key, and the learner gets the list, the table and the lines',async()=>{
  const a=await inSection('level-two');
  typeItem(a,'ordering',{'guide-item-stem':ORDER.stem,'guide-item-order-items':ORDER.orderItems.join('\n'),'guide-item-why':'Feelings first, then facts.'});
  await submit(a,'guide-item-form');
  typeItem(a,'performance',{'guide-item-stem':PERF.stem,'guide-item-criteria':'Listens without interrupting | 2\nRestates the other person\'s point'});
  await submit(a,'guide-item-form');
  await click(a,'guide-sec-level-one');
  typeItem(a,'semantic-differential',{'guide-item-stem':SEMANTIC.stem,'guide-item-sd-pairs':'Boring | Interesting'});
  await submit(a,'guide-item-form');
  typeItem(a,'action-plan',{'guide-item-stem':ACTION.stem,'guide-item-actions':'1'});
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64')),docXml=await zip.file('word/document.xml').async('string'),stylesXml=await zip.file('word/styles.xml').async('string');
  const styles=readStyles(parse(stylesXml));
  const rows=Array.from(parse(docXml).getElementsByTagNameNS(W,'p')).map(p=>`${styles.get(Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val')||'Normal')?.name}: ${textOf(p)}`.trim());
  for(const gone of ['Answer key','Criterion','Listens without interrupting','Feelings first'])assert.ok(!rows.some(r=>r.includes(gone)),`${gone} is not in the workbook`);
  assert.ok(rows.includes('OSG Option: Ask for a change')&&rows.includes('OSG Answer Line: Your order, by letter:'),'the mixed-up list and its line');
  assert.ok(rows.includes('OSG Answer Line: Your response or notes:'));
  assert.ok(rows.includes('OSG Table Head: Left word')&&rows.includes('OSG Rating Box: ☐'));
  assert.ok(rows.includes('OSG Item Note: Action 1')&&rows.includes('OSG Answer Line: By when:'));
  assert.equal(a.n('guide-workbook-problem').hidden,true);
});
