// Matching and rating statements: the two assessment items that are tables.
// The tables, check boxes and read-back are proven here against stand-ins for Word, not against Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import {makeContext,buildPackage} from '../src/guide-ooxml.mjs';
import {sectionBlock,itemBlock,EMPTY_BOX} from '../src/guide-assess-ooxml.mjs';
import {ITEM_TYPES,LIMITS,RATING_SCALES,parsePairs,parseLines,parseLabels,matchingChoices} from '../src/guide-items.mjs';
import {readItem,readAssessments,itemPlacement} from '../src/guide-model.mjs';
import {parse,partXml,paragraphsOf,packageToDocx} from './helpers/guide-package.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {pane,submit,started,WORD_QUIRKS,click,moveTo,startsWith,addModule,addLesson} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',W14='http://schemas.microsoft.com/office/word/2010/wordml';
const make=body=>{const ctx=makeContext(),xml=body(ctx);return {ctx,xml,pkg:buildPackage(xml,ctx,{})};};
const rowsOf=pkg=>paragraphsOf(pkg).map(p=>`${p.style}: ${p.text.trim()}`);
const tablesOf=pkg=>Array.from(parse(partXml(pkg,'/word/document.xml')).getElementsByTagNameNS(W,'tbl'));

const MATCH={type:'matching',number:1,stem:'Match each term with its meaning.',
  pairs:[{item:'Active listening',match:'Fully attending to the speaker'},{item:'Paraphrasing',match:'Restating in your own words'},{item:'Probing',match:'Asking follow-up questions'}],
  extras:['Interrupting to advise'],why:'These are the core skills.',points:3};
const RATE={type:'rating-statements',stem:'Trainer. Mark one box for each statement.',statements:['The trainer was prepared.','The trainer answered questions clearly.'],scale:'agree'};

test('matching is a question, a table with a header row, lettered choices in alphabetical order, and a key',()=>{
  const {ctx,pkg}=make(c=>itemBlock(MATCH,c));
  assert.deepEqual(rowsOf(pkg),[
    'OSG Question: Question 1. Match each term with its meaning. Write the letter of the best match for each item.',
    'OSG Table Head: No.','OSG Table Head: Item','OSG Table Head: Your match',
    'OSG Table Cell: 1','OSG Table Cell: Active listening','OSG Table Cell: Letter:',
    'OSG Table Cell: 2','OSG Table Cell: Paraphrasing','OSG Table Cell: Letter:',
    'OSG Table Cell: 3','OSG Table Cell: Probing','OSG Table Cell: Letter:',
    'OSG Item Note: Choices:',
    'OSG Option: Asking follow-up questions','OSG Option: Fully attending to the speaker','OSG Option: Interrupting to advise','OSG Option: Restating in your own words',
    'OSG Answer Key: Answer key (matching): 1 = b; 2 = d; 3 = a.',
    'OSG Answer Key: Why: These are the core skills.','OSG Answer Key: Points: 3'
  ]);
  assert.equal(ctx.lists,1,'the choices are a real lettered list that starts at a');
});

test('both tables are real tables: a header row Word knows about, no merged cells, no blank cells, a caption and a description',()=>{
  const {pkg}=make(c=>itemBlock(MATCH,c)+itemBlock(RATE,c));
  const tables=tablesOf(pkg);
  assert.equal(tables.length,2);
  for(const t of tables){
    const rows=Array.from(t.childNodes).filter(n=>n.localName==='tr');
    assert.equal(rows[0].getElementsByTagNameNS(W,'tblHeader').length,1,'the first row is marked as the header row');
    assert.equal(t.getElementsByTagNameNS(W,'tblHeader').length,1,'and only the first');
    assert.equal(t.getElementsByTagNameNS(W,'gridSpan').length+t.getElementsByTagNameNS(W,'vMerge').length+t.getElementsByTagNameNS(W,'hMerge').length,0,'no merged cells');
    assert.equal(t.getElementsByTagNameNS(W,'tbl').length,0,'no nested table');
    assert.equal(t.getElementsByTagNameNS(W,'tblCaption').length,1);assert.equal(t.getElementsByTagNameNS(W,'tblDescription').length,1);
    for(const tc of Array.from(t.getElementsByTagNameNS(W,'tc')))assert.notEqual(Array.from(tc.getElementsByTagNameNS(W,'t')).map(x=>x.textContent).join('').trim(),'','every cell holds words or a box');
    const widths=Array.from(t.getElementsByTagNameNS(W,'gridCol')).map(g=>Number(g.getAttributeNS(W,'w')));
    assert.equal(widths.reduce((a,b)=>a+b,0),9360,'the columns fill the text width');
    assert.equal(rows.every(r=>Array.from(r.childNodes).filter(n=>n.localName==='tc').length===widths.length),true,'every row has a cell for every column');
  }
});

test('matching: pairs can share a match, choices are listed once each, and the key follows the sorted choices',()=>{
  const {pkg}=make(c=>itemBlock({type:'matching',number:1,stem:'S',pairs:[{item:'One',match:'Yes'},{item:'Two',match:'no'},{item:'Three',match:'yes'}],extras:['NO','Maybe']},c));
  const rows=rowsOf(pkg);
  assert.deepEqual(rows.filter(r=>r.startsWith('OSG Option')),['OSG Option: Maybe','OSG Option: no','OSG Option: Yes']);
  assert.ok(rows.includes('OSG Answer Key: Answer key (matching): 1 = c; 2 = b; 3 = c.'));
  assert.deepEqual(matchingChoices([{match:'b'},{match:'A'}],['a']),['A','b']);
});

test('matching says what is missing instead of building half a table',()=>{
  const bad=(pairs,extras,pattern)=>assert.throws(()=>make(c=>itemBlock({...MATCH,pairs,extras},c)),pattern);
  bad([{item:'One',match:'x'}],[],/at least two pairs/);
  bad([{item:'One',match:'x'},{item:'Two',match:''}],[],/Line 2 is missing its match/);
  bad([{item:'',match:'x'},{item:'Two',match:'y'}],[],/Line 1 is missing the item/);
  bad(Array.from({length:LIMITS.pairs+1},(_,i)=>({item:`i${i}`,match:`m${i}`})),[],/up to 10 pairs/);
  bad(MATCH.pairs,['1','2','3','4','5','6'],/Up to 5 extra choices/);
  assert.throws(()=>make(c=>itemBlock({...MATCH,stem:'  '},c)),/Type the question first/);
});

test('rating statements are a table: one row to a statement and a real check box under each rating',()=>{
  const {pkg}=make(c=>itemBlock(RATE,c));
  assert.deepEqual(rowsOf(pkg).slice(0,8),[
    'OSG Question: Trainer. Mark one box for each statement.',
    'OSG Table Head: Statement','OSG Table Head: Strongly disagree','OSG Table Head: Disagree','OSG Table Head: Neutral','OSG Table Head: Agree','OSG Table Head: Strongly agree',
    'OSG Table Cell: The trainer was prepared.'
  ]);
  const rows=rowsOf(pkg);
  assert.equal(rows.filter(r=>r===`OSG Rating Box: ${EMPTY_BOX}`).length,10,'five boxes for each of the two statements');
  assert.equal(rows.some(r=>r.startsWith('OSG Answer')),false,'a survey has no key and no line to write on');
  const doc=parse(partXml(pkg,'/word/document.xml')),boxes=Array.from(doc.getElementsByTagNameNS(W14,'checkbox'));
  assert.equal(boxes.length,10);
  const sdts=Array.from(doc.getElementsByTagNameNS(W,'sdt')),ids=sdts.map(s=>s.getElementsByTagNameNS(W,'id')[0].getAttributeNS(W,'val'));
  assert.equal(new Set(ids).size,10,'each content control has its own number');
  assert.equal(sdts[0].getElementsByTagNameNS(W,'alias')[0].getAttributeNS(W,'val'),'The trainer was prepared.: Strongly disagree','a screen reader hears the statement and the rating');
  assert.equal(sdts[9].getElementsByTagNameNS(W,'alias')[0].getAttributeNS(W,'val'),'The trainer answered questions clearly.: Strongly agree');
  assert.equal(boxes[0].getElementsByTagNameNS(W14,'checked')[0].getAttributeNS(W14,'val'),'0','empty to begin with');
});

test('rating statements take a preset scale or the author\'s own labels, and check the labels',()=>{
  const labels=pkg=>rowsOf(pkg).filter(r=>r.startsWith('OSG Table Head')).map(r=>r.replace('OSG Table Head: ',''));
  assert.deepEqual(labels(make(c=>itemBlock({...RATE,scale:'success'},c)).pkg),['Statement',...RATING_SCALES[1].labels]);
  assert.deepEqual(labels(make(c=>itemBlock({...RATE,scale:'custom',labels:['Yes','Sometimes','No']},c)).pkg),['Statement','Yes','Sometimes','No']);
  const bad=(extra,pattern)=>assert.throws(()=>make(c=>itemBlock({...RATE,...extra},c)),pattern);
  bad({scale:'custom',labels:['Only one']},/between 2 and 7 labels/);
  bad({scale:'custom',labels:[]},/between 2 and 7 labels/);
  bad({scale:'custom',labels:Array.from({length:8},(_,i)=>`L${i}`)},/between 2 and 7 labels/);
  bad({statements:[]},/at least one statement/);
  bad({statements:Array.from({length:LIMITS.statements+1},(_,i)=>`S${i}`)},/up to 20 statements/);
  const wide=make(c=>itemBlock({...RATE,scale:'custom',labels:['1','2','3','4','5','6','7']},c));
  assert.equal(Array.from(tablesOf(wide.pkg)[0].getElementsByTagNameNS(W,'gridCol')).length,8,'seven ratings and the statement column');
});

test('pairs, statements and labels are read from what the author types, even pasted from a spreadsheet',()=>{
  assert.deepEqual(parsePairs('Term\tMeaning\nOther | Thing\n\n  Lone  '),[{item:'Term',match:'Meaning'},{item:'Other',match:'Thing'},{item:'Lone',match:''}]);
  assert.deepEqual(parsePairs('A | B | C'),[{item:'A',match:'B | C'}]);
  assert.deepEqual(parseLines(' one \r\n\r\n two '),['one','two']);
  assert.deepEqual(parseLabels('Yes, No,\n Maybe,'),['Yes','No','Maybe']);
});

// ---- reading an item back so it can be edited
const section=(...items)=>make(c=>sectionBlock('level-two')+items.map(i=>itemBlock(i,c)).join(''));
const questionAt=(paras,n=0)=>paras.map((p,i)=>p.style==='OSG Question'?i:-1).filter(i=>i>=0)[n];

test('a matching item reads back into the fields it was built from',()=>{
  const paras=paragraphsOf(section(MATCH).pkg),it=readItem(paras,questionAt(paras));
  assert.equal(it.type,'matching');assert.equal(it.stem,MATCH.stem);assert.equal(it.number,1);
  assert.deepEqual(it.pairs,MATCH.pairs);
  assert.deepEqual(it.extras,['Interrupting to advise']);
  assert.equal(it.why,'These are the core skills.');assert.equal(it.points,'3');
  assert.equal(it.label,'Question 1 (matching)');
  const {sections}=readAssessments(paras);
  assert.deepEqual(sections[0].items.map(i=>[i.type,i.label,i.hasKey]),[['matching','Question 1 (matching)',true]]);
});

test('a rating item reads back, with its preset scale or its own labels',()=>{
  for(const [extra,scale,labels] of [[{scale:'agree'},'agree',RATING_SCALES[0].labels],[{scale:'success'},'success',RATING_SCALES[1].labels],[{scale:'custom',labels:['Yes','No']},'custom',['Yes','No']]]){
    const paras=paragraphsOf(section({...RATE,...extra}).pkg),it=readItem(paras,questionAt(paras));
    assert.equal(it.type,'rating-statements');assert.equal(it.stem,RATE.stem);assert.equal(it.number,null);
    assert.deepEqual(it.statements,RATE.statements);assert.deepEqual(it.labels,[...labels]);assert.equal(it.scale,scale);
    assert.equal(it.label,'Likert scale');
  }
});

test('an item built from what read back is the same item: nothing is lost by editing',()=>{
  const paras=paragraphsOf(section(MATCH,RATE).pkg);
  for(const original of [MATCH,RATE]){
    const it=readItem(paras,questionAt(paras,original===MATCH?0:1));
    const again=make(c=>itemBlock({...it,number:it.number},c)),first=make(c=>itemBlock(original,c));
    assert.deepEqual(rowsOf(again.pkg),rowsOf(first.pkg));
  }
});

test('the whole table is part of its item: placement from any cell goes after the key, or after the last box',()=>{
  const paras=paragraphsOf(section(MATCH,RATE).pkg),q1=questionAt(paras,0),q2=questionAt(paras,1);
  const inMatch=paras.findIndex(p=>p.text==='Probing'),inRate=paras.findIndex(p=>p.text==='The trainer was prepared.');
  assert.equal(itemPlacement(paras,inMatch).anchor,q2-1,'after the matching item\'s key');
  assert.equal(itemPlacement(paras,inRate).anchor,paras.length-1,'after the last check box');
  assert.equal(itemPlacement(paras,inMatch).number,2,'numbered after the matching item');
  assert.equal(itemPlacement(paras,inRate).number,2,'a rating item is not numbered, so the next question is number 2');
  assert.ok(q1<q2);
});

test('a whole section with every kind of item passes the structure check, with a header on both tables',async()=>{
  const {pkg}=make(c=>sectionBlock('level-two')+itemBlock({type:'multiple-choice',number:1,stem:'Q?',options:['a','b'],correct:[0]},c)+itemBlock({...MATCH,number:2},c)+itemBlock({type:'true-false',number:3,stem:'T?',truth:true},c)+itemBlock(RATE,c)+itemBlock({type:'open-comment',stem:'More?'},c));
  const result=await checkDocx(await packageToDocx(pkg,'Every item'));
  assert.deepEqual(result.findings.filter(f=>f.level==='fail'),[]);
  const bad=result.findings.filter(f=>['TABLE-BLANK','TABLE-MERGED','TABLE-NESTED','EMPTY-PARAGRAPHS','TYPED-LIST'].includes(f.rule));
  assert.deepEqual(bad,[]);
  assert.equal(result.findings.filter(f=>f.rule==='TABLE-HEADER'&&f.level==='pass').length,2);
});

test('the item list knows both new types and keeps the survey ones last',()=>{
  const ids=ITEM_TYPES.map(t=>t.id);
  assert.ok(ids.includes('matching')&&ids.includes('rating-statements'));
  assert.equal(ITEM_TYPES.find(t=>t.id==='matching').group,'check');
  assert.equal(ITEM_TYPES.find(t=>t.id==='rating-statements').group,'survey');
});

// ---- the pane
const assessTab=a=>click(a,'guide-tab-assess');
const typeItem=(a,type,fields)=>{
  a.itemRadio(type).checked=true;a.itemRadio(type).onchange();
  for(const [id,v] of Object.entries(fields))a.n(id).value=v;
};
async function inSection(level){
  const a=await started('Tables',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await assessTab(a);await click(a,`guide-sec-${level}`);
  assert.equal(a.err(),'');
  return a;
}
const sectionRows=a=>a.rows().slice(a.rows().findIndex(startsWith('OSG Assessment')));
const PAIRS='Active listening | Fully attending to the speaker\nParaphrasing\tRestating in your own words\nProbing | Asking follow-up questions';
const addMatching=async a=>{typeItem(a,'matching',{'guide-item-stem':MATCH.stem,'guide-item-pairs':PAIRS,'guide-item-extras':'Interrupting to advise','guide-item-points':'3'});await submit(a,'guide-item-form');};

test('choosing matching or rating statements shows only the fields they need',async()=>{
  const a=await started('Fields',{quirks:WORD_QUIRKS});
  const shown=()=>['guide-item-options','guide-item-truefalse','guide-item-answer-box','guide-item-lines-box','guide-item-matching','guide-item-rating','guide-item-key-box'].filter(id=>!a.n(id).hidden);
  a.itemRadio('matching').onchange();
  assert.deepEqual(shown(),['guide-item-matching','guide-item-key-box']);
  assert.equal(a.n('guide-item-stem-label').textContent,'Question');
  a.itemRadio('rating-statements').onchange();
  assert.deepEqual(shown(),['guide-item-rating'],'a survey has no key');
  assert.equal(a.n('guide-item-stem-label').textContent,'Prompt');
  a.itemRadio('multiple-choice').onchange();
  assert.deepEqual(shown(),['guide-item-options','guide-item-key-box']);
});

test('matching goes into a Level Two section from one line to a pair',async()=>{
  const a=await inSection('level-two');
  await addMatching(a);
  assert.equal(a.err(),'');
  const rows=sectionRows(a);
  assert.ok(rows.includes('OSG Table Head: Your match'));
  assert.ok(rows.includes('OSG Table Cell: Paraphrasing'));
  assert.ok(rows.includes('OSG Answer Key: Answer key (matching): 1 = b; 2 = d; 3 = a.'));
  assert.equal(rows.at(-1),'OSG Answer Key: Points: 3','no stray empty paragraph after the item');
  assert.equal(a.n('guide-item-pairs').value,'','the form is cleared');
  assert.equal(a.n('guide-item-where').textContent,'Items go in: Level Two Assessment, after Question 1 (matching).');
});

test('rating statements go into a Level One survey, with the scale the author picked',async()=>{
  const a=await inSection('level-one');
  typeItem(a,'rating-statements',{'guide-item-stem':RATE.stem,'guide-item-statements':RATE.statements.join('\n')});
  a.n('guide-item-scale-success').checked=true;a.n('guide-item-scale-agree').checked=false;
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  const rows=sectionRows(a);
  assert.ok(rows.includes('OSG Table Head: Completely'));assert.equal(rows.includes('OSG Table Head: Strongly agree'),false);
  assert.equal(rows.filter(r=>r.startsWith('OSG Rating Box')).length,10);
  assert.equal(a.n('guide-item-where').textContent,'Items go in: Level One Assessment, after Likert scale.');
  assert.equal(a.n('guide-item-scale-agree').checked,true,'the scale goes back to the first choice for the next item');
  typeItem(a,'rating-statements',{'guide-item-stem':'Logistics.','guide-item-statements':'The room was comfortable.'});
  a.n('guide-item-scale-custom').checked=true;a.n('guide-item-scale-agree').checked=false;a.n('guide-item-labels').value='Yes, Sometimes, No';
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.ok(sectionRows(a).includes('OSG Table Head: Sometimes'));
});

test('a form that is not finished says what is missing, above Insert item, and inserts nothing',async()=>{
  const a=await inSection('level-two');
  const before=a.rows();
  typeItem(a,'matching',{'guide-item-stem':'Match.','guide-item-pairs':'Only one | pair'});
  await submit(a,'guide-item-form');
  assert.match(a.n('guide-item-problem').textContent,/at least two pairs/);
  assert.equal(a.n('guide-item-problem').hidden,false);assert.equal(a.n('guide-item-offer').hidden,true);assert.equal(a.n('guide-error').hidden,true);
  a.n('guide-item-pairs').value='One | a\nTwo';
  await submit(a,'guide-item-form');
  assert.match(a.n('guide-item-problem').textContent,/Line 2 is missing its match/);
  typeItem(a,'rating-statements',{'guide-item-stem':'Prompt.','guide-item-statements':''});
  await submit(a,'guide-item-form');
  assert.match(a.n('guide-item-problem').textContent,/at least one statement/);
  a.n('guide-item-statements').value='A statement.';a.n('guide-item-scale-custom').checked=true;a.n('guide-item-labels').value='Only one';
  await submit(a,'guide-item-form');
  assert.match(a.n('guide-item-problem').textContent,/between 2 and 7 labels/);
  assert.deepEqual(a.rows(),before,'nothing was inserted');
});

test('Edit this item brings a matching item back into the form, and Update item replaces the whole table',async()=>{
  const a=await inSection('level-two');
  await addMatching(a);
  await moveTo(a,r=>r==='OSG Table Cell: Probing');
  assert.equal(a.n('guide-assess-selected-text').textContent,'Your cursor is in Question 1 (matching).');
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-stem').value,MATCH.stem);
  assert.equal(a.n('guide-item-pairs').value,'Active listening | Fully attending to the speaker\nParaphrasing | Restating in your own words\nProbing | Asking follow-up questions');
  assert.equal(a.n('guide-item-extras').value,'Interrupting to advise');
  assert.equal(a.n('guide-item-points').value,'3');
  a.n('guide-item-pairs').value='Active listening | Fully attending to the speaker\nProbing | Asking follow-up questions';
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  const rows=sectionRows(a);
  assert.equal(rows.filter(r=>r==='OSG Table Head: No.').length,1,'one table, not two');
  assert.equal(rows.includes('OSG Table Cell: Paraphrasing'),false);
  assert.ok(rows.includes('OSG Answer Key: Answer key (matching): 1 = b; 2 = a.'));
  assert.equal(rows.filter(r=>r.startsWith('OSG Question')).length,1);
  assert.equal(a.adapter.undo(),true,'one Ctrl+Z puts the old item back');
  assert.ok(sectionRows(a).includes('OSG Table Cell: Paraphrasing'));
});

test('Edit this item brings a rating item back with its scale, and Update item replaces the whole table',async()=>{
  const a=await inSection('level-one');
  typeItem(a,'rating-statements',{'guide-item-stem':RATE.stem,'guide-item-statements':RATE.statements.join('\n')});
  await submit(a,'guide-item-form');
  await moveTo(a,r=>r==='OSG Rating Box: ☐');
  assert.equal(a.n('guide-assess-selected-text').textContent,'Your cursor is in Likert scale.');
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-statements').value,RATE.statements.join('\n'));
  assert.equal(a.n('guide-item-scale-agree').checked,true);
  a.n('guide-item-scale-custom').checked=true;a.n('guide-item-scale-agree').checked=false;a.n('guide-item-labels').value='Yes, No';
  a.n('guide-item-statements').value='Only this one.';
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  const rows=sectionRows(a);
  assert.equal(rows.filter(r=>r==='OSG Table Head: Statement').length,1);
  assert.deepEqual(rows.filter(r=>r.startsWith('OSG Table Head')),['OSG Table Head: Statement','OSG Table Head: Yes','OSG Table Head: No']);
  assert.equal(rows.filter(r=>r.startsWith('OSG Rating Box')).length,2);
  await moveTo(a,r=>r==='OSG Rating Box: ☐');
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-scale-custom').checked,true);assert.equal(a.n('guide-item-labels').value,'Yes, No');
});

test('the next question is numbered after a matching item, and a rating item does not use a number',async()=>{
  const a=await inSection('level-two');
  await addMatching(a);
  typeItem(a,'rating-statements',{'guide-item-stem':'Rate.','guide-item-statements':'One.'});
  await submit(a,'guide-item-form');
  typeItem(a,'true-false',{'guide-item-stem':'Is it?'});a.n('guide-item-true').checked=true;
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.deepEqual(sectionRows(a).filter(r=>r.startsWith('OSG Question')),[
    'OSG Question: Question 1. Match each term with its meaning. Write the letter of the best match for each item.','OSG Question: Rate.','OSG Question: Question 2. Is it?'
  ]);
});

test('an item added with the cursor in a table goes after the whole item, key and all',async()=>{
  const a=await inSection('level-two');
  await addMatching(a);
  await moveTo(a,r=>r==='OSG Table Cell: Probing');
  typeItem(a,'open-comment',{'guide-item-stem':'Comments?'});
  await submit(a,'guide-item-form');
  const rows=sectionRows(a),q=rows.findIndex(r=>r.startsWith('OSG Question: Comments'));
  assert.equal(rows[q-1],'OSG Answer Key: Points: 3');
});

test('Undo takes a whole matching item out, table and all',async()=>{
  const a=await inSection('level-two');
  const before=a.rows();
  await addMatching(a);
  assert.equal(a.adapter.undo(),true);
  await a.p.quietRefresh();
  assert.deepEqual(a.rows(),before);
});

test('Review counts nothing wrong with a section of tables, and Update guide leaves the tables alone',async()=>{
  const a=await inSection('level-two');
  await addMatching(a);
  typeItem(a,'rating-statements',{'guide-item-stem':'Rate.','guide-item-statements':'One.\nTwo.'});
  await submit(a,'guide-item-form');
  const before=a.rows();
  await click(a,'guide-update');
  assert.equal(a.err(),'');
  assert.deepEqual(a.rows().filter(r=>!r.startsWith('TOC')),before.filter(r=>!r.startsWith('TOC')));
});

test('the semantic differential\'s instructions come before the box the pairs are typed in, not after it (author, 2026-11-07)',async()=>{
  const {readFileSync}=await import('node:fs');
  const html=readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');
  const hint=html.indexOf('id="guide-item-sd-hint"'),box=html.indexOf('id="guide-item-sd-pairs"');
  assert.ok(hint>0&&box>0);
  assert.ok(hint<box,'the hint is above the text box');
  assert.match(html,/id="guide-item-sd-pairs" rows="5" aria-describedby="guide-item-sd-hint"/,'and still described by it, for a screen reader');
});
