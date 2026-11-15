import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {DOMParser} from '@xmldom/xmldom';
import {makeContext,buildPackage} from '../src/guide-ooxml.mjs';
import {sectionBlock,itemBlock} from '../src/guide-assess-ooxml.mjs';
import {ITEM_TYPES,LEVELS} from '../src/guide-items.mjs';
import {readItem,readAssessments,itemPlacement,sectionPlacement,updatePlan,guideIssues,describeAt} from '../src/guide-model.mjs';
import {checkParts} from '../src/guide-check.mjs';
import {NS,stylesPart} from '../src/guide-styles.mjs';
import {paragraphsOf,partXml,packageToDocx,TINY} from './helpers/guide-package.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {focusedId} from './helpers/app-dom.mjs';
import {pane,submit,started,WORD_QUIRKS,click,startsWith,moveTo,addModule,addLesson,addBlock,guideRows} from './helpers/guide-harness.mjs';

const build=body=>{const ctx=makeContext();return {ctx,xml:body(ctx)};};
const rowsOf=body=>{const {ctx,xml}=build(body);return {ctx,rows:paragraphsOf(buildPackage(xml,ctx,{})).map(p=>`${p.style}: ${p.text.trim()}`)};};

test('a multiple-choice item is a question, a lettered list, an answer line and a key in its own style',()=>{
  const {ctx,rows}=rowsOf(c=>itemBlock({type:'multiple-choice',number:1,stem:'Which step comes first?',options:['Name the feeling','Describe the facts','Ask for a change'],correct:[1],why:'Facts keep the first words neutral.',points:1,objective:'Reframe a disagreement.'},c));
  assert.deepEqual(rows,[
    'OSG Question: Question 1. Which step comes first?',
    'OSG Option: Name the feeling','OSG Option: Describe the facts','OSG Option: Ask for a change',
    'OSG Answer Line: Your answer:',
    'OSG Answer Key: Answer key (multiple choice): b.',
    'OSG Answer Key: Why: Facts keep the first words neutral.',
    'OSG Answer Key: Points: 1',
    'OSG Answer Key: Objective: Reframe a disagreement.'
  ]);
  assert.equal(ctx.lists,1,'one lettered list');
});

test('the options are a real Word list that starts again at a for every item, never typed letters',()=>{
  const {ctx,xml}=build(c=>itemBlock({type:'multiple-choice',number:1,stem:'A?',options:['x','y'],correct:[0]},c)+itemBlock({type:'multiple-response',number:2,stem:'B?',options:['p','q','r'],correct:[0,2]},c));
  const pkg=buildPackage(xml,ctx,{}),numbering=partXml(pkg,'/word/numbering.xml');
  assert.equal(ctx.lists,2);
  assert.match(numbering,/<w:numFmt w:val="lowerLetter"\/>/);
  assert.equal([...numbering.matchAll(/<w:num w:numId="\d+">/g)].length,2);
  assert.equal([...numbering.matchAll(/<w:startOverride w:val="1"\/>/g)].length,2);
  assert.match(partXml(pkg,'/word/_rels/document.xml.rels'),/relationships\/numbering" Target="numbering\.xml"/);
  assert.equal([...xml.matchAll(/<w:numId w:val="1"\/>/g)].length,2);
  assert.equal([...xml.matchAll(/<w:numId w:val="2"\/>/g)].length,3);
  assert.doesNotMatch(xml,/>\s*[a-z]\.\s/,'no typed a. b. c.');
});

test('content with no lettered list carries no numbering part',()=>{
  const {ctx,xml}=build(c=>sectionBlock('level-two'));
  assert.doesNotMatch(buildPackage(xml,ctx,{}),/numbering\.xml/);
});

test('each type builds the paragraphs it should, with the key naming the type',()=>{
  const cases={
    'multiple-response':{item:{stem:'Pick the helpful ones.',options:['a1','b1','c1','d1'],correct:[0,2]},expect:['OSG Question: Question 3. Pick the helpful ones. Choose all that apply.','OSG Option: a1','OSG Option: b1','OSG Option: c1','OSG Option: d1','OSG Answer Line: Your answers:','OSG Answer Key: Answer key (multiple response): a, c.']},
    'true-false':{item:{stem:'Tone changes meaning.',truth:true},expect:['OSG Question: Question 3. Tone changes meaning.','OSG Answer Line: True or false:','OSG Answer Key: Answer key (true or false): True.']},
    'fill-blank':{item:{stem:'The first step is to [blank] the facts.',answer:'describe; state'},expect:['OSG Question: Question 3. The first step is to [blank] the facts.','OSG Answer Line: Your answer:','OSG Answer Key: Answer key (fill in the blank): describe; state']},
    'short-answer':{item:{stem:'Why is tone important?',answer:'It changes how a question is heard.'},expect:['OSG Question: Question 3. Why is tone important?','OSG Answer Line: Your answer:','OSG Answer Line: ','OSG Answer Line: ','OSG Answer Key: Answer key (short answer): It changes how a question is heard.']},
    'reflection':{item:{stem:'Describe a time you reframed.',answer:'Look for a neutral rewording.',lines:2},expect:['OSG Question: Question 3. Describe a time you reframed.','OSG Answer Line: Your response:','OSG Answer Line: ','OSG Answer Key: Answer key (reflection): Look for a neutral rewording.']},
    'open-comment':{item:{stem:'What else would help?',lines:3},expect:['OSG Question: What else would help?','OSG Answer Line: Your comments:','OSG Answer Line: ','OSG Answer Line: ']}
  };
  for(const [type,{item,expect}] of Object.entries(cases)){
    const {rows}=rowsOf(c=>itemBlock({type,number:3,...item},c));
    assert.deepEqual(rows,expect,type);
  }
});

test('an item says what is wrong instead of building half of itself',()=>{
  const bad=(item,message)=>assert.throws(()=>itemBlock({type:'multiple-choice',number:1,stem:'Q?',options:['a','b'],correct:[0],...item},makeContext()),message);
  bad({stem:'  '},/Type the question/);
  bad({options:['only one'],correct:[0]},/at least two options/);
  bad({options:['a','b','c'],correct:[]},/Mark the one correct option/);
  bad({options:['a','b','c'],correct:[0,1]},/Mark the one correct option/);
  assert.throws(()=>itemBlock({type:'multiple-response',number:1,stem:'Q?',options:['a','b'],correct:[]},makeContext()),/at least one correct/);
  assert.throws(()=>itemBlock({type:'true-false',number:1,stem:'Q?'},makeContext()),/true or false/);
  assert.throws(()=>itemBlock({type:'fill-blank',number:1,stem:'Q [blank]'},makeContext()),/accepted answers/i);
  assert.throws(()=>itemBlock({type:'nonsense',number:1,stem:'Q?'},makeContext()),/Choose a type of item/);
});

test('the key letters follow the options that were typed, even when a middle option was left empty',()=>{
  const {rows}=rowsOf(c=>itemBlock({type:'multiple-choice',number:1,stem:'Q?',options:['first','','third'],correct:[2]},c));
  assert.deepEqual(rows.filter(r=>r.startsWith('OSG Option')),['OSG Option: first','OSG Option: third']);
  assert.ok(rows.includes('OSG Answer Key: Answer key (multiple choice): b.'),'third is shown as b, so the key says b');
});

test('a section is a heading in the contents, a note to the learner and the lines to fill in',()=>{
  for(const level of LEVELS){
    const {rows}=rowsOf(()=>sectionBlock(level.id));
    assert.equal(rows[0],`OSG Assessment: ${level.title}`);
    assert.equal(rows[1],`Normal: ${level.intro}`);
    assert.deepEqual(rows.slice(2),level.info.map(l=>`Normal: ${l}`));
  }
  assert.throws(()=>sectionBlock('level-nine'),/Level One, Two or Three/);
});

// ---- reading an item back
const itemRows=item=>{const {ctx,xml}=build(c=>sectionBlock('level-two')+itemBlock(item,c));return paragraphsOf(buildPackage(xml,ctx,{}));};

test('every type reads back into the fields it was built from',()=>{
  const items=[
    {type:'multiple-choice',number:1,stem:'Which?',options:['a1','b1','c1'],correct:[2],why:'Because.',points:'2',objective:'Obj.'},
    {type:'multiple-response',number:1,stem:'Which ones?',options:['a1','b1','c1','d1'],correct:[0,3]},
    {type:'true-false',number:1,stem:'Is it?',truth:false,points:'1'},
    {type:'fill-blank',number:1,stem:'It is [blank].',answer:'one; two'},
    {type:'short-answer',number:1,stem:'Why?',answer:'Because tone.',lines:3},
    {type:'reflection',number:1,stem:'Reflect.',answer:'Look for tone.',lines:6},
    {type:'open-comment',number:1,stem:'Comments?',lines:4}
  ];
  for(const item of items){
    const paras=itemRows(item),at=paras.findIndex(p=>p.style==='OSG Question'),read=readItem(paras,at);
    assert.equal(read.type,item.type,item.type);
    assert.equal(read.stem,item.stem,item.type);
    assert.equal(read.number,item.type==='open-comment'?null:1,item.type);
    assert.deepEqual(read.options,item.options||[],item.type);
    assert.deepEqual(read.correct,item.correct||[],item.type);
    if('truth' in item)assert.equal(read.truth,item.truth);
    assert.equal(read.answer,item.answer||'');
    assert.equal(read.why,item.why||'');assert.equal(read.points,item.points||'');assert.equal(read.objective,item.objective||'');
    if(['short-answer','reflection','open-comment'].includes(item.type))assert.equal(read.lines,item.lines,item.type);
  }
});

test('the same item reads back from any of its paragraphs, and the key belongs to it',()=>{
  const paras=itemRows({type:'multiple-choice',number:1,stem:'Q?',options:['x','y'],correct:[0],why:'W'});
  const q=paras.findIndex(p=>p.style==='OSG Question');
  for(let i=q;i<paras.length;i++)assert.deepEqual(readItem(paras,i).span,{start:q,end:paras.length-1},`from paragraph ${i}`);
  assert.equal(readItem(paras,q-1),null,'the section note is not an item');
  assert.equal(describeAt(paras,q+2).label,'Question 1 (multiple choice)');
});

test('a multiple-choice item whose key is missing is still read, so it can be fixed, and Review says so',()=>{
  const paras=[{text:'Guide',style:'OSG Guide Title'},{text:'Level Two Assessment',style:'OSG Assessment'},{text:'Question 1. Q?',style:'OSG Question'},{text:'one',style:'OSG Option'},{text:'two',style:'OSG Option'},{text:'Your answer:',style:'OSG Answer Line'}];
  assert.equal(readItem(paras,2).type,'multiple-choice');
  const plan=updatePlan(paras);
  assert.ok(guideIssues(plan).some(i=>i.rule==='ITEM-KEY'&&/Question 1 \(multiple choice\) has no answer key/.test(i.message)));
});

test('Review reports a section with no items',()=>{
  const plan=updatePlan([{text:'Guide',style:'OSG Guide Title'},{text:'Level One Assessment',style:'OSG Assessment'},{text:'Dear Participant',style:'Normal'}]);
  assert.ok(guideIssues(plan).some(i=>i.rule==='ASSESS-EMPTY'&&/Level One Assessment has no items yet/.test(i.message)));
});

// ---- placement and numbering
const twoItems=()=>{
  const {ctx,xml}=build(c=>sectionBlock('level-two')+itemBlock({type:'true-false',number:1,stem:'One?',truth:true},c)+itemBlock({type:'true-false',number:2,stem:'Two?',truth:false},c)+itemBlock({type:'open-comment',stem:'Comments?'},c));
  return [{text:'Guide',style:'OSG Guide Title'},...paragraphsOf(buildPackage(xml,ctx,{}))];
};

test('a new item follows the item the cursor is in, and is numbered to follow the items before it',()=>{
  const paras=twoItems(),q=paras.map((p,i)=>p.style==='OSG Question'?i:-1).filter(i=>i>=0);
  const first=itemPlacement(paras,q[0]);
  assert.equal(first.number,2);assert.equal(first.anchor,q[1]-1,'after the whole of item 1, key included');
  assert.equal(itemPlacement(paras,q[1]+1).number,3);
  const fromHeading=itemPlacement(paras,1);
  assert.equal(fromHeading.anchor,paras.length-1,'from the section heading it goes after the last item');
  assert.equal(fromHeading.number,3,'the open comment has no number, so the next numbered item is 3');
});

test('with no items yet, the first goes after the note and the information lines',()=>{
  const paras=[{text:'Guide',style:'OSG Guide Title'},...paragraphsOf(buildPackage(sectionBlock('level-one'),makeContext(),{}))];
  const p=itemPlacement(paras,1);
  assert.equal(p.anchor,paras.length-1);assert.equal(p.number,1);
});

test('an item cannot be placed outside an assessment section',()=>{
  assert.throws(()=>itemPlacement([{text:'Guide',style:'OSG Guide Title'},{text:'Hello',style:'Normal'}],1),/Items go inside an assessment section, and your cursor is not in one/);
});

test('a new section goes after the section the cursor is in, else after the last module or section, else at the end',()=>{
  const paras=twoItems();
  assert.equal(sectionPlacement(paras,3).anchor,paras.length-1);
  const guide=[{text:'Guide',style:'OSG Guide Title'},{text:'Module 1: A',style:'OSG Module'},{text:'Lesson 1.1: B',style:'OSG Lesson'},{text:'Time: 1 minute. Cumulative time: 0 minutes.',style:'OSG Time'},{text:'Transition Go.',style:'OSG Cue Facilitator'},{text:'',style:'Normal'}];
  assert.equal(sectionPlacement(guide,2).anchor,4,'after the module and its Transition');
  assert.equal(sectionPlacement([{text:'Guide',style:'OSG Guide Title'},{text:'Notes',style:'Normal'},{text:'',style:'Normal'}],1).anchor,1);
});

test('Update guide renumbers questions in each section, ignoring open comments',()=>{
  const paras=twoItems();
  paras[paras.findIndex(p=>/^Question 2\./.test(p.text))].text='Question 5. Two?';
  const plan=updatePlan(paras),qs=plan.edits.filter(e=>e.kind==='question');
  assert.deepEqual(qs.map(e=>[e.from,e.to]),[['Question 5.','Question 2.']]);
});

test('a section\'s items stop at the next heading, so a module after it is not part of it',()=>{
  const paras=[...twoItems(),{text:'Module 2: Next',style:'OSG Module'},{text:'Lesson 2.1: X',style:'OSG Lesson'}];
  const s=readAssessments(paras).sections[0];
  assert.equal(s.items.length,3);
  assert.equal(s.end,paras.length-3);
});

test('answer lines are not reported as empty paragraphs in a row',()=>{
  const {ctx,xml}=build(c=>itemBlock({type:'reflection',number:1,stem:'Reflect.',answer:'x',lines:6},c));
  const root=new DOMParser().parseFromString(`<w:document ${NS}><w:body>${xml}<w:sectPr/></w:body></w:document>`,'application/xml');
  const found=checkParts({root,stylesRoot:new DOMParser().parseFromString(stylesPart(),'application/xml'),title:'T',language:'en-US'}).findings;
  assert.equal(found.find(f=>f.rule==='EMPTY-PARAGRAPHS'),undefined);
});

test('a whole assessment section passes the structure check with nothing to fix',async()=>{
  const {ctx,xml}=build(c=>sectionBlock('level-two')+itemBlock({type:'multiple-choice',number:1,stem:'Q?',options:['a','b'],correct:[0]},c)+itemBlock({type:'reflection',number:2,stem:'R.',answer:'x'},c));
  const file=await packageToDocx(buildPackage(xml,ctx,{}),'An assessment');
  const result=await checkDocx(file);
  assert.deepEqual(result.findings.filter(f=>f.level==='fail'),[]);
});

// ---- the pane, against a Word that drops the last paragraph's style
const assessTab=a=>click(a,'guide-tab-assess');
const typeItem=(a,type,fields)=>{
  a.itemRadio(type).checked=true;a.itemRadio(type).onchange();
  for(const [id,v] of Object.entries(fields))a.n(id).value=v;
};
const checkOption=(a,n)=>{a.n(`guide-item-correct-${n}`).checked=true;};
async function withSection(level='level-two'){
  const a=await started('Assess',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await assessTab(a);
  await click(a,`guide-sec-${level}`);
  assert.equal(a.err(),'');
  return a;
}
const sectionRows=a=>a.rows().slice(a.rows().findIndex(startsWith('OSG Assessment')));

test('the switch between Guide and Assessments shows on those screens, and the assessment screen opens on its heading',async()=>{
  const a=await started('Switch',{quirks:WORD_QUIRKS});
  assert.equal(a.n('guide-switch').hidden,false);
  assert.equal(a.n('guide-tab-guide').attributes['aria-pressed'],'true');
  await assessTab(a);
  assert.equal(a.n('guide-assess').hidden,false);
  assert.equal(a.n('guide-build').hidden,true);
  assert.equal(a.n('guide-tab-assess').attributes['aria-pressed'],'true');
  await click(a,'guide-tab-guide');
  assert.equal(a.n('guide-build').hidden,false);
  assert.equal(a.n('guide-assess').hidden,true);
});

test('a Level Two section goes after the last module, and its items go inside it, in their styles, with a Word that drops the last style',async()=>{
  const a=await withSection('level-two');
  typeItem(a,'multiple-choice',{'guide-item-stem':'Which step comes first?','guide-item-opt-1':'Name the feeling','guide-item-opt-2':'Describe the facts','guide-item-opt-3':'Ask for a change','guide-item-why':'Facts keep it neutral.','guide-item-points':'1'});
  checkOption(a,2);
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  const rows=sectionRows(a);
  assert.equal(rows[0],'OSG Assessment: Level Two Assessment');
  assert.ok(rows.indexOf('OSG Question: Question 1. Which step comes first?')>0);
  assert.deepEqual(rows.slice(-8),[
    'OSG Question: Question 1. Which step comes first?','OSG Option: Name the feeling','OSG Option: Describe the facts','OSG Option: Ask for a change',
    'OSG Answer Line: Your answer:','OSG Answer Key: Answer key (multiple choice): b.','OSG Answer Key: Why: Facts keep it neutral.','OSG Answer Key: Points: 1'
  ]);
  assert.match(a.n('guide-status').textContent,/Multiple choice item added\..*choose Edit, then Undo/);
  assert.equal(a.n('guide-item-stem').value,'','the form is cleared for the next item');
});

test('the section goes after the last module and its Transition, not into a lesson',async()=>{
  const a=await withSection('level-one');
  const rows=a.rows(),heading=rows.findIndex(startsWith('OSG Assessment'));
  assert.equal(rows[heading-1].startsWith('OSG Cue Facilitator: Transition'),true);
  assert.deepEqual(rows.slice(heading,heading+2),['OSG Assessment: Level One Assessment',`Normal: ${LEVELS[0].intro}`]);
});

test('items are numbered in order, and one added in the middle brings the next Update guide to renumber',async()=>{
  const a=await withSection('level-two');
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  typeItem(a,'true-false',{'guide-item-stem':'Two?'});a.n('guide-item-false').checked=true;await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.deepEqual(a.rows().filter(r=>r.startsWith('OSG Question')),['OSG Question: Question 1. One?','OSG Question: Question 2. Two?']);
  await moveTo(a,r=>r==='OSG Question: Question 1. One?');
  typeItem(a,'true-false',{'guide-item-stem':'Between?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  assert.deepEqual(a.rows().filter(r=>r.startsWith('OSG Question')),['OSG Question: Question 1. One?','OSG Question: Question 2. Between?','OSG Question: Question 2. Two?']);
  assert.match(a.n('guide-update-note').textContent,/1 number or time is out of date\./);
  await click(a,'guide-update');
  assert.deepEqual(a.rows().filter(r=>r.startsWith('OSG Question')),['OSG Question: Question 1. One?','OSG Question: Question 2. Between?','OSG Question: Question 3. Two?']);
});

test('an item added with the cursor inside another item goes after the whole of it, key and all',async()=>{
  const a=await withSection('level-two');
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  await moveTo(a,startsWith('OSG Answer Line'));
  typeItem(a,'open-comment',{'guide-item-stem':'Comments?'});await submit(a,'guide-item-form');
  const rows=sectionRows(a),q=rows.findIndex(r=>r.startsWith('OSG Question: Comments'));
  assert.equal(rows[q-1],'OSG Answer Key: Answer key (true or false): True.');
});

test('an item cannot be added outside an assessment section; the message sits above Insert item, with a way forward',async()=>{
  const a=await started('Outside',{quirks:WORD_QUIRKS});
  await assessTab(a);
  const before=a.rows();
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  assert.match(a.n('guide-item-problem').textContent,/Items go inside an assessment section, and your cursor is not in one\. What you typed is still here\./);
  assert.equal(a.n('guide-item-msg-box').hidden,false);
  assert.equal(a.n('guide-item-problem').hidden,false);
  assert.equal(a.n('guide-item-offer').hidden,false,'a Level Two section can be started from the message');
  assert.equal(a.n('guide-error').hidden,true,'not at the top of the pane, where it cannot be seen from Insert item');
  assert.equal(a.n('guide-item-stem').value,'One?','what was typed is kept');
  assert.equal(a.n('guide-item-true').checked,true);
  assert.deepEqual(a.rows(),before,'nothing was inserted');
});

test('starting the Level Two section from the message keeps the typed item, and Insert item then puts it in',async()=>{
  const a=await started('Outside',{quirks:WORD_QUIRKS});
  await assessTab(a);
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  await click(a,'guide-item-start-two');
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-item-problem').hidden,true);assert.equal(a.n('guide-item-offer').hidden,true);
  assert.equal(a.n('guide-item-ok').hidden,false);
  assert.match(a.n('guide-item-ok').textContent,/Level Two Assessment added\. What you typed is still in the form\. Choose Insert item/);
  assert.equal(a.n('guide-item-stem').value,'One?');
  assert.equal(focusedId(),'guide-item-go','focus is on Insert item, which is enabled again by then');
  assert.equal(a.n('guide-item-where').attributes['data-state'],'ready');
  assert.equal(a.n('guide-item-where').textContent,'Items go in: Level Two Assessment.');
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  assert.ok(sectionRows(a).includes('OSG Question: Question 1. One?'));
  assert.equal(a.n('guide-item-stem').value,'','the form is cleared once the item is in');
  assert.equal(a.n('guide-item-msg-box').hidden,true);
});

test('switching between the Guide and Assessments screens drops the last screen\'s message',async()=>{
  const a=await started('Switch messages',{quirks:WORD_QUIRKS});
  await addModule(a,'One');
  assert.notEqual(a.n('guide-status').textContent,'');
  await assessTab(a);
  assert.equal(a.n('guide-status').textContent,'');
  await click(a,'guide-sec-level-two');
  assert.notEqual(a.n('guide-status').textContent,'');
  await click(a,'guide-tab-guide');
  assert.equal(a.n('guide-status').textContent,'');
});

test('before anything is typed the pane says where an item would go',async()=>{
  const a=await started('Where',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await assessTab(a);
  const where=()=>({text:a.n('guide-item-where').textContent,state:a.n('guide-item-where').attributes['data-state']});
  assert.equal(where().state,'none');
  assert.match(where().text,/^No assessment section here yet\. Start one above, or click inside one on the page/);
  await click(a,'guide-sec-level-two');
  assert.deepEqual(where(),{text:'Items go in: Level Two Assessment.',state:'ready'});
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  assert.equal(where().text,'Items go in: Level Two Assessment, after Question 1 (true or false).');
  await moveTo(a,startsWith('OSG Lesson'));
  assert.equal(where().state,'none','back in the guide, outside any section');
});

test('the message above Insert item goes away when the cursor reaches a section, or another type is chosen',async()=>{
  const a=await started('Clears',{quirks:WORD_QUIRKS});
  await assessTab(a);
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  assert.equal(a.n('guide-item-msg-box').hidden,false);
  a.itemRadio('open-comment').onchange();
  assert.equal(a.n('guide-item-msg-box').hidden,true,'choosing another type clears it');
  await submit(a,'guide-item-form');
  assert.equal(a.n('guide-item-problem').hidden,false);
  await click(a,'guide-sec-level-three');
  assert.equal(a.n('guide-item-problem').hidden,true,'the cursor is in a section now');
  assert.equal(a.n('guide-item-msg-box').hidden,true);
});

test('the three section buttons say which is the quiz, and the page matches the data',()=>{
  const html=readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');
  assert.deepEqual(LEVELS.map(l=>l.button),['Level One: reaction survey','Level Two: knowledge check (quiz)','Level Three: follow-up survey']);
  for(const l of LEVELS){
    assert.match(html,new RegExp(`<button id="guide-sec-${l.id}" type="button" aria-describedby="guide-sec-${l.id}-hint">${l.button.replace(/[()]/g,'\\$&')}</button>`));
    assert.ok(html.includes(`<p id="guide-sec-${l.id}-hint" class="hint">${l.blurb}</p>`),`${l.id} has its one-line description`);
  }
  assert.match(LEVELS[1].blurb,/quiz/i);
});

test('the pane header is compact: no tagline, and the build check stays visible in the header',()=>{
  const html=readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');
  const header=html.match(/<header\b[^>]*>([\s\S]*?)<\/header>/)[1];
  assert.doesNotMatch(header,/tagline|Prepare to teach/);
  assert.match(header,/id="build-loaded"/);assert.match(header,/id="build-check"/);
});

test('a form that is not finished says what is missing and inserts nothing',async()=>{
  const a=await withSection('level-two');
  const before=a.rows();
  await submit(a,'guide-item-form');assert.match(a.err(),/Type the question/);
  assert.equal(a.n('guide-item-problem').hidden,false,'the message sits above Insert item');
  assert.equal(a.n('guide-error').hidden,true);
  assert.equal(a.n('guide-item-offer').hidden,true,'no section is missing here, so no section is offered');
  typeItem(a,'multiple-choice',{'guide-item-stem':'Q?','guide-item-opt-1':'a','guide-item-opt-2':'b'});
  await submit(a,'guide-item-form');assert.match(a.err(),/Mark the one correct option/);
  a.n('guide-item-correct-1').checked=true;a.n('guide-item-points').value='lots';
  await submit(a,'guide-item-form');assert.match(a.err(),/Points must be a number/);
  assert.deepEqual(a.rows(),before);
});

test('choosing a type shows only the fields it needs',async()=>{
  const a=await started('Fields',{quirks:WORD_QUIRKS});
  const shown=()=>['guide-item-options','guide-item-truefalse','guide-item-answer-box','guide-item-lines-box','guide-item-key-box'].filter(id=>!a.n(id).hidden);
  a.itemRadio('multiple-choice').onchange();
  assert.deepEqual(shown(),['guide-item-options','guide-item-key-box']);
  a.itemRadio('true-false').onchange();
  assert.deepEqual(shown(),['guide-item-truefalse','guide-item-key-box']);
  a.itemRadio('short-answer').onchange();
  assert.deepEqual(shown(),['guide-item-answer-box','guide-item-lines-box','guide-item-key-box']);
  assert.equal(a.n('guide-item-lines').value,'3');
  a.itemRadio('open-comment').onchange();
  assert.deepEqual(shown(),['guide-item-lines-box'],'a survey question has no key');
  assert.equal(a.n('guide-item-stem-label').textContent,'Prompt');
});

test('Edit this item puts the item back in the form and Update item replaces it where it stands',async()=>{
  const a=await withSection('level-two');
  typeItem(a,'multiple-choice',{'guide-item-stem':'Which?','guide-item-opt-1':'one','guide-item-opt-2':'two','guide-item-opt-3':'three','guide-item-why':'Because.'});checkOption(a,1);
  await submit(a,'guide-item-form');
  typeItem(a,'true-false',{'guide-item-stem':'After?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  await moveTo(a,startsWith('OSG Option'));
  assert.equal(a.n('guide-assess-selected').hidden,false);
  assert.equal(a.n('guide-assess-selected-text').textContent,'Your cursor is in Question 1 (multiple choice).');
  await click(a,'guide-assess-edit');
  assert.equal(a.n('guide-item-stem').value,'Which?');
  assert.equal(a.n('guide-item-opt-3').value,'three');
  assert.equal(a.n('guide-item-correct-1').checked,true);assert.equal(a.n('guide-item-correct-2').checked,false);
  assert.equal(a.n('guide-item-why').value,'Because.');
  assert.equal(a.n('guide-item-heading').textContent,'Edit an item');
  assert.equal(a.n('guide-item-go').textContent,'Update item');
  a.n('guide-item-correct-1').checked=false;checkOption(a,3);a.n('guide-item-stem').value='Which is best?';a.n('guide-item-opt-4').value='four';
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  const rows=sectionRows(a).filter(r=>/^OSG (Question|Option|Answer Key)/.test(r));
  assert.deepEqual(rows,[
    'OSG Question: Question 1. Which is best?','OSG Option: one','OSG Option: two','OSG Option: three','OSG Option: four',
    'OSG Answer Key: Answer key (multiple choice): c.','OSG Answer Key: Why: Because.',
    'OSG Question: Question 2. After?','OSG Answer Key: Answer key (true or false): True.'
  ]);
  assert.equal(a.n('guide-item-go').textContent,'Insert item');
  assert.match(a.n('guide-status').textContent,/Multiple choice item updated\./);
});

test('Cancel edit leaves an item as it was',async()=>{
  const a=await withSection('level-two');
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  const before=a.rows();
  await moveTo(a,startsWith('OSG Question'));
  await click(a,'guide-assess-edit');
  a.n('guide-item-stem').value='Changed';
  await click(a,'guide-item-cancel');
  assert.deepEqual(a.rows(),before);
  assert.equal(a.n('guide-item-stem').value,'');
});

test('Update item refuses if the question was changed in Word after it was chosen',async()=>{
  const a=await withSection('level-two');
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  await moveTo(a,startsWith('OSG Question'));
  await click(a,'guide-assess-edit');
  a.adapter.typeOver(a.rows().findIndex(startsWith('OSG Question')),'Question 1. Someone retyped this');
  a.n('guide-item-stem').value='Mine';
  await submit(a,'guide-item-form');
  assert.match(a.err(),/has changed since you chose it/);
});

test('the Build screen leaves items to the Assessments screen, and Contents lists each section',async()=>{
  const a=await withSection('level-two');
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  await moveTo(a,startsWith('OSG Question'));
  assert.equal(a.n('guide-selected').hidden,true,'no "Edit this block" for an item');
  assert.equal(a.n('guide-assess-selected').hidden,false);
  await click(a,'guide-tab-guide');
  const rows=a.n('guide-contents').children.map(li=>li.children[0].textContent);
  assert.ok(rows.includes('Level Two Assessment'));
});

test('Review says a section has no items yet, and Go there reaches it',async()=>{
  const a=await withSection('level-three');
  await click(a,'guide-tab-guide');await click(a,'guide-open-review');
  const look=a.n('guide-review-look').children.find(li=>/Level Three Assessment has no items yet/.test(li.children[0].textContent));
  assert.ok(look);
  await look.children[1].onclick();
  assert.equal(a.rows()[a.adapter.cursor()],'OSG Assessment: Level Three Assessment');
});

test('Undo takes a whole item out and the pane follows',async()=>{
  const a=await withSection('level-two');
  a.p.watch();
  typeItem(a,'true-false',{'guide-item-stem':'One?'});a.n('guide-item-true').checked=true;await submit(a,'guide-item-form');
  assert.ok(a.rows().some(startsWith('OSG Question')));
  a.adapter.undo();
  await a.adapter.fireSelection();
  assert.ok(!a.rows().some(startsWith('OSG Question')));
  assert.ok(a.rows().some(startsWith('OSG Assessment')),'the section is still there');
});
