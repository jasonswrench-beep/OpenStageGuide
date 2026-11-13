// The module overview: the page under a module's heading, built by a button from the module's lessons. Proven against
// stand-ins for Word, not Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {makeContext} from '../src/guide-ooxml.mjs';
import {overviewFacts,overviewBody,overviewState,overviewPlacement,overviewsSummary,moduleAt,OVERVIEW_HEAD} from '../src/guide-pages.mjs';
import {readGuide,lessonPlacement,moduleIntroEnd,kindOf} from '../src/guide-model.mjs';
import {summarize} from '../src/guide-workbook.mjs';
import {textOf} from '../src/guide-check.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {parse,packageToDocx} from './helpers/guide-package.mjs';
import {pane,started,WORD_QUIRKS,click,addModule,addLesson,addBlock,moveTo,startsWith} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const P=(text,style='Normal')=>({text,style});

// What lies under the module heading, up to its first lesson.
const intro=(a,heading)=>{
  const r=a.rows(),i=r.findIndex(x=>x.startsWith(`OSG Module: ${heading}`));
  let j=i+1;
  while(j<r.length&&!/^(OSG Lesson|OSG Module|OSG Cue|OSG Time|heading)/.test(r[j]))j++;
  return r.slice(i+1,j);
};
const count=(a,row)=>a.rows().filter(r=>r===row).length;

// Module 1: lessons "Listen" (20 min, with an objective) and "Reframe" (15 min, none). Module 2: "Reply" (10 min).
async function guide(){
  const a=await started('Reframing conflict',{quirks:WORD_QUIRKS});
  await addModule(a,'Recognize');
  await addLesson(a,'Listen','20');
  await addBlock(a,'objective','Name the feeling.');
  await addLesson(a,'Reframe','15');
  await addModule(a,'Respond');
  await addLesson(a,'Reply','10');
  await moveTo(a,startsWith('OSG Module: Module 1'));
  return a;
}
const TABLE=[
  'OSG Overview: Time: 35 minutes in 2 lessons.',
  'OSG Table Head: Lesson','OSG Table Head: Minutes','OSG Table Head: Objective',
  'OSG Table Cell: Lesson 1.1: Listen','OSG Table Cell: 20','OSG Table Cell: Name the feeling.',
  'OSG Table Cell: Lesson 1.2: Reframe','OSG Table Cell: 15','OSG Table Cell: No objective yet.'
];

test('the facts: each lesson with its minutes and its objective, and the module\'s total',async()=>{
  const a=await guide();
  const list=await a.adapter.paragraphs(),mod=readGuide(list).modules[0];
  assert.deepEqual(overviewFacts(list,mod),{name:'Module 1: Recognize',total:35,lessons:[
    {name:'Lesson 1.1: Listen',minutes:20,objective:'Name the feeling.'},
    {name:'Lesson 1.2: Reframe',minutes:15,objective:'No objective yet.'}
  ]});
});

test('the first build writes the author\'s lines to type over, the time line and the lessons table, in one page under the heading',async()=>{
  const a=await guide();
  assert.match(a.n('guide-overview-note').textContent,/The overview of Module 1: Recognize has not been built yet\. 0 of 2 modules have an overview\./);
  await click(a,'guide-build-overview');
  assert.equal(a.err(),'');
  assert.deepEqual(intro(a,'Module 1'),[
    'OSG Overview: Goal: Say what learners will be able to do by the end of this module.',
    'OSG Overview: Overview: Describe what happens in this module and how it connects to the rest of the program.',
    ...TABLE,
    'OSG Overview: Materials needed: List the materials, equipment and room set-up this module needs.',
    'OSG Overview: References and resources: List any readings, links or other resources.'
  ]);
  assert.match(a.n('guide-status').textContent,/The overview of Module 1: Recognize is built from its 2 lessons\. Type over the sentence after Goal/);
  assert.match(a.n('guide-overview-note').textContent,/matches your guide\. 1 of 2 modules has an overview\./);
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Overview')).length,5,'only module 1 has one');
});

test('a module with an overview still reads as a guide: the lessons, time blocks and totals are unchanged',async()=>{
  const a=await guide();
  await click(a,'guide-build-overview');
  const g=readGuide(await a.adapter.paragraphs());
  assert.deepEqual(g.modules.map(m=>[m.text,m.lessons.length]),[['Module 1: Recognize',2],['Module 2: Respond',1]]);
  assert.equal(a.n('guide-total').textContent,'Total teaching time: 45 min');
});

test('building again rewrites the time line and table only: the author\'s own words stay, and no second table appears',async()=>{
  const a=await guide();
  await click(a,'guide-build-overview');
  const goal=a.rows().findIndex(r=>r.startsWith('OSG Overview: Goal:'));
  a.adapter.typeOver(goal,'Goal: Learners can name a feeling in one sentence.');
  const mats=a.rows().findIndex(r=>r.startsWith('OSG Overview: Materials needed:'));
  a.adapter.typeOver(mats,'Materials needed: Flip chart and markers.');
  await moveTo(a,startsWith('OSG Lesson: Lesson 1.2'));
  await addLesson(a,'Close','10');
  await moveTo(a,startsWith('OSG Module: Module 1'));
  assert.match(a.n('guide-overview-note').textContent,/Module 1: Recognize is out of date\. Choose its button to build it again\. 1 of 2 modules has an overview, 1 out of date\./);
  await click(a,'guide-build-overview');
  assert.equal(a.err(),'');
  assert.equal(count(a,'OSG Table Head: Lesson'),1);
  const body=intro(a,'Module 1');
  assert.equal(body[0],'OSG Overview: Goal: Learners can name a feeling in one sentence.');
  assert.ok(body.includes('OSG Overview: Time: 45 minutes in 3 lessons.'));
  assert.ok(body.includes('OSG Table Cell: Lesson 1.3: Close'));
  assert.equal(body.at(-2),'OSG Overview: Materials needed: Flip chart and markers.');
  assert.match(a.n('guide-status').textContent,/Your own lines were not touched/);
  assert.match(a.n('guide-overview-note').textContent,/matches your guide/);
});

test('typing over the table\'s minutes makes the note say it is out of date',async()=>{
  const a=await guide();
  await click(a,'guide-build-overview');
  a.adapter.typeOver(a.rows().indexOf('OSG Table Cell: 20'),'99');await a.p.quietRefresh();
  assert.match(a.n('guide-overview-note').textContent,/out of date/);
});

test('words the author typed among the overview\'s lines do not hide it from the button or move it',async()=>{
  const a=await guide();
  await click(a,'guide-build-overview');
  const goal=a.rows().findIndex(r=>r.startsWith('OSG Overview: Goal:'));
  await a.adapter.insert(()=>'<w:p><w:r><w:t>A note of my own.</w:t></w:r></w:p>',{where:'after',after:goal});
  await a.p.quietRefresh();
  assert.match(a.n('guide-overview-note').textContent,/matches your guide/);
  await click(a,'guide-build-overview');
  assert.equal(count(a,'OSG Table Head: Lesson'),1);
  assert.equal(count(a,'Normal: A note of my own.'),1);
});

test('the button works on the module the cursor is in, and says so when the cursor is in none',async()=>{
  const a=await guide();
  a.adapter.place(0);await a.p.quietRefresh();
  assert.match(a.n('guide-overview-note').textContent,/Put your cursor in a module to build its overview\./);
  const before=a.rows();
  await click(a,'guide-build-overview');
  assert.match(a.err(),/Put your cursor inside a module first/);
  assert.deepEqual(a.rows(),before);
  await moveTo(a,startsWith('OSG Lesson: Lesson 2.1'));
  await click(a,'guide-build-overview');
  assert.equal(a.err(),'');
  assert.equal(intro(a,'Module 2').filter(r=>r.startsWith('OSG Table Cell: Lesson')).join('|'),'OSG Table Cell: Lesson 2.1: Reply');
  assert.equal(intro(a,'Module 1').length,0,'module 1 is untouched');
});

test('with no lesson in the module the button says what to do and changes nothing',async()=>{
  const a=await started('Empty guide',{quirks:WORD_QUIRKS});
  await addModule(a,'Alone');
  await moveTo(a,startsWith('OSG Module: Module 1'));
  const before=a.rows();
  await click(a,'guide-build-overview');
  assert.match(a.err(),/Add a lesson to this module first/);
  assert.deepEqual(a.rows(),before);
  assert.equal(a.n('guide-build-overview').disabled,false);
});

test('if only the author\'s lines are left, the button adds just the time line and table after them',async()=>{
  const a=pane([P('Module 1: A','OSG Module'),P('Goal: Mine.','OSG Overview'),P('Lesson 1.1: B','OSG Lesson'),P('Time: 10 minutes. Cumulative time: 0 minutes.','OSG Time')],{quirks:WORD_QUIRKS});
  await a.p.refresh();a.adapter.place(0);await a.p.quietRefresh();
  await click(a,'guide-build-overview');
  assert.equal(a.err(),'');
  assert.deepEqual(a.rows().slice(0,4),['OSG Module: Module 1: A','OSG Overview: Goal: Mine.','OSG Overview: Time: 10 minutes in 1 lesson.','OSG Table Head: Lesson']);
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Overview: Goal:')).length,1);
});

test('a new lesson in a module with no lessons goes after its overview, not above it',()=>{
  const paras=[P('T','OSG Guide Title'),P('Module 1: A','OSG Module'),P('Goal: Mine.','OSG Overview'),P('Time: 0 minutes in 0 lessons.','OSG Overview'),
    P('Lesson','OSG Table Head'),P('Minutes','OSG Table Head'),P('Objective','OSG Table Head'),P('Transition   Next','OSG Cue Facilitator')];
  assert.equal(moduleIntroEnd(paras,1),6);
  assert.equal(lessonPlacement(paras,1).anchor,6);
  assert.equal(moduleIntroEnd([P('Module 1: A','OSG Module'),P('Lesson 1.1: B','OSG Lesson')],0),0,'no overview: right after the heading, as before');
});

test('the overview body has a header row, no merged or empty cells, and says what the table is',()=>{
  const facts={name:'Module 1: A',total:10,lessons:[{name:'Lesson 1.1: B',minutes:10,objective:'No objective yet.'}]};
  const xml=overviewBody(facts,{withAuthored:true});
  assert.match(xml,/<w:tblHeader\/>/);
  assert.match(xml,/<w:tblCaption w:val="Lessons in Module 1: A"\/>/);
  assert.doesNotMatch(xml,/gridSpan|vMerge/);
  const doc=parse(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`);
  const cells=Array.from(doc.getElementsByTagNameNS(W,'tc'));
  assert.equal(cells.length,6);
  assert.ok(cells.every(c=>textOf(c).trim()));
  assert.deepEqual(OVERVIEW_HEAD,['Lesson','Minutes','Objective']);
  assert.throws(()=>overviewBody({name:'X',total:0,lessons:[]}),/Add a lesson to this module first/);
  assert.equal(makeContext().used.size,0);
});

// ---- the review, the structure check, undo, and the learner workbook
test('Undo takes the whole overview out in one step',async()=>{
  const a=await guide();
  const before=a.rows();
  await click(a,'guide-build-overview');
  assert.ok(a.adapter.undo());
  assert.deepEqual(a.rows(),before);
});

test('with an overview built the review has nothing to fix and the structure check finds the table headed',async()=>{
  const a=await guide();
  await click(a,'guide-build-overview');
  await click(a,'guide-open-review');
  assert.deepEqual(a.n('guide-review-fix').children.map(li=>li.textContent),['Nothing to fix.']);
  const result=await checkDocx(await packageToDocx(await a.adapter.sourcePackage(),'Reframing conflict'));
  assert.deepEqual(result.findings.filter(f=>f.level==='fail'),[]);
  assert.deepEqual(result.findings.filter(f=>['TABLE-BLANK','TABLE-MERGED','EMPTY-PARAGRAPHS','TYPED-LIST'].includes(f.rule)&&f.level!=='pass'),[]);
  assert.equal(result.findings.filter(f=>f.rule==='TABLE-HEADER'&&f.level==='pass').length,1);
});

test('the overview is for the facilitator: the workbook leaves it out, says so, and lists none of it as the author\'s own text',async()=>{
  const a=await guide();
  await addBlock(a,'ask','What is hard about it?');
  await moveTo(a,startsWith('OSG Module: Module 1'));
  await click(a,'guide-build-overview');
  const s=summarize(await a.adapter.paragraphs());
  assert.equal(s.left.overview,1);
  assert.deepEqual(s.leftOut,[],'the overview is not "your own text"');
  await click(a,'guide-tab-workbook');
  assert.match(a.n('guide-workbook-left').textContent,/1 module overview/);
  await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true,'the leak check passes');
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64')),doc=parse(await zip.file('word/document.xml').async('string'));
  const text=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>textOf(p)).join('\n');
  for(const gone of ['Say what learners will be able to do','Materials needed','References and resources','Time: 35 minutes','No objective yet','Minutes'])assert.ok(!text.includes(gone),`${gone} is not in the workbook`);
  assert.ok(text.includes('What is hard about it?'),'the learner block is');
});

test('the kinds and the summary read the overview from its style, not its words',()=>{
  assert.equal(kindOf(P('Goal: x','OSG Overview')),'overview');
  assert.equal(kindOf(P('Goal: x')),'other');
  const paras=[P('Module 1: A','OSG Module'),P('Goal: x','OSG Overview'),P('Lesson 1.1: B','OSG Lesson')];
  assert.deepEqual(overviewsSummary(paras),{total:1,built:0,stale:0});
  assert.equal(moduleAt(paras,2).text,'Module 1: A');
  assert.equal(moduleAt([P('Intro text.'),paras[0]],0),null);
  assert.deepEqual(overviewState(paras,readGuide(paras).modules[0]),{built:false,stale:false});
  assert.deepEqual(overviewPlacement(paras,readGuide(paras).modules[0]).after,1);
});
