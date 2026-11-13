// The two generated Getting Started pages: Training at a Glance and Graphic Cues. Proven against stand-ins for Word, not Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {makeContext,buildPackage,para,run} from '../src/guide-ooxml.mjs';
import {glanceRows,glanceBody,glanceState,glancePlacement,cuesUsed,cuesBody,cuesState,cuesPlacement,GLANCE_HEAD,CUES_HEAD} from '../src/guide-pages.mjs';
import {readGuide,cueFromText} from '../src/guide-model.mjs';
import {readStyles,textOf} from '../src/guide-check.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {parse,paragraphsOf,packageToDocx} from './helpers/guide-package.mjs';
import {pane,started,WORD_QUIRKS,click,submit,addModule,addLesson,addBlock,moveTo,startsWith} from './helpers/guide-harness.mjs';
import {focusedId} from './helpers/app-dom.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const P=(text,style='Normal')=>({text,style});

// What lies under a Heading 2 of this name, up to the next heading or module.
const under=(a,name)=>{
  const r=a.rows(),i=r.indexOf(`heading 2: ${name}`);
  if(i<0)return null;
  let j=i+1;
  while(j<r.length&&!/^(heading [12]|OSG Module|OSG Lesson|OSG Assessment)/.test(r[j]))j++;
  return r.slice(i+1,j);
};
const count=(a,row)=>a.rows().filter(r=>r===row).length;

// Two modules: lesson 1.1 (20 min) and 1.2 (15 min), then lesson 2.1 whose time block says 30 but whose blocks add up to 10.
async function guide(opts={quirks:WORD_QUIRKS}){
  const a=await started('Reframing conflict',opts);
  await addModule(a,'Recognize');
  await addLesson(a,'Listen','20');
  await addLesson(a,'Reframe','15');
  await addModule(a,'Respond');
  await addLesson(a,'Reply','30');
  await addBlock(a,'say','Say it neutrally.',{minutes:'5'});
  await addBlock(a,'ask','What makes it hard?',{minutes:'5'});
  return a;
}
const GLANCE_ROWS=[
  'OSG Table Head: Module or lesson','OSG Table Head: Minutes','OSG Table Head: Starts at minute',
  'OSG Table Cell: Module 1: Recognize','OSG Table Cell: 35','OSG Table Cell: 0',
  'OSG Table Cell: Lesson 1.1: Listen','OSG Table Cell: 20','OSG Table Cell: 0',
  'OSG Table Cell: Lesson 1.2: Reframe','OSG Table Cell: 15','OSG Table Cell: 20',
  'OSG Table Cell: Module 2: Respond','OSG Table Cell: 10','OSG Table Cell: 35',
  'OSG Table Cell: Lesson 2.1: Reply','OSG Table Cell: 10','OSG Table Cell: 35'
];

// ---- Training at a Glance: the table
test('the glance lists each module with its lessons\' minutes added up, and each lesson with the minute it starts',async()=>{
  const a=await guide();
  const g=readGuide(await a.adapter.paragraphs()),{rows,total}=glanceRows(g);
  assert.deepEqual(rows.map(r=>[r.kind,r.name,r.minutes,r.starts]),[
    ['module','Module 1: Recognize',35,0],['lesson','Lesson 1.1: Listen',20,0],['lesson','Lesson 1.2: Reframe',15,20],
    ['module','Module 2: Respond',10,35],['lesson','Lesson 2.1: Reply',10,35]
  ],'a lesson whose blocks have minutes takes its time from them, as Update guide does');
  assert.equal(total,45);
});

test('the glance table has a real header row, no merged or blank cells, and says what it is for a screen reader',()=>{
  const g=readGuide([P('Module 1: A','OSG Module'),P('Lesson 1.1: B','OSG Lesson'),P('Time: 10 minutes. Cumulative time: 0 minutes.','OSG Time')]);
  const ctx=makeContext(),xml=glanceBody(g);
  assert.match(xml,/<w:tblHeader\/>/);
  assert.match(xml,/<w:tblCaption w:val="Training at a Glance"\/>/);
  assert.match(xml,/<w:tblDescription w:val="[^"]{20,}"\/>/);
  assert.doesNotMatch(xml,/gridSpan|vMerge|hMerge/);
  const doc=parse(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`);
  const cells=Array.from(doc.getElementsByTagNameNS(W,'tc'));
  assert.equal(cells.length,3*3);
  assert.ok(cells.every(c=>textOf(c).trim()),'every cell has words');
  assert.equal(ctx.used.size,0,'no pictures');
});

test('the glance will not build until there is a lesson to list',()=>{
  assert.throws(()=>glanceBody(readGuide([P('Module 1: A','OSG Module')])),/Add a module and a lesson first/);
});

// ---- Training at a Glance: the button
test('the button swaps the placeholder under the heading for the table, an intro and the total',async()=>{
  const a=await guide();
  assert.match(a.n('guide-glance-note').textContent,/has not been built yet/);
  await click(a,'guide-build-glance');
  assert.equal(a.err(),'');
  const body=under(a,'Training at a Glance');
  assert.deepEqual(body,[
    'Normal: The minutes each module and lesson takes, and the minute of the session at which each one starts.',
    ...GLANCE_ROWS,
    'Normal: Total teaching time: 45 minutes.'
  ]);
  assert.ok(!a.rows().some(r=>r.includes('Give a short overview')),'the placeholder is gone');
  assert.match(a.n('guide-glance-note').textContent,/matches your guide/);
  assert.match(a.n('guide-status').textContent,/built from 3 lessons\. It took the place of the placeholder text\./);
  assert.equal(a.n('guide-build-glance').disabled,false);
  assert.equal(focusedId(),'guide-build-glance','focus comes back to the button');
});

test('the note says the page is out of date once the guide changes, and building again replaces the table, not adds another',async()=>{
  const a=await guide();
  await click(a,'guide-build-glance');
  // The cursor is on the new page; the author clicks back into the last lesson to add another.
  await moveTo(a,startsWith('OSG Lesson: Lesson 2.1'));
  await addLesson(a,'Close','10');
  assert.match(a.n('guide-glance-note').textContent,/Training at a Glance is out of date/);
  await click(a,'guide-build-glance');
  assert.equal(a.err(),'');
  assert.equal(count(a,'OSG Table Head: Module or lesson'),1,'one table');
  assert.equal(a.rows().filter(r=>r.startsWith('Normal: Total teaching time')).length,1,'one total');
  assert.ok(under(a,'Training at a Glance').includes('OSG Table Cell: Lesson 2.2: Close'));
  assert.match(a.n('guide-glance-note').textContent,/matches your guide/);
  assert.match(a.n('guide-status').textContent,/The earlier one was replaced\./);
});

test('typing over the table\'s numbers makes the note say it is out of date',async()=>{
  const a=await guide();
  await click(a,'guide-build-glance');
  const i=a.rows().indexOf('OSG Table Cell: Lesson 1.1: Listen');
  a.adapter.typeOver(i+1,'99');await a.p.quietRefresh();
  assert.match(a.n('guide-glance-note').textContent,/out of date/);
});

test('the author\'s own text under the heading is kept: the table goes above it',async()=>{
  const a=await guide();
  const i=a.rows().findIndex(r=>r.startsWith('Normal: Give a short overview'));
  a.adapter.typeOver(i,'My own overview of the course.');await a.p.quietRefresh();
  await click(a,'guide-build-glance');
  const body=under(a,'Training at a Glance');
  assert.equal(body.at(-1),'Normal: My own overview of the course.');
  assert.equal(count(a,'OSG Table Head: Module or lesson'),1);
  assert.match(a.n('guide-status').textContent,/above your own text/);
  // Building again replaces only its own page.
  await click(a,'guide-build-glance');
  assert.equal(count(a,'OSG Table Head: Module or lesson'),1);
  assert.equal(under(a,'Training at a Glance').at(-1),'Normal: My own overview of the course.');
});

test('text the author added after the total line stays when the page is built again',async()=>{
  const a=await guide();
  await click(a,'guide-build-glance');
  const i=a.rows().findIndex(r=>r.startsWith('Normal: Total teaching time'));
  await a.adapter.insert(()=>para(null,run('Breaks are not counted.')),{where:'after',after:i});
  await a.p.quietRefresh();
  await click(a,'guide-build-glance');
  assert.equal(a.err(),'');
  assert.equal(count(a,'OSG Table Head: Module or lesson'),1);
  const body=under(a,'Training at a Glance');
  assert.deepEqual(body.slice(-2),['Normal: Total teaching time: 45 minutes.','Normal: Breaks are not counted.']);
});

test('words the author typed between the heading and the table do not hide the page from the button',async()=>{
  const a=await guide();
  await click(a,'guide-build-cues');
  const h=a.rows().indexOf('heading 2: Graphic Cues');
  await a.adapter.insert(()=>para(null,run('Icons are from the training set.')),{where:'after',after:h});
  await a.p.quietRefresh();
  assert.match(a.n('guide-cues-note').textContent,/matches your guide/);
  await moveTo(a,startsWith('OSG Time'));
  await addBlock(a,'discussion','Talk about the case.');
  assert.match(a.n('guide-cues-note').textContent,/out of date/);
  await click(a,'guide-build-cues');
  assert.equal(count(a,'OSG Table Head: Cue'),1,'replaced, not added');
  assert.equal(under(a,'Graphic Cues')[0],'Normal: Icons are from the training set.');
});

test('with no lesson yet the button says what to do and changes nothing',async()=>{
  const a=await started('Empty guide',{quirks:WORD_QUIRKS});
  const before=a.rows();
  await click(a,'guide-build-glance');
  assert.match(a.err(),/Add a module and a lesson first/);
  assert.deepEqual(a.rows(),before);
  assert.equal(a.n('guide-build-glance').disabled,false,'the buttons are given back');
});

test('with no Training at a Glance heading, one is added before the first module',async()=>{
  const a=pane(undefined,{quirks:WORD_QUIRKS});await a.p.refresh();
  a.type('guide-title','Bare guide');a.n('guide-standard').checked=false;
  await submit(a,'guide-start-form');
  await addModule(a,'One');await addLesson(a,'First','10');
  await click(a,'guide-build-glance');
  assert.equal(a.err(),'');
  const r=a.rows(),h=r.indexOf('heading 2: Training at a Glance');
  assert.ok(h>0&&h<r.findIndex(x=>x.startsWith('OSG Module')),'the heading comes before the first module');
  assert.match(a.n('guide-status').textContent,/no Training at a Glance heading, so one was added/);
  assert.equal(count(a,'OSG Table Head: Module or lesson'),1);
});

test('a guide that opens with its first module still gets the page, at the very top',async()=>{
  const a=pane([P('Module 1: A','OSG Module'),P('Lesson 1.1: B','OSG Lesson'),P('Time: 10 minutes. Cumulative time: 0 minutes.','OSG Time')],{quirks:WORD_QUIRKS});
  await a.p.refresh();
  await click(a,'guide-build-glance');
  assert.equal(a.err(),'');
  assert.equal(a.rows()[0],'heading 2: Training at a Glance');
  assert.ok(a.rows().includes('OSG Table Cell: Lesson 1.1: B'));
});

test('the placement is read from the document: replace a built page, replace only the untouched placeholder, else go under the heading',()=>{
  const heading=P('Training at a Glance','Heading 2');
  assert.equal(glancePlacement([P('Title','OSG Guide Title'),heading,P('Give a short overview of the modules, the time each takes and what happens in each.')]).replaces,'placeholder');
  const own=glancePlacement([heading,P('Their words.')]);
  assert.deepEqual([own.replaces,own.remove,own.after],[null,null,0]);
  assert.deepEqual(glanceState([heading,P('Their words.')]),{built:false,stale:false});
});

// ---- Graphic Cues
test('the cues page lists the cues the guide uses, in the catalogue\'s order with Time first, and nothing else',async()=>{
  const a=await guide();
  await click(a,'guide-build-cues');
  assert.equal(a.err(),'');
  const body=under(a,'Graphic Cues');
  assert.equal(body[0].startsWith('Normal: These pictures mark the kind of block'),true);
  const cells=body.filter(r=>r.startsWith('OSG Table Cell:')).map(r=>r.slice('OSG Table Cell: '.length));
  const names=cells.filter((_,i)=>i%2===0);
  assert.deepEqual(names,['Time','Say','Ask','Transition']);
  assert.deepEqual(body.filter(r=>r.startsWith('OSG Table Head:')),CUES_HEAD.map(h=>`OSG Table Head: ${h}`));
  assert.ok(cells[1].length>20,'each has a meaning');
  assert.match(a.n('guide-status').textContent,/built with 4 cues/);
});

test('the cues page goes before Program Preparation, after About This Guide',async()=>{
  const a=await guide();
  await click(a,'guide-build-cues');
  const r=a.rows(),about=r.indexOf('heading 2: About This Guide'),cues=r.indexOf('heading 2: Graphic Cues'),prep=r.indexOf('heading 2: Program Preparation');
  assert.ok(about>=0&&about<cues&&cues<prep,`${about} < ${cues} < ${prep}`);
  assert.match(a.n('guide-status').textContent,/no Graphic Cues heading, so one was added/);
});

test('the cue pictures are decorative, so the name beside each one is what a screen reader reads',()=>{
  const entries=cuesUsed([P('Time: 5 minutes. Cumulative time: 0 minutes.','OSG Time'),P('Say   Hello','OSG Cue Facilitator'),P('Ask   Why?','OSG Cue')]);
  const ctx=makeContext(),xml=cuesBody(entries,ctx);
  assert.deepEqual([...ctx.used].sort(),['questions','say','time-on-task']);
  assert.equal((xml.match(/adec:decorative/g)||[]).length,3);
  assert.equal((xml.match(/descr=""/g)||[]).length,3);
  assert.match(xml,/<w:tblHeader\/>/);
  assert.doesNotMatch(xml,/gridSpan|vMerge/);
  const icons=Object.fromEntries([...ctx.used].map(i=>[i,'AAAA']));
  const rows=paragraphsOf(buildPackage(xml,ctx,icons)).map(p=>`${p.style}: ${p.text.trim()}`);
  assert.ok(rows.includes('OSG Table Cell: Time')&&rows.includes('OSG Table Cell: Say')&&rows.includes('OSG Table Cell: Ask'));
});

test('a cue is told from one with a longer name that starts the same way',()=>{
  assert.equal(cueFromText('  Time on task   Work alone (10 min)').id,'time-on-task');
  assert.equal(cueFromText('Key point   Neutral questions').id,'key-point');
  assert.equal(cueFromText('Nothing like a cue'),null);
});

test('a block added later makes the cues page out of date, and building again swaps it, with one table',async()=>{
  const a=await guide();
  await click(a,'guide-build-cues');
  assert.match(a.n('guide-cues-note').textContent,/matches your guide/);
  await moveTo(a,startsWith('OSG Time'));
  await addBlock(a,'discussion','Talk about the case.');
  assert.match(a.n('guide-cues-note').textContent,/out of date/);
  await click(a,'guide-build-cues');
  assert.equal(a.err(),'');
  assert.equal(count(a,'OSG Table Head: Cue'),1);
  assert.ok(under(a,'Graphic Cues').includes('OSG Table Cell: Discussion'));
  assert.match(a.n('guide-cues-note').textContent,/matches your guide/);
  assert.match(a.n('guide-status').textContent,/The earlier one was replaced\./);
});

test('with no blocks yet the cues button says what to do and changes nothing',async()=>{
  const a=await started('Empty guide',{quirks:WORD_QUIRKS});
  const before=a.rows();
  await click(a,'guide-build-cues');
  assert.match(a.err(),/Add some blocks first/);
  assert.deepEqual(a.rows(),before);
});

test('the two pages sit side by side without touching each other',async()=>{
  const a=await guide();
  await click(a,'guide-build-cues');await click(a,'guide-build-glance');await click(a,'guide-build-cues');await click(a,'guide-build-glance');
  assert.equal(a.err(),'');
  assert.equal(count(a,'OSG Table Head: Cue'),1);
  assert.equal(count(a,'OSG Table Head: Module or lesson'),1);
  assert.equal(under(a,'Training at a Glance').filter(r=>r.startsWith('OSG Table Cell')).length,GLANCE_ROWS.length-3);
  assert.match(a.n('guide-glance-note').textContent,/matches/);
  assert.match(a.n('guide-cues-note').textContent,/matches/);
  assert.deepEqual(GLANCE_HEAD,['Module or lesson','Minutes','Starts at minute']);
});

test('Undo takes a built page out in one step',async()=>{
  const a=await guide();
  const before=a.rows();
  await click(a,'guide-build-glance');
  assert.ok(a.adapter.undo());
  assert.deepEqual(a.rows(),before.slice(0,before.length));
});

// ---- both pages through the review, the structure check and the learner workbook
test('with both pages built the review has nothing to fix and no table or picture problem',async()=>{
  const a=await guide();
  await click(a,'guide-build-glance');await click(a,'guide-build-cues');
  await click(a,'guide-open-review');
  assert.equal(a.err(),'');
  assert.deepEqual(a.n('guide-review-fix').children.map(li=>li.textContent),['Nothing to fix.']);
  assert.deepEqual(a.n('guide-review-look').children.map(li=>li.textContent).filter(t=>/table|header|picture|alt text|blank|merged/i.test(t)),[]);
});

test('with both pages built the document passes the structure check, both tables headed',async()=>{
  const a=await guide();
  await click(a,'guide-build-glance');await click(a,'guide-build-cues');
  const result=await checkDocx(await packageToDocx(await a.adapter.sourcePackage(),'Reframing conflict'));
  assert.deepEqual(result.findings.filter(f=>f.level==='fail'),[]);
  assert.deepEqual(result.findings.filter(f=>['TABLE-BLANK','TABLE-MERGED','EMPTY-PARAGRAPHS','TYPED-LIST','IMAGE-ALT'].includes(f.rule)&&f.level!=='pass'),[]);
  assert.equal(result.findings.filter(f=>f.rule==='TABLE-HEADER'&&f.level==='pass').length,2,'both tables have a header row');
});

test('the learner workbook leaves both pages out, and still passes the leak check',async()=>{
  const a=await guide();
  await addBlock(a,'objective','Reframe a disagreement.');
  await click(a,'guide-build-glance');await click(a,'guide-build-cues');
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true);
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64')),doc=parse(await zip.file('word/document.xml').async('string'));
  const text=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>textOf(p)).join('\n');
  for(const gone of ['Module or lesson','Starts at minute','Total teaching time','Graphic Cues','What it means','These pictures mark'])assert.ok(!text.includes(gone),`${gone} is not in the workbook`);
  assert.ok(text.includes('Reframe a disagreement.'),'the learner blocks are');
});

test('cuesState and cuesPlacement read the document, not a record',()=>{
  const prep=P('Program Preparation','Heading 2');
  assert.equal(cuesPlacement([P('T','OSG Guide Title'),P('About This Guide','Heading 2'),P('Words.'),prep,P('Prep text.')]).after,2);
  assert.deepEqual(cuesState([P('Graphic Cues','Heading 2'),P('Their own words about the icons.')]),{built:false,stale:false});
  assert.equal(cuesPlacement([P('Graphic Cues','Heading 2'),P('Their own words.')]).remove,null,'their text is never removed');
});
