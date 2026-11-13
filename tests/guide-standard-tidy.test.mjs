// The standard pages (Acknowledgements, The Program in Perspective, the Appendix) and the tidy tools (move and delete a whole
// lesson or module). These run against the in-memory document, not native Word: the real Word adapter is exercised against a
// stand-in in guide-word-tidy.test.mjs, and neither proves what Word itself does.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {pane,started,click,moveTo,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';
import {missingStandardPages,standardPageState,appendixPlacement,tidyTarget,tidyMove,tidyDeletion,lessonSpan,topSpan} from '../src/guide-model.mjs';
import {classify,summarize} from '../src/guide-workbook.mjs';
import {standardPages} from '../src/guide-ooxml.mjs';

const headings=a=>a.rows().filter(r=>/^OSG (Module|Lesson|Assessment): /.test(r)).map(r=>r.replace(/^OSG \w+: /,''));
const pageHeads=a=>a.rows().filter(r=>/^heading [12]: /i.test(r)).map(r=>r.replace(/^heading [12]: /i,''));
const times=a=>a.rows().filter(r=>r.startsWith('OSG Time: ')).map(r=>r.slice('OSG Time: '.length));
const idx=(a,row)=>a.rows().indexOf(row);

// Two modules; the first has three lessons, two of them with a block.
async function twoModules(){
  const a=await started();
  await addModule(a,'One');
  await addLesson(a,'Alpha','10');
  await addBlock(a,'ask','Alpha question?',{minutes:'4'});
  await addLesson(a,'Beta','6');
  await addBlock(a,'say','Beta talk.',{minutes:'3'});
  await addLesson(a,'Gamma','5');
  await addModule(a,'Two');
  await addLesson(a,'Delta','8');
  return a;
}

// ---- the standard pages
test('a new guide has Acknowledgements before the Contents, and The Program in Perspective before Program Preparation',async()=>{
  const a=await started();
  assert.deepEqual(pageHeads(a),['Acknowledgements','Contents','Getting Started','About This Guide','The Program in Perspective','Program Preparation','Training at a Glance']);
  assert.match(a.n('guide-standard-note').textContent,/Acknowledgements and The Program in Perspective are in your guide\. It has no Appendix yet/);
  assert.match(a.rows().join('\n'),/Thank the people and organizations/);
});

test('the Graphic Cues page goes between About This Guide and The Program in Perspective',async()=>{
  const a=await started();
  await addModule(a,'One');await addLesson(a,'Alpha');await addBlock(a,'ask','Why?');
  await click(a,'guide-build-cues');
  assert.equal(a.err(),'');
  assert.deepEqual(pageHeads(a).slice(2),['Getting Started','About This Guide','Graphic Cues','The Program in Perspective','Program Preparation','Training at a Glance']);
});

const OLD=[
  {text:'Old guide',style:'OSG Guide Title'},{text:'Facilitator guide',style:'Normal'},
  {text:'Contents',style:'heading 1'},{text:'Update this table to list the modules and lessons.',style:'Normal'},
  {text:'Getting Started',style:'heading 1'},{text:'About This Guide',style:'heading 2'},{text:'About text',style:'Normal'},
  {text:'Program Preparation',style:'heading 2'},{text:'Preparation text',style:'Normal'},
  {text:'Training at a Glance',style:'heading 2'},{text:'Glance text',style:'Normal'},
  {text:'Module 1: First',style:'OSG Module'}
];

test('a guide started before these pages existed can be given the ones it lacks, each in its place',async()=>{
  const a=pane(OLD);await a.p.refresh();
  assert.match(a.n('guide-standard-note').textContent,/^Missing: Acknowledgements and The Program in Perspective\./);
  await click(a,'guide-add-standard');
  assert.equal(a.err(),'');
  assert.deepEqual(pageHeads(a),['Acknowledgements','Contents','Getting Started','About This Guide','The Program in Perspective','Program Preparation','Training at a Glance']);
  const rows=a.rows();
  assert.equal(rows[idx(a,'heading 1: Acknowledgements')-1],'Normal: Facilitator guide','Acknowledgements follows the title lines');
  assert.match(a.n('guide-status').textContent,/Added Acknowledgements and The Program in Perspective\./);
  assert.match(a.n('guide-status').textContent,/Ctrl\+Z/);
  assert.match(a.n('guide-standard-note').textContent,/Acknowledgements and The Program in Perspective are in your guide\./);
  assert.equal(a.rows().filter(r=>/Thank the people/.test(r)).length,1);
  const before=a.rows();
  await click(a,'guide-add-standard');
  assert.match(a.n('guide-status').textContent,/All the standard pages are already in your guide\./);
  assert.deepEqual(a.rows(),before,'a second press changes nothing');
});

test('a Graphic Cues page that is already there stays above The Program in Perspective',async()=>{
  const rows=[...OLD.slice(0,7),{text:'Graphic Cues',style:'heading 2'},{text:'Cue intro',style:'Normal'},...OLD.slice(7)];
  const a=pane(rows);await a.p.refresh();
  await click(a,'guide-add-standard');
  assert.deepEqual(pageHeads(a),['Acknowledgements','Contents','Getting Started','About This Guide','Graphic Cues','The Program in Perspective','Program Preparation','Training at a Glance']);
  assert.equal(a.rows()[idx(a,'heading 2: The Program in Perspective')-1],'Normal: Cue intro');
});

test('the pages that cannot be placed are named, and nothing is invented for a guide with none of the pages around them',()=>{
  const bare=[{text:'Title',style:'OSG Guide Title'},{text:'Module 1: A',style:'OSG Module'}];
  const found=missingStandardPages(bare);
  assert.deepEqual(found.pages.map(p=>[p.page,p.after]),[['acknowledgements',0]],'only the title is needed for Acknowledgements');
  assert.deepEqual(found.cannot,['The Program in Perspective']);
  const none=missingStandardPages([{text:'Module 1: A',style:'OSG Module'}]);
  assert.deepEqual(none.pages,[]);
  assert.deepEqual(none.cannot,['Acknowledgements','The Program in Perspective']);
});

test('The Program in Perspective goes after About This Guide when the other pages are gone',()=>{
  const paras=OLD.filter(p=>!/Program Preparation|Preparation text|Training at a Glance|Glance text/.test(p.text));
  const found=missingStandardPages(paras);
  const perspective=found.pages.find(p=>p.page==='perspective');
  assert.equal(paras[perspective.after].text,'About text');
});

test('the Appendix is added at the end of the guide, once, and later modules still go before it',async()=>{
  const a=await twoModules();
  await click(a,'guide-add-appendix');
  assert.equal(a.err(),'');
  const rows=a.rows();
  assert.equal(rows.at(-2),'heading 1: Appendix');
  assert.match(rows.at(-1),/^Normal: Add reference material here/);
  assert.match(a.n('guide-standard-note').textContent,/It has an Appendix\./);
  const once=a.rows();
  await click(a,'guide-add-appendix');
  assert.deepEqual(a.rows(),once,'a second press adds nothing');
  assert.match(a.n('guide-status').textContent,/already has an Appendix/);
  assert.equal(a.rows()[a.adapter.cursor()],'heading 1: Appendix','and takes the cursor there');
  await addModule(a,'Three');
  const after=a.rows();
  assert.ok(after.indexOf('OSG Module: Module 3: Three')<after.indexOf('heading 1: Appendix'));
  assert.equal(after.at(-2),'heading 1: Appendix');
});

test('the Acknowledgements and the Appendix are left out of the learner workbook with the other front and own pages',async()=>{
  const paras=[
    {style:'OSG Guide Title',text:'Guide'},{style:'Normal',text:'Facilitator guide'},{style:'heading 1',text:'Acknowledgements'},{style:'Normal',text:'Thank the people.'},
    {style:'heading 1',text:'Contents'},{style:'heading 1',text:'Getting Started'},{style:'heading 2',text:'The Program in Perspective'},{style:'Normal',text:'Where it fits.'},
    {style:'OSG Module',text:'Module 1: A'},{style:'OSG Lesson',text:'Lesson 1.1: B'},{style:'OSG Cue',text:' Ask Why?'},
    {style:'heading 1',text:'Appendix'},{style:'Normal',text:'A job aid for the facilitator.'}
  ];
  const {decisions}=classify(paras);
  assert.deepEqual(decisions.slice(2,8).map(d=>d.keep),[false,false,false,false,false,false]);
  assert.equal(decisions[10].keep,true,'the learner block is kept');
  assert.equal(decisions[11].keep,false);assert.equal(decisions[12].keep,false,'the appendix is not for learners');
  assert.equal(summarize(paras).blocks,1);
});

test('the standard pages the sample guide shows include the new ones',()=>{
  const xml=standardPages('T');
  assert.ok(xml.indexOf('Acknowledgements')<xml.indexOf('Contents'));
  assert.ok(xml.indexOf('About This Guide')<xml.indexOf('The Program in Perspective')&&xml.indexOf('The Program in Perspective')<xml.indexOf('Program Preparation'));
});

// ---- what the tidy buttons act on
test('the tidy buttons show for a lesson, for a block in it, and for a module, and not elsewhere',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='heading 2: About This Guide');
  assert.equal(a.n('guide-tidy').hidden,true,'not in the Getting Started pages');
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.2: Beta');
  assert.equal(a.n('guide-tidy').hidden,false);
  assert.equal(a.n('guide-move-up').textContent,'Move lesson up');assert.equal(a.n('guide-delete').textContent,'Delete lesson');
  assert.match(a.n('guide-tidy-text').textContent,/Lesson 1\.2: Beta/);
  await moveTo(a,r=>/^OSG Cue Facilitator: .*Beta talk/.test(r));
  assert.match(a.n('guide-tidy-text').textContent,/Lesson 1\.2: Beta/,'a block belongs to its lesson');
  await moveTo(a,r=>r==='OSG Module: Module 1: One');
  assert.equal(a.n('guide-move-up').textContent,'Move module up');
  await moveTo(a,r=>/^OSG Cue Facilitator: .*Transition/.test(r));
  assert.match(a.n('guide-tidy-text').textContent,/Module 1: One/,'the Transition belongs to the module');
});

// ---- moving
test('Move lesson up swaps it with the lesson above, takes its blocks along, renumbers and re-times',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.2: Beta');
  await click(a,'guide-move-up');
  assert.equal(a.err(),'');
  assert.deepEqual(headings(a),['Module 1: One','Lesson 1.1: Beta','Lesson 1.2: Alpha','Lesson 1.3: Gamma','Module 2: Two','Lesson 2.1: Delta']);
  const beta=idx(a,'OSG Lesson: Lesson 1.1: Beta'),alpha=idx(a,'OSG Lesson: Lesson 1.2: Alpha');
  assert.match(a.rows().slice(beta,alpha).join('\n'),/Beta talk\./,'its block went with it');
  assert.doesNotMatch(a.rows().slice(beta,alpha).join('\n'),/Alpha question/);
  assert.match(a.rows().slice(alpha,idx(a,'OSG Lesson: Lesson 1.3: Gamma')).join('\n'),/Alpha question\?/);
  assert.deepEqual(times(a),['Time: 3 minutes. Cumulative time: 0 minutes.','Time: 4 minutes. Cumulative time: 3 minutes.','Time: 5 minutes. Cumulative time: 7 minutes.','Time: 8 minutes. Cumulative time: 12 minutes.']);
  assert.match(a.rows()[idx(a,'OSG Module: Module 2: Two')-1],/Transition/,'the Transition still ends the module');
  assert.equal(a.rows()[a.adapter.cursor()],'OSG Lesson: Lesson 1.1: Beta','the cursor is on the moved lesson, ready to move it again');
  assert.match(a.n('guide-status').textContent,/Lesson 1\.1: Beta moved up\. The numbers and times are up to date/);
  assert.match(a.n('guide-update-note').textContent,/^Numbers and times are right\./);
  assert.equal(a.n('guide-tidy-text').textContent.includes('Lesson 1.1: Beta'),true,'and the buttons follow it');
});

test('Move lesson down, then up again, puts everything back',async()=>{
  const a=await twoModules();
  await click(a,'guide-update'); // the times start out of date (blocks have minutes), and a move brings them up to date too
  const start=a.rows();
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.1: Alpha');
  await click(a,'guide-move-down');
  assert.deepEqual(headings(a).slice(1,4),['Lesson 1.1: Beta','Lesson 1.2: Alpha','Lesson 1.3: Gamma']);
  await click(a,'guide-move-up');
  assert.deepEqual(a.rows(),start);
});

test('a lesson stays in its own module: the first and last lesson say so and nothing changes',async()=>{
  const a=await twoModules();
  const start=a.rows();
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.1: Alpha');
  await click(a,'guide-move-up');
  assert.match(a.n('guide-status').textContent,/Lesson 1\.1: Alpha is already the first lesson in its module/);
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.3: Gamma');
  await click(a,'guide-move-down');
  assert.match(a.n('guide-status').textContent,/Lesson 1\.3: Gamma is already the last lesson in its module/);
  assert.deepEqual(a.rows(),start);
  assert.equal(a.err(),'');
});

test('Move module up swaps whole modules, with their lessons, overview and Transition, and renumbers every heading',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='OSG Module: Module 2: Two');
  await click(a,'guide-move-up');
  assert.equal(a.err(),'');
  assert.deepEqual(headings(a),['Module 1: Two','Lesson 1.1: Delta','Module 2: One','Lesson 2.1: Alpha','Lesson 2.2: Beta','Lesson 2.3: Gamma']);
  assert.equal(a.rows().filter(r=>/Transition/.test(r)&&r.startsWith('OSG Cue')).length,2);
  assert.match(a.rows()[idx(a,'OSG Module: Module 2: One')-1],/Transition/);
  assert.deepEqual(times(a)[0],'Time: 8 minutes. Cumulative time: 0 minutes.');
  assert.equal(a.rows()[a.adapter.cursor()],'OSG Module: Module 1: Two');
  await click(a,'guide-move-up');
  assert.match(a.n('guide-status').textContent,/Module 1: Two is already the first module/);
});

test('the last module cannot move down, and a module will not move into the Appendix',async()=>{
  const a=await twoModules();
  await click(a,'guide-add-appendix');
  const start=a.rows();
  await moveTo(a,r=>r==='OSG Module: Module 2: Two');
  await click(a,'guide-move-down');
  assert.match(a.n('guide-status').textContent,/Module 2: Two is already the last module/);
  assert.deepEqual(a.rows(),start);
});

test('moving with the cursor in no module or lesson says where to put it',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='heading 1: Contents');
  assert.equal(a.n('guide-tidy').hidden,true);
  await a.n('guide-move-up').onclick();
  assert.match(a.err(),/cursor inside a module or lesson/);
});

// ---- deleting
test('Delete asks first, names what will go, and Keep it leaves everything as it was',async()=>{
  const a=await twoModules();
  const start=a.rows();
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.2: Beta');
  await click(a,'guide-delete');
  assert.equal(a.n('guide-delete-confirm').hidden,false);
  assert.match(a.n('guide-delete-text').textContent,/Delete Lesson 1\.2: Beta\? This removes the lesson with its time block, its blocks and anything you typed in it: 3 paragraphs in all\. You can take it back: click in your document and press Ctrl\+Z several times\./);
  assert.deepEqual(a.rows(),start,'nothing is deleted yet');
  await click(a,'guide-delete-no');
  assert.equal(a.n('guide-delete-confirm').hidden,true);
  assert.match(a.n('guide-status').textContent,/Nothing was deleted/);
  assert.deepEqual(a.rows(),start);
});

test('Yes, delete it removes the lesson, renumbers and re-times, and the guide stays in step',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.2: Beta');
  await click(a,'guide-delete');await click(a,'guide-delete-yes');
  assert.equal(a.err(),'');
  assert.deepEqual(headings(a),['Module 1: One','Lesson 1.1: Alpha','Lesson 1.2: Gamma','Module 2: Two','Lesson 2.1: Delta']);
  assert.ok(!a.rows().some(r=>/Beta talk/.test(r)));
  assert.deepEqual(times(a),['Time: 4 minutes. Cumulative time: 0 minutes.','Time: 5 minutes. Cumulative time: 4 minutes.','Time: 8 minutes. Cumulative time: 9 minutes.']);
  assert.match(a.n('guide-status').textContent,/Lesson 1\.2: Beta deleted\. The numbers and times are up to date/);
  assert.equal(a.n('guide-delete-confirm').hidden,true);
  assert.match(a.n('guide-update-note').textContent,/^Numbers and times are right\./);
});

test('deleting a module takes its lessons, overview and Transition, and the next module becomes Module 1',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='OSG Module: Module 1: One');
  await click(a,'guide-delete');
  assert.match(a.n('guide-delete-text').textContent,/Delete Module 1: One\? This removes the module with its 3 lessons, its Transition block and anything you typed in it: /);
  await click(a,'guide-delete-yes');
  assert.deepEqual(headings(a),['Module 1: Two','Lesson 1.1: Delta']);
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Cue Facilitator')&&/Transition/.test(r)).length,1);
  assert.ok(!a.rows().some(r=>/Alpha question|Beta talk/.test(r)));
  assert.deepEqual(times(a),['Time: 8 minutes. Cumulative time: 0 minutes.']);
});

test('if the cursor moved after Delete was asked, nothing is deleted',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.2: Beta');
  await click(a,'guide-delete');
  const start=a.rows();
  a.adapter.place(a.rows().findIndex(r=>r==='OSG Lesson: Lesson 1.1: Alpha'));
  await click(a,'guide-delete-yes');
  assert.match(a.err(),/cursor has moved since you chose Delete, so nothing was deleted/);
  assert.deepEqual(a.rows(),start);
  assert.equal(a.n('guide-delete-confirm').hidden,true,'and the question is withdrawn');
});

test('moving the cursor to something else withdraws a Delete question',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.2: Beta');
  await click(a,'guide-delete');
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.3: Gamma');
  assert.equal(a.n('guide-delete-confirm').hidden,true);
});

// ---- the model: spans and moves over assessment sections and the Appendix
const P=(style,text)=>({style,text});
const MIXED=[
  P('OSG Guide Title','G'),P('heading 1','Getting Started'),
  P('OSG Module','Module 1: A'),P('OSG Overview','Goal: x'),P('OSG Lesson','Lesson 1.1: B'),P('OSG Time','Time: 5 minutes. Cumulative time: 0 minutes.'),P('OSG Cue',' Ask Why?'),P('Normal','typed by the author'),P('heading 3','a lower heading the author added'),P('Normal','more'),
  P('OSG Cue Facilitator',' Transition Next.'),P('Normal','typed after the Transition'),
  P('OSG Assessment','Level Two Assessment'),P('OSG Question','Question 1. Which?'),
  P('OSG Module','Module 2: C'),P('OSG Lesson','Lesson 2.1: D'),P('OSG Time','Time: 5 minutes. Cumulative time: 5 minutes.'),
  P('heading 1','Appendix'),P('Normal','job aid')
];

test('a lesson runs to the next lesson or the Transition, and a module to the next module, section or top-level heading',()=>{
  assert.deepEqual(lessonSpan(MIXED,4),{start:4,end:9},'the author\'s lines and a lower heading stay inside the lesson; the Transition does not');
  assert.deepEqual(topSpan(MIXED,2),{start:2,end:11},'the module takes its overview, lesson, Transition and typed text, not the assessment section');
  assert.deepEqual(topSpan(MIXED,12),{start:12,end:13});
  assert.deepEqual(topSpan(MIXED,14),{start:14,end:16},'the last module stops at the Appendix');
});

test('the cursor decides the target: the lesson it is in, else the module',()=>{
  assert.equal(tidyTarget(MIXED,7).kind,'lesson');
  assert.equal(tidyTarget(MIXED,3).kind,'module','the overview is the module\'s');
  assert.equal(tidyTarget(MIXED,10).kind,'module','and so is the Transition');
  assert.equal(tidyTarget(MIXED,2).kind,'module');
  assert.equal(tidyTarget(MIXED,12),null,'an assessment section is not moved by these buttons');
  assert.equal(tidyTarget(MIXED,1),null);
  assert.equal(tidyTarget(MIXED,-1),null);
});

test('a module moves down past an assessment section, and up past it again; nothing moves into the Appendix or above the title pages',()=>{
  const m1=tidyTarget(MIXED,2),m2=tidyTarget(MIXED,14);
  assert.deepEqual(tidyMove(MIXED,m1,1),{span:{start:2,end:11},after:13,over:'Level Two Assessment'});
  assert.deepEqual(tidyMove(MIXED,m2,-1),{span:{start:14,end:16},after:11,over:'Level Two Assessment'});
  assert.match(tidyMove(MIXED,m1,-1).message,/Module 1: A is already the first module/);
  assert.match(tidyMove(MIXED,m2,1).message,/Module 2: C is already the last module/);
});

test('deleting counts what goes',()=>{
  assert.deepEqual(tidyDeletion(MIXED,tidyTarget(MIXED,2)),{span:{start:2,end:11},paragraphs:10,lessons:1,text:'Module 1: A'});
  assert.equal(tidyDeletion(MIXED,tidyTarget(MIXED,7)).paragraphs,6);
});

test('standard page and appendix positions are read from the headings, ignoring case and punctuation',()=>{
  const s=standardPageState([P('heading 1','ACKNOWLEDGEMENTS'),P('Heading 2','The program in perspective!'),P('heading 1','Appendix')]);
  assert.deepEqual([s.acknowledgements,s.perspective,s.appendix],[0,1,2]);
  assert.deepEqual(appendixPlacement([P('Normal','a'),P('Normal',''),P('Normal','b'),P('Normal',''),P('Normal','')]),{exists:false,anchor:2});
  assert.deepEqual(appendixPlacement([P('heading 1','Appendix')]),{exists:true,at:0});
});

// ---- feedback the author can see (a button halfway down the pane used to answer in a line at the very top, out of sight)
test('on a new guide the build buttons say, beside themselves, what they are waiting for',async()=>{
  const a=await started();
  assert.match(a.n('guide-glance-note').textContent,/^Training at a Glance has not been built yet\. It needs a module and a lesson: add them first\./);
  assert.match(a.n('guide-cues-note').textContent,/^The Graphic Cues page has not been built yet\. It needs some blocks: add them first\./);
  await addModule(a,'One');
  assert.match(a.n('guide-glance-note').textContent,/needs a module and a lesson/,'a module alone is not enough for the Glance');
  assert.doesNotMatch(a.n('guide-cues-note').textContent,/needs some blocks/,'the module\'s Transition block is a cue, so the Graphic Cues page can be built');
  await addLesson(a,'Alpha');
  assert.doesNotMatch(a.n('guide-glance-note').textContent,/needs a module/);
});

test('pressing a build button too early still says why, in the message line of the pane',async()=>{
  const a=await started();
  await click(a,'guide-build-glance');assert.match(a.err(),/Add a module and a lesson first/);
  await click(a,'guide-build-cues');assert.match(a.err(),/Add some blocks first/);
  await click(a,'guide-build-overview');assert.match(a.err(),/Put your cursor inside a module first/);
  await click(a,'guide-add-standard');assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/All the standard pages are already in your guide\./);
});

test('the status and error lines are sticky at the top of the scrolling pane, so they are never out of sight',()=>{
  const css=fs.readFileSync(new URL('../public/style.css',import.meta.url),'utf8');
  assert.match(css,/\.guide-status:not\(:empty\),\.guide-error:not\(\[hidden\]\)\{position:sticky;top:0;/);
  assert.match(css,/main\{scroll-padding-top:/,'and a focused control does not land underneath them');
});

test('the empty paragraph that ends the document belongs to the document, not to the last module or lesson',()=>{
  const ENDING=[P('OSG Guide Title','G'),P('OSG Module','Module 1: A'),P('OSG Lesson','Lesson 1.1: B'),P('OSG Time','Time: 5 minutes. Cumulative time: 0 minutes.'),
    P('OSG Cue Facilitator',' Transition Next.'),P('OSG Module','Module 2: C'),P('OSG Lesson','Lesson 2.1: D'),P('OSG Time','Time: 5 minutes. Cumulative time: 5 minutes.'),P('Normal',''),P('Normal','')];
  assert.deepEqual(topSpan(ENDING,5),{start:5,end:7},'the last module stops before the two empty paragraphs that end the document');
  assert.deepEqual(lessonSpan(ENDING,6),{start:6,end:7});
  assert.deepEqual(tidyMove(ENDING,tidyTarget(ENDING,1),1),{span:{start:1,end:4},after:7,over:'Module 2: C'},'a module moved down goes in before them');
  assert.deepEqual(tidyMove(ENDING,tidyTarget(ENDING,5),-1),{span:{start:5,end:7},after:0,over:'Module 1: A'});
  // An empty paragraph in the middle of the guide is the author's spacer and stays inside its span.
  const MIDDLE=[P('OSG Module','Module 1: A'),P('OSG Lesson','Lesson 1.1: B'),P('Normal',''),P('OSG Module','Module 2: C')];
  assert.deepEqual(topSpan(MIDDLE,0),{start:0,end:2});
});

// ---- the contents note says why it thinks the table is out of date
import {contentsState} from '../src/guide-model.mjs';
test('an out-of-date contents table says whether it has too few lines or a line that differs, and a current one says nothing',()=>{
  const toc=(...texts)=>texts.map(t=>P('TOC 1',`${t}\t2`)),heads=[P('heading 1','Contents'),P('OSG Module','Module 1: A'),P('OSG Lesson','Lesson 1.1: B')];
  assert.deepEqual(contentsState([...toc('Contents','Module 1: A','Lesson 1.1: B'),...heads]),{present:true,stale:false});
  const few=contentsState([...toc('Contents','Module 1: A'),...heads]).why;
  assert.match(few,/^The table has 2 lines and the guide has 3 headings\. Just above its first line are /);
  assert.match(few,/ The first heading without a matching line is "Lesson 1\.1: B"/);
  assert.match(contentsState([...toc('Module 1: A','Lesson 1.1: B'),...heads]).why,/The first heading without a matching line is "Contents", where the table has "Module 1: A/);
  const differs=contentsState([...toc('Contents','Module 1: Old name','Lesson 1.1: B'),...heads]);
  assert.equal(differs.stale,true);
  assert.equal(differs.why,'Line 2 of the table reads "Module 1: Old name\t2" where the heading is "Module 1: A".'.replace('\t',' '));
});

test('the pane\'s note carries that reason after the sentence it always had',async()=>{
  const a=await started();
  await addModule(a,'One');await addLesson(a,'Alpha');
  a.adapter.place(a.rows().findIndex(r=>r==='OSG Lesson: Lesson 1.1: Alpha'));await a.p.quietRefresh();
  assert.match(a.n('guide-update-note').textContent,/The contents table is out of date\. The table has \d+ lines and the guide has \d+ headings\./);
});

test('the first line of the contents table is counted although the pane API of Word reports it with no style name',()=>{
  const paras=[P('heading 1','Contents'),P('','Contents	3'),P('TOC 1','Module 1: A	4'),P('TOC 2','Lesson 1.1: B	4'),P('Normal',''),P('OSG Module','Module 1: A'),P('OSG Lesson','Lesson 1.1: B')];
  assert.deepEqual(contentsState(paras),{present:true,stale:false});
  // A styleless line that is not a table line (no page number), or that is not beside the table, is not counted.
  assert.equal(contentsState([P('heading 1','Contents'),P('','just words'),P('TOC 1','Module 1: A	4'),P('OSG Module','Module 1: A')]).stale,true);
  assert.equal(contentsState([P('','Elsewhere	9'),P('Normal','x'),P('TOC 1','Module 1: A	4'),P('OSG Module','Module 1: A')]).stale,false,'the table is complete without it');
});

test('a move or a delete says to press Ctrl+Z several times, and the other actions still say once',async()=>{
  const a=await twoModules();
  await moveTo(a,r=>r==='OSG Lesson: Lesson 1.2: Beta');
  await click(a,'guide-move-up');
  assert.match(a.n('guide-status').textContent,/press Ctrl\+Z several times \(a move or a delete is done in several steps\) until it is as it was\./);
  await click(a,'guide-delete');await click(a,'guide-delete-yes');
  assert.match(a.n('guide-status').textContent,/deleted\..*press Ctrl\+Z several times/);
  a.type('guide-new-title','New');await a.n('guide-add-module').onclick();
  assert.match(a.n('guide-status').textContent,/click in your document and press Ctrl\+Z\.$/,'adding a module is one step');
});
test('the Start screen lists the standard pages by their real names',()=>{
  const html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');
  assert.match(html,/Add the standard pages <span class="hint">Acknowledgements, contents, Getting Started, Training at a Glance<\/span>/);
  assert.doesNotMatch(html,/Cover, contents/);
});
