import test from 'node:test';
import assert from 'node:assert/strict';
import {pane,submit,started,WORD_QUIRKS,click,startsWith,moveTo,addModule,addLesson,addBlock,guideRows} from './helpers/guide-harness.mjs';

// Editing what is already in the guide, and the pane keeping up with Word (first native test, 2026-10-31).
// The in-memory document stands in for Word; it is not Word.
const selectedText=a=>a.n('guide-selected').hidden?'':a.n('guide-selected-text').textContent;
async function guideWithBlock(){
  const a=await started('Editing',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First','10');
  await addBlock(a,'ask','What makes it hard?',{notes:'Listen for blame.',minutes:'4'});
  return a;
}

test('the pane says what the cursor is in and offers to edit it',async()=>{
  const a=await guideWithBlock();
  await moveTo(a,startsWith('OSG Cue: Ask'));
  assert.equal(selectedText(a),'Your cursor is in the Ask block.');
  assert.equal(a.n('guide-edit-selected').textContent,'Edit this block');
  await moveTo(a,startsWith('OSG Facilitator Notes'));
  assert.equal(selectedText(a),'Your cursor is in the Ask block.','the notes belong to the same block');
  await moveTo(a,startsWith('OSG Lesson'));
  assert.equal(selectedText(a),'Your cursor is in Lesson 1.1: First.');
  assert.equal(a.n('guide-edit-selected').textContent,'Edit this lesson');
  await moveTo(a,startsWith('heading 1: Contents'));
  assert.equal(selectedText(a),'');
});

test('Edit this block puts the block back in the form, and Update block replaces it in place',async()=>{
  const a=await guideWithBlock();
  await addBlock(a,'say','Name the feeling.');
  await moveTo(a,startsWith('OSG Cue: Ask'));
  await click(a,'guide-edit-selected');
  assert.equal(a.n('guide-block-text').value,'What makes it hard?');
  assert.equal(a.n('guide-block-notes').value,'Listen for blame.');
  assert.equal(a.n('guide-block-minutes').value,'4');
  assert.equal(a.n('guide-block-heading').textContent,'Edit a block');
  assert.equal(a.n('guide-block-go').textContent,'Update block');
  assert.equal(a.n('guide-block-cancel').hidden,false);
  assert.equal(a.radio('ask').checked,true);
  a.type('guide-block-text','What makes it hard to reframe?');a.type('guide-block-minutes','6');a.type('guide-block-notes','Listen for blame and tone.');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  assert.deepEqual(guideRows(a).slice(3,6),['OSG Cue: Ask What makes it hard to reframe? (6 min)','OSG Facilitator Notes: Facilitator notes: Listen for blame and tone.','OSG Cue Facilitator: Say Name the feeling.']);
  assert.equal(a.n('guide-block-heading').textContent,'Add a block');
  assert.equal(a.n('guide-block-go').textContent,'Insert block');
  assert.equal(a.n('guide-block-text').value,'');
  assert.match(a.n('guide-status').textContent,/Ask block updated\..*choose Edit, then Undo/);
  assert.match(a.n('guide-update-note').textContent,/1 number or time is out of date\./,'the lesson\'s time follows the new minutes');
});

test('an edited block can be taken back with Undo, and the pane follows',async()=>{
  const a=await guideWithBlock();
  await moveTo(a,startsWith('OSG Cue: Ask'));
  await click(a,'guide-edit-selected');
  a.type('guide-block-text','Changed');await submit(a,'guide-block-form');
  assert.ok(a.rows().some(r=>r.includes('Ask Changed')));
  a.adapter.undo();
  await a.adapter.fireSelection();
  assert.ok(a.rows().some(r=>r.includes('Ask What makes it hard?')));
});

test('Cancel edit puts the form back and changes nothing',async()=>{
  const a=await guideWithBlock();
  const before=a.rows();
  await moveTo(a,startsWith('OSG Cue: Ask'));
  await click(a,'guide-edit-selected');
  a.type('guide-block-text','Never mind');
  await click(a,'guide-block-cancel');
  assert.deepEqual(a.rows(),before);
  assert.equal(a.n('guide-block-go').textContent,'Insert block');
  assert.equal(a.n('guide-block-text').value,'');
  assert.match(a.n('guide-status').textContent,/Edit cancelled/);
});

test('Update block refuses if the block was changed in Word after it was chosen',async()=>{
  const a=await guideWithBlock();
  await moveTo(a,startsWith('OSG Cue: Ask'));
  await click(a,'guide-edit-selected');
  a.adapter.typeOver(a.rows().findIndex(startsWith('OSG Cue: Ask')),' Ask Someone retyped this');
  a.type('guide-block-text','Mine');await submit(a,'guide-block-form');
  assert.match(a.err(),/has changed since you chose it/);
  assert.ok(a.rows().some(r=>r.includes('Someone retyped this')));
});

test('Update block still finds the block if lines were added above it after it was chosen',async()=>{
  const a=await guideWithBlock();
  await moveTo(a,startsWith('OSG Cue: Ask'));
  await click(a,'guide-edit-selected');
  const snap=a.adapter.snapshot();snap.splice(1,0,{text:'A line typed above',style:'Normal'},{text:'Another',style:'Normal'});
  a.adapter.seed(snap);a.adapter.place(snap.findIndex(r=>r.style==='OSG Cue'));
  a.type('guide-block-text','Found again');await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  assert.ok(a.rows().some(r=>r.includes('Ask Found again')));
});

test('Edit this lesson changes the title and minutes, keeps the numbering, and brings every time up to date',async()=>{
  const a=await started('Editing',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First','10');await addLesson(a,'Second','5');
  await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));
  await click(a,'guide-edit-selected');
  assert.equal(a.n('guide-edit-form').hidden,false);
  assert.equal(a.n('guide-edit-title').value,'First');
  assert.equal(a.n('guide-edit-minutes').value,'10');
  a.type('guide-edit-title','Opening');a.type('guide-edit-minutes','20');
  await a.n('guide-edit-form').onsubmit({preventDefault(){}});
  assert.equal(a.err(),'');
  const rows=guideRows(a);
  assert.ok(rows.includes('OSG Lesson: Lesson 1.1: Opening'));
  assert.ok(rows.includes('OSG Time: Time: 20 minutes. Cumulative time: 0 minutes.'));
  assert.ok(rows.includes('OSG Time: Time: 5 minutes. Cumulative time: 20 minutes.'),'the next lesson\'s clock moved with it');
  assert.equal(a.n('guide-edit-form').hidden,true);
  assert.doesNotMatch(a.n('guide-update-note').textContent,/numbers? or times? (is|are) out of date/,'the lesson\'s own edit already brought every number and time up to date');
});

test('a lesson whose time comes from its blocks cannot have its minutes typed over, and says why',async()=>{
  const a=await guideWithBlock();
  await moveTo(a,startsWith('OSG Lesson'));
  await click(a,'guide-edit-selected');
  assert.equal(a.n('guide-edit-minutes').disabled,true);
  assert.match(a.n('guide-edit-note').textContent,/total of its blocks' minutes \(4\)/);
});

test('Edit this module renames it and leaves its number',async()=>{
  const a=await started('Editing',{quirks:WORD_QUIRKS});
  await addModule(a,'One');
  await moveTo(a,startsWith('OSG Module'));
  await click(a,'guide-edit-selected');
  assert.equal(a.n('guide-edit-minutes').hidden,true);
  a.type('guide-edit-title','Recognize the conflict');
  await a.n('guide-edit-form').onsubmit({preventDefault(){}});
  assert.ok(a.rows().includes('OSG Module: Module 1: Recognize the conflict'));
});

test('typing over a time block in Word is read back as it is, and Update guide puts the right words in',async()=>{
  const a=await guideWithBlock();
  a.adapter.typeOver(a.rows().findIndex(startsWith('OSG Time')),'Time: 10 minutes. Cumulative time: 99 minutes.');
  await a.p.quietRefresh();
  assert.match(a.n('guide-update-note').textContent,/1 number or time is out of date\./);
  await click(a,'guide-update');
  assert.ok(a.rows().includes('OSG Time: Time: 4 minutes. Cumulative time: 0 minutes.'));
  assert.match(a.n('guide-status').textContent,/Updated 1 number or time and the contents table\./);
});

test('Update guide also brings the contents table up to date, and the note says when it is out of date',async()=>{
  const a=await guideWithBlock();
  assert.match(a.n('guide-update-note').textContent,/The contents table is out of date\./,'the table was made before the module and lesson existed');
  assert.equal(a.n('guide-update').disabled,false);
  assert.match(a.n('guide-status').textContent,/Choose Update guide to bring the numbers, times and contents table up to date/);
  await click(a,'guide-update');
  const toc=a.rows().filter(r=>/^TOC \d:/.test(r));
  assert.ok(toc.some(r=>r.startsWith('TOC 1: Module 1: One')));
  assert.ok(toc.some(r=>r.startsWith('TOC 2: Lesson 1.1: First')));
  assert.doesNotMatch(a.n('guide-update-note').textContent,/contents table is out of date/);
  assert.match(a.n('guide-status').textContent,/Updated 1 number or time and the contents table\./);
});

test('Update guide still refreshes the contents table when the numbers and times are already right',async()=>{
  const a=await started('Contents',{quirks:WORD_QUIRKS});
  await addModule(a,'One');
  assert.equal(a.n('guide-update').disabled,false,'the button is never greyed out just because the numbers are fine');
  await click(a,'guide-update');
  assert.match(a.n('guide-status').textContent,/Numbers and times were already right\. The contents table is refreshed\./);
  assert.ok(a.rows().some(r=>r.startsWith('TOC 1: Module 1: One')));
});

test('if Word cannot refresh the contents table, Update guide says how to do it by hand',async()=>{
  const a=await guideWithBlock();
  a.adapter.refreshFields=async()=>false;
  await click(a,'guide-update');
  assert.match(a.n('guide-status').textContent,/right-click it and choose Update Field/);
});

test('numbers and times out of date are counted in plain words',async()=>{
  const a=await started('Counting',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First','5');await addLesson(a,'Second','6');
  await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));
  await addLesson(a,'Between','7');
  assert.match(a.n('guide-update-note').textContent,/2 numbers or times are out of date\./);
});

// When Word does not answer, or the cursor cannot be read.
test('a cursor that cannot be read does not stop the pane showing the guide',async()=>{
  const a=await guideWithBlock();
  a.adapter.cursorIndex=async()=>{throw new Error('The selection is not available.');};
  await a.p.refresh();
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-build').hidden,false);
  assert.equal(a.n('guide-selected').hidden,true);
});

test('if the first read of the document fails, the pane offers Try again, and Try again works',async()=>{
  const a=pane();
  const real=a.adapter.paragraphs;
  a.adapter.paragraphs=async()=>{throw new Error('Word did not answer.');};
  await a.p.refresh();
  assert.match(a.err(),/Word did not answer/);
  assert.equal(a.n('guide-retry-box').hidden,false);
  a.adapter.paragraphs=real;
  await a.n('guide-retry').onclick();
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-retry-box').hidden,true);
  assert.equal(a.n('guide-start').hidden,false);
});

test('a Word call that never answers gives the buttons back with a plain message',async()=>{
  const a=pane(undefined,{timeoutMs:20});
  await a.p.refresh();
  a.type('guide-title','Slow');
  a.adapter.insert=()=>new Promise(()=>{});
  await submit(a,'guide-start-form');
  assert.match(a.err(),/Word did not answer in time/);
  assert.equal(a.n('guide-start-go').disabled,false);
});

test('opening a guide that carries a reopen request from an earlier build removes it, so Word stops reporting a missing task pane',async()=>{
  const a=await guideWithBlock();
  const fresh=pane(a.adapter.snapshot());
  fresh.adapter.setAutoShow(true);
  await fresh.p.refresh();
  assert.equal(fresh.adapter.autoShow(),false);
});

test('the pane no longer asks Word to reopen it with the document, and removes an earlier request',async()=>{
  const a=pane();await a.p.refresh();
  a.adapter.setAutoShow(true);
  a.type('guide-title','Reopen');await submit(a,'guide-start-form');
  assert.equal(a.adapter.autoShow(),false,'a request left in the document by an earlier build is taken out');
  const b=await started('Fresh');
  assert.equal(b.adapter.autoShow(),false,'and a new guide never asks, because this add-in cannot open the pane that way yet');
});

// The pane keeps up with Word.
test('after Undo removes the whole guide, the pane goes back to Start a guide',async()=>{
  const a=await started('Undo me',{quirks:WORD_QUIRKS});
  a.p.watch();
  assert.equal(a.n('guide-build').hidden,false);
  a.adapter.undo();
  await a.adapter.fireSelection();
  assert.equal(a.n('guide-start').hidden,false);
  assert.equal(a.n('guide-build').hidden,true);
});

test('moving the cursor in Word updates what the pane says it is in',async()=>{
  const a=await guideWithBlock();
  a.p.watch();
  a.adapter.place(a.rows().findIndex(startsWith('OSG Lesson')));
  await a.adapter.fireSelection();
  assert.equal(selectedText(a),'Your cursor is in Lesson 1.1: First.');
  a.adapter.place(a.rows().findIndex(startsWith('OSG Cue: Ask')));
  await a.adapter.fireSelection();
  assert.equal(selectedText(a),'Your cursor is in the Ask block.');
});

test('the Contents list is not rebuilt when nothing in it changed, so the focus stays where it is',async()=>{
  const a=await guideWithBlock();
  const items=a.n('guide-contents').children;
  await a.p.quietRefresh();
  assert.equal(a.n('guide-contents').children,items,'the same list items, not new ones');
});

test('Go there follows the words when the document has changed since the review',async()=>{
  const a=pane([
    {text:'My guide',style:'OSG Guide Title'},{text:'Module 1: A',style:'OSG Module'},{text:'Lesson 1.1: B',style:'OSG Lesson'},{text:'Ask Why?',style:'OSG Cue'}
  ]);
  await a.p.refresh();await click(a,'guide-open-review');
  const go=a.n('guide-review-fix').children.find(li=>/has no time block/.test(li.children[0].textContent)).children[1];
  a.adapter.seed([{text:'Two new lines',style:'Normal'},{text:'Another',style:'Normal'},...a.adapter.snapshot()]);
  await go.onclick();
  assert.equal(a.err(),'');
  assert.equal(a.rows()[a.adapter.cursor()],'OSG Lesson: Lesson 1.1: B','it went to the lesson, not to whatever is now at the old number');
});

test('Go there says so plainly when the words it was sent to are gone',async()=>{
  const a=pane([
    {text:'My guide',style:'OSG Guide Title'},{text:'Module 1: A',style:'OSG Module'},{text:'Lesson 1.1: B',style:'OSG Lesson'},{text:'Ask Why?',style:'OSG Cue'}
  ]);
  await a.p.refresh();await click(a,'guide-open-review');
  const go=a.n('guide-review-fix').children.find(li=>/has no time block/.test(li.children[0].textContent)).children[1];
  a.adapter.seed([{text:'Nothing like it now',style:'Normal'}]);
  await go.onclick();
  assert.match(a.err(),/has changed since this list was made/);
});

test('Review reports only what it can see in the author\'s saved file, where Word had joined the time block to the Transition',async()=>{
  // Test 1.docx after the first native test, as the pane reads it. The pane cannot tell the paragraph is joined,
  // so it says the lesson has no time block; the real fix is that the join no longer happens.
  const a=pane([
    {text:'Test',style:'OSG Guide Title'},{text:'Module 1: Module 1',style:'OSG Module'},{text:'Lesson 1.1: Lesson 1',style:'OSG Lesson'},
    {text:' Time: 10 minutes. Cumulative time: 5 minutes. Transition Say how this module leads into the next one.',style:'Normal'},
    {text:' Ask test (5 min)',style:'OSG Cue'},{text:'Facilitator notes: test',style:'Normal'}
  ]);
  await a.p.refresh();await click(a,'guide-open-review');
  const fixes=a.n('guide-review-fix').children.map(li=>li.children[0].textContent);
  assert.ok(fixes.some(t=>/has no time block/.test(t)));
});
