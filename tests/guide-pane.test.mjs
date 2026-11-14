import test from 'node:test';
import assert from 'node:assert/strict';
import {CUES} from '../src/guide-cues.mjs';
import {moduleHeading,cueBlock} from '../src/guide-ooxml.mjs';
import {pane,submit,started,WORD_QUIRKS,click,startsWith,moveTo,addModule,addLesson,addBlock,guideRows} from './helpers/guide-harness.mjs';

test('a document with no guide shows the Start screen and nothing else',async()=>{
  const a=pane();await a.p.refresh();
  assert.equal(a.n('guide-start').hidden,false);
  assert.equal(a.n('guide-build').hidden,true);
  assert.equal(a.n('guide-review').hidden,true);
});

test('Start needs a title, and refusing changes nothing',async()=>{
  const a=pane();await a.p.refresh();
  await submit(a,'guide-start-form');
  assert.match(a.err(),/Type a title for the guide/);
  assert.deepEqual(a.rows(),[]);
  assert.equal(a.n('guide-start').hidden,false);
});

test('Start puts the standard pages at the top, sets the document title and opens the Build screen',async()=>{
  const a=await started();
  assert.equal(a.rows()[0],'OSG Guide Title: Reframing conflict');
  assert.ok(a.rows().some(r=>r==='heading 1: Contents'));
  assert.equal(await a.adapter.title(),'Reframing conflict');
  assert.equal(a.n('guide-build').hidden,false);
  assert.equal(a.n('guide-start').hidden,true);
  assert.equal(a.n('guide-build-heading').textContent,'Reframing conflict');
  assert.match(a.n('guide-status').textContent,/Guide started/);
  assert.equal(a.err(),'');
});

test('existing text stays below the new pages, and the pane says so',async()=>{
  const a=pane([{text:'My earlier notes',style:'Normal'}]);await a.p.refresh();
  a.type('guide-title','New guide');await submit(a,'guide-start-form');
  assert.equal(a.rows().at(-1),'Normal: My earlier notes');
  assert.equal(a.rows()[0],'OSG Guide Title: New guide');
  assert.match(a.n('guide-status').textContent,/existing text is below/);
});

test('the sample guide is a finished example, and the pane can read it',async()=>{
  const a=pane();await a.p.refresh();
  await a.n('guide-sample').onclick();
  assert.equal(a.n('guide-build').hidden,false);
  assert.match(a.n('guide-total').textContent,/20 min/);
  assert.match(a.n('guide-update-note').textContent,/^Numbers and times are right\./,'the sample\'s numbers, times and contents table are all current');
});

test('add module, lesson and block build a guide in order, leaving the cursor where the next insert belongs',async()=>{
  const a=await started();
  await a.n('guide-add-module').onclick();
  assert.match(a.err(),/title for the module/);
  a.type('guide-new-title','Recognize the conflict');await a.n('guide-add-module').onclick();
  assert.deepEqual(a.rows().slice(-2).map(r=>r.split(':')[0]),['OSG Module','OSG Cue Facilitator']);
  assert.equal(a.rows().at(-2),'OSG Module: Module 1: Recognize the conflict');

  a.type('guide-lesson-title','Reframe the conversation');a.type('guide-new-minutes','10');await a.n('guide-add-lesson').onclick();
  const lessonAt=a.rows().findIndex(r=>r.startsWith('OSG Lesson'));
  assert.equal(a.rows()[lessonAt],'OSG Lesson: Lesson 1.1: Reframe the conversation');
  assert.equal(a.rows()[lessonAt+1],'OSG Time: Time: 10 minutes. Cumulative time: 0 minutes.');
  assert.equal(a.rows()[lessonAt+2].split(':')[0],'OSG Cue Facilitator','the module Transition stays after the lesson');

  a.pick('ask');a.type('guide-block-text','What makes it hard?');a.type('guide-block-notes','Listen for blame.');a.type('guide-block-minutes','4');
  await submit(a,'guide-block-form');
  assert.equal(a.rows()[lessonAt+2],'OSG Cue: Ask What makes it hard? (4 min)');
  assert.equal(a.rows()[lessonAt+3],'OSG Facilitator Notes: Facilitator notes: Listen for blame.');
  assert.equal(a.rows()[lessonAt+4].split(':')[0],'OSG Cue Facilitator');
  assert.equal(a.n('guide-block-text').value,'');
  assert.equal(a.err(),'');
});

test('choosing a cue sets its default audience, and the author can change it',async()=>{
  const a=await started();
  a.pick('say');
  assert.equal(a.n('guide-audience-facilitator').checked,true);
  a.pick('ask');
  assert.equal(a.n('guide-audience-everyone').checked,true);
  a.n('guide-audience-everyone').checked=false;a.n('guide-audience-facilitator').checked=true;
  a.type('guide-block-text','Who goes first?');await submit(a,'guide-block-form');
  assert.equal(a.rows().at(-1),'OSG Cue Facilitator: Ask Who goes first?');
});

test('a block needs text, and minutes must be a whole number',async()=>{
  const a=await started();
  await submit(a,'guide-block-form');assert.match(a.err(),/Type the text for this block/);
  a.type('guide-block-text','x');a.type('guide-block-minutes','soon');await submit(a,'guide-block-form');
  assert.match(a.err(),/whole number/);
  assert.ok(!a.rows().some(r=>r.includes(': Say x')));
});

test('a lesson needs a module around the cursor',async()=>{
  const a=await started();
  a.type('guide-lesson-title','Loose');await a.n('guide-add-lesson').onclick();
  assert.match(a.err(),/cursor inside a module/);
});

test('Update guide fixes cumulative times and numbers, once, and says how to undo',async()=>{
  const a=await started();
  a.type('guide-new-title','One');await a.n('guide-add-module').onclick();
  a.type('guide-lesson-title','First lesson');a.type('guide-new-minutes','10');await a.n('guide-add-lesson').onclick();
  a.pick('ask');a.type('guide-block-text','Why?');a.type('guide-block-minutes','4');await submit(a,'guide-block-form');
  assert.equal(a.n('guide-update').disabled,false);
  assert.match(a.n('guide-update-note').textContent,/1 number or time is out of date\./);
  await a.n('guide-update').onclick();
  assert.ok(a.rows().includes('OSG Time: Time: 4 minutes. Cumulative time: 0 minutes.'));
  assert.doesNotMatch(a.n('guide-update-note').textContent,/out of date/);
  assert.match(a.n('guide-status').textContent,/Ctrl\+Z/);
  a.type('guide-lesson-title','Second lesson');a.type('guide-new-minutes','6');await a.n('guide-add-lesson').onclick();
  assert.ok(a.rows().includes('OSG Time: Time: 6 minutes. Cumulative time: 4 minutes.'));
  assert.match(a.n('guide-total').textContent,/10 min/);
});

test('Contents lists modules and lessons, and choosing one moves the cursor there',async()=>{
  const a=await started();
  a.type('guide-new-title','One');await a.n('guide-add-module').onclick();
  a.type('guide-lesson-title','Lesson A');a.type('guide-new-minutes','10');await a.n('guide-add-lesson').onclick();
  const list=a.n('guide-contents').children;
  assert.equal(list.length,2);
  assert.match(list[1].children[0].textContent,/Lesson 1\.1: Lesson A \(10 min\)/);
  await list[0].children[0].onclick();
  assert.equal(a.adapter.cursor(),a.rows().findIndex(r=>r.startsWith('OSG Module')));
});

test('the cue list offers every cue, and the find box narrows it',async()=>{
  const a=await started();
  const labels=a.n('guide-cues').children;
  assert.equal(labels.length,CUES.length);
  a.type('guide-cue-find','tr');a.n('guide-cue-find').oninput();
  const shown=labels.filter(l=>!l.hidden).map(l=>l.dataset.label);
  assert.deepEqual(shown.sort(),['trainer notes','training instrument','transition']);
  assert.match(a.n('guide-cue-count').textContent,/3 cues match/);
  a.type('guide-cue-find','');a.n('guide-cue-find').oninput();
  assert.equal(labels.filter(l=>l.hidden).length,0);
});

test('Review lists what to fix and look at, with a way to go to each, and never claims to be Word\'s checker',async()=>{
  const a=await started();
  a.type('guide-new-title','One');await a.n('guide-add-module').onclick();
  a.type('guide-lesson-title','Empty lesson');a.type('guide-new-minutes','5');await a.n('guide-add-lesson').onclick();
  await a.n('guide-open-review').onclick();
  assert.equal(a.n('guide-review').hidden,false);
  assert.equal(a.n('guide-build').hidden,true);
  assert.match(a.n('guide-review-summary').textContent,/0 to fix, 1 to look at, 1 to check yourself\. The checks of the document itself .* are in Accessibility\./);
  const look=a.n('guide-review-look').children;
  assert.match(look[0].children[0].textContent,/Lesson 1\.1: Empty lesson has no blocks yet/);
  await look[0].children[1].onclick();
  assert.equal(a.rows()[a.adapter.cursor()],'OSG Lesson: Lesson 1.1: Empty lesson');
  await a.n('guide-review-back').onclick();
  assert.equal(a.n('guide-build').hidden,false);
});

test('Open Designer and its shortcut bring the Guide screen back from Review, Assessments and Workbook',async()=>{
  const a=await started();
  a.type('guide-new-title','One');await a.n('guide-add-module').onclick();
  await a.n('guide-open-review').onclick();
  assert.equal(a.n('guide-review').hidden,false);
  await a.p.showGuide();
  assert.equal(a.n('guide-review').hidden,true);
  assert.equal(a.n('guide-build').hidden,false);
  for(const tab of ['guide-tab-assess','guide-tab-workbook']){
    await a.n(tab).onclick();
    assert.equal(a.n('guide-build').hidden,true,`${tab} leaves the Guide screen`);
    await a.p.showGuide();
    assert.equal(a.n('guide-build').hidden,false,`${tab} then Open Designer returns to the Guide screen`);
  }
  // Already on the Guide screen, it only reads the document again.
  await a.p.showGuide();
  assert.equal(a.n('guide-build').hidden,false);
});

test('with no guide in the document, Open Designer from Review goes to the Start screen',async()=>{
  const a=pane();await a.p.refresh();
  await a.p.showReview();
  assert.equal(a.n('guide-review').hidden,false);
  await a.p.showGuide();
  assert.equal(a.n('guide-review').hidden,true);
  assert.equal(a.n('guide-start').hidden,false);
});

test('Review reports a missing time block as something to fix',async()=>{
  const a=pane([
    {text:'My guide',style:'OSG Guide Title'},{text:'Module 1: A',style:'OSG Module'},{text:'Lesson 1.1: B',style:'OSG Lesson'},{text:'Ask Why?',style:'OSG Cue'}
  ]);
  await a.p.refresh();await a.n('guide-open-review').onclick();
  const fixes=a.n('guide-review-fix').children.map(li=>li.children[0].textContent);
  assert.ok(fixes.some(t=>/Lesson 1\.1: B has no time block/.test(t)),fixes.join(' | '));
});

test('every action button is disabled while one is running, then restored',async()=>{
  const a=await started();
  let during;
  const original=a.adapter.insert.bind(a.adapter);
  a.adapter.insert=async(...args)=>{during=['guide-add-module','guide-block-go','guide-update'].map(id=>a.n(id).disabled);return original(...args);};
  a.type('guide-new-title','One');await a.n('guide-add-module').onclick();
  assert.deepEqual(during,[true,true,true]);
  assert.equal(a.n('guide-add-module').disabled,false);
});

test('a failed insert shows a plain message and leaves the buttons usable',async()=>{
  const a=await started();
  a.adapter.insert=async()=>{throw new Error('The picture for the ask cue could not be loaded. Check that the OpenStage Guide server is running, then try again.');};
  a.type('guide-block-text','x');await submit(a,'guide-block-form');
  assert.match(a.err(),/server is running/);
  assert.equal(a.n('guide-block-go').disabled,false);
});

// ---- the first native test (2026-10-31): Word dropped the last paragraph's style, blocks landed inside other blocks,
// ---- blocks could not be edited, and the pane did not follow Undo.

test('with a Word that drops the last paragraph\'s style, a whole guide still comes out in its styles and in order',async()=>{
  const a=await started('Reframing',{quirks:WORD_QUIRKS});
  a.adapter.place(0);
  await addModule(a,'One');
  await addLesson(a,'First','10');
  await addBlock(a,'ask','What makes it hard?',{notes:'Listen for blame.',minutes:'4'});
  await addBlock(a,'say','Name the feeling.');
  await addBlock(a,'key-point','Neutral words help.');
  assert.deepEqual(guideRows(a),[
    'OSG Module: Module 1: One',
    'OSG Lesson: Lesson 1.1: First',
    'OSG Time: Time: 10 minutes. Cumulative time: 0 minutes.',
    'OSG Cue: Ask What makes it hard? (4 min)',
    'OSG Facilitator Notes: Facilitator notes: Listen for blame.',
    'OSG Cue Facilitator: Say Name the feeling.',
    'OSG Cue: Key point Neutral words help.',
    'OSG Cue Facilitator: Transition Say how this module leads into the next one.'
  ]);
});

test('the in-memory Word reproduces the author\'s file when there is no spare paragraph (so the tests above would have caught it)',async()=>{
  const a=pane(undefined,{quirks:WORD_QUIRKS});
  await a.adapter.insert(c=>moduleHeading(1,'One')+cueBlock({cue:'transition',text:'Lead on.'},c),{where:'start',spare:false});
  assert.deepEqual(a.rows(),['OSG Module: Module 1: One','Normal: Transition Lead on.'],'the Transition keeps its words and loses its style');
});

test('a new block goes after the whole block the cursor is in, never between a cue and its notes',async()=>{
  const a=await started('Placement',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await addBlock(a,'ask','Why?',{notes:'Listen.',minutes:'3'});
  await moveTo(a,startsWith('OSG Cue: Ask'));
  await addBlock(a,'say','Name it.');
  const rows=guideRows(a);
  assert.deepEqual(rows.slice(3,6),['OSG Cue: Ask Why? (3 min)','OSG Facilitator Notes: Facilitator notes: Listen.','OSG Cue Facilitator: Say Name it.']);
  // With the cursor on the notes, the new block still goes after the Ask block (its cue and its notes), ahead of Say.
  await moveTo(a,startsWith('OSG Facilitator Notes'));
  await addBlock(a,'activity','In pairs.');
  assert.deepEqual(guideRows(a).slice(3,7),['OSG Cue: Ask Why? (3 min)','OSG Facilitator Notes: Facilitator notes: Listen.','OSG Cue: Activity In pairs.','OSG Cue Facilitator: Say Name it.']);
});

test('a block added with the cursor on a lesson heading goes after that lesson\'s time block',async()=>{
  const a=await started('Placement',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await moveTo(a,startsWith('OSG Lesson'));
  await addBlock(a,'objective','Reframe a disagreement.');
  assert.deepEqual(guideRows(a).slice(1,4),['OSG Lesson: Lesson 1.1: First','OSG Time: Time: 10 minutes. Cumulative time: 0 minutes.','OSG Cue: Objective Reframe a disagreement.']);
});

test('with the cursor on the module\'s Transition, a new block still goes in the last lesson, before the Transition',async()=>{
  const a=await started('Placement',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await addBlock(a,'ask','Why?');
  await moveTo(a,startsWith('OSG Cue Facilitator: Transition'));
  await addBlock(a,'key-point','Neutral words help.');
  assert.deepEqual(guideRows(a).slice(-3),['OSG Cue: Ask Why?','OSG Cue: Key point Neutral words help.','OSG Cue Facilitator: Transition Say how this module leads into the next one.']);
});

test('a new lesson goes after the lesson the cursor is in, not into the middle of it or after later lessons',async()=>{
  const a=await started('Placement',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First','5');
  await addBlock(a,'ask','Why?',{minutes:'5'});
  await addLesson(a,'Second','6');
  await moveTo(a,startsWith('OSG Cue: Ask'));
  await addLesson(a,'Between','7');
  const rows=guideRows(a),between=rows.findIndex(startsWith('OSG Lesson: Lesson 1.2: Between'));
  assert.deepEqual(rows.filter(startsWith('OSG Lesson')),['OSG Lesson: Lesson 1.1: First','OSG Lesson: Lesson 1.2: Between','OSG Lesson: Lesson 1.2: Second'],'the old second lesson is renumbered by Update guide');
  assert.equal(rows[between-1],'OSG Cue: Ask Why? (5 min)','it follows the whole of lesson 1, including its blocks');
  await click(a,'guide-update');
  assert.deepEqual(guideRows(a).filter(startsWith('OSG Lesson')),['OSG Lesson: Lesson 1.1: First','OSG Lesson: Lesson 1.2: Between','OSG Lesson: Lesson 1.3: Second']);
  assert.ok(guideRows(a).includes('OSG Time: Time: 7 minutes. Cumulative time: 5 minutes.'));
  assert.ok(guideRows(a).includes('OSG Time: Time: 6 minutes. Cumulative time: 12 minutes.'));
});

test('a lesson added with the cursor on the module heading joins the end of the module, before the Transition',async()=>{
  const a=await started('Placement',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await moveTo(a,startsWith('OSG Module'));
  await addLesson(a,'Second');
  const rows=guideRows(a);
  assert.deepEqual(rows.filter(startsWith('OSG Lesson')),['OSG Lesson: Lesson 1.1: First','OSG Lesson: Lesson 1.2: Second']);
  assert.equal(rows.at(-1).split(':')[0],'OSG Cue Facilitator','the Transition is still last');
});

test('a new module goes after the module the cursor is in, and the first goes after the standard pages wherever the cursor is',async()=>{
  const a=await started('Placement',{quirks:WORD_QUIRKS});
  a.adapter.place(0);await a.p.quietRefresh();
  await addModule(a,'One');
  const rows=a.rows(),glance=rows.findIndex(r=>r==='heading 2: Training at a Glance');
  assert.equal(rows[glance+1].startsWith('Normal: Give a short overview'),true);
  assert.equal(rows[glance+2],'OSG Module: Module 1: One','after the last standard page');
  await addLesson(a,'First');
  await addModule(a,'Two');
  await moveTo(a,startsWith('OSG Lesson'));
  await addModule(a,'Between');
  assert.deepEqual(a.rows().filter(startsWith('OSG Module')),['OSG Module: Module 1: One','OSG Module: Module 2: Between','OSG Module: Module 2: Two']);
  const between=a.rows().findIndex(startsWith('OSG Module: Module 2: Between'));
  assert.equal(a.rows()[between-1].startsWith('OSG Cue Facilitator: Transition'),true,'after module 1\'s Transition');
});

test('existing text below the standard pages stays below the first module',async()=>{
  const a=pane([{text:'My earlier notes',style:'Normal'}],{quirks:WORD_QUIRKS});await a.p.refresh();
  a.type('guide-title','New guide');await submit(a,'guide-start-form');
  await addModule(a,'One');
  assert.equal(a.rows().at(-1),'Normal: My earlier notes');
  assert.ok(a.rows().findIndex(startsWith('OSG Module'))<a.rows().length-1);
});

test('a Transition goes at the end of its module wherever the cursor is, and a second one is refused',async()=>{
  const a=await started('Placement',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');await addBlock(a,'ask','Why?');
  // Take the automatic Transition out by hand so the module has none, then ask for one with the cursor in the lesson.
  const kept=a.adapter.snapshot().filter(r=>!r.text.includes('Transition'));
  a.adapter.seed(kept);
  await moveTo(a,startsWith('OSG Time'));
  await addBlock(a,'transition','On to the next module.');
  assert.equal(a.rows().at(-1),'OSG Cue Facilitator: Transition On to the next module.');
  a.pick('transition');a.type('guide-block-text','A second one');await submit(a,'guide-block-form');
  assert.match(a.err(),/already has a Transition block/);
});
