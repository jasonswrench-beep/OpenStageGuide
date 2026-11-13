import test from 'node:test';
import assert from 'node:assert/strict';
import {readGuide,blockSpan,blockAnchor,lessonPlacement,modulePlacement,transitionAnchor,readBlock,describeAt} from '../src/guide-model.mjs';

const P=(text,style='Normal')=>({text,style});
// A small guide: two lessons in one module, the first with an Ask block (cue + two lines + notes), a Transition at the end.
const guide=[
  P('Cover','OSG Guide Title'),                                   // 0
  P('Training at a Glance','Heading 2'),                          // 1
  P('Give a short overview.'),                                    // 2
  P('Module 1: One','OSG Module'),                                // 3
  P('Lesson 1.1: First','OSG Lesson'),                            // 4
  P('\u0001 Time: 10 minutes. Cumulative time: 0 minutes.','OSG Time'), // 5
  P('\u0001 Ask\tWhat makes it hard? (4 min)','OSG Cue'),         // 6
  P('A second line of the question.','OSG Cue Body'),            // 7
  P('Facilitator notes: Listen for blame.','OSG Facilitator Notes'), // 8
  P('Some text the author wrote by hand.'),                       // 9
  P('Lesson 1.2: Second','OSG Lesson'),                           // 10
  P('\u0001 Time: 5 minutes. Cumulative time: 4 minutes.','OSG Time'), // 11
  P('\u0001 Transition\tOn to the next module.','OSG Cue Facilitator'), // 12
  P('Text below the guide that the author kept.')                 // 13
];

test('a block is its cue paragraph, any further lines and its notes, from whichever one the cursor is in',()=>{
  for(const at of [6,7,8])assert.deepEqual(blockSpan(guide,at),{start:6,end:8},`cursor at ${at}`);
  assert.equal(blockSpan(guide,5),null);
  assert.equal(blockSpan(guide,9),null);
});

test('the module, each lesson and the Transition know where they end',()=>{
  const g=readGuide(guide),mod=g.modules[0];
  assert.equal(mod.transitionIndex,12);
  assert.deepEqual(mod.lessons.map(l=>l.end),[8,11],'a lesson ends at its last block, not at text the author wrote after it');
  assert.equal(mod.end,12);
});

test('a new block goes after the whole block, after a lesson\'s time block, and before the Transition',()=>{
  assert.equal(blockAnchor(guide,6),8);
  assert.equal(blockAnchor(guide,7),8);
  assert.equal(blockAnchor(guide,4),5,'on a lesson heading: after its time block');
  assert.equal(blockAnchor(guide,5),5);
  assert.equal(blockAnchor(guide,9),9,'plain text the author wrote: right after it');
  assert.equal(blockAnchor(guide,12),11,'on the module\'s Transition: in the last lesson, before it');
});

test('a Transition goes at the module\'s end, and a module that has one refuses another',()=>{
  assert.throws(()=>transitionAnchor(guide,6),/already has a Transition block/);
  const without=guide.filter((_,i)=>i!==12);
  assert.equal(transitionAnchor(without,6).anchor,11);
  assert.throws(()=>transitionAnchor(guide.slice(0,3),1),/inside a module/);
});

test('a new lesson goes after the lesson the cursor is in, or at the end of the module from the module heading',()=>{
  assert.deepEqual(lessonPlacement(guide,7),{anchor:8,m:1,l:2,before:4,after:'Lesson 1.1: First'});
  assert.deepEqual(lessonPlacement(guide,11),{anchor:11,m:1,l:3,before:9,after:'Lesson 1.2: Second'});
  const onModule=lessonPlacement(guide,3);
  assert.equal(onModule.anchor,11);assert.equal(onModule.l,3);
  assert.throws(()=>lessonPlacement(guide,2),/inside a module/);
});

test('a new module goes after the module the cursor is in, or after the standard pages when there is none yet',()=>{
  assert.deepEqual(modulePlacement(guide,6),{anchor:12,n:2});
  assert.deepEqual(modulePlacement(guide,0),{anchor:12,n:2},'with the cursor in the front pages it still goes after the last module');
  assert.deepEqual(modulePlacement(guide.slice(0,3),0),{anchor:2,n:1},'the first module follows the "Training at a Glance" prompt');
  assert.deepEqual(modulePlacement([P('Cover','OSG Guide Title'),P('Hello')],1),{anchor:1,n:1},'without those pages it goes at the cursor');
});

test('a block is read back into its fields, including the picture placeholder Word leaves at the start',()=>{
  const b=readBlock(guide,7);
  assert.equal(b.cue,'ask');
  assert.equal(b.text,'What makes it hard?\nA second line of the question.');
  assert.equal(b.notes,'Listen for blame.');
  assert.equal(b.minutes,'4');
  assert.equal(b.audience,'everyone');
  assert.deepEqual(b.span,{start:6,end:8});
  const t=readBlock(guide,12);
  assert.equal(t.cue,'transition');assert.equal(t.audience,'facilitator');assert.equal(t.minutes,'');
  assert.equal(readBlock(guide,5),null);
});

test('a cue whose name starts another cue\'s name is not mistaken for it',()=>{
  const g=[P('Cover','OSG Guide Title'),P('\u0001 Training instrument\tComplete the survey.','OSG Cue'),P('\u0001 Trainer notes\tKeep time.','OSG Cue Facilitator')];
  assert.equal(readBlock(g,1).cue,'training-instrument');
  assert.equal(readBlock(g,2).cue,'trainer-notes');
});

test('what the cursor is in: a block, a lesson (heading or time block), a module, or nothing',()=>{
  assert.equal(describeAt(guide,8).label,'Ask block');
  assert.equal(describeAt(guide,4).type,'lesson');
  assert.equal(describeAt(guide,5).type,'lesson');
  assert.equal(describeAt(guide,3).type,'module');
  assert.equal(describeAt(guide,2),null);
});

// The author's recording of 2026-11-08: a module with no lesson, slides put in while the cursor was on its Transition, then Add module.
const noLesson=[
  P('Cover','OSG Guide Title'),                                          // 0
  P('Module 1: One','OSG Module'),                                       // 1
  P('\u0001 Transition\tOn to the next module.','OSG Cue Facilitator'), // 2
];
test('with the cursor on the Transition of a module that has no lesson, a new block goes inside the module, before the Transition',()=>{
  assert.equal(blockAnchor(noLesson,2),1,'straight after the module heading, so the next module cannot land above it');
  const withBlock=[...noLesson.slice(0,2),P('\u0001 Say\tWelcome.','OSG Cue'),noLesson[2]];
  assert.equal(blockAnchor(withBlock,3),2,'after the module\'s own blocks, still before the Transition');
});

test('a new module goes after the slides that sit straight after the previous module\'s Transition, not above them',()=>{
  const slides=[
    ...noLesson,
    P('\u0001 Slide\tSlide 1: Intro','OSG Cue Facilitator'),   // 3
    P('\uFFFC','OSG Cue Body'),                                 // 4 the picture
    P('Slide text: Hello','OSG Facilitator Notes'),             // 5
    P('Facilitator notes: Say hi.','OSG Facilitator Notes'),    // 6
    P('\u0001 Slide\tSlide 2: Next','OSG Cue Facilitator'),    // 7
    P('Facilitator notes: Go on.','OSG Facilitator Notes')      // 8
  ];
  for(const at of [2,3,6,8])assert.deepEqual(modulePlacement(slides,at),{anchor:8,n:2},`cursor at ${at}`);
  // Plain text the author wrote after the slides is theirs and does not move the module's end.
  const withText=[...slides,P('The author wrote this.')];
  assert.deepEqual(modulePlacement(withText,6),{anchor:8,n:2});
  // A module that follows is still a boundary: the new one goes before it, after the slides.
  const next=[...slides,P('Module 2: Two','OSG Module'),P('\u0001 Transition\tEnd.','OSG Cue Facilitator')];
  assert.deepEqual(modulePlacement(next,6),{anchor:8,n:2});
});
