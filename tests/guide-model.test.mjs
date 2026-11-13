import test from 'node:test';
import assert from 'node:assert/strict';
import {readGuide,updatePlan,guideIssues,outline,kindOf} from '../src/guide-model.mjs';
import {makeContext,sampleGuide,buildPackage} from '../src/guide-ooxml.mjs';
import {ICON_FILES} from '../src/guide-cues.mjs';
import {paragraphsOf,TINY} from './helpers/guide-package.mjs';

const S={module:'OSG Module',lesson:'OSG Lesson',time:'OSG Time',cue:'OSG Cue',fac:'OSG Cue Facilitator'};
const P=(style,text)=>({style,text});
const iconData=Object.fromEntries(ICON_FILES.map(i=>[i,TINY]));

test('the sample guide reads back as one module and one lesson whose times already agree',()=>{
  const ctx=makeContext(),paras=paragraphsOf(buildPackage(sampleGuide(ctx),ctx,iconData));
  const plan=updatePlan(paras);
  assert.equal(plan.guide.isGuide,true);
  assert.equal(plan.guide.title,'Sample guide: Handling workplace conflict');
  assert.equal(plan.guide.modules.length,1);
  assert.equal(plan.guide.modules[0].lessons.length,1);
  assert.equal(plan.guide.modules[0].lessons[0].blockMinutes,20);
  assert.equal(plan.totalMinutes,20);
  assert.deepEqual(plan.edits,[]);
  assert.deepEqual(guideIssues(plan),[]);
});

test('a document with no guide styles is not a guide',()=>{
  const plan=updatePlan([P('Normal','Hello'),P('heading 1','A heading')]);
  assert.equal(plan.guide.isGuide,false);
  assert.deepEqual(guideIssues(plan),[]);
  assert.deepEqual(plan.edits,[]);
});

test('cumulative time is the minutes taught before each lesson starts',()=>{
  const plan=updatePlan([
    P(S.module,'Module 1: A'),
    P(S.lesson,'Lesson 1.1: One'),P(S.time,'Time: 0 minutes. Cumulative time: 99 minutes.'),P(S.cue,'Ask What? (4 min)'),P(S.cue,'Say More (6 min)'),
    P(S.lesson,'Lesson 1.2: Two'),P(S.time,'Time: 5 minutes. Cumulative time: 0 minutes.'),P(S.cue,'Ask Why?'),
    P(S.lesson,'Lesson 1.3: Three'),P(S.time,'Time: 7 minutes. Cumulative time: 15 minutes.'),P(S.cue,'Ask How?')
  ]);
  assert.deepEqual(plan.edits,[
    {kind:'time',index:2,minutes:10,cumulative:0},
    {kind:'time',index:6,minutes:5,cumulative:10}
  ]);
  assert.equal(plan.totalMinutes,22);
  assert.deepEqual(outline(plan).map(r=>[r.kind,r.minutes,r.cumulative]),[['module',undefined,undefined],['lesson',10,0],['lesson',5,10],['lesson',7,15]]);
});

test('headings are renumbered in order, and a heading not in the numbered form is left alone',()=>{
  const plan=updatePlan([
    P(S.module,'Module 3: Reframe'),P(S.lesson,'Lesson 3.4: Questions'),P(S.time,'Time: 1 minute. Cumulative time: 0 minutes.'),P(S.cue,'Ask Why? (1 min)'),
    P(S.module,'Conflict'),P(S.lesson,'Lesson 9.9: Tone'),P(S.time,'Time: 2 minutes. Cumulative time: 1 minute.'),P(S.cue,'Ask Who? (2 min)')
  ]);
  assert.deepEqual(plan.edits,[
    {kind:'heading',index:0,text:'Module 1: Reframe'},
    {kind:'heading',index:1,text:'Lesson 1.1: Questions'},
    {kind:'heading',index:5,text:'Lesson 2.1: Tone'}
  ]);
  assert.ok(guideIssues(plan).some(i=>i.rule==='MODULE-NAME'&&i.paragraph===4));
});

test('a lesson with no block minutes keeps the time already typed in its time block',()=>{
  const plan=updatePlan([P(S.module,'Module 1: A'),P(S.lesson,'Lesson 1.1: One'),P(S.time,'Time: 12 minutes. Cumulative time: 0 minutes.'),P(S.cue,'Say Text')]);
  assert.deepEqual(plan.edits,[]);
  assert.equal(plan.totalMinutes,12);
});

test('a placeholder character before the time text does not stop it being read',()=>{
  const plan=updatePlan([P(S.module,'Module 1: A'),P(S.lesson,'Lesson 1.1: One'),P(S.time,'/ Time: 3 minutes. Cumulative time: 0 minutes.'),P(S.cue,'/ Transition Next')]);
  assert.equal(plan.guide.modules[0].lessons[0].timeMinutes,3);
  assert.equal(plan.guide.modules[0].hasTransition,true);
});

test('issues name what to fix and where',()=>{
  const plan=updatePlan([
    P('Title','My guide'),
    P(S.module,'Module 1: Empty'),
    P(S.module,'Module 2: Bare'),P(S.lesson,'Lesson 2.1: No time'),
    P(S.module,'Module 3: Fine'),P(S.lesson,'Lesson 3.1: Good'),P(S.time,'Time: 1 minute. Cumulative time: 0 minutes.'),P(S.cue,'Transition Next (1 min)')
  ]);
  const issues=guideIssues(plan),rules=issues.map(i=>i.rule+':'+i.paragraph);
  assert.ok(issues.some(i=>i.level==='fail'&&i.rule==='LESSON-TIME'&&i.paragraph===3));
  assert.ok(rules.includes('MODULE-EMPTY:1'));
  assert.ok(rules.includes('MODULE-TRANSITION:1'));
  assert.ok(rules.includes('MODULE-TRANSITION:2'));
  assert.ok(rules.includes('LESSON-EMPTY:3'));
  assert.ok(!rules.some(r=>r.endsWith(':4')&&r.startsWith('MODULE-TRANSITION')));
});

test('lessons that come before any module are reported, not lost',()=>{
  const plan=updatePlan([P(S.lesson,'Lesson 1.1: Loose'),P(S.time,'Time: 1 minute. Cumulative time: 0 minutes.'),P(S.cue,'Ask Why?')]);
  assert.equal(plan.guide.modules[0].index,-1);
  assert.ok(guideIssues(plan).some(i=>i.rule==='LESSON-NO-MODULE'));
});

test('paragraph kinds come from style names, with the built-in Title recognised',()=>{
  assert.equal(kindOf({style:S.fac}),'cue');
  assert.equal(kindOf({style:'Titel',builtIn:'Title'}),'title');
  assert.equal(kindOf({style:'Normal'}),'other');
});
