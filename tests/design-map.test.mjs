import test from 'node:test';
import assert from 'node:assert/strict';
import {planMap,guideMap,mapSummaryText,mapTable,mapCsv,MAP_COLUMNS} from '../src/design-map.mjs';
import {design} from './helpers/design-fixture.mjs';

test('the plan\'s map has a row for each objective, with its level, lesson, activities and assessments', () => {
  const map=planMap(design());
  assert.equal(map.kind,'plan');
  assert.equal(map.rows.length,4);
  const [o1,o2,o3,o4]=map.rows;
  assert.equal(o1.levelLabel,'Understand','a level from the verb when the author chose none');
  assert.equal(o2.levelLabel,'Apply','the author\'s own choice');
  assert.equal(o1.lesson,'The coaching conversation');
  assert.equal(o1.module,'Talk first');
  assert.deepEqual(o1.activities.map(a=>a.title),['Story of a hard conversation']);
  assert.deepEqual(o1.assessments.map(a=>a.title),['Short quiz']);
  assert.equal(o3.levelLabel,'','"understand" is not a level a verb gives');
  assert.deepEqual(map.summary,{objectives:4,withActivity:2,withAssessment:3,withBoth:2,withIssues:3});
});

test('a gap is said in words: no activity, no assessment, a weak verb, an assessment below the objective, no lesson', () => {
  const [o1,o2,o3,o4]=planMap(design()).rows;
  assert.deepEqual(o1.issues,[]);
  assert.deepEqual(o2.issues.map(i=>i.rule),['MAP-LEVEL-LOW']);
  assert.match(o2.issues[0].message,/"Skills check" checks at Remember, below this objective's Apply/);
  assert.deepEqual(o3.issues.map(i=>i.rule),['MAP-NO-LEVEL','MAP-WEAK-VERB','MAP-NO-ACTIVITY','MAP-NO-ASSESSMENT']);
  assert.match(o3.issues[1].message,/"understand" cannot be seen or measured/);
  assert.deepEqual(o4.issues.map(i=>i.rule),['MAP-NO-ACTIVITY','MAP-LEVEL-LOW'],'the quiz that checks it asks for less than it says');
});

test('an activity or an assessment that is for no objective is listed apart', () => {
  const {loose}=planMap(design());
  assert.deepEqual(loose.activities.map(a=>a.title),['Gallery walk']);
  assert.deepEqual(loose.assessments.map(a=>a.title),['Unlinked check']);
});

test('an objective with no words is said to have none, and one that is in no lesson says so when there is a storyboard', () => {
  const d=design();
  d.objectives.push({id:'o5',text:'',level:''});
  d.objectives.push({id:'o6',text:'Learners will be able to design a handout.',level:''});
  const rows=planMap(d).rows;
  assert.deepEqual(rows[4].issues.map(i=>i.rule).slice(0,1),['MAP-EMPTY']);
  assert.ok(rows[5].issues.some(i=>i.rule==='MAP-NO-LESSON'));
  d.storyboard=[];
  assert.ok(!planMap(d).rows[5].issues.some(i=>i.rule==='MAP-NO-LESSON'),'no storyboard, no lesson to be missing from');
});

test('the summary says how many have what, and what to do when there is nothing yet', () => {
  assert.match(mapSummaryText(planMap(design())),/^4 objectives: 2 with an activity, 3 with an assessment, 2 with both\.$/);
  assert.match(mapSummaryText(planMap({...design(),objectives:[]})),/Write them on the Needs and objectives screen/);
  assert.match(mapSummaryText(guideMap([])),/The guide has no Objective blocks yet/);
});

// A guide with two lessons: one objective practised and checked, one with neither.
const P=(text,style)=>({text,style});
const guide=[
  P('Cover','OSG Guide Title'),
  P('Module 1: One','OSG Module'),
  P('Lesson 1.1: First','OSG Lesson'),
  P('\u0001 Time: 30 minutes. Cumulative time: 0 minutes.','OSG Time'),
  P('\u0001 Objective\tLearners will be able to explain the three steps.','OSG Cue'),
  P('\u0001 Activity\tPractise the steps','OSG Cue'),
  P('Objective: Learners will be able to explain the three steps.','OSG Cue Body'),
  P('Instructions: Do it.','OSG Cue Body'),
  P('Lesson 1.2: Second','OSG Lesson'),
  P('\u0001 Time: 20 minutes. Cumulative time: 30 minutes.','OSG Time'),
  P('\u0001 Objective\tLearners will understand the policy.','OSG Cue'),
];

test('the guide\'s map has each Objective block with its level, lesson, module and minutes, and what alignment found for it', () => {
  const map=guideMap(guide);
  assert.equal(map.kind,'guide');
  assert.equal(map.rows.length,2);
  const [a,b]=map.rows;
  assert.equal(a.levelLabel,'Understand');
  assert.equal(a.lesson,'Lesson 1.1: First');
  assert.equal(a.module,'Module 1: One');
  assert.equal(a.minutes,30);
  assert.deepEqual(a.activities.map(x=>x.title),['Practise the steps']);
  assert.deepEqual(a.issues.map(i=>i.rule),['MAP-NO-ASSESSMENT']);
  assert.deepEqual(b.issues.map(i=>i.rule),['MAP-WEAK-VERB','MAP-NO-ACTIVITY','MAP-NO-ASSESSMENT']);
  assert.equal(b.minutes,20);
});

test('the map is a table, and a CSV a spreadsheet opens as UTF-8, with quotes where a cell needs them', () => {
  const t=mapTable(planMap(design()));
  assert.deepEqual(t.head,[...MAP_COLUMNS]);
  assert.equal(t.body.length,4);
  assert.deepEqual(t.body[0],['Talk first','The coaching conversation','Learners will be able to explain the three steps of a coaching conversation.','Understand','Story of a hard conversation','Short quiz']);
  const d=design();
  d.objectives[0].text='Learners will be able to explain "the steps", then list them.';
  const csv=mapCsv(planMap(d));
  assert.ok(csv.startsWith('﻿Module,Lesson,Objective,Level,Activities,Assessments\r\n'));
  assert.ok(csv.includes('"Learners will be able to explain ""the steps"", then list them."'));
  assert.equal(csv.split('\r\n').length,6,'a heading, four rows, and the end of the last line');
});
