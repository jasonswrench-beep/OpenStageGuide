import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {checkPstat,readPstat} from '@openstage/pstat';
import {evaluationProject,evaluationIssues,evaluationTable,scalePoints,kindsForLevel,MEASURE_KINDS,EVALUATION_COLUMNS} from '../src/design-evaluation.mjs';
import {pstatBytes} from '../src/guide-pstat.mjs';
import {design} from './helpers/design-fixture.mjs';

test('a scale is its points: numbers need no labels, words become the labels of 1 to the last, and there are two to seven', () => {
  assert.deepEqual(scalePoints('1, 2, 3, 4, 5'),{labels:[],low:1,high:5,points:['1','2','3','4','5']});
  assert.deepEqual(scalePoints('Rarely, Sometimes, Often'),{labels:[{value:1,label:'Rarely'},{value:2,label:'Sometimes'},{value:3,label:'Often'}],low:1,high:3,points:['Rarely','Sometimes','Often']});
  assert.deepEqual(scalePoints('').points,['1','2','3','4','5'],'the usual five when none are given');
  assert.throws(()=>scalePoints('1'),/2 to 7 points/);
  assert.throws(()=>scalePoints('1,2,3,4,5,6,7,8'),/2 to 7 points/);
});

test('each level offers the kinds of measure that fit it', () => {
  assert.deepEqual(kindsForLevel(1).map(k=>k.id),['rating']);
  assert.deepEqual(kindsForLevel(2).map(k=>k.id),['pre-post','score']);
  assert.deepEqual(kindsForLevel(3).map(k=>k.id),['rating','score']);
  assert.deepEqual(kindsForLevel(4).map(k=>k.id),['metric']);
  assert.equal(MEASURE_KINDS.length,4);
});

test('the file has a variable for what each measure gives, labelled with its level, and empty rows for the people', () => {
  const p=evaluationProject(design(),{guideTitle:'Coaching first',rows:5,appVersion:'OpenStage Guide 0.1.0-test'});
  assert.deepEqual(p.dataset.variables.map(v=>v.name),['respondent','l1_end_of_day_q01','l1_end_of_day_q02','l2_escalation_quiz_pre','l2_escalation_quiz_post','l3_conversations_observed_q01','l4_escalations_per_month_before','l4_escalations_per_month_after']);
  assert.deepEqual(p.dataset.variables.map(v=>v.id),[1,2,3,4,5,6,7,8]);
  assert.equal(p.dataset.next_id,9);
  assert.equal(p.dataset.variables[1].label,'L1: The session met its goals.');
  assert.equal(p.dataset.variables[3].label,'L2: Escalation quiz, before');
  assert.deepEqual(p.dataset.variables[5].value_labels,[{value:1,label:'Rarely'},{value:2,label:'Sometimes'},{value:3,label:'Often'}]);
  assert.deepEqual(p.dataset.variables[1].value_labels,[]);
  assert.deepEqual(p.dataset.variables.map(v=>v.measure),['scale','ordinal','ordinal','scale','scale','ordinal','scale','scale']);
  assert.deepEqual(p.dataset.columns[0].values,[1,2,3,4,5]);
  assert.ok(p.dataset.columns.slice(1).every(c=>c.values.length===5&&c.values.every(v=>v===null)));
  assert.equal(p.app_version,'OpenStage Guide 0.1.0-test');
  assert.deepEqual([p.operations,p.results,'profile' in p],[[],[],false]);
});

test('the plan is in the notes the file carries: each level, who and when, the objectives, and what the author wrote last', () => {
  const notes=evaluationProject(design(),{guideTitle:'Coaching first',rows:2}).source.report.notes;
  assert.match(notes[0],/The evaluation plan of the guide "Coaching first", by Kirkpatrick's four levels\. The answers are empty/);
  assert.ok(notes.some(n=>/^Level 1: Reaction, "End of day": 2 statements rated 1, 2, 3, 4, 5 \(l1_end_of_day_q01 to l1_end_of_day_q02\); when: At the end of the second half day; of: Every learner\.$/.test(n)),notes.join(' | '));
  assert.ok(notes.some(n=>/^Level 2: Learning, "Escalation quiz": a score before \(l2_escalation_quiz_pre\) and after \(l2_escalation_quiz_post\); the gain is the second minus the first/.test(n)));
  assert.ok(notes.some(n=>/^"Escalation quiz" is for: Learners will be able to justify when to escalate\.$/.test(n)));
  assert.ok(notes.some(n=>/^Level 4: Results, "Escalations per month": a measure of the work before/.test(n)));
  assert.equal(notes.at(-1),'Ask again after three months.');
});

test('the file passes the shared core\'s checker and reads back whole', async () => {
  const bytes=await pstatBytes(evaluationProject(design(),{guideTitle:'Coaching first',rows:3}),JSZip);
  assert.deepEqual(checkPstat(bytes),[]);
  const back=readPstat(bytes);
  assert.equal(back.dataset.variables.length,8);
  assert.equal(back.source.report.rows,3);
  assert.equal(back.source.report.variables,8);
});

test('names that would meet are kept apart, and a measure with no name still gets variables', () => {
  const d=design();
  d.evaluation.measures.push({...d.evaluation.measures[0],id:'e5'});
  d.evaluation.measures.push({...d.evaluation.measures[3],id:'e6',name:''});
  const names=evaluationProject(d,{rows:0}).dataset.variables.map(v=>v.name);
  assert.equal(new Set(names.map(n=>n.toLowerCase())).size,names.length);
  assert.ok(names.includes('l1_end_of_day_q01_2'));
  assert.ok(names.includes('l4_measure_before'));
  assert.ok(names.every(n=>/^[A-Za-z][A-Za-z0-9_]{0,63}$/.test(n)));
});

test('the number of people is checked, and a plan with no measures says what to add', () => {
  assert.throws(()=>evaluationProject(design(),{rows:-1}),/whole number from 0 to 999/);
  assert.throws(()=>evaluationProject(design(),{rows:1000}),/whole number from 0 to 999/);
  const d=design();
  d.evaluation.measures=[];
  assert.throws(()=>evaluationProject(d),/no measures yet/);
});

test('what is missing from the plan is said in words: no measures, a measure with no statements or no time, a level left out', () => {
  assert.deepEqual(evaluationIssues(design()),[]);
  const d=design();
  d.evaluation.measures[0].statements=[];
  d.evaluation.measures[1].when='';
  d.evaluation.measures.pop();
  const issues=evaluationIssues(d);
  assert.ok(issues.includes('"End of day" is statements rated on a scale, and has no statements.'));
  assert.ok(issues.includes('"Escalation quiz" does not say when it is measured.'));
  assert.ok(issues.includes('Nothing is planned for level 4, Results.'));
  assert.deepEqual(evaluationIssues({...d,evaluation:{measures:[],notes:''}}),['There are no measures yet. Add one for each level you will evaluate.']);
});

test('the plan is a table for the guide, by level', () => {
  const t=evaluationTable(design());
  assert.deepEqual(t.head,[...EVALUATION_COLUMNS]);
  assert.deepEqual(t.body[0],['Level 1: Reaction','End of day','Statements rated on a scale','At the end of the second half day','Every learner']);
  assert.deepEqual(t.body.map(r=>r[0].slice(0,7)),['Level 1','Level 2','Level 3','Level 4']);
});
