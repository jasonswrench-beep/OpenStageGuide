import test from 'node:test';
import assert from 'node:assert/strict';
import {emptyDesign,cleanDesign,designText,readDesign,nextId,removeObjective,removeActivity,removeAssessment,storyboardMinutes,LIMITS,KIRKPATRICK} from '../src/design-model.mjs';
import {raw,design} from './helpers/design-fixture.mjs';

test('empty notes are what an unwritten guide has, and read back the same', () => {
  const d=emptyDesign();
  assert.equal(d.version,1);
  assert.deepEqual(cleanDesign(null),d);
  assert.deepEqual(cleanDesign({}),d);
  assert.equal(designText(d),null,'empty notes are not kept at all');
  assert.deepEqual(readDesign(null),d);
  assert.deepEqual(readDesign(''),d);
});

test('notes survive the text the project file keeps, whole', () => {
  const d=design();
  assert.deepEqual(readDesign(designText(d)),d);
  assert.equal(d.objectives.length,4);
  assert.equal(d.storyboard[0].lessons[1].activities[0],'a2');
});

test('what is not known is dropped, and what points at nothing is taken out', () => {
  const r=raw();
  r.surprise='x';r.needs.extra='y';
  r.activities[0].objectives=['o1','o99'];
  r.storyboard[0].lessons[0].activities=['a1','a99'];
  r.storyboard[0].lessons[0].assessments=['s99'];
  r.objectives[1].level='nonsense';
  r.assessments[0].kind='nonsense';
  const d=cleanDesign(r);
  assert.equal('surprise' in d,false);
  assert.equal('extra' in d.needs,false);
  assert.deepEqual(d.activities[0].objectives,['o1']);
  assert.deepEqual(d.storyboard[0].lessons[0].activities,['a1']);
  assert.deepEqual(d.storyboard[0].lessons[0].assessments,[]);
  assert.equal(d.objectives[1].level,'');
  assert.equal(d.assessments[0].kind,'quiz');
});

test('numbers are whole and in range, text is trimmed and limited, and a list has its limit', () => {
  const r=raw();
  r.activities[0].minutes=-5;
  r.activities[1].minutes=1000;
  r.activities[2].minutes='12';
  r.assessments[1].slots=51;
  r.objectives[0].text='  spaced   ';
  const d=cleanDesign(r);
  assert.deepEqual(d.activities.map(a=>a.minutes),[0,0,12]);
  assert.equal(d.assessments[1].slots,0);
  assert.equal(d.objectives[0].text,'spaced');
  const long=raw();
  long.objectives[0].text='x'.repeat(LIMITS.text+1);
  assert.throws(()=>cleanDesign(long),/An objective can be 4,000 characters at most/);
  const many=raw();
  many.objectives=Array.from({length:LIMITS.objectives+1},(_,i)=>({id:`o${i}`,text:'a'}));
  assert.throws(()=>cleanDesign(many),/There can be 80 objectives at most/);
});

test('notes that are not in the right form are refused with words, and newer notes ask for a newer program', () => {
  assert.throws(()=>cleanDesign('text'),/not in the right form/);
  assert.throws(()=>cleanDesign({objectives:'x'}),/not in the right form/);
  assert.throws(()=>cleanDesign({objectives:[1]}),/Objective 1 is not in the right form/);
  assert.throws(()=>cleanDesign({version:0}),/no usable version/);
  assert.throws(()=>cleanDesign({version:2}),/newer OpenStage Guide/);
  assert.throws(()=>readDesign('{not json'),/not readable, so the design desk was not opened/);
  const twice=raw();
  twice.objectives[1].id='o1';
  assert.throws(()=>cleanDesign(twice),/Two objectives have the same id/);
});

test('a new id is the first that is free', () => {
  assert.equal(nextId('o',[]),'o1');
  assert.equal(nextId('o',[{id:'o1'},{id:'o2'}]),'o3');
  assert.equal(nextId('o',[{id:'o1'},{id:'o3'}]),'o2');
  assert.equal(nextId('e',[{id:'e1'}],[{id:'e2'}]),'e3');
});

test('taking out an objective takes it out of every list that pointed at it, and leaves the rest', () => {
  const before=design();
  const d=removeObjective(before,'o4');
  assert.equal(d.objectives.length,3);
  assert.deepEqual(d.assessments[1].objectives,['o1']);
  assert.deepEqual(d.storyboard[1].lessons[0].objectives,['o3']);
  assert.deepEqual(d.evaluation.measures[1].objectives,[]);
  assert.equal(before.objectives.length,4,'the notes it was given are not changed');
  assert.deepEqual(cleanDesign(d),d,'what is left is still sound');
});

test('taking out an activity or an assessment takes it out of the lessons that held it', () => {
  const a=removeActivity(design(),'a1');
  assert.deepEqual(a.storyboard[0].lessons[0].activities,[]);
  assert.equal(a.activities.length,2);
  const s=removeAssessment(design(),'s2');
  assert.deepEqual(s.storyboard[0].lessons[0].assessments,[]);
  assert.equal(s.assessments.length,2);
});

test('the storyboard adds up its minutes, by module', () => {
  assert.deepEqual(storyboardMinutes(design()),{total:95,modules:[{id:'m1',title:'Talk first',minutes:75},{id:'m2',title:'Know the policy',minutes:20}]});
});

test('Kirkpatrick has four levels in order', () => {
  assert.deepEqual(KIRKPATRICK.map(k=>k.id),[1,2,3,4]);
  assert.ok(KIRKPATRICK.every(k=>k.label.startsWith(`Level ${k.id}: `)&&k.asks.endsWith('?')&&k.kinds.length));
});
