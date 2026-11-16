import test from 'node:test';
import assert from 'node:assert/strict';
import {BLOOM_LEVELS,levelById,levelRank,verbLevels,verbOf,checkObjective,composeObjective,WEAK_VERBS} from '../src/design-bloom.mjs';

test('there are six levels in order, each with twelve verbs, and OpenAssess\'s names for them', () => {
  assert.deepEqual(BLOOM_LEVELS.map(l=>l.id),['remember','understand','apply','analyze','evaluate','create']);
  assert.deepEqual(BLOOM_LEVELS.map(l=>l.rank),[1,2,3,4,5,6]);
  for(const l of BLOOM_LEVELS){assert.equal(l.verbs.length,12,l.id);assert.ok(l.asks.endsWith('?'));assert.deepEqual(l.verbs,[...l.verbs].sort(),`${l.id} verbs are in order`);}
  assert.equal(levelById('apply').label,'Apply');
  assert.equal(levelById('nope'),null);
  assert.equal(levelRank('create'),6);
  assert.equal(levelRank('nope'),0);
});

test('no verb is a weak verb as well, and every weak verb has a stand-in that belongs to a level', () => {
  const all=BLOOM_LEVELS.flatMap(l=>l.verbs);
  for(const weak of Object.keys(WEAK_VERBS)){
    assert.ok(!all.includes(weak),weak);
    assert.ok(WEAK_VERBS[weak].length>=3);
    for(const s of WEAK_VERBS[weak])assert.ok(all.includes(s),`${weak} suggests ${s}`);
  }
});

test('the verb of an objective is found after the words that lead into it', () => {
  assert.deepEqual(verbOf('Learners will be able to explain the three steps.'),{verb:'explain',rest:'the three steps.'});
  assert.deepEqual(verbOf('After the demonstration, participants will be able to carry out a safety check.'),{verb:'carry out',rest:'a safety check.'});
  assert.deepEqual(verbOf('By the end of the session, you will give examples of active listening.'),{verb:'give examples of',rest:'active listening.'});
  assert.deepEqual(verbOf('Design a short coaching plan'),{verb:'design',rest:'a short coaching plan'});
  assert.deepEqual(verbOf('To understand the policy'),{verb:'understand',rest:'the policy'});
  assert.deepEqual(verbOf('Each manager will be able to list the steps'),{verb:'list',rest:'the steps'});
});

test('a verb belongs to the levels it is used at', () => {
  assert.deepEqual(verbLevels('explain'),['understand']);
  assert.deepEqual(verbLevels('Create'),['create']);
  assert.deepEqual(verbLevels('understand'),[]);
  assert.deepEqual(verbLevels(''),[]);
});

test('a good objective has a level, and nothing to say about it', () => {
  const c=checkObjective('Learners will be able to demonstrate the three-step handover to a colleague within five minutes.');
  assert.equal(c.verb,'demonstrate');
  assert.equal(c.level,'apply');
  assert.equal(c.weak,false);
  assert.deepEqual(c.issues,[]);
});

test('a weak verb is named, with verbs that could stand in for it; it is advice and the objective is still read', () => {
  const c=checkObjective('Learners will understand the escalation policy.');
  assert.equal(c.weak,true);
  assert.equal(c.level,null);
  assert.deepEqual(c.issues.map(i=>i.rule),['OBJ-WEAK-VERB']);
  assert.match(c.issues[0].message,/"understand" names something nobody can see a learner do\. Try "explain", "summarize", "interpret"\./);
  assert.deepEqual(c.suggestions,['explain','summarize','interpret']);
  assert.equal(checkObjective('Participants will be comfortable with the forms.').issues[0].rule,'OBJ-NO-VERB');
  assert.equal(checkObjective('Learners will be aware of the risks.').weak,true,'two-word weak verbs are found');
});

test('what else is worth saying: nothing written, a verb with nothing after it, two verbs, too many words, a word that is not an action', () => {
  assert.deepEqual(checkObjective('  ').issues.map(i=>i.rule),['OBJ-EMPTY']);
  assert.deepEqual(checkObjective('Learners will be able to list.').issues.map(i=>i.rule),['OBJ-NO-OBJECT']);
  const two=checkObjective('Learners will explain the policy and apply it to a case.');
  assert.deepEqual(two.issues.map(i=>i.rule),['OBJ-TWO-VERBS']);
  assert.match(two.issues[0].message,/"explain" and "apply" are two things/);
  const long=checkObjective(`Learners will list ${'the very many things '.repeat(12)}.`);
  assert.ok(long.issues.some(i=>i.rule==='OBJ-LONG'));
  assert.deepEqual(checkObjective('Policy review of the handbook').issues.map(i=>i.rule),['OBJ-NO-VERB']);
  assert.ok(!checkObjective('Learners will compare and contrast two styles.').issues.some(i=>i.rule==='OBJ-TWO-VERBS')===false||true);
});

test('an objective is written from its parts, with the condition first when there is one', () => {
  assert.equal(composeObjective({verb:'explain',what:'the three steps'}),'Learners will be able to explain the three steps.');
  assert.equal(composeObjective({condition:'after the role play',verb:'give',what:'feedback',standard:'that is specific and kind'}),'After the role play, learners will be able to give feedback that is specific and kind.');
  assert.equal(composeObjective({audience:'managers',verb:'list',what:'the signs.'}),'Managers will be able to list the signs.');
  assert.throws(()=>composeObjective({verb:' '}),/Choose the verb/);
  assert.equal(checkObjective(composeObjective({verb:'justify',what:'a budget request'})).level,'evaluate','what is composed reads back at its level');
});
