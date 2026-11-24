// What the guide remembers about the materials and decks made from it (src/session-record.mjs): fingerprints of the activities' words, what changed since an output was made,
// the status shown for it, and the member as the project file keeps it.
import test from 'node:test';
import assert from 'node:assert/strict';
import {makeContext,buildPackage,moduleHeading,lessonHeading,timeBlock,cueBlock} from '../src/guide-ooxml.mjs';
import {hashText,activityBlocks,fingerprints,diffSources,describeDiff,readSessions,sessionsText,emptySessions,recordOutput,outputStatus,setExcluded,removeOutput,saveDeck,removeDeck,OUTPUT_KINDS,LIMITS} from '../src/session-record.mjs';
import {paragraphsOf} from './helpers/guide-package.mjs';

const guideOf=(...blocks)=>{
  const ctx=makeContext();
  const xml=moduleHeading(1,'M')+lessonHeading(1,1,'L')+timeBlock({minutes:20,cumulative:0},ctx)+blocks.map(b=>cueBlock({cue:'activity',...b},ctx)).join('')+cueBlock({cue:'transition',text:'On.'},ctx);
  return paragraphsOf(buildPackage(xml,ctx,Object.fromEntries([...ctx.used].map(i=>[i,'AAAA']))));
};
const A={text:'Opening round',minutes:5,parts:{instructions:'Say your name.',output:'A list of hopes',materials:'Flip chart',steps:'5 min: Go around'}};
const B={text:'Rehearsal',minutes:10,parts:{instructions:'Run it.',grouping:'Triads',debrief:'What changed?'},notes:'Keep a spare.'};

test('the same words make the same fingerprint, and any change to them makes another',()=>{
  assert.equal(hashText('abc'),hashText('abc'));
  assert.notEqual(hashText('abc'),hashText('abd'));
  assert.match(hashText(''),/^[0-9a-f]{8}$/);
  const one=fingerprints(guideOf(A,B),'runsheet'),again=fingerprints(guideOf(A,B),'runsheet');
  assert.deepEqual(one,again);
  assert.deepEqual(one.map(s=>s.key),['opening round#1','rehearsal#1']);
  assert.ok(Object.keys(one[0].parts).includes('instructions')&&Object.keys(one[1].parts).includes('notes'));
});

test('a same-titled activity twice gets two keys, so each keeps its own fingerprint',()=>{
  const blocks=activityBlocks(guideOf(A,{...A,minutes:6}));
  assert.deepEqual(blocks.map(b=>b.key),['opening round#1','opening round#2']);
});

test('an output is only flagged by the words that matter to it',()=>{
  const before=guideOf(A,B);
  // the workbook shows learners' parts: a change to the facilitator's steps does not touch it, a change to the instructions does
  const stepsChanged=guideOf({...A,parts:{...A.parts,steps:'5 min: Go around twice'}},B);
  assert.ok(diffSources(fingerprints(before,'workbook'),fingerprints(stepsChanged,'workbook')).same);
  assert.ok(!diffSources(fingerprints(before,'runsheet'),fingerprints(stepsChanged,'runsheet')).same);
  const instructionsChanged=guideOf({...A,parts:{...A.parts,instructions:'Say your name and your role.'}},B);
  const d=diffSources(fingerprints(before,'workbook'),fingerprints(instructionsChanged,'workbook'));
  assert.deepEqual(d.changed,[{title:'Opening round',parts:['Instructions']}]);
  assert.deepEqual(describeDiff(d),['Opening round: Instructions changed.']);
  // only the minutes changed: the run sheet and the slides show them, the workbook and the checklist do not
  const longer=guideOf({...A,minutes:7},B);
  assert.ok(!diffSources(fingerprints(before,'runsheet'),fingerprints(longer,'runsheet')).same);
  assert.ok(!diffSources(fingerprints(before,'slides'),fingerprints(longer,'slides')).same);
  assert.ok(diffSources(fingerprints(before,'workbook'),fingerprints(longer,'workbook')).same);
  assert.ok(diffSources(fingerprints(before,'checklist'),fingerprints(longer,'checklist')).same);
  assert.throws(()=>fingerprints(before,'nope'),/not known/);
});

test('added, removed and reordered activities are named',()=>{
  const was=fingerprints(guideOf(A,B),'runsheet');
  const added=diffSources(was,fingerprints(guideOf(A,B,{text:'Exit check',minutes:5,parts:{instructions:'Write one thing.'}}),'runsheet'));
  assert.deepEqual(added.added,['Exit check']);assert.deepEqual(describeDiff(added),['Exit check is new in the guide.']);
  const removed=diffSources(was,fingerprints(guideOf(A),'runsheet'));
  assert.deepEqual(removed.removed,['Rehearsal']);assert.deepEqual(describeDiff(removed),['Rehearsal is not in the guide any more.']);
  const moved=diffSources(was,fingerprints(guideOf(B,A),'runsheet'));
  assert.equal(moved.reordered,true);assert.deepEqual(describeDiff(moved),['The activities are in a different order.']);
  assert.deepEqual(describeDiff(diffSources(was,was)),['Nothing in it has changed since it was made.']);
});

test('an output is Suggested before it is made, Ready after, Review needed when the guide changes, and Excluded when left out; one lesson only watches its own',()=>{
  const paras=guideOf(A,B);
  const s0=recordOutput(emptySessions(),{kind:'runsheet',name:'Run sheet',paras,now:()=>'2026-11-23T09:00:00Z'});
  assert.equal(s0.outputs.length,1);
  const o=s0.outputs[0];
  assert.equal(o.builtAt,'2026-11-23T09:00:00.000Z');
  assert.equal(outputStatus(o,paras).status,'Ready');
  assert.equal(outputStatus({...o,builtAt:''},paras).status,'Suggested');
  const edited=guideOf({...A,parts:{...A.parts,instructions:'Say your name and role.'}},B);
  const st=outputStatus(o,edited);
  assert.equal(st.status,'Review needed');assert.deepEqual(describeDiff(st.diff),['Opening round: Instructions changed.']);
  assert.equal(outputStatus({...o,excluded:true},edited).status,'Excluded');
  // made again, it is the same record refreshed, not a second one
  const s1=recordOutput(s0,{kind:'runsheet',name:'Run sheet',paras:edited,now:()=>'2026-11-23T10:00:00Z'});
  assert.equal(s1.outputs.length,1);assert.equal(s1.outputs[0].id,o.id);
  assert.equal(outputStatus(s1.outputs[0],edited).status,'Ready');
  // another kind or name is another record, newest first
  const s2=recordOutput(s1,{kind:'checklist',name:'Checklist',paras});
  assert.deepEqual(s2.outputs.map(x=>x.kind),['checklist','runsheet']);
  assert.equal(setExcluded(s2,s2.outputs[0].id,true).outputs[0].excluded,true);
  assert.equal(removeOutput(s2,s2.outputs[0].id).outputs.length,1);
  assert.deepEqual(Object.keys(OUTPUT_KINDS),['runsheet','checklist','workbook','cards','slides']);
});

test('the member is kept as text, read back whole, and anything not whole is left out rather than half used',()=>{
  const paras=guideOf(A,B);
  const s=recordOutput(emptySessions(),{kind:'workbook',name:'Workbook (Skills Practice)',scope:'The whole guide',paras});
  const saved=saveDeck(s,{title:'Scenario cards',type:'activity',paper:'A4',perPage:6,cropMarks:true,flip:'short',cards:[{front:['One'],back:['Back']},{front:['Two'],back:[]}]});
  const text=sessionsText(saved.sessions);
  assert.deepEqual(readSessions(text),saved.sessions);
  assert.equal(saved.sessions.decks[0].paper,'A4');assert.equal(saved.sessions.decks[0].cards[1].back.length,0);
  assert.deepEqual(readSessions(null),emptySessions());
  assert.deepEqual(readSessions('not json'),emptySessions());
  assert.deepEqual(readSessions(JSON.stringify({version:99,outputs:[{}]})),emptySessions());
  const dirty=readSessions(JSON.stringify({version:1,outputs:[{kind:'nope'},null,{id:'o1',kind:'runsheet',name:'  Run   sheet ',sources:[{key:'a',title:'A',hash:'abcd1234',parts:{ok:'abcd1234',bad:'zz',__proto__:'abcd1234'}},{key:1}]}],decks:[{cards:[]},{id:'d1',title:'T',perPage:3,cards:[{front:['x'.repeat(400)]},{front:[]}]}]}));
  assert.equal(dirty.outputs.length,1);assert.equal(dirty.outputs[0].name,'Run sheet');
  assert.deepEqual(dirty.outputs[0].sources[0].parts,{ok:'abcd1234'});assert.equal(dirty.outputs[0].sources.length,1);
  assert.equal(dirty.decks.length,2);assert.equal(dirty.decks[1].perPage,4,'a card count that is not allowed becomes the usual one');
  assert.equal(dirty.decks[1].cards.length,1);assert.equal(dirty.decks[1].cards[0].front[0].length,300);
});

test('a deck is saved, replaced by its id, removed, and refused when it has no card with words',()=>{
  const first=saveDeck(emptySessions(),{title:'A',cards:[{front:['x']}]});
  assert.equal(first.id,'d1');
  const again=saveDeck(first.sessions,{id:'d1',title:'A renamed',cards:[{front:['y']}]});
  assert.equal(again.sessions.decks.length,1);assert.equal(again.sessions.decks[0].title,'A renamed');
  const second=saveDeck(again.sessions,{title:'B',cards:[{front:['z']}]});
  assert.deepEqual(second.sessions.decks.map(d=>d.id),['d1','d2']);
  assert.deepEqual(removeDeck(second.sessions,'d1').decks.map(d=>d.id),['d2']);
  assert.throws(()=>saveDeck(emptySessions(),{title:'Empty',cards:[]}),/at least one card/);
  assert.throws(()=>saveDeck({...emptySessions(),decks:Array.from({length:LIMITS.decks},(_,i)=>({id:`d${i+1}`,title:'x',cards:[{front:['a']}]}))},{title:'Too many',cards:[{front:['a']}]}),/up to 30 decks/);
});
