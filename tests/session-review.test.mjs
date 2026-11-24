// What the Review and the agenda say about Activity blocks written as parts (src/guide-model.mjs activityIssues, src/guide-agenda.mjs): steps that do not add up, a step with no
// minutes, a need on an activity that is not there, a catalog version that has been replaced, and the phase beside the time.
import test from 'node:test';
import assert from 'node:assert/strict';
import {makeContext,buildPackage,moduleHeading,lessonHeading,timeBlock,cueBlock} from '../src/guide-ooxml.mjs';
import {updatePlan,guideIssues} from '../src/guide-model.mjs';
import {guideAgenda,agendaFiles,hasPhases} from '../src/guide-agenda.mjs';
import {needFrom,readSteps,stepLine,readSource,sourceText,activityById} from '../src/session-text.mjs';
import {paragraphsOf} from './helpers/guide-package.mjs';

const guideWith=(...blocks)=>{
  const ctx=makeContext();
  const xml=moduleHeading(1,'M')+lessonHeading(1,1,'L')+timeBlock({minutes:20,cumulative:0},ctx)+blocks.map(b=>cueBlock({cue:'activity',...b},ctx)).join('')+cueBlock({cue:'transition',text:'On.'},ctx);
  return paragraphsOf(buildPackage(xml,ctx,Object.fromEntries([...ctx.used].map(i=>[i,'AAAA']))));
};
const rules=paras=>guideIssues(updatePlan(paras)).filter(i=>/^ACTIVITY-/.test(i.rule));

test('timed steps are read as minutes and words, and the lines with no minutes are told apart',()=>{
  assert.deepEqual(readSteps('3 min: Hand out cards\n2 min  Explain\n1 minute - Check\nno minutes here'),{steps:[{min:3,text:'Hand out cards'},{min:2,text:'Explain'},{min:1,text:'Check'}],without:['no minutes here'],minutes:6});
  assert.deepEqual(readSteps(''),{steps:[],without:[],minutes:0});
  assert.equal(stepLine({min:4,text:'Pair up'}),'4 min: Pair up');
  assert.deepEqual(readSource('Scenario rehearsal (A13, version 1)'),{id:'A13',version:1});
  assert.equal(readSource('made it myself'),null);
  assert.equal(sourceText(activityById('A13')),'Scenario rehearsal (A13, version 1)');
});

test('steps that add up to the activity\'s minutes raise nothing; steps that do not are named, with both totals',()=>{
  const ok=guideWith({text:'Round',minutes:5,parts:{instructions:'Go.',steps:'2 min: One\n3 min: Two'}});
  assert.deepEqual(rules(ok),[]);
  const short=guideWith({text:'Round',minutes:4,parts:{instructions:'Go.',steps:'2 min: One\n3 min: Two'}});
  const [issue]=rules(short);
  assert.equal(issue.rule,'ACTIVITY-STEPS');
  assert.match(issue.message,/"Round" add up to 5 minutes, but the activity is 4 minutes/);
  assert.equal(short[issue.paragraph].text.includes('Round'),true,'it points at the activity\'s own line');
  const bare=guideWith({text:'Round',minutes:5,parts:{instructions:'Go.',steps:'2 min: One\nthen something else'}});
  assert.deepEqual(rules(bare).map(i=>i.rule),['ACTIVITY-STEP-MINUTES']);
  // an activity with no minutes of its own, or with no steps, has nothing to compare
  assert.deepEqual(rules(guideWith({text:'Round',parts:{instructions:'Go.',steps:'2 min: One'}})),[]);
  assert.deepEqual(rules(guideWith({text:'Round',minutes:5,parts:{instructions:'Go.'}})),[]);
});

test('a need that names an earlier activity must find it: a renamed, moved or removed one is reported',()=>{
  const first={text:'Priority vote',minutes:5,parts:{instructions:'Vote.'}},second=name=>({text:'Decision record',minutes:5,parts:{instructions:'Write.',needs:`${needFrom(name)}: Ranked shortlist`}});
  assert.deepEqual(rules(guideWith(first,second('Priority vote'))),[]);
  assert.deepEqual(rules(guideWith(first,second('priority  VOTE'))),[],'case and spacing do not matter');
  const [gone]=rules(guideWith(first,second('Dot vote')));
  assert.equal(gone.rule,'ACTIVITY-NEEDS');
  assert.match(gone.message,/"Decision record" needs the output of "Dot vote"/);
  // after it, not before it
  assert.equal(rules(guideWith(second('Priority vote'),first)).length,1);
  // a plain need in words is not a reference and is left alone
  assert.deepEqual(rules(guideWith({text:'X',minutes:5,parts:{instructions:'Go.',needs:'The ranked shortlist'}})),[]);
});

test('an activity made from an older version of the catalog says so; one made from the current version does not',()=>{
  const a=activityById('A13');
  assert.deepEqual(rules(guideWith({text:a.title,minutes:5,parts:{instructions:'Go.',source:sourceText(a)}})),[]);
  const older=activityById('A13');
  const [issue]=rules(guideWith({text:a.title,minutes:5,parts:{instructions:'Go.',source:`${older.title} (A13, version 0)`}}));
  assert.equal(issue.rule,'ACTIVITY-SOURCE');
  assert.match(issue.message,/made from version 0.*now has version 1/);
  assert.deepEqual(rules(guideWith({text:'Mine',minutes:5,parts:{instructions:'Go.',source:'Some other book (page 4)'}})),[]);
});

test('the agenda shows each activity\'s phase beside its time, and the Word file gains a Phase column only when there is one',()=>{
  const phased=guideWith({text:'Warm up',minutes:5,parts:{instructions:'Go.',phase:'Open'}},{text:'Try it',minutes:15,parts:{instructions:'Do.',phase:'Practice'}});
  const a=guideAgenda(phased,'09:00');
  const blocks=a.rows.filter(r=>r.kind==='block');
  assert.deepEqual(blocks.map(r=>[r.text,r.phase,r.start,r.end]),[['Activity: Warm up','Open','09:00','09:05'],['Activity: Try it','Practice','09:05','09:20']]);
  assert.equal(hasPhases(a.rows),true);
  const doc=agendaFiles(a).find(f=>f.name==='word/document.xml').text;
  assert.match(doc,/<w:t[^>]*>Phase<\/w:t>/);
  assert.match(doc,/<w:t[^>]*>Practice<\/w:t>/);
  assert.equal((doc.match(/<w:gridCol /g)||[]).length,4);
  const plain=guideWith({text:'Warm up',minutes:5,parts:{instructions:'Go.'}});
  const b=guideAgenda(plain,'09:00');
  assert.equal(hasPhases(b.rows),false);
  const docPlain=agendaFiles(b).find(f=>f.name==='word/document.xml').text;
  assert.equal((docPlain.match(/<w:gridCol /g)||[]).length,3);
  assert.doesNotMatch(docPlain,/>Phase</);
});
