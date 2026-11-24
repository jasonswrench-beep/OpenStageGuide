// The part of optional AI help that does not depend on a model (src/session-assist.mjs): what is asked, what is checked, and that a helper that is missing, fails or answers badly changes nothing.
import test from 'node:test';
import assert from 'node:assert/strict';
import {assistContext,checkProposals,proposeChanges,withEdits,NO_PROVIDER} from '../src/session-assist.mjs';
import {planSlideImport,activityRewrite} from '../src/session-slide-import.mjs';
import {readBlock} from '../src/guide-model.mjs';
import {makeContext,buildPackage,moduleHeading,lessonHeading,timeBlock,cueBlock} from '../src/guide-ooxml.mjs';
import {paragraphsOf} from './helpers/guide-package.mjs';

const guideOf=()=>{
  const ctx=makeContext();
  const xml=moduleHeading(1,'M')+lessonHeading(1,1,'L')+timeBlock({minutes:15,cumulative:0},ctx)
    +cueBlock({cue:'activity',text:'Opening round',minutes:5,parts:{instructions:'Say your name.',output:'A list of hopes',source:'Opening check-in (A01, version 1)'}},ctx)
    +cueBlock({cue:'activity',text:'Rehearsal',minutes:10,parts:{instructions:'Run it.\nSwap roles.',grouping:'Triads'},notes:'Keep a spare.'},ctx)+cueBlock({cue:'transition',text:'On.'},ctx);
  return paragraphsOf(buildPackage(xml,ctx,Object.fromEntries([...ctx.used].map(i=>[i,'AAAA']))));
};

test('the helper is given only what the author chose, as plain text, without the source line or the private notes',()=>{
  const paras=guideOf();
  const c=assistContext(paras,{keys:['rehearsal#1'],request:'  Make the steps shorter.  '});
  assert.equal(c.request,'Make the steps shorter.');
  assert.equal(c.activities.length,1);
  assert.equal(c.activities[0].title,'Rehearsal');
  assert.deepEqual(Object.keys(c.activities[0].parts).sort(),['grouping','instructions']);
  assert.doesNotMatch(JSON.stringify(c),/Keep a spare|Opening|A01/);
  const all=assistContext(paras,{request:'Check the wording.'});
  assert.equal(all.activities.length,2);
  assert.ok(!JSON.stringify(all).includes('A01'),'the source line is never given');
  assert.throws(()=>assistContext(paras,{request:'  '}),/Say what you would like help with/);
  assert.throws(()=>assistContext(paras,{keys:['nope#1'],request:'x'}),/Choose at least one activity/);
});

test('with no helper set up, asking says so in words, sends nothing and changes nothing',async()=>{
  const paras=guideOf();
  const before=JSON.stringify(paras);
  for(const provider of [undefined,null,'a service']){
    const r=await proposeChanges(provider,paras,{request:'Make it shorter.'});
    assert.deepEqual(r,{ok:false,message:NO_PROVIDER});
  }
  assert.equal(JSON.stringify(paras),before);
  assert.match(NO_PROVIDER,/No AI helper is set up in this program, so nothing was sent anywhere and nothing was changed/);
});

test('a helper that works gives changes for the author to look at; each is checked, and an answer that is not whole is refused entirely',async()=>{
  const paras=guideOf();
  const provider=async ctx=>{
    assert.deepEqual(Object.keys(ctx),['request','activities']);
    return {proposals:[{key:'rehearsal#1',part:'instructions',to:'Run it once.\nSwap roles.'},{key:'opening round#1',part:'output',to:'A list of hopes and a theme'},{key:'rehearsal#1',part:'grouping',to:'Triads'}]};
  };
  const r=await proposeChanges(provider,paras,{request:'Tighten the wording.'});
  assert.equal(r.ok,true);
  assert.deepEqual(r.changes.map(c=>c.id),['rehearsal#1:instructions','opening round#1:output'],'a proposal that changes nothing is left out');
  assert.equal(r.changes[0].from,'Run it.\nSwap roles.');assert.equal(r.changes[0].label,'Instructions');
  assert.equal(r.changes.every(c=>c.conflict===false),true);
  const bad=answer=>proposeChanges(async()=>answer,paras,{request:'x'});
  assert.match((await bad('words')).message,/did not answer with a list of proposals, so nothing was changed/);
  assert.match((await bad([{key:'ghost#1',part:'output',to:'x'}])).message,/not given to the helper/);
  assert.match((await bad([{key:'rehearsal#1',part:'source',to:'x'}])).message,/"source", which a proposal may not change/);
  assert.match((await bad([{key:'rehearsal#1',part:'output',to:'  '}])).message,/has no words/);
  assert.match((await bad([{key:'rehearsal#1',part:'output',to:'x'.repeat(4001)}])).message,/longer than 4000 letters/);
  assert.match((await bad(Array.from({length:31},()=>({key:'rehearsal#1',part:'output',to:'x'})))).message,/more than can be looked at together/);
  assert.match((await bad([])).message,/had nothing to change/);
  assert.equal((await bad([null])).ok,false);
});

test('a helper that fails or cannot be reached leaves the session as it was, and says why',async()=>{
  const paras=guideOf();
  const before=JSON.stringify(paras);
  const r=await proposeChanges(async()=>{throw new Error('connection refused');},paras,{request:'x'});
  assert.equal(r.ok,false);assert.match(r.message,/could not be reached or failed \(connection refused\)\. Your session is as it was\./);
  assert.equal(JSON.stringify(paras),before);
});

test('the author can edit a proposal before accepting it, and what is accepted goes in through the same path as a slide import',async()=>{
  const paras=guideOf();
  const r=await proposeChanges(async()=>[{key:'rehearsal#1',part:'instructions',to:'Run it once.\nSwap roles.'}],paras,{request:'x'});
  const edited=withEdits(r.changes,{'rehearsal#1:instructions':'Run it once.\nSwap roles after three minutes.'});
  assert.equal(edited[0].to,'Run it once.\nSwap roles after three minutes.');
  assert.equal(withEdits(r.changes,{'rehearsal#1:instructions':'   '})[0].to,'Run it once.\nSwap roles.','an empty edit is ignored');
  const rw=activityRewrite(paras,'rehearsal#1',{[edited[0].part]:edited[0].to});
  assert.equal(rw.args.parts.instructions,'Run it once.\nSwap roles after three minutes.');
  assert.equal(rw.args.parts.grouping,'Triads');assert.equal(rw.args.notes,'Keep a spare.');
  assert.equal(readBlock(paras,rw.span.start).parts.instructions,'Run it.\nSwap roles.','proposing and editing write nothing');
  assert.ok(planSlideImport);
});

test('every proposal is for a part an author can edit, and the check has no way to touch the source line or a part that does not exist',()=>{
  const context={request:'x',activities:[{key:'a#1',title:'A',minutes:5,parts:{instructions:'Go.'}}]};
  for(const part of ['source','nope','__proto__','constructor'])assert.throws(()=>checkProposals([{key:'a#1',part,to:'x'}],context),/may not change/);
  assert.equal(checkProposals([{key:'a#1',part:'instructions',to:'Go again.'}],context)[0].id,'a#1:instructions');
});
