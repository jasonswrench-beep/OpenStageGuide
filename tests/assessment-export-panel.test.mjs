import test from 'node:test';
import assert from 'node:assert/strict';
import {mountAssessmentExports} from '../src/assessment-export-panel.mjs';
import {mountAssessmentPanel} from '../src/assessment-panel.mjs';
import {assessmentDom,descendants} from './helpers/assessment-dom.mjs';
import {validateAssessmentDefinition} from '../src/assessment-scoring.mjs';
const definition=()=>validateAssessmentDefinition({schemaVersion:1,id:'synthetic',version:1,title:'Synthetic export',evidenceStatus:'not-established',items:[{id:'private-item',type:'semantic-differential',prompt:'Rate clarity',reverse:true,target:'A synthetic message',leftEndpoint:'Unclear',rightEndpoint:'Clear',highScoreMeaning:'PRIVATE_HIGH',options:[{id:'PRIVATE_A',label:'First label',code:1},{id:'PRIVATE_B',label:'Second label',code:2}]}],subscales:[]});
function setup(packageFiles=async()=>new Blob(['private immutable package bytes'])){const f=assessmentDom();let source={definition:definition(),fingerprint:'same-definition'},wordBusy=false;const downloads=[];const panel=mountAssessmentExports(f.document,{getSource:()=>source,isBusy:()=>wordBusy,download:(blob,name)=>downloads.push({blob,name}),now:()=> '2026-10-28T14:00:00.000Z',packageFiles});return {...f,panel,downloads,source:s=>{source=s;},busy:b=>{wordBusy=b;panel.renderBusy();},e:id=>f.n('export-'+id)};}
const click=(f,id)=>f.e(id).onclick();
const check=(f,id,value=true)=>{f.e(id).checked=value;f.e(id).onchange();};
function review(f,route='respondent'){if(f.e('route').value!==route){f.e('route').value=route;f.e('route').onchange();}click(f,'preview');assert.equal(f.e('error').hidden,true,f.e('error').textContent);check(f,'reviewed');if(route==='author')check(f,'private');}

test('respondent and separate PRIVATE guide reviews require authentic route consent and exclude responses from UI dependencies',async()=>{
  const f=setup();review(f);assert.doesNotMatch(f.e('output').textContent,/PRIVATE_HIGH|PRIVATE_A|Reverse keyed/u);await click(f,'download');assert.equal(f.downloads.length,1);const stamp=f.e('output').textContent;await click(f,'download');assert.equal(f.downloads.length,2);assert.notEqual(f.downloads[0].name,f.downloads[1].name);assert.equal(f.e('output').textContent,stamp);
  f.e('route').value='author';f.e('route').onchange();assert.equal(f.e('reviewed').checked,false);assert.equal(f.e('private').checked,false);click(f,'preview');check(f,'reviewed');await click(f,'download');assert.equal(f.downloads.length,2);assert.match(f.e('error').textContent,/PRIVATE/u);check(f,'private');await click(f,'download');assert.equal(f.downloads.length,3);assert.match(f.downloads[2].name,/PRIVATE/u);assert.match(f.e('output').textContent,/PRIVATE_HIGH/u);
});

test('eventless raw choices validity definition display and visibility tampering refuse while drafts survive',async()=>{
  for(const mutate of [f=>f.e('instructions').value='eventless',f=>f.e('language').value='de-DE',f=>f.e('route').value='author',f=>f.e('instructions').validity={customError:true},f=>f.source({definition:definition(),fingerprint:'changed unfinished buffer'}),f=>f.e('output').textContent+=' forged',f=>f.e('output').hidden=true,f=>f.e('output').setAttribute('style','display:none'),f=>f.e('section').hidden=true,f=>f.e('output').append(f.document.createElement('span'))]){
    const f=setup();review(f);mutate(f);await click(f,'download');assert.equal(f.downloads.length,0);assert.equal(f.e('error').hidden,false);if(f.e('instructions').value)assert.equal(f.e('instructions').value,'eventless');
  }
  const f=setup();click(f,'preview');f.e('reviewed').checked=true;await click(f,'download');assert.equal(f.downloads.length,0);assert.match(f.e('error').textContent,/Acknowledge/u);
});

test('async failed interrupted repeated and stale assembly never offers a file and retains raw choices',async()=>{
  for(const change of [f=>{f.e('instructions').value='retain me';f.e('instructions').oninput();},f=>{f.busy(true);f.busy(false);},f=>{f.document.getElementById('view-assessments').hidden=true;f.panel.renderBusy();f.document.getElementById('view-assessments').hidden=false;},f=>{check(f,'reviewed',false);check(f,'reviewed',true);},f=>f.e('output').setAttribute('style','opacity:0')]){
    let release;const f=setup(()=>new Promise(resolve=>{release=resolve;}));review(f);const running=click(f,'download');assert.equal(f.e('download').disabled,true);change(f);release(new Blob(['bytes']));await running;assert.equal(f.downloads.length,0);assert.equal(f.e('error').hidden,false);
  }
  let release;const f=setup(()=>new Promise(resolve=>{release=resolve;}));review(f);const first=click(f,'download');await click(f,'download');release(new Blob(['bytes']));await first;assert.equal(f.downloads.length,1);
  const failure=setup(async()=>{throw new Error('synthetic ZIP failure');});review(failure);await click(failure,'download');assert.equal(failure.downloads.length,0);assert.match(failure.e('error').textContent,/ZIP failure/u);assert.equal(failure.e('reviewed').checked,true);
});

test('actual assessment capture integration excludes entered answers, guards pending choices and retains unrelated drafts',async()=>{
  const f=assessmentDom(),downloads=[];const panel=mountAssessmentPanel(f.document,{download:(blob,name)=>downloads.push({blob,name}),packageFiles:async()=>new Blob(['bytes']),now:()=> '2026-10-28T14:00:00.000Z'});
  f.n('title').value='Captured synthetic';f.n('title').oninput();f.n('item-prompt').value='Public prompt';f.n('item-prompt').oninput();f.n('item-save').onclick();f.n('review').onclick();f.n('reviewed').checked=true;f.n('reviewed').onchange();f.n('capture').onclick();
  f.n('export-preview').onclick();f.n('export-reviewed').checked=true;f.n('export-reviewed').onchange();const text=f.n('export-output').textContent;
  const radio=descendants(f.n('responses')).find(n=>n.type==='radio');radio.checked=true;radio.onchange();panel.renderBusy();assert.equal(f.n('export-reviewed').checked,true);assert.equal(f.n('export-output').textContent,text);await f.n('export-download').onclick();assert.equal(downloads.length,1);assert.equal(radio.checked,true);
  f.n('new').onclick();assert.equal(f.n('confirm').hidden,false);await f.n('export-download').onclick();assert.equal(downloads.length,1);assert.match(f.n('export-error').textContent,/pending|review changed/u);f.n('keep').onclick();assert.equal(f.n('title').value,'Captured synthetic');assert.equal(radio.checked,true);
  f.n('item-prompt').value='UNSAVED';f.n('item-prompt').oninput();f.n('export-preview').onclick();assert.match(f.n('export-error').textContent,/unfinished|matching/u);assert.equal(f.n('item-prompt').value,'UNSAVED');
});

test('actual tester answer and scoring during async export preserve review because export never reads responses',async()=>{
  const f=assessmentDom(),downloads=[];let release;
  mountAssessmentPanel(f.document,{download:(blob,name)=>downloads.push({blob,name}),packageFiles:()=>new Promise(resolve=>{release=resolve;}),now:()=> '2026-10-28T14:00:00.000Z'});
  f.n('title').value='Async response isolation';f.n('title').oninput();f.n('item-prompt').value='A public prompt';f.n('item-prompt').oninput();f.n('item-save').onclick();f.n('review').onclick();f.n('reviewed').checked=true;f.n('reviewed').onchange();f.n('capture').onclick();
  f.n('export-preview').onclick();f.n('export-reviewed').checked=true;f.n('export-reviewed').onchange();const review=f.n('export-output').textContent,running=f.n('export-download').onclick();
  const radio=descendants(f.n('responses')).find(n=>n.type==='radio');radio.checked=true;radio.onchange();f.n('score').onclick();assert.match(f.n('score-output').textContent,/keyed 1/u);assert.equal(f.n('export-reviewed').checked,true);release(new Blob(['bytes']));await running;assert.equal(downloads.length,1);assert.equal(f.n('export-output').textContent,review);assert.equal(radio.checked,true);
});

test('new assessment confirmation explicitly retains export option drafts while clearing only reviewed association',()=>{
  const f=assessmentDom();mountAssessmentPanel(f.document);
  f.n('export-instructions').value='KEEP organization instructions';f.n('export-instructions').oninput();f.n('export-route').value='author';f.n('export-route').onchange();f.n('export-language').value='fr-CA';f.n('export-language').onchange();
  f.n('new').onclick();assert.equal(f.n('confirm').hidden,false);assert.match(f.n('confirm-text').textContent,/Export instructions, copy choice and language are retained/u);f.n('discard').onclick();assert.equal(f.n('export-instructions').value,'KEEP organization instructions');assert.equal(f.n('export-route').value,'author');assert.equal(f.n('export-language').value,'fr-CA');assert.equal(f.n('export-reviewed').checked,false);assert.equal(f.n('export-private').checked,false);assert.match(f.n('status').textContent,/retained; review them/u);
});
