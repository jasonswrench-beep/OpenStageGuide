import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {mountAssessmentPanel} from '../src/assessment-panel.mjs';
import {assessmentDom,descendants} from './helpers/assessment-dom.mjs';

function fixture(){const ui=assessmentDom();let busy=false;const panel=mountAssessmentPanel(ui.document,{isBusy:()=>busy});return {...ui,panel,busy:value=>{busy=value;panel.renderBusy();}};}
function field(f,id,value,event='oninput'){f.n(id).value=value;f.n(id)[event]?.();}
function click(f,id){f.n(id).onclick();}
function check(f,id,value=true){f.n(id).checked=value;f.n(id).onchange?.();}
function anchors(f){return f.n('anchors').children.map(row=>({row,label:row.children.find(n=>n.id.startsWith('measure-anchor-label-')),code:row.children.find(n=>n.id.startsWith('measure-anchor-code-')),up:row.children.find(n=>n.id.startsWith('measure-anchor-up-')),down:row.children.find(n=>n.id.startsWith('measure-anchor-down-')),remove:row.children.find(n=>n.id.startsWith('measure-anchor-remove-'))}));}
function savedItemIds(f){return f.n('item-list').children.filter(n=>n.value).map(n=>n.value);}
function savedRuleIds(f){return f.n('rule-list').children.filter(n=>n.value).map(n=>n.value);}
function members(f){return f.n('members').children.map(row=>row.children.find(n=>n.type==='checkbox'));}
function item(f,prompt,{type='likert',reverse=false}={}){field(f,'item-prompt',prompt);field(f,'item-type',type,'onchange');check(f,'item-reverse',reverse);if(type==='semantic-differential'){field(f,'item-target','The synthetic explanation');field(f,'item-left',reverse?'Clear':'Unclear');field(f,'item-right',reverse?'Unclear':'Clear');field(f,'item-high','Greater perceived clarity');field(f,'anchor-preset','semantic','onchange');click(f,'anchor-apply');click(f,'discard');}click(f,'item-save');assert.equal(f.n('error').hidden,true,f.n('error').textContent);return savedItemIds(f).at(-1);}
function newItem(f,prompt,options){click(f,'item-new');if(!f.n('confirm').hidden)click(f,'discard');return item(f,prompt,options);}
function subscale(f,name,{method='mean',mode='complete-case',k=''}={}){field(f,'rule-name',name);for(const n of members(f)){n.checked=true;n.onchange();}field(f,'rule-method',method,'onchange');field(f,'rule-missing',mode,'onchange');field(f,'rule-k',k);click(f,'rule-save');assert.equal(f.n('error').hidden,true,f.n('error').textContent);}
function capture(f){click(f,'review');assert.equal(f.n('error').hidden,true,f.n('error').textContent);check(f,'reviewed');click(f,'capture');if(!f.n('confirm').hidden)click(f,'discard');assert.equal(f.n('error').hidden,true,f.n('error').textContent);}
function groups(f){return f.n('responses').children.map(group=>({group,radios:descendants(group).filter(n=>n.type==='radio'),clear:descendants(group).find(n=>n.id.startsWith('measure-clear-response-'))}));}
function answer(f,index,optionIndex){const n=groups(f)[index].radios[optionIndex];n.checked=true;n.onchange();}
function base(){const f=fixture();field(f,'title','Synthetic session measure');return f;}

test('session workspace starts without a course, answers, persistence or global side effects and attempts only an unload warning',()=>{
  const f=fixture();assert.equal(f.n('title').value,'');assert.equal(f.n('responses').children.length,0);assert.equal(f.n('score').disabled,true);assert.equal(f.panel.hasSessionWork(),false);
  field(f,'title','Only this session');let prevented=0;const e={preventDefault(){prevented++;}};f.panel.beforeUnload(e);assert.equal(prevented,1);assert.equal(e.returnValue,'');
  const module=fs.readFileSync(new URL('../src/assessment-panel.mjs',import.meta.url),'utf8');assert.doesNotMatch(module,/\b(?:fetch|localStorage|sessionStorage|XMLHttpRequest)\s*[.(]|\b(?:Office|Word)\./u);
});

test('eventless definition metadata and validity changes count as session work before reset/unload',()=>{
  for(const mutate of [f=>f.n('version').value='2',f=>f.n('version').value='1.',f=>f.n('title').validity={customError:true}]){const f=fixture();mutate(f);assert.equal(f.panel.hasSessionWork(),true);let prevented=0;f.panel.beforeUnload({preventDefault(){prevented++;}});assert.equal(prevented,1);click(f,'new');assert.equal(f.n('confirm').hidden,false);click(f,'keep');assert.equal(f.panel.hasSessionWork(),true);}
});

test('closing a confirmation returns focus to a visible control or the safe status instead of hidden Keep',()=>{
  const f=base();f.n('item-new').focus();field(f,'item-prompt','Unfinished');click(f,'item-new');assert.equal(f.document.focused,f.n('keep'));click(f,'keep');assert.equal(f.document.focused,f.n('item-new'));assert.equal(f.n('confirm').hidden,true);
  click(f,'item-new');click(f,'discard');assert.notEqual(f.document.focused,f.n('keep'));assert.notEqual(f.document.focused,f.n('discard'));assert.equal(f.document.focused.hidden,false);
});

test('destructive anchor confirmation never returns focus to the detached removed button',()=>{
  const f=base(),remove=anchors(f)[0].remove;remove.focus();remove.onclick();click(f,'discard');assert.equal(anchors(f).length,4);assert.equal(f.document.focused,f.n('status'));assert.notEqual(f.document.focused,remove);
});

test('Likert items and explicit reverse-keyed mean score raw1/reverse2 as2.5 with visible provenance',()=>{
  const f=base();item(f,'First direct');newItem(f,'Second reverse',{reverse:true});subscale(f,'Authored group');capture(f);
  assert.equal(groups(f).length,2);assert.ok(groups(f).every(g=>g.radios.every(n=>!n.checked)));
  answer(f,0,0);answer(f,1,1);click(f,'score');assert.match(f.n('score-output').textContent,/Authored group.*2\.5; mean; complete-case; n\/N 2\/2; complete/u);assert.match(f.n('score-output').textContent,/raw 2; reverse L\+U-x; keyed 4/u);assert.match(f.n('score-output').textContent,/validation evidence is not established/u);
});

test('semantic differential endpoints and position labels are explicit and direct6/reverse2 both key6',()=>{
  const f=base();item(f,'Rate clarity',{type:'semantic-differential'});newItem(f,'Rate reversed clarity',{type:'semantic-differential',reverse:true});subscale(f,'Clarity');capture(f);
  assert.match(f.n('responses').textContent,/Left endpoint: Unclear\. Right endpoint: Clear/u);assert.match(f.n('responses').textContent,/Scale position 6 of 7/u);for(const g of groups(f)){const ref=g.group.getAttribute('aria-describedby');assert.ok(ref);assert.match(f.document.getElementById(ref).textContent,/Left endpoint/u);assert.ok(g.radios.every(n=>n.getAttribute('aria-describedby')===ref));}answer(f,0,5);answer(f,1,1);click(f,'score');assert.match(f.n('score-output').textContent,/Clarity.*6; mean/u);
});

test('agreement frequency importance and custom shifted anchors are editable and replacement is guarded',()=>{
  const f=base();const first=anchors(f)[0].label;first.value='Authored anchor';first.oninput();
  field(f,'anchor-preset','frequency','onchange');click(f,'anchor-apply');assert.equal(f.n('confirm').hidden,false);assert.equal(f.document.focused,f.n('keep'));click(f,'keep');assert.equal(first.value,'Authored anchor');
  click(f,'anchor-apply');click(f,'discard');assert.equal(anchors(f)[0].label.value,'Never');field(f,'anchor-preset','importance','onchange');click(f,'anchor-apply');click(f,'discard');assert.equal(anchors(f)[0].label.value,'Not at all important');
  field(f,'anchor-preset','custom','onchange');field(f,'anchor-count','3');field(f,'anchor-min','-1');click(f,'anchor-apply');click(f,'discard');assert.deepEqual(anchors(f).map(a=>a.code.value),['-1','0','1']);field(f,'item-prompt','Custom');click(f,'item-save');assert.equal(f.n('error').hidden,true);
});

test('raw unfinished item and rule buffers survive navigation-like hiding, renders, validation refusal and busy actions',()=>{
  const f=base();field(f,'item-prompt','Keep item');anchors(f)[0].code.value='1.';anchors(f)[0].code.validity={badInput:true};anchors(f)[0].code.oninput();field(f,'rule-k','2.');const input=anchors(f)[0].code;
  f.document.getElementById('view-assessments').hidden=true;f.panel.renderBusy();f.document.getElementById('view-assessments').hidden=false;f.panel.renderBusy();assert.equal(anchors(f)[0].code,input);assert.equal(input.value,'1.');assert.equal(input.validity.badInput,true);assert.equal(f.n('rule-k').value,'2.');click(f,'item-save');assert.equal(f.n('error').hidden,false);assert.equal(input.value,'1.');
  f.busy(true);click(f,'new');click(f,'item-new');click(f,'item-save');assert.equal(f.n('item-prompt').value,'Keep item');assert.equal(f.n('title').value,'Synthetic session measure');assert.equal(f.n('confirm').hidden,true);f.busy(false);assert.equal(input.value,'1.');
});

test('item-only definitions review with pristine empty subscale editor and do not infer an aggregate',()=>{
  const f=base();item(f,'Item only');capture(f);answer(f,0,0);click(f,'score');assert.match(f.n('score-output').textContent,/No aggregate requested/u);assert.equal(savedRuleIds(f).length,0);
});

test('selection alone never retargets saved item draft and loading dirty work requires explicit choice',()=>{
  const f=base(),a=item(f,'A'),b=newItem(f,'B');f.n('item-list').value=a;f.n('item-list').onchange();field(f,'item-prompt','Edited B');click(f,'item-save');assert.match(f.n('items').textContent,/Edited B/u);assert.match(f.n('items').textContent,/A/u);
  field(f,'item-prompt','UNSAVED B');click(f,'item-load');click(f,'keep');assert.equal(f.n('item-prompt').value,'UNSAVED B');click(f,'item-load');click(f,'discard');assert.equal(f.n('item-prompt').value,'A');assert.match(f.n('item-id').textContent,new RegExp(a));assert.notEqual(a,b);
});

test('stale pending intents refuse eventless raw/selection/validity changes and Escape keeps current work',()=>{
  for(const mutate of [f=>f.n('item-prompt').value='NEW EVENTLESS',f=>f.n('item-list').value='unknown',f=>anchors(f)[0].code.validity={badInput:true}]){
    const f=base();field(f,'item-prompt','Original');click(f,'item-new');mutate(f);click(f,'discard');assert.equal(f.n('confirm').hidden,true);assert.match(f.n('error').textContent,/changed/u);assert.notEqual(f.n('item-prompt').value,'');
  }
  const f=base();field(f,'item-prompt','Keep via Escape');click(f,'item-new');let prevented=0;f.n('confirm').onkeydown({key:'Escape',preventDefault(){prevented++;}});assert.equal(prevented,1);assert.equal(f.n('item-prompt').value,'Keep via Escape');assert.equal(f.n('confirm').hidden,true);
});

test('captured version identity is immutable locally; changed content needs newversion and replacement keeps responses until confirmed',()=>{
  const f=base();item(f,'Original');capture(f);answer(f,0,1);field(f,'item-prompt','Changed');click(f,'item-save');click(f,'review');assert.match(f.n('error').textContent,/new authored version/u);assert.equal(groups(f)[0].radios[1].checked,true);field(f,'version','2');click(f,'review');check(f,'reviewed');click(f,'capture');assert.equal(f.n('confirm').hidden,false);click(f,'keep');assert.equal(groups(f)[0].radios[1].checked,true);
  click(f,'capture');click(f,'discard');assert.match(f.n('test-status').textContent,/version 2/u);assert.ok(groups(f)[0].radios.every(n=>!n.checked));
});

test('same-definition recapture preserves answers; checked flag alone and tampered review display never count as review',()=>{
  const f=base();item(f,'One');click(f,'review');f.n('reviewed').checked=true;click(f,'capture');assert.match(f.n('error').textContent,/Acknowledge/u);check(f,'reviewed');click(f,'capture');answer(f,0,2);capture(f);assert.equal(groups(f)[0].radios[2].checked,true);
  for(const change of [()=>f.n('review-output').textContent+=' forged',()=>f.n('review-output').hidden=true,()=>f.n('review-output').attributes.style='display:none']){click(f,'review');check(f,'reviewed');change();click(f,'capture');assert.match(f.n('error').textContent,/displayed definition changed/u);f.n('review-output').hidden=false;f.n('review-output').attributes.style='';}
});

test('pending captured replacement revalidates acknowledgement state, generation and actual review visibility',()=>{
  for(const uncheck of [f=>f.n('reviewed').checked=false,f=>check(f,'reviewed',false),f=>f.n('review-output').hidden=true,f=>f.n('review-output').attributes.style='display:none',f=>f.n('review-output').textContent+=' altered',f=>f.n('review-output').append(f.document.createElement('span'))]){const f=base();item(f,'One');capture(f);answer(f,0,0);field(f,'title','New title');field(f,'version','2');click(f,'review');check(f,'reviewed');click(f,'capture');uncheck(f);click(f,'discard');assert.match(f.n('error').textContent,/changed/u);assert.equal(groups(f)[0].radios[0].checked,true);assert.match(f.n('test-status').textContent,/version 1/u);}
});

test('score reads current radio states and rejects multiple, changed, hidden, extra or missing controls without stale result',()=>{
  for(const mutate of [f=>{groups(f)[0].radios[0].checked=groups(f)[0].radios[1].checked=true;},f=>groups(f)[0].radios[0].value='unknown',f=>groups(f)[0].radios[0].hidden=true,f=>groups(f)[0].group.append(f.document.createElement('input')),f=>groups(f)[0].radios[0].remove()]){
    const f=base();item(f,'One');capture(f);answer(f,0,0);click(f,'score');assert.match(f.n('score-output').textContent,/keyed 1/u);mutate(f);click(f,'score');assert.equal(f.n('error').hidden,false);assert.doesNotMatch(f.n('score-output').textContent,/keyed 1/u);
  }
});

test('semantic response description references are part of displayed scoring provenance',()=>{
  for(const mutate of [g=>g.radios[0].setAttribute('aria-describedby','missing-context'),g=>g.group.setAttribute('aria-describedby','changed')]){const f=base();item(f,'Rate',{type:'semantic-differential'});capture(f);answer(f,0,0);const g=groups(f)[0];mutate(g);click(f,'score');assert.match(f.n('error').textContent,/displayed response controls changed/u);assert.doesNotMatch(f.n('score-output').textContent,/keyed 1/u);}
});

test('eventless response/draft edits invalidate prior score on ordinary render and preserve raw choices',()=>{
  const f=base();item(f,'One');capture(f);answer(f,0,0);click(f,'score');const g=groups(f)[0];g.radios[0].checked=false;g.radios[2].checked=true;f.panel.renderBusy();assert.match(f.n('score-output').textContent,/out of date/u);assert.equal(g.radios[2].checked,true);click(f,'score');assert.match(f.n('score-output').textContent,/raw 3/u);f.n('title').value='EVENTLESS CHANGE';f.panel.renderBusy();assert.match(f.n('score-output').textContent,/out of date/u);click(f,'score');assert.match(f.n('error').textContent,/draft differs/u);assert.equal(g.radios[2].checked,true);
});

test('missing and NA remain distinct in plannedN and partial means need explicitk; clear responses requires confirmation',()=>{
  const f=base();item(f,'One');newItem(f,'Two');newItem(f,'Three');subscale(f,'Partial',{mode:'minimum-answered',k:'2'});capture(f);answer(f,0,0);answer(f,1,4);answer(f,2,5);click(f,'score');assert.match(f.n('score-output').textContent,/3; mean; minimum-answered; k=2; n\/N 2\/3; partial; missing 0; N\/A 1/u);groups(f)[2].clear.onclick();click(f,'score');assert.match(f.n('score-output').textContent,/missing 1; N\/A 0/u);click(f,'clear-responses');click(f,'keep');assert.equal(groups(f)[0].radios[0].checked,true);click(f,'clear-responses');click(f,'discard');click(f,'score');assert.match(f.n('score-output').textContent,/NO SCORE.*n\/N 0\/3; no-responses/u);
});

test('partial sums and mixed numeric ranges refuse locally with rule/item drafts preserved',()=>{
  const f=base();item(f,'One');newItem(f,'Two');field(f,'rule-name','Invalid sum');for(const n of members(f)){n.checked=true;n.onchange();}field(f,'rule-method','sum','onchange');field(f,'rule-missing','minimum-answered','onchange');field(f,'rule-k','1');click(f,'rule-save');assert.match(f.n('error').textContent,/sum requires/u);assert.equal(f.n('rule-k').value,'1');assert.equal(f.n('rule-name').value,'Invalid sum');
  const g=base();item(g,'One');newItem(g,'Two');for(const a of anchors(g)){a.code.value=String(Number(a.code.value)+1);a.code.oninput();}click(g,'item-save');field(g,'rule-name','Mixed');for(const n of members(g)){n.checked=true;n.onchange();}click(g,'rule-save');assert.match(g.n('error').textContent,/mixed numeric ranges/u);assert.equal(g.n('rule-name').value,'Mixed');assert.ok(members(g).every(n=>n.checked));
});

test('item deletion never silently drops saved or unfinished draft rule membership',()=>{
  const f=base(),id=item(f,'Member');const m=members(f)[0];m.checked=true;m.onchange();f.n('item-list').value=id;click(f,'item-remove');assert.match(f.n('error').textContent,/current draft subscale memberships/u);assert.deepEqual(savedItemIds(f),[id]);assert.equal(m.checked,true);field(f,'rule-name','Saved');click(f,'rule-save');m.checked=false;m.onchange();click(f,'item-remove');assert.match(f.n('error').textContent,/saved and current/u);assert.deepEqual(savedItemIds(f),[id]);
});

test('hidden semantic and inactive partialmean raw fields survive savedbuffer load/type switches',()=>{
  const f=base();item(f,'Semantic',{type:'semantic-differential'});field(f,'item-type','likert','onchange');click(f,'item-save');const first=savedItemIds(f)[0];newItem(f,'Other');f.n('item-list').value=first;click(f,'item-load');assert.equal(f.n('item-left').value,'Unclear');field(f,'item-type','semantic-differential','onchange');assert.equal(f.n('item-right').value,'Clear');
  field(f,'rule-k','2.');field(f,'rule-missing','complete-case','onchange');assert.equal(f.n('rule-k').value,'2.');f.panel.renderBusy();assert.equal(f.n('rule-k').value,'2.');
});

test('cached detached anchor controls refuse safely; anchor ordering changes stable order without coercingcodes',()=>{
  const f=base(),old=anchors(f)[0].down;field(f,'anchor-preset','frequency','onchange');click(f,'anchor-apply');click(f,'discard');const before=anchors(f).map(a=>a.label.value);old.onclick();assert.match(f.n('error').textContent,/replaced/u);assert.deepEqual(anchors(f).map(a=>a.label.value),before);anchors(f)[0].down.onclick();assert.deepEqual(anchors(f).slice(0,2).map(a=>a.label.value),['Rarely','Never']);
});

test('forms prevent default and reset/delete choices keep assessment work until explicit confirmation',()=>{
  const f=base();field(f,'item-prompt','<b>literal authored text</b>');let prevented=0;f.n('item-form').onsubmit({preventDefault(){prevented++;}});assert.equal(prevented,1);assert.match(f.n('items').textContent,/<b>literal authored text<\/b>/u);click(f,'new');click(f,'keep');assert.equal(savedItemIds(f).length,1);click(f,'new');click(f,'discard');assert.equal(savedItemIds(f).length,0);assert.equal(f.n('title').value,'');
});

test('static new-control accessibility/session semantics and minimal app integration seams are explicit, not nativequalification',()=>{
  const f=fixture(),html=f.html,css=fs.readFileSync(new URL('../public/style.css',import.meta.url),'utf8'),app=fs.readFileSync(new URL('../src/app.mjs',import.meta.url),'utf8');
  assert.match(html,/id="view-assessments" aria-labelledby="measure-heading"/u);assert.match(html,/id="measure-confirm"[^>]*role="group"[^>]*aria-labelledby="measure-confirm-text"/u);assert.match(html,/id="measure-status" role="status" aria-live="polite"/u);assert.match(html,/id="measure-error" role="alert"/u);assert.match(html,/id="measure-item-form" autocomplete="off"/u);assert.match(html,/switching Word documents may start another pane session/u);
  for(const id of ['title','version','item-type','item-prompt','item-reverse','item-target','item-left','item-right','item-high','anchor-preset','anchor-count','anchor-min','rule-name','rule-method','rule-missing','rule-k','reviewed'])assert.match(html,new RegExp(`for="measure-${id}"`));
  assert.match(css,/#view-assessments input\[type=radio\].*width:auto/u);assert.match(css,/#view-assessments\{overflow-wrap:anywhere\}/u);assert.match(app,/assessmentPanel\?\.renderBusy\(\);/u);assert.match(app,/mountAssessmentPanel\(document,\{isBusy:\(\)=>busy,download:downloadBlob,getZip:\(\)=>globalThis.JSZip\}\)/u);assert.match(app,/'materials','assessments'\]\.includes\(name\)/u);
});


test('local validation alerts scroll into view without clearing raw drafts or taking keyboard focus',()=>{
  const f=base();field(f,'item-prompt','KEEP invalid draft');const code=anchors(f)[0].code;code.value='1.';code.validity={badInput:true};code.oninput();
  const focus=f.n('item-save');focus.focus();let options=null;f.n('error').scrollIntoView=value=>{options=value;};click(f,'item-save');
  assert.equal(f.n('error').hidden,false);assert.match(f.n('error').textContent,/invalid/u);assert.deepEqual(options,{block:'nearest'});
  assert.equal(f.document.focused,focus);assert.equal(f.n('item-prompt').value,'KEEP invalid draft');assert.equal(code.value,'1.');assert.equal(code.validity.badInput,true);
});
