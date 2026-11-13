import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {mountAssessmentPanel} from '../src/assessment-panel.mjs';
import {assessmentDom,descendants} from './helpers/assessment-dom.mjs';
import {createAssessmentFile,serializeAssessmentFile,parseAssessmentFile,assessmentDraftFromDefinition} from '../src/assessment-file.mjs';
import {assessmentFileDefinition} from './helpers/assessment-file-fixture.mjs';

function fixture(){const ui=assessmentDom(),offers=[];let hostBusy=false;const panel=mountAssessmentPanel(ui.document,{isBusy:()=>hostBusy,download:(blob,name)=>offers.push({blob,name})});return {...ui,panel,offers,busy:value=>{hostBusy=value;panel.renderBusy();}};}
const click=(f,id)=>f.n(id).onclick();
const field=(f,id,value,event='oninput')=>{f.n(id).value=value;f.n(id)[event]?.();};
const check=(f,id,value=true)=>{f.n(id).checked=value;f.n(id).onchange?.();};
const rows=f=>f.n('anchors').children.map(row=>({label:row.children.find(n=>n.id.startsWith('measure-anchor-label-')),code:row.children.find(n=>n.id.startsWith('measure-anchor-code-'))}));
const groups=f=>f.n('responses').children.map(group=>({group,radios:descendants(group).filter(n=>n.type==='radio')}));
function item(f,text='Local prompt'){field(f,'item-prompt',text);click(f,'item-save');assert.equal(f.n('error').hidden,true,f.n('error').textContent);}
function capture(f){click(f,'review');assert.equal(f.n('error').hidden,true,f.n('error').textContent);check(f,'reviewed');click(f,'capture');if(!f.n('confirm').hidden)click(f,'discard');assert.equal(f.n('error').hidden,true,f.n('error').textContent);}
function answer(f){const radio=groups(f)[0].radios[1];radio.checked=true;radio.onchange();return radio;}
function base(){const f=fixture();field(f,'title','Local PRIVATE authoring');item(f);return f;}
const bytes=file=>serializeAssessmentFile(file);
function localFile(data,{chunk=65536,beforeRead,afterRead,declared=data.length,name='PRIVATE-assessment.json'}={}){let index=0,reads=0;return {name,size:declared,stream(){return {getReader(){return {async read(){await beforeRead?.(++reads);if(index>=data.length)return {done:true};const value=data.subarray(index,Math.min(data.length,index+chunk));index+=value.length;const result={done:false,value};queueMicrotask(()=>afterRead?.(reads,value));return result;},async cancel(){},releaseLock(){}};}};}};}
function choose(f,file){f.n('file-input').files=file?[file]:[];f.n('file-input').onchange?.();}
async function load(f,file){choose(f,file);await click(f,'file-load');}
function definitionFile(def=assessmentFileDefinition()){return createAssessmentFile('definition',def);}
async function lastFile(f){return parseAssessmentFile(new Uint8Array(await f.offers.at(-1).blob.arrayBuffer()));}
function authored(f){return {title:f.n('title').value,version:f.n('version').value,itemId:f.n('item-id').textContent,ruleId:f.n('rule-id').textContent,itemFields:['item-type','item-prompt','item-target','item-left','item-right','item-high','anchor-preset','anchor-count','anchor-min'].map(id=>[id,f.n(id).value]),reverse:f.n('item-reverse').checked,anchors:rows(f).map(row=>[row.label.id,row.label.value,row.code.value]),ruleFields:['rule-name','rule-method','rule-missing','rule-k'].map(id=>[id,f.n(id).value]),members:f.n('members').children.map(row=>{const c=row.children.find(n=>n.type==='checkbox');return [c.id,c.checked];}),items:f.n('items').textContent,rules:f.n('rules').textContent,itemSelection:f.n('item-list').value,ruleSelection:f.n('rule-list').value,responses:groups(f).map(group=>group.radios.map(n=>[n.value,n.checked])),testStatus:f.n('test-status').textContent};}

test('PRIVATE draft download preserves unfinished raw editor state and excludes tester, consent and separate export preferences',async()=>{
 const f=base();capture(f);answer(f);field(f,'item-target','Inactive semantic field retained');field(f,'item-high','Inactive meaning');field(f,'rule-k','099');field(f,'version','1.');rows(f)[0].code.value='-';rows(f)[0].code.oninput();f.n('export-instructions').value='EXCLUDED_INSTRUCTIONS';f.n('export-route').value='author';f.n('export-language').value='fr-CA';
 const before=authored(f);click(f,'file-save-draft');assert.equal(f.offers.length,1,f.n('file-error').textContent);const file=await lastFile(f);assert.equal(file.kind,'draft');assert.equal(file.payload.version,'1.');assert.equal(file.payload.activeItem.buffer.anchors[0].code,'-');assert.equal(file.payload.activeItem.buffer.fields['item-high'],'Inactive meaning');assert.equal(file.payload.activeRule.buffer.fields['rule-k'],'099');assert.deepEqual(authored(f),before);const text=await f.offers[0].blob.text();assert.doesNotMatch(text,/EXCLUDED_INSTRUCTIONS|capturedKey|responses|consent|scoreBinding|invalid|baseline|serial/u);assert.match(f.offers[0].name,/^OpenStage-PRIVATE-Assessment-Draft-.*\.json$/u);
 const restored=fixture();await load(restored,localFile(bytes(file),{chunk:17}));assert.equal(restored.n('confirm').hidden,false,restored.n('file-error').textContent);click(restored,'discard');assert.equal(restored.n('error').hidden,true,restored.n('error').textContent);assert.equal(restored.n('version').value,'1.');assert.equal(rows(restored)[0].code.value,'-');assert.equal(restored.n('item-high').value,'Inactive meaning');click(restored,'file-save-definition');assert.equal(restored.offers.length,0);assert.match(restored.n('file-error').textContent,/unfinished/u);
});

test('validated definition saves require clean valid work, repeat exact snapshots with unique names, and reserve offered identity',async()=>{
 const f=base();click(f,'file-save-definition');assert.equal(f.offers.length,1,f.n('file-error').textContent);click(f,'file-save-definition');assert.equal(f.offers.length,2);assert.notEqual(f.offers[0].name,f.offers[1].name);assert.deepEqual(await f.offers[0].blob.text(),await f.offers[1].blob.text());assert.equal((await lastFile(f)).kind,'definition');
 field(f,'item-prompt','Changed saved content');click(f,'item-save');click(f,'file-save-definition');assert.equal(f.offers.length,2);assert.match(f.n('file-error').textContent,/new authored version/u);click(f,'review');assert.match(f.n('error').textContent,/new authored version/u);field(f,'version','2');click(f,'file-save-definition');assert.equal(f.offers.length,3,f.n('file-error').textContent);field(f,'item-prompt','Unfinished');click(f,'file-save-definition');assert.equal(f.offers.length,3);assert.match(f.n('file-error').textContent,/unfinished/u);
});

test('checked file Keep retains drafts and live response nodes; Replace clears capture and consent while retaining export choices',async()=>{
 const f=base();capture(f);const radio=answer(f);f.n('export-instructions').value='Retained export instructions';f.n('export-route').value='author';f.n('export-language').value='de-DE';field(f,'item-high','Keep inactive raw');const before=authored(f),incoming=definitionFile();
 await load(f,localFile(bytes(incoming)));assert.equal(f.n('confirm').hidden,false,f.n('file-error').textContent);assert.equal(groups(f)[0].radios[1],radio);click(f,'keep');assert.deepEqual(authored(f),before);assert.equal(radio.checked,true);
 await load(f,localFile(bytes(incoming)));click(f,'discard');assert.equal(f.n('error').hidden,true,f.n('error').textContent);assert.equal(f.n('title').value,incoming.payload.title);assert.equal(f.n('responses').children.length,0);assert.equal(f.n('reviewed').checked,false);assert.equal(f.n('export-reviewed').checked,false);assert.equal(f.n('export-private').checked,false);assert.equal(f.n('export-instructions').value,'Retained export instructions');assert.equal(f.n('export-route').value,'author');assert.equal(f.n('export-language').value,'de-DE');assert.match(f.n('test-status').textContent,/No definition captured/u);click(f,'file-save-definition');assert.equal(f.offers.length,1,f.n('file-error').textContent);assert.deepEqual((await lastFile(f)).payload,incoming.payload);
});

test('file cancel, malformed, unsupported and incomplete reads preserve authoring and synthetic responses',async()=>{
 const bad=[new Uint8Array([0xff,0xfe]),new TextEncoder().encode('{'),new TextEncoder().encode('{"format":"openstage-private-assessment","version":99,"kind":"definition","payload":{}}'),new TextEncoder().encode('null')];
 for(const value of bad){const f=base();capture(f);const radio=answer(f),before=authored(f);await load(f,localFile(value));assert.equal(f.n('file-error').hidden,false);assert.deepEqual(authored(f),before);assert.equal(groups(f)[0].radios[1],radio);assert.equal(f.n('confirm').hidden,true);}
 const f=base();click(f,'review');check(f,'reviewed');const before=authored(f);choose(f,null);await click(f,'file-load');f.n('file-input').oncancel();assert.equal(f.n('reviewed').checked,true);assert.deepEqual(authored(f),before);const data=bytes(definitionFile());await load(f,localFile(data,{declared:data.length+1}));assert.match(f.n('file-error').textContent,/incomplete/u);assert.deepEqual(authored(f),before);await load(f,localFile(data,{declared:4*1024*1024+1}));assert.match(f.n('file-error').textContent,/4 MiB/u);assert.deepEqual(authored(f),before);
});

test('async local reads reject source, file, busy and workspace away-back changes without replacement',async()=>{
 for(const change of [f=>field(f,'title','New raw text during read'),f=>{f.document.getElementById('view-assessments').hidden=true;f.panel.renderBusy();f.document.getElementById('view-assessments').hidden=false;f.panel.renderBusy();},f=>f.busy(true),f=>{f.busy(true);f.busy(false);},f=>choose(f,localFile(new Uint8Array([1])))]){
  const f=base();capture(f);const radio=answer(f);let release;const wait=new Promise(resolve=>release=resolve),file=localFile(bytes(definitionFile()),{beforeRead:async n=>{if(n===1)await wait;}});choose(f,file);const operation=click(f,'file-load');change(f);const after=authored(f);release();await operation;assert.equal(f.n('file-error').hidden,false);assert.deepEqual(authored(f),after);assert.equal(groups(f)[0].radios[1],radio);assert.equal(f.n('confirm').hidden,true);
 }
});

test('cancel or empty-then-same-file selection invalidates a held read without losing current work',async()=>{
 for(const interrupt of [f=>f.n('file-input').oncancel(),(f,file)=>{choose(f,null);choose(f,file);}]){const f=base();capture(f);const radio=answer(f),before=authored(f);let release;const wait=new Promise(resolve=>release=resolve),file=localFile(bytes(definitionFile()),{beforeRead:async n=>{if(n===1)await wait;}});choose(f,file);const operation=click(f,'file-load');interrupt(f,file);release();await operation;assert.equal(f.n('file-error').hidden,false);assert.equal(f.n('confirm').hidden,true);assert.deepEqual(authored(f),before);assert.equal(groups(f)[0].radios[1],radio);}
});

test('candidate payload and confirmation structure are bound before Replace; stale changes preserve current work',async()=>{
 for(const [index,change] of [f=>f.n('file-output').textContent+=' altered',f=>f.n('file-output').hidden=true,f=>f.n('confirm').hidden=true,f=>f.n('discard').textContent='Different action',f=>f.n('confirm').append(f.document.createElement('span')),f=>field(f,'item-prompt','New unsaved raw work')].entries()){
  const f=base();capture(f);const radio=answer(f);await load(f,localFile(bytes(definitionFile())));change(f);const after=authored(f);click(f,'discard');assert.equal(f.n('error').hidden,false,`candidate mutation case ${index}: ${f.n('file-error').textContent}`);assert.deepEqual(authored(f),after);assert.equal(groups(f)[0].radios[1],radio);
 }
});

test('unexpected commit failure restores current controls, dynamic nodes and retained raw/tester state',async()=>{
 const f=base();capture(f);const radio=answer(f);const before=authored(f);await load(f,localFile(bytes(definitionFile())));const node=f.n('members'),original=node.replaceChildren.bind(node);let once=true;node.replaceChildren=(...args)=>{if(once){once=false;throw new Error('Synthetic DOM commit failure');}return original(...args);};click(f,'discard');assert.match(f.n('error').textContent,/Synthetic DOM/u);assert.deepEqual(authored(f),before);assert.equal(groups(f)[0].radios[1],radio);assert.equal(radio.checked,true);
});

test('post-restore exact authoring readback catches current setter changes and rolls back before clearing the tester',async()=>{
 const f=base();capture(f);const radio=answer(f),before=authored(f),incoming=definitionFile();await load(f,localFile(bytes(incoming)));const node=f.n('title');let value=node.value;Object.defineProperty(node,'value',{get:()=>value,set:v=>{value=v===incoming.payload.title?v+' CHANGED BY CURRENT CONTROL':v;},configurable:true});click(f,'discard');assert.match(f.n('error').textContent,/restore this PRIVATE file exactly/u);assert.deepEqual(authored(f),before);assert.equal(groups(f)[0].radios[1],radio);
});

test('cached same-ID anchor controls cannot edit a replaced or rolled-back row',async()=>{
 const f=base();click(f,'file-save-draft');const saved=await lastFile(f),oldRow=f.n('anchors').children[0],oldDown=oldRow.children.find(n=>n.id.startsWith('measure-anchor-down-')),oldRemove=oldRow.children.find(n=>n.id.startsWith('measure-anchor-remove-'));await load(f,localFile(bytes(saved)));click(f,'discard');const after=authored(f);oldDown.onclick();assert.match(f.n('error').textContent,/replaced/u);assert.deepEqual(authored(f),after);oldRemove.onclick();assert.match(f.n('error').textContent,/replaced/u);assert.equal(f.n('confirm').hidden,true);assert.deepEqual(authored(f),after);
 const g=base();capture(g);const radio=answer(g),before=authored(g);let detachedDown;await load(g,localFile(bytes(definitionFile())));const responses=g.n('responses'),original=responses.replaceChildren.bind(responses);let once=true;responses.replaceChildren=(...nodes)=>{if(once){once=false;detachedDown=g.n('anchors').children[0].children.find(n=>n.id.startsWith('measure-anchor-down-'));throw new Error('Late synthetic commit failure');}return original(...nodes);};click(g,'discard');assert.match(g.n('error').textContent,/Late synthetic/u);assert.deepEqual(authored(g),before);assert.equal(groups(g)[0].radios[1],radio);detachedDown.onclick();assert.match(g.n('error').textContent,/replaced/u);assert.deepEqual(authored(g),before);
});

test('stream chunks are copied before subsequent reader mutation and uncertain definition offers retain identity checks',async()=>{
 const data=bytes(definitionFile()),owned=new Uint8Array(data);let reads=0;const file={size:data.length,stream(){return {getReader(){return {async read(){if(++reads===1)return {done:false,value:owned};owned.fill(0);return {done:true};},async cancel(){},releaseLock(){}};}};}};const f=fixture();await load(f,file);click(f,'discard');assert.equal(f.n('error').hidden,true,f.n('error').textContent);click(f,'file-save-definition');assert.deepEqual((await lastFile(f)).payload,definitionFile().payload);
 const ui=assessmentDom();let attempts=0;mountAssessmentPanel(ui.document,{download:()=>{attempts++;throw new Error('Synthetic uncertain local offer');}});const g={...ui};field(g,'title','Offer identity');item(g);click(g,'file-save-definition');assert.equal(attempts,1);assert.match(g.n('file-error').textContent,/outcome is unconfirmed/u);field(g,'item-prompt','Changed same version');click(g,'item-save');click(g,'file-save-definition');assert.equal(attempts,1);assert.match(g.n('file-error').textContent,/new authored version/u);
});

test('oversized reader chunks are refused before a private byte copy is allocated',async()=>{
 const f=base(),before=authored(f),Original=globalThis.Uint8Array,poison=new Original(20);let copies=0;globalThis.Uint8Array=new Proxy(Original,{construct(target,args,newTarget){if(args[0]===poison)copies++;return Reflect.construct(target,args,newTarget);}});
 try{const file={size:1,stream(){return {getReader(){return {async read(){return {done:false,value:poison};},async cancel(){},releaseLock(){}};}};}};await load(f,file);assert.equal(copies,0);assert.match(f.n('file-error').textContent,/exceeded its declared/u);assert.deepEqual(authored(f),before);}finally{globalThis.Uint8Array=Original;}
});

test('detached current-control representability refuses imported text sanitization before changing current state',async()=>{
 const f=base();capture(f);const before=authored(f),definition=assessmentFileDefinition();definition.title='Imported\nmultiline title';const original=f.document.createElement.bind(f.document);f.document.createElement=tag=>{const node=original(tag);if(tag==='input'){let value='';Object.defineProperty(node,'value',{get:()=>value,set:v=>{value=String(v).replace(/[\r\n]/gu,'');},configurable:true});}return node;};await load(f,localFile(bytes(definitionFile(definition))));assert.match(f.n('file-error').textContent,/represent exactly/u);assert.deepEqual(authored(f),before);assert.equal(f.n('confirm').hidden,true);
});

test('membership order and full roster survive definition load, new-item save and confirmed unused-item removal',async()=>{
 const f=fixture(),incoming=definitionFile();await load(f,localFile(bytes(incoming)));click(f,'discard');click(f,'file-save-definition');assert.deepEqual((await lastFile(f)).payload.subscales[0].itemIds,incoming.payload.subscales[0].itemIds);
 click(f,'item-new');field(f,'item-prompt','Unused third item');click(f,'item-save');click(f,'file-save-draft');let saved=await lastFile(f);assert.equal(saved.payload.items.length,3);assert.ok(saved.payload.ruleBuffers.every(row=>row.buffer.members.length===3));const id=saved.payload.items.at(-1).id;f.n('item-list').value=id;f.n('item-list').onchange();click(f,'item-remove');click(f,'discard');click(f,'file-save-draft');saved=await lastFile(f);assert.equal(saved.payload.items.length,2);assert.ok(saved.payload.ruleBuffers.every(row=>row.buffer.members.length===2&&!row.buffer.members.some(member=>member.id===id)));
 const round=fixture();await load(round,localFile(bytes(saved)));click(round,'discard');click(round,'file-save-draft');assert.deepEqual((await lastFile(round)).payload,saved.payload);
});

test('conflicting ID/version loads recover editable data but never bypass session identity ledger through New or repeated Keep',async()=>{
 const f=base();capture(f);const original=await(()=>{click(f,'file-save-definition');return lastFile(f);})();const changed=structuredClone(original);changed.payload.title='Conflicting same-version title';const badFile=createAssessmentFile('definition',changed.payload);click(f,'new');click(f,'discard');
 await load(f,localFile(bytes(badFile)));assert.match(f.n('file-output').textContent,/VERSION CONFLICT/u);click(f,'keep');assert.equal(f.n('title').value,'');await load(f,localFile(bytes(badFile)));click(f,'discard');assert.equal(f.n('title').value,'Conflicting same-version title');click(f,'review');assert.match(f.n('error').textContent,/new authored version/u);const offers=f.offers.length;click(f,'file-save-definition');assert.equal(f.offers.length,offers);click(f,'file-save-draft');assert.equal(f.offers.length,offers+1);field(f,'version','2');click(f,'file-save-definition');assert.equal(f.offers.length,offers+2,f.n('file-error').textContent);
 const g=fixture();await load(g,localFile(bytes(original)));click(g,'discard');field(g,'title','Changed after definition import');click(g,'file-save-definition');assert.equal(g.offers.length,0);assert.match(g.n('file-error').textContent,/new authored version/u);
});

test('imported IDs are reserved and Unicode stream boundaries retain exact authored code points',async()=>{
 const definition=assessmentFileDefinition(11,2,-3);definition.id='assessment-1';definition.items[0].id='item-8';definition.items[0].options[0].id='option-9';for(const rule of definition.subscales)rule.itemIds=rule.itemIds.map(id=>id==='item-1'?'item-8':id);definition.items[0].prompt='Unicode \u{1f642} 漢字 e\u0301 '.repeat(600);const f=fixture();await load(f,localFile(bytes(definitionFile(definition)),{chunk:16383}));click(f,'discard');click(f,'file-save-definition');assert.deepEqual((await lastFile(f)).payload,definition);const occupied=new Set([definition.id,...definition.items.flatMap(i=>[i.id,...i.options.map(o=>o.id)]),...definition.subscales.map(s=>s.id)]);click(f,'item-new');const active=f.n('item-id').textContent.split(': ').at(-1);assert.equal(occupied.has(active),false);for(const row of rows(f))assert.equal(occupied.has(row.label.id.replace('measure-anchor-label-','')),false);
});

test('fresh pane identities are distinct UUIDs, imports retain identity, and duplicate UUID New refuses before clearing work',async()=>{
 const a=base(),b=base();click(a,'file-save-draft');click(b,'file-save-draft');const af=await lastFile(a),bf=await lastFile(b);assert.match(af.payload.definitionId,/^assessment-[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/iu);assert.notEqual(af.payload.definitionId,bf.payload.definitionId);
 const loaded=fixture();await load(loaded,localFile(bytes(af)));click(loaded,'discard');click(loaded,'file-save-draft');assert.equal((await lastFile(loaded)).payload.definitionId,af.payload.definitionId);
 capture(a);const radio=answer(a),before=authored(a),original=globalThis.crypto.randomUUID;try{globalThis.crypto.randomUUID=()=>af.payload.definitionId.slice('assessment-'.length);click(a,'new');assert.equal(a.n('confirm').hidden,false);click(a,'discard');assert.match(a.n('error').textContent,/already used/u);assert.deepEqual(authored(a),before);assert.equal(groups(a)[0].radios[1],radio);}finally{globalThis.crypto.randomUUID=original;}
});

test('PRIVATE file controls expose explicit labels and no persistence, document, network or response dependency',()=>{
 const f=fixture();assert.match(f.n('file-save-draft').textContent||f.html,/PRIVATE authoring draft/u);assert.match(f.html,/Keep my current work/u);assert.match(f.html,/Replace|Discard/u);assert.match(f.n('file-notice').textContent||f.html,/export instructions|Word-export instructions/u);assert.match(f.html,/role="status" aria-live="polite"/u);const module=fs.readFileSync(new URL('../src/assessment-files-panel.mjs',import.meta.url),'utf8');assert.doesNotMatch(module,/\b(?:fetch|localStorage|sessionStorage|XMLHttpRequest)\s*[.(]|\b(?:Office|Word)\./u);assert.doesNotMatch(module,/getResponses|scoreAssessment|createAssessmentResponse/u);
});
