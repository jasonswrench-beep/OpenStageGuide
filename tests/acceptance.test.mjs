import test from 'node:test';
import assert from 'node:assert/strict';
import {createAcceptanceFaults,ACCEPTANCE_KEY,ACCEPTANCE_RUN,ACCEPTANCE_TITLE} from '../src/acceptance-faults.mjs';
import {fixture,makeHost,adapter} from './helpers/mock-office.mjs';
import {KEY,PENDING} from '../src/document-store.mjs';
import {addModule} from '../src/model.mjs';
import {attachAcceptancePanel} from '../src/acceptance-panel.mjs';
function markedHost(){const p=fixture();p.title=ACCEPTANCE_TITLE;const h=makeHost(p);h.memory[ACCEPTANCE_KEY]={disposable:true,run:ACCEPTANCE_RUN,projectId:p.id};h.durable=structuredClone(h.memory);return {p,h};}
test('Inspect cannot refresh the native settings bag during a controller transaction',async()=>{
 const {p,h}=markedHost(),w=await adapter(h),{next,item}=addModule(p,'Locked','');
 let release,entered;const started=new Promise(ok=>entered=ok),gate=new Promise(ok=>release=ok);
 const original=h.settings.saveAsync;h.settings.saveAsync=cb=>{entered();gate.then(()=>original(cb));};
 const operation=w.insert(p,next,'module',item);await started;
 const refreshes=h.refreshes;await assert.rejects(w.acceptanceSnapshot(),/already running/);assert.equal(h.refreshes,refreshes);
 release();await operation;
});
test('overlapping Inspect and case clicks cannot arm, clear or inspect during an active panel action',async()=>{
 class E{constructor(){this.children=[];this.dataset={};this.style={};}append(...x){this.children.push(...x);}}
 const main=new E();globalThis.document={createElement:()=>new E(),querySelector:()=>main};
 let release,snapshots=0,clears=0,arms=0;const gate=new Promise(ok=>release=ok);
 const word={clearAcceptance:()=>clears++,armAcceptance:async()=>arms++,acceptanceSnapshot:async()=>{snapshots++;await gate;return {};}};
 attachAcceptancePanel({word,run:async a=>a(),refresh:async()=>{},commit:async()=>{},getProject:fixture,isBusy:()=>false});
 const buttons=main.children[0].children.filter(x=>x.dataset.acceptanceAction==='1');
 const inspecting=buttons.at(-1).onclick();await buttons[0].onclick();await buttons.at(-1).onclick();
 assert.equal(snapshots,1);assert.equal(clears,0);assert.equal(arms,0);assert.ok(buttons.every(b=>b.disabled));
 release();await inspecting;assert.ok(buttons.every(b=>!b.disabled));
});
test('acceptance hooks cannot arm on ordinary documents and abort if document identity changes',async()=>{
 const h=makeHost(fixture()),f=createAcceptanceFaults(()=>h.settings);
 assert.throws(()=>f.arm({kind:'save',when:'before'}),/marked disposable/);
 const m=markedHost(),g=createAcceptanceFaults(()=>m.h.settings);g.arm({kind:'save',when:'before'});
 m.h.memory[KEY].id='different-document';
 assert.throws(()=>g.settings().saveAsync(()=>{}),/identity changed/);
 assert.equal(m.h.saves,0);assert.ok(g.events().some(e=>e.event==='scope-lost'));
});
test('marker loss aborts synchronization; stale pre-write rejection cannot leave an inherited arm',async()=>{
 const {p,h}=markedHost(),w=await adapter(h),{next,item}=addModule(p,'Stale module','');
 await w.armAcceptance({kind:'save',when:'before',nth:1});
 h.durable[KEY].revision++;await assert.rejects(w.insert(p,next,'module',item),/changed/);
 w.clearAcceptance();
 const f=createAcceptanceFaults(()=>h.settings);f.arm({kind:'write-sync',when:'after'});delete h.memory[ACCEPTANCE_KEY];
 let synced=false;await assert.rejects(f.sync({sync:async()=>{synced=true;}},'write'),/identity changed/);assert.equal(synced,false);
 assert.throws(()=>f.arm({kind:'save',when:'before'}),/marked disposable/);
});
for(const [label,point,fields,recoverable] of [
 ['before prepare save',{kind:'save',when:'before',nth:1},7,false],
 ['after real commit save',{kind:'save',when:'after',nth:2},9,true],
 ['after partial write',{kind:'write-sync',when:'after',nth:1},8,false],
 ['after complete write',{kind:'write-sync',when:'after',nth:2},9,true]
])test(`synthetic boundary ${label} exercises actual adapter/controller and remains one-shot`,async()=>{
 const {p,h}=markedHost(),w=await adapter(h),{next,item}=addModule(p,'Synthetic module','');
 await w.armAcceptance(point);await assert.rejects(w.insert(p,next,'module',item),/paused/);
 assert.equal(w.isBlocked(),true);assert.equal(h.records.length,fields);
 const before=structuredClone(h.records);
 await assert.rejects(w.readProject(),/paused/);
 if(recoverable){assert.deepEqual(await w.recover(),next);assert.deepEqual(await w.recover(),next);assert.equal(h.durable[PENDING].phase,'committed');}
 else await assert.rejects(w.recover(),/deleted/);
 assert.deepEqual(h.records,before);
 assert.ok((await w.acceptanceSnapshot()).events.some(e=>e.event==='synthetic-fault'));
});

for(const when of ['before','after'])test(`synthetic settings failure ${when} duration write has operation-specific recovery proof`,async()=>{
  const {p,h}=markedHost(),w=await adapter(h),id=p.modules[0].activities[0].id;
  await w.armAcceptance({kind:'write-sync',when,nth:1});
  await assert.rejects(w.updateActivitySettings(p,id,{minutes:12.5,lines:8}),/paused/);
  assert.equal(h.cues[id],`Learning activity · ${when==='before'?10:12.5} min`);
  assert.equal(h.cueWrites,when==='before'?0:1);
  if(when==='before')await assert.rejects(w.recover(),/cue/);
  else{const result=await w.recover();assert.equal(result.modules[0].activities[0].minutes,12.5);assert.equal(result.modules[0].activities[0].lines,8);await w.recover();assert.equal(h.cueWrites,1);}
  assert.ok((await w.acceptanceSnapshot()).events.some(e=>e.event==='synthetic-fault'&&e.kind==='write-sync'&&e.when===when));
});
