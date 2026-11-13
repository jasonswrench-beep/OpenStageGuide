import test from 'node:test';
import assert from 'node:assert/strict';
import {controllerFixture} from './helpers/icon-controller-fixture.mjs';
import {createIconController} from '../src/icon-controller.mjs';
import {validateIconJournal} from '../src/icon-transition.mjs';
const copy=x=>structuredClone(x);
function paused(r){assert.equal(r.status,'paused');assert.equal(r.replayContent,false);assert.equal(r.nativeVerified,false);assert.equal(r.writeAuthorized,false);}
test('disconnected controller completes one initial physical request with full retained committed evidence and unchanged manual data',async()=>{
  const h=await controllerFixture(),base=copy(h.request.project),manual=h.manual,result=await h.controller.start(h.request);assert.equal(result.status,'completed');assert.equal(h.calls.saves,2);assert.equal(h.calls.mutations,1);assert.equal(h.calls.inspections,3);assert.equal(h.manual,manual);assert.equal(result.nativeVerified,false);assert.equal(result.writeAuthorized,false);
  const receipt=validateIconJournal(h.snapshot.journal);assert.deepEqual(receipt.base,base);assert.equal(receipt.phase,'committed');assert.deepEqual(receipt.project,h.snapshot.project);assert.deepEqual(Object.keys(receipt).sort(),['base','checkpoint','icon','kind','phase','project','schema','txId']);assert.deepEqual(result.retainedJournal,receipt);
});
test('healthy same-icon no-op has no checkpoint inspection persistence or mutation',async()=>{
  const h=await controllerFixture();h.request.icon=h.request.project.modules[0].activities[0].icon;assert.equal((await h.controller.start(h.request)).status,'noop');assert.deepEqual(h.calls,{reads:1,checkpoints:0,proofs:0,inspections:0,saves:0,mutations:0});
});
test('initial and recovery calls share an exclusive guard and busy calls invoke no adapters',async()=>{
  let release,entered;const gate=new Promise(ok=>release=ok),start=new Promise(ok=>entered=ok),h=await controllerFixture();h.hooks.checkpoint=async()=>{entered();await gate;};const pending=h.controller.start(h.request);await start;const counts=copy(h.calls);assert.equal((await h.controller.start(h.request)).status,'busy');assert.equal((await h.controller.recover()).status,'busy');assert.deepEqual(h.calls,counts);release();assert.equal((await pending).status,'completed');
});
test('full stale base or forged checkpoint stops before any preparation or physical request',async()=>{
  for(const defect of ['metadata','checkpoint','capture','manual']){const h=await controllerFixture();if(defect==='metadata')h.request.project.modules[0].activities[0].privateUnknown.preserve='stale';if(defect==='checkpoint')h.checkpointOverride={...h.binding};if(defect==='capture')h.hooks.checkpoint=async(s)=>{s.captureId=crypto.randomUUID();};if(defect==='manual')h.hooks.checkpoint=async()=>h.changeManual();paused(await h.controller.start(h.request));assert.equal(h.calls.saves,0);assert.equal(h.calls.mutations,0);}
});
test('failed uncertain or throwing preparation callbacks never permit a physical request even if records persisted',async()=>{
  for(const status of ['failed','uncertain','throw'])for(const persist of ['none','full','journal','project']){const h=await controllerFixture();h.saveFailures.set(1,{status,persist});const r=await h.controller.start(h.request);paused(r);assert.equal(h.calls.saves,1);assert.equal(h.calls.mutations,0);assert.ok(r.retainedJournal);assert.equal(h.manual,'KEEP MANUAL CONTENT');const before=copy(h.calls);paused(await h.controller.start(h.request));assert.deepEqual(h.calls,before);}
});
test('wrong prepared proof or changed state during proof verification stops before physical request',async()=>{
  for(const defect of ['copied-proof','manual','project','journal','document']){const h=await controllerFixture();h.hooks.proof=async()=>{if(defect==='copied-proof')h.proofOverride={...h.preparation};if(defect==='manual')h.changeManual();if(defect==='project')h.snapshot.project.title+='changed';if(defect==='journal')h.snapshot.journal.checkpoint.id=crypto.randomUUID();if(defect==='document')h.snapshot.documentId=crypto.randomUUID();};paused(await h.controller.start(h.request));assert.equal(h.calls.saves,1);assert.equal(h.calls.mutations,0);}
});
test('affirmative exact-before image label ownership and preservation are mandatory for initial mutation',async()=>{
  for(const defect of ['image','label','after','ownership','placement','dimensions','format','unrelated']){const h=await controllerFixture();h.hooks.inspect=async(c,j)=>{if(defect==='image'||defect==='after')h.imageHash=j.icon.after.assetHash;if(defect==='label'||defect==='after')h.label='Assessment · 10 min';if(defect==='ownership')h.ownership='unowned';if(['placement','dimensions','format','unrelated'].includes(defect))h[defect]='unknown';};paused(await h.controller.start(h.request));assert.equal(h.calls.saves,1);assert.equal(h.calls.mutations,0);}
});
test('physical failed uncertain or throwing callbacks never replay any image or label change',async()=>{
  for(const status of ['failed','uncertain','throw'])for(const change of ['none','image','label','full']){const h=await controllerFixture();h.physicalFailure={status,change};paused(await h.controller.start(h.request));assert.equal(h.calls.mutations,1);assert.equal(h.calls.saves,1);const before=copy(h.calls);paused(await h.controller.start(h.request));assert.deepEqual(h.calls,before);assert.equal(h.calls.mutations,1);}
});
test('a confirmed callback with partial physical state cannot commit metadata',async()=>{
  for(const change of ['none','image','label']){const h=await controllerFixture();h.physicalFailure={status:'confirmed',change};paused(await h.controller.start(h.request));assert.equal(h.calls.saves,1);assert.equal(h.calls.mutations,1);assert.equal(h.snapshot.journal.phase,'prepared');}
});
test('late manual storage identity or inspection changes after the physical request pause without a committed save',async()=>{
  for(const defect of ['manual','metadata','identity','inspection']){const h=await controllerFixture();h.hooks.inspect=async()=>{if(h.calls.inspections!==2)return;if(defect==='manual')h.changeManual();if(defect==='metadata')h.snapshot.project.modules[0].activities[0].answers+='late';if(defect==='identity')h.snapshot.documentId=crypto.randomUUID();if(defect==='inspection')throw new Error('After inspection unavailable');};paused(await h.controller.start(h.request));assert.equal(h.calls.mutations,1);assert.equal(h.calls.saves,1);}
});
test('commit callback uncertainty preserves actual records and recovery recognizes or conditionally completes without physical replay',async()=>{
  for(const status of ['failed','uncertain','throw'])for(const persist of ['none','full','project','journal']){const h=await controllerFixture();h.saveFailures.set(2,{status,persist});paused(await h.controller.start(h.request));assert.equal(h.calls.mutations,1);assert.equal(h.calls.saves,2);const actual=copy(h.snapshot),r=await h.controller.recover();
    if(persist==='full'){assert.equal(r.status,'recognized');assert.equal(h.calls.saves,2);}else if(persist==='none'){assert.equal(r.status,'completed');assert.equal(h.calls.saves,3);assert.equal((await h.controller.recover()).status,'recognized');assert.equal(h.calls.saves,3);}else{paused(r);assert.equal(h.calls.saves,2);assert.deepEqual(h.snapshot,actual);}assert.equal(h.calls.mutations,1);
  }
});
test('exact-after prepared recovery completes metadata once and repeated recovery only recognizes full receipt',async()=>{
  const h=await controllerFixture();h.physicalFailure={status:'uncertain',change:'full'};paused(await h.controller.start(h.request));assert.equal((await h.controller.recover()).status,'completed');const committed=copy(h.snapshot.journal);for(let i=0;i<3;i++)assert.equal((await h.controller.recover()).status,'recognized');assert.equal(h.calls.mutations,1);assert.equal(h.calls.saves,2);assert.deepEqual(h.snapshot.journal,committed);assert.deepEqual(validateIconJournal(committed),committed);
});
test('recovery receives no physical capability and exact committed recognition has zero saves',async()=>{
  const h=await controllerFixture();assert.equal((await h.controller.start(h.request)).status,'completed');let physicalReads=0;const adapters={...h.adapters};Object.defineProperty(adapters,'mutatePhysicalOnce',{get(){physicalReads++;throw new Error('Recovery must not obtain this capability');}});const recovery=createIconController(adapters),saves=h.calls.saves;for(let i=0;i<2;i++)assert.equal((await recovery.recover(h.capabilities())).status,'recognized');assert.equal(h.calls.saves,saves);assert.equal(h.calls.mutations,1);assert.equal(physicalReads,0);
});
test('recovery exact-before mixed unknown manual or mismatched full stored state remains paused',async()=>{
  for(const defect of ['before','image-only','label-only','manual','metadata','ownership','dimensions','journal','checkpoint']){const h=await controllerFixture();h.physicalFailure={status:'uncertain',change:'full'};paused(await h.controller.start(h.request));if(defect==='before'||defect==='label-only')h.imageHash=h.snapshot.journal.icon.before.assetHash;if(defect==='before'||defect==='image-only')h.label='Learning activity · 10 min';if(defect==='manual')h.changeManual();if(defect==='metadata')h.snapshot.project.modules[0].activities[0].privateUnknown.preserve='late';if(defect==='ownership')h.ownership='ambiguous';if(defect==='dimensions')h.dimensions='unknown';if(defect==='journal')h.snapshot.journal.txId=crypto.randomUUID();if(defect==='checkpoint')h.snapshot.journal.checkpoint.id=crypto.randomUUID();const saves=h.calls.saves;paused(await h.controller.recover());assert.equal(h.calls.saves,saves);assert.equal(h.calls.mutations,1);}
});
test('late changes at conditional recovery persistence are refused and never overwrite latest manual work',async()=>{
  const h=await controllerFixture();h.physicalFailure={status:'uncertain',change:'full'};paused(await h.controller.start(h.request));h.hooks.persist=async(n)=>{if(n===2)h.changeManual();};paused(await h.controller.recover());assert.equal(h.snapshot.journal.phase,'prepared');assert.match(h.manual,/CHANGED/);assert.equal(h.calls.mutations,1);assert.equal(h.calls.saves,2);paused(await h.controller.recover());assert.equal(h.calls.saves,2);
});
test('failed recovery save after full durable completion is recognized on repeated fresh recovery without another save',async()=>{
  const h=await controllerFixture();h.physicalFailure={status:'uncertain',change:'full'};paused(await h.controller.start(h.request));h.saveFailures.set(2,{status:'failed',persist:'full'});paused(await h.controller.recover());assert.equal(h.snapshot.journal.phase,'committed');assert.equal((await h.controller.recover()).status,'recognized');assert.equal(h.calls.saves,2);assert.equal(h.calls.mutations,1);
});
test('storage read failures at every successful initial-path boundary halt with at most one physical request',async()=>{
  const control=await controllerFixture();assert.equal((await control.controller.start(control.request)).status,'completed');const boundaries=control.calls.reads;
  for(let at=1;at<=boundaries;at++){const h=await controllerFixture();h.hooks.read=async n=>{if(n===at)throw new Error('Read unavailable');};paused(await h.controller.start(h.request));assert.ok(h.calls.mutations<=1);assert.ok(h.calls.saves<=2);assert.equal(h.manual,'KEEP MANUAL CONTENT');}
});
test('unconfirmed storage or unavailable inspection prevents recovery completion and physical replay',async()=>{
  for(const defect of ['read-unknown','read-failed','inspect-failed','inspect-unavailable','missing-capability']){const h=await controllerFixture();h.physicalFailure={status:'uncertain',change:'full'};paused(await h.controller.start(h.request));if(defect==='read-unknown')h.snapshot.persistence='unknown';if(defect==='read-failed')h.hooks.read=async()=>{throw new Error('Read failure');};if(defect==='inspect-failed')h.hooks.inspect=async()=>{throw new Error('Inspection failed');};let controller=h.controller,args={};if(defect==='inspect-unavailable'){controller=createIconController({...h.adapters,inspectPhysical:async()=>null});args=h.capabilities();}if(defect==='missing-capability')controller=createIconController(h.adapters);paused(await controller.recover(args));assert.equal(h.calls.saves,1);assert.equal(h.calls.mutations,1);}
});

test('confirmed acknowledgments with incomplete persisted records are refused by exact reread',async()=>{
  for(const stage of [1,2])for(const persist of stage===1?['none','project']:['none','project','journal']){const h=await controllerFixture();h.saveFailures.set(stage,{status:'confirmed',persist});paused(await h.controller.start(h.request));assert.equal(h.calls.saves,stage);assert.equal(h.calls.mutations,stage===1?0:1);assert.ok(h.snapshot.journal===null||validateIconJournal(h.snapshot.journal));}
});
test('conditional initial mutation boundary refuses a late manual change and never dispatches a second attempt',async()=>{
  const h=await controllerFixture(),before=h.imageHash;h.hooks.mutate=async()=>h.changeManual();paused(await h.controller.start(h.request));assert.equal(h.imageHash,before);assert.equal(h.calls.mutations,1);assert.equal(h.calls.saves,1);const counts=copy(h.calls);paused(await h.controller.start(h.request));assert.deepEqual(h.calls,counts);paused(await h.controller.recover());assert.equal(h.calls.mutations,1);
});
test('unavailable read exceptions and malformed checkpoint envelopes fail closed without invoking envelope getters',async()=>{
  let getters=0;for(const defect of ['null-error','getter','extra','symbol']){const h=await controllerFixture();if(defect==='null-error')h.hooks.read=async()=>{throw null;};else h.adapters.captureCheckpoint=async s=>{const e={documentId:s.documentId,captureId:s.captureId,binding:h.binding};if(defect==='getter')Object.defineProperty(e,'binding',{enumerable:true,get(){getters++;return h.binding;}});if(defect==='extra')e.extra=true;if(defect==='symbol')e[Symbol('hidden')]=true;return e;};paused(await h.controller.start(h.request));assert.equal(h.calls.mutations,0);assert.equal(h.calls.saves,0);}assert.equal(getters,0);
});

test('opaque checkpoint preparation cannot be reassigned to another document session even with identical course and transaction data',async()=>{
  const h=await controllerFixture();h.physicalFailure={status:'uncertain',change:'full'};paused(await h.controller.start(h.request));const journal=copy(h.snapshot.journal);h.snapshot.documentId=crypto.randomUUID();h.touch();const saves=h.calls.saves;paused(await h.controller.recover());paused(await createIconController(h.adapters).recover(h.capabilities()));assert.equal(h.calls.saves,saves);assert.equal(h.calls.mutations,1);assert.deepEqual(h.snapshot.journal,journal);
});

test('initial completion reinspects latest committed capture and pauses on content changed after save before acknowledgment',async()=>{
  for(const defect of ['manual','image','label','unknown-inspection']){const h=await controllerFixture();let savedCapture;
    h.hooks.afterPersist=async n=>{if(n!==2)return;savedCapture=h.snapshot.captureId;if(defect==='manual')h.changeManual();if(defect==='image'){h.imageHash='f'.repeat(64);h.touch();}if(defect==='label'){h.label='Changed after save';h.touch();}if(defect==='unknown-inspection'){h.unrelated='unknown';h.touch();}};
    paused(await h.controller.start(h.request));assert.equal(h.snapshot.journal.phase,'committed');assert.deepEqual(h.snapshot.project,h.snapshot.journal.project);assert.notEqual(h.snapshot.captureId,savedCapture);assert.equal(h.calls.saves,2);assert.equal(h.calls.mutations,1);assert.equal(h.calls.inspections,3);const records=copy({project:h.snapshot.project,journal:h.snapshot.journal});paused(await h.controller.recover());assert.equal(h.calls.saves,2);assert.equal(h.calls.mutations,1);assert.deepEqual({project:h.snapshot.project,journal:h.snapshot.journal},records);
  }
});
test('recovery metadata completion reinspects the new committed capture without saving or replaying again after a post-save change',async()=>{
  for(const defect of ['manual','image','label','unavailable']){const h=await controllerFixture();h.physicalFailure={status:'uncertain',change:'full'};paused(await h.controller.start(h.request));let committedCapture;
    h.hooks.afterPersist=async n=>{if(n!==2)return;committedCapture=h.snapshot.captureId;if(defect==='manual')h.changeManual();if(defect==='image'){h.imageHash='e'.repeat(64);h.touch();}if(defect==='label'){h.label='Changed after recovery save';h.touch();}if(defect==='unavailable'){h.touch();h.hooks.inspect=async()=>{throw new Error('Committed inspection unavailable');};}};
    paused(await h.controller.recover());assert.equal(h.snapshot.journal.phase,'committed');assert.notEqual(h.snapshot.captureId,committedCapture);assert.equal(h.calls.saves,2);assert.equal(h.calls.mutations,1);assert.equal(h.calls.inspections,3);const saves=h.calls.saves;paused(await h.controller.recover());assert.equal(h.calls.saves,saves);assert.equal(h.calls.mutations,1);
  }
});

test('changes during the latest committed inspection fail its exact reread and preserve durable records',async()=>{
  const h=await controllerFixture();h.hooks.inspect=async()=>{if(h.calls.inspections===3)h.changeManual();};paused(await h.controller.start(h.request));assert.equal(h.snapshot.journal.phase,'committed');assert.equal(h.calls.saves,2);assert.equal(h.calls.mutations,1);assert.equal(h.calls.inspections,3);assert.match(h.manual,/CHANGED/);
});

test('disconnected controller conservatively refuses legacy committed receipt admission without changing the live store',async()=>{
  for(const schema of [2,3]){const h=await controllerFixture();h.snapshot.journal={schema,kind:'activity',phase:'committed',txId:crypto.randomUUID(),project:copy(h.snapshot.project)};paused(await h.controller.start(h.request));assert.equal(h.calls.saves,0);assert.equal(h.calls.mutations,0);assert.equal(h.calls.checkpoints,0);assert.equal(h.snapshot.journal.schema,schema);}
});
