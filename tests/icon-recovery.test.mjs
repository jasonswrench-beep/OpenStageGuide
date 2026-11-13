import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {ownedFixture} from './helpers/owned-fixture.mjs';
import {installParser,assets} from './helpers/checkpoint-fixture.mjs';
import {createCheckpoint,bindSavedCheckpoint,verifyPreparedCheckpointCapture,matchesCheckpointPreparation} from '../src/checkpoint.mjs';
import {createIconJournal} from '../src/icon-transition.mjs';
import {createIconRecoveryContract} from '../src/icon-recovery.mjs';
import {xml} from '../src/worksheet.mjs';
installParser();
function report(context,journal){const item=journal.project.modules[0].activities[0];return {...context,ownerTag:journal.icon.ownerTag,imageHash:journal.icon.after.assetHash,displayedLabel:`${item.icon} · ${item.minutes} min`,ownership:'verified',placement:'unchanged',dimensions:'unchanged',format:'unchanged',unrelated:'unchanged'};}
async function setup({phase='prepared',inspect}={}){
  const f=await ownedFixture(),master=await f.zip.generateAsync({type:'uint8array'}),backup=await createCheckpoint(master,f.project,JSZip,assets),binding=await bindSavedCheckpoint(backup.data,master,f.project,JSZip);
  const prepared=createIconJournal(f.project,f.item.id,'Assessment',{id:binding.checkpointId,masterHash:binding.masterHash,backupHash:binding.backupHash}),path='word/webextensions/webextension1.xml';f.zip.file(path,(await f.zip.file(path).async('string')).replace('</we:properties>',`<we:property name="openstage.guide.pending.v1" value="${xml(JSON.stringify(prepared))}"/></we:properties>`));
  const preparation=await verifyPreparedCheckpointCapture(binding,await f.zip.generateAsync({type:'uint8array'}),prepared,JSZip),journal={...prepared,phase};let calls=0;
  const contract=createIconRecoveryContract({inspect:async(...args)=>{calls++;return inspect?inspect(...args):report(...args);}}),documentId=crypto.randomUUID(),token=contract.begin(documentId,journal.txId);
  const args={token,journal,storedProject:phase==='prepared'?journal.base:journal.project,persistence:'confirmed',binding,preparation};
  return {contract,args,prepared,documentId,calls:()=>calls};
}
test('disabled recovery contract correlates real opaque checkpoint/capture proofs and repeats eligibility without any effects',async()=>{
  const f=await setup(),before=structuredClone(f.args.journal);f.args.observation=await f.contract.observe(f.args.token,f.args.journal);
  assert.equal(matchesCheckpointPreparation(f.args.binding,f.args.preparation,f.prepared),true);
  for(let i=0;i<3;i++){const result=f.contract.classify(f.args);assert.equal(result.decision,'eligible-metadata-completion');assert.equal(result.retainJournal,true);assert.equal(result.replayContent,false);assert.equal(result.writeAuthorized,false);assert.equal(result.nativeVerified,false);}
  assert.equal(f.calls(),1);assert.deepEqual(f.args.journal,before);assert.deepEqual(f.args.storedProject,before.base);assert.equal(f.args.observation.nativeVerified,false);
});
test('complete committed declaration recognizes repeatedly without saving while incompatible stored phases remain paused',async()=>{
  const f=await setup({phase:'committed'});f.args.observation=await f.contract.observe(f.args.token,f.args.journal);
  for(let i=0;i<3;i++)assert.equal(f.contract.classify(f.args).decision,'recognize-committed');assert.equal(f.calls(),1);
  assert.equal(f.contract.classify({...f.args,storedProject:f.args.journal.base}).decision,'pause');
  const g=await setup();g.args.observation=await g.contract.observe(g.args.token,g.args.journal);assert.equal(g.contract.classify({...g.args,storedProject:g.args.journal.project}).decision,'pause');
});
test('exact before, mixed image-label, lost ownership and changed or unknown preservation evidence always pause',async()=>{
  const changes=[(r,j)=>{r.imageHash=j.icon.before.assetHash;r.displayedLabel=`${j.icon.before.icon} · 10 min`;},(r,j)=>r.imageHash=j.icon.before.assetHash,(r,j)=>r.displayedLabel=`${j.icon.before.icon} · 10 min`,r=>r.imageHash='f'.repeat(64),r=>r.displayedLabel='Other label',...['placement','dimensions','format','unrelated'].flatMap(key=>['changed','unknown'].map(value=>r=>r[key]=value)),...['missing','ambiguous','unowned','unknown'].map(value=>r=>r.ownership=value)];
  for(const change of changes){const f=await setup({inspect:(c,j)=>{const r=report(c,j);change(r,j);return r;}});f.args.observation=await f.contract.observe(f.args.token,f.args.journal);const result=f.contract.classify(f.args);assert.equal(result.decision,'pause');assert.equal(result.retainJournal,true);assert.equal(result.replayContent,false);}
});
test('forged copied cross-session and stale observations or tokens are refused',async()=>{
  const f=await setup();f.args.observation=await f.contract.observe(f.args.token,f.args.journal);
  for(const observation of [{...f.args.observation},JSON.parse(JSON.stringify(f.args.observation)),undefined])assert.equal(f.contract.classify({...f.args,observation}).decision,'pause');
  assert.equal(f.contract.classify({...f.args,token:{...f.args.token}}).decision,'pause');
  const other=f.contract.begin(crypto.randomUUID(),f.args.journal.txId);assert.equal(f.contract.classify({...f.args,token:other}).decision,'pause');
  const foreign=createIconRecoveryContract({inspect:report});assert.equal(foreign.classify(f.args).decision,'pause');
  f.contract.begin(f.documentId,f.args.journal.txId);assert.equal(f.contract.classify(f.args).decision,'pause');await assert.rejects(f.contract.observe(f.args.token,f.args.journal),/stale/);
});
test('wrong transaction capture or document inspection context is refused rather than becoming physical authority',async()=>{
  for(const key of ['documentId','txId','captureId','generation','ownerTag']){const f=await setup({inspect:(c,j)=>({...report(c,j),[key]:key==='generation'?c.generation+1:crypto.randomUUID()})});await assert.rejects(f.contract.observe(f.args.token,f.args.journal));assert.equal(f.contract.classify(f.args).decision,'pause');}
});
test('persistence uncertainty changed full metadata journals and fake checkpoint or delta receipts remain paused',async()=>{
  const f=await setup();f.args.observation=await f.contract.observe(f.args.token,f.args.journal);
  for(const persistence of ['failed','unknown',undefined,'committed'])assert.equal(f.contract.classify({...f.args,persistence}).decision,'pause');
  for(const key of ['binding','preparation'])assert.equal(f.contract.classify({...f.args,[key]:{...f.args[key]}}).decision,'pause');
  const stored=structuredClone(f.args.storedProject);stored.modules[0].activities[0].privateUnknown.preserve='Changed unrelated data';assert.equal(f.contract.classify({...f.args,storedProject:stored}).decision,'pause');
  for(const key of ['txId','checkpoint','project']){const journal=structuredClone(f.args.journal);if(key==='txId')journal.txId=crypto.randomUUID();else if(key==='checkpoint')journal.checkpoint.id=crypto.randomUUID();else journal.project.modules[0].activities[0].answers+=' unrelated';assert.equal(f.contract.classify({...f.args,journal}).decision,'pause');assert.equal(matchesCheckpointPreparation(f.args.binding,f.args.preparation,journal),false);}
});
test('inspection failures unavailable evidence asynchronous staleness and changed journal never produce an eligible observation',async()=>{
  for(const inspect of [async()=>{throw new Error('Inspection unavailable');},async()=>null,async()=>({})]){const f=await setup({inspect});await assert.rejects(f.contract.observe(f.args.token,f.args.journal));assert.equal(f.contract.classify(f.args).decision,'pause');}
  let release,entered;const gate=new Promise(ok=>release=ok),start=new Promise(ok=>entered=ok),f=await setup({inspect:async(c,j)=>{entered();await gate;return report(c,j);}}),pending=f.contract.observe(f.args.token,f.args.journal);await start;f.contract.begin(f.documentId,f.args.journal.txId);release();await assert.rejects(pending,/stale/);
  const g=await setup({inspect:async(c,j)=>{g.args.journal.project.modules[0].activities[0].answers+='late';return report(c,j);}});await assert.rejects(g.contract.observe(g.args.token,g.args.journal));
});

test('every newer inspection attempt invalidates a prior good observation even on failure mixed result or invalid journal',async()=>{
  for(const next of ['failure','mixed','invalid']){
    let calls=0;const f=await setup({inspect:async(c,j)=>{if(++calls===1)return report(c,j);if(next==='failure')throw new Error('New inspection failed');return {...report(c,j),displayedLabel:'Mixed label'};}});f.args.observation=await f.contract.observe(f.args.token,f.args.journal);assert.equal(f.contract.classify(f.args).decision,'eligible-metadata-completion');
    if(next==='invalid')await assert.rejects(f.contract.observe(f.args.token,{...f.args.journal,schema:3}));
    else if(next==='failure')await assert.rejects(f.contract.observe(f.args.token,f.args.journal));
    else {const latest=await f.contract.observe(f.args.token,f.args.journal);assert.equal(f.contract.classify({...f.args,observation:latest}).decision,'pause');}
    assert.equal(f.contract.classify(f.args).decision,'pause');
  }
});
test('pending and out-of-order inspections cannot revive an earlier observation and latest unchanged result repeats without effects',async()=>{
  const waits=[];let calls=0;const f=await setup({inspect:(c,j)=>++calls===1?report(c,j):new Promise(resolve=>waits.push(()=>resolve(report(c,j))))});f.args.observation=await f.contract.observe(f.args.token,f.args.journal);
  const second=f.contract.observe(f.args.token,f.args.journal);assert.equal(f.contract.classify(f.args).decision,'pause');const secondRejected=assert.rejects(second,/stale/);
  const third=f.contract.observe(f.args.token,f.args.journal);assert.equal(waits.length,2);waits[1]();const latest=await third;waits[0]();await secondRejected;
  assert.equal(f.contract.classify(f.args).decision,'pause');for(let i=0;i<3;i++){const decision=f.contract.classify({...f.args,observation:latest});assert.equal(decision.decision,'eligible-metadata-completion');assert.equal(decision.writeAuthorized,false);assert.equal(decision.replayContent,false);}
  assert.equal(calls,3);
});
test('exported checkpoint correlation helper rejects invisible malformed data without invoking getters',async()=>{
  const f=await setup();let getters=0;assert.equal(matchesCheckpointPreparation(f.args.binding,f.args.preparation,f.prepared),true);
  for(const defect of ['undefined','symbol','hidden','getter','sparse-extra','nan','boxed']){const j=structuredClone(f.prepared);
    if(defect==='undefined')j.extra=undefined;
    if(defect==='symbol')j[Symbol('invisible')]='hidden';
    if(defect==='hidden')Object.defineProperty(j,'extra',{value:'hidden',enumerable:false});
    if(defect==='getter')Object.defineProperty(j,'base',{enumerable:true,get(){getters++;return f.prepared.base;}});
    if(defect==='sparse-extra'){j.extra=Array(1);j.extra.foo='hidden';}
    if(defect==='nan')j.extra=NaN;
    if(defect==='boxed')j.txId=new String(j.txId);
    assert.equal(matchesCheckpointPreparation(f.args.binding,f.args.preparation,j),false,defect);
  }
  assert.equal(getters,0);assert.equal(matchesCheckpointPreparation(f.args.binding,f.args.preparation,f.prepared),true);
});
