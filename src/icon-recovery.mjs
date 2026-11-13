// Disabled read-only adapter contract, not a Word adapter or mutation authority.
// The injected inspector must eventually supply independently verified physical
// evidence. Tests exercise its contract only; no native freshness is inferred.
import {validateIconJournal,validateIconProject,cloneIconData} from './icon-transition.mjs';
import {matchesCheckpointPreparation} from './checkpoint.mjs';
const uuid=x=>typeof x==='string'&&/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/.test(x);
const stable=x=>Array.isArray(x)?x.map(stable):x&&typeof x==='object'?Object.fromEntries(Object.keys(x).sort().map(k=>[k,stable(x[k])])):x;
const same=(a,b)=>JSON.stringify(stable(a))===JSON.stringify(stable(b));
const result=(decision,reason)=>Object.freeze({decision,reason,retainJournal:true,replayContent:false,writeAuthorized:false,nativeVerified:false});
const pause=reason=>result('pause',reason);
const fail=()=>{throw new Error('Unsupported or stale physical-inspection contract. No recovery write is authorized.');};
export function createIconRecoveryContract({inspect}={}){
  const tokens=new WeakMap(),observations=new WeakMap(),attempts=new WeakMap(),current=new Map();
  function begin(documentId,txId,captureId=crypto.randomUUID()){
    if(!uuid(documentId)||!uuid(txId)||!uuid(captureId))fail();
    const context={documentId,txId,captureId,generation:(current.get(documentId)?.generation||0)+1};
    current.set(documentId,context);const token=Object.freeze({scope:'disabled-icon-recovery-capture',nativeVerified:false});tokens.set(token,context);return token;
  }
  const fresh=context=>current.get(context.documentId)===context;
  async function observe(token,journal){
    const context=tokens.get(token),attempt={};
    // Starting any newer attempt invalidates all older observations immediately,
    // including when validation/inspection fails or remains pending.
    if(context)attempts.set(token,attempt);
    const expected=validateIconJournal(journal);
    if(!context||!fresh(context)||context.txId!==expected.txId||typeof inspect!=='function')fail();
    // Inspect receives isolated values; it cannot mutate private session state.
    const data=cloneIconData(await inspect(structuredClone(context),structuredClone(expected)));
    const names=['documentId','txId','captureId','generation','ownerTag','imageHash','displayedLabel','ownership','placement','dimensions','format','unrelated'];
    if(!data||!same(Object.keys(data).sort(),names.sort())||!same(Object.fromEntries(['documentId','txId','captureId','generation'].map(k=>[k,data[k]])),context)||data.ownerTag!==expected.icon.ownerTag||typeof data.imageHash!=='string'||!/^[a-f0-9]{64}$/.test(data.imageHash)||typeof data.displayedLabel!=='string'||!['verified','missing','ambiguous','unowned','unknown'].includes(data.ownership)||['placement','dimensions','format','unrelated'].some(k=>!['unchanged','changed','unknown'].includes(data[k])))fail();
    if(!fresh(context)||attempts.get(token)!==attempt||!same(validateIconJournal(journal),expected))fail();
    const observation=Object.freeze({scope:'disabled-physical-inspector-contract',nativeVerified:false});observations.set(observation,{context,attempt,journal:expected,data:structuredClone(data)});return observation;
  }
  function assess({token,observation,journal,storedProject,persistence,binding,preparation}={},phase='after'){
    try{
      const context=tokens.get(token),seen=observations.get(observation),checked=validateIconJournal(journal);
      if(!context||!seen||seen.context!==context||attempts.get(token)!==seen.attempt||!fresh(context)||context.txId!==checked.txId||!same(seen.journal,checked))return pause('Wrong document, transaction or stale observation.');
      const prepared={...checked,phase:'prepared'};
      if(!matchesCheckpointPreparation(binding,preparation,prepared))return pause('Reverify the exact checkpoint and prepared capture in this session.');
      if(persistence!=='confirmed')return pause('Persistence is failed or uncertain.');
      const stored=validateIconProject(storedProject),data=seen.data;
      if(data.ownership!=='verified'||['placement','dimensions','format','unrelated'].some(k=>data[k]!=='unchanged'))return pause('Physical ownership or preservation evidence is incomplete.');
      const item=checked.project.modules.flatMap(m=>m.activities).find(a=>a.id===checked.icon.id);
      if(data.imageHash!==checked.icon[phase].assetHash||data.displayedLabel!==`${checked.icon[phase].icon} · ${item.minutes} min`)return pause('Physical image and displayed label do not prove the complete expected state.');
      if(phase==='before')return checked.phase==='prepared'&&same(stored,checked.base)?result('eligible-initial-mutation','Before-state declarations agree; no mutation authority is conferred.'):pause('Initial mutation requires the exact prepared base.');
      if(checked.phase==='prepared'&&same(stored,checked.base))return result('eligible-metadata-completion','Data-contract prerequisites agree; a future verified controller must recheck before any write.');
      if(checked.phase==='committed'&&same(stored,checked.project))return result('recognize-committed','Exact committed metadata and declared physical after-state agree; no save or content replay.');
      return pause('Stored metadata differs from the exact expected phase.');
    }catch{return pause('Invalid or incomplete recovery evidence.');}
  }
  return Object.freeze({begin,observe,classify:args=>assess(args),assessBefore:args=>assess(args,'before')});
}
