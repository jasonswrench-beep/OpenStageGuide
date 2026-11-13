// Disconnected fixture-adapter controller only. No Office/app/store integration.
// Exact reads/conditional persistence are injected contracts, not Word CAS APIs.
import {cloneIconData,validateIconProject,validateIconJournal,createIconJournal,proposeIconTransition} from './icon-transition.mjs';
import {matchesCheckpointBase,matchesCheckpointPreparation} from './checkpoint.mjs';
import {createIconRecoveryContract} from './icon-recovery.mjs';
const stable=x=>Array.isArray(x)?x.map(stable):x&&typeof x==='object'?Object.fromEntries(Object.keys(x).sort().map(k=>[k,stable(x[k])])):x;
const same=(a,b)=>JSON.stringify(stable(a))===JSON.stringify(stable(b));
const uuid=x=>typeof x==='string'&&/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/.test(x);
function refuse(message){throw new Error(message);}
// Same-module session correlation only, not native document authentication.
// Opaque capture proofs cannot be reassigned to a different document-session ID.
const documentScopes=new WeakMap();
export function createIconController(adapters={}){
  let running=false,retained=null,verified=null,stage='idle';
  const physical=createIconRecoveryContract({inspect:(...args)=>adapters.inspectPhysical(...args)});
  const answer=(status,reason,journal=retained?.journal)=>Object.freeze({status,reason,stage,retainedJournal:journal?cloneIconData(journal):null,replayContent:false,writeAuthorized:false,nativeVerified:false});
  async function read(){
    const s=cloneIconData(await adapters.readState());
    if(!s||!same(Object.keys(s).sort(),['captureId','documentId','journal','persistence','project'])||!uuid(s.documentId)||!uuid(s.captureId)||s.persistence!=='confirmed')refuse('Current storage or capture is unavailable or uncertain.');
    validateIconProject(s.project);if(s.journal!==null)validateIconJournal(s.journal);return s;
  }
  async function exact(expected){const s=await read();if(!same(s,expected))refuse('The exact document, capture or stored state changed.');return s;}
  function evidence(envelope,state,key){
    if(!envelope||typeof envelope!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(envelope)))refuse('Unsupported capture evidence envelope.');
    const keys=Reflect.ownKeys(envelope),descriptors=Object.getOwnPropertyDescriptors(envelope);
    if(keys.some(k=>typeof k!=='string')||!same([...keys].sort(),['captureId','documentId',key].sort())||keys.some(k=>!descriptors[k].enumerable||!Object.hasOwn(descriptors[k],'value')))refuse('Unsupported capture evidence envelope.');
    if(envelope.documentId!==state.documentId||envelope.captureId!==state.captureId||!envelope[key])refuse('Capture evidence belongs to another document or stale capture.');return envelope[key];
  }
  async function inspect(state,journal,binding,preparation,before=false){
    const token=physical.begin(state.documentId,journal.txId,state.captureId),observation=await physical.observe(token,journal);
    await exact(state);
    const args={token,observation,journal,storedProject:state.project,persistence:state.persistence,binding,preparation};
    const result=before?physical.assessBefore(args):physical.classify(args);
    if(result.decision!== (before?'eligible-initial-mutation':journal.phase==='committed'?'recognize-committed':'eligible-metadata-completion'))refuse(result.reason);
    return {token,observation,captureId:state.captureId};
  }
  async function persist(expected,project,journal){
    await exact(expected);let outcome;
    try{outcome=cloneIconData(await adapters.persistExact(cloneIconData(expected),{project:cloneIconData(project),journal:validateIconJournal(journal)}));}catch{outcome={status:'uncertain'};}
    const actual=await read();
    if(!outcome||!same(Object.keys(outcome),['status'])||outcome.status!=='confirmed')refuse('Persistence failed or is uncertain; actual state must be recovered afresh.');
    if(actual.documentId!==expected.documentId||!same(actual.project,project)||!same(actual.journal,journal))refuse('Persistence did not retain the exact complete project and journal.');
    return actual;
  }
  async function commit(state,journal,binding,preparation){
    if(!matchesCheckpointPreparation(binding,preparation,{...journal,phase:'prepared'}))refuse('The complete prepared capture proof no longer matches.');
    await exact(state);const receipt=validateIconJournal({...journal,phase:'committed'});
    stage='commit';retained={journal:receipt,binding,preparation};
    const committed=await persist(state,receipt.project,receipt);
    // Persistence changes the captured package and may yield while manual or
    // physical content changes. Verify this latest committed capture afresh;
    // the earlier prepared/after observation cannot seal completion.
    stage='inspect-committed';await inspect(committed,receipt,binding,preparation);await exact(committed);
    retained=null;return answer('completed','Full committed metadata and latest after-state reread confirmed.',receipt);
  }
  async function exclusive(action){
    if(running)return answer('busy','Another icon operation is running.');running=true;
    try{return await action();}catch(error){return answer('paused',error instanceof Error?error.message:'An adapter operation failed or is uncertain.');}finally{running=false;}
  }
  async function start({project,id,icon}={}){
    return exclusive(async()=>{
      stage='read-base';if(retained)refuse('A prior operation is unresolved; recovery must not replay its physical request.');
      const base=await read(),reviewed=validateIconProject(project);
      if(!same(base.project,reviewed))refuse('The full reviewed guide differs from current storage.');
      if(base.journal&&(base.journal.phase!=='committed'||!same(base.journal.project,base.project)))refuse('An unresolved operation already exists.');
      // No-op before checkpoint/persistence; validates the complete proposal.
      if(!proposeIconTransition(base.project,id,icon))return answer('noop','The selected icon is unchanged.',base.journal);
      stage='checkpoint';const binding=evidence(await adapters.captureCheckpoint(cloneIconData(base)),base,'binding');
      if(!matchesCheckpointBase(binding,base.project))refuse('The saved checkpoint is not the exact current base.');await exact(base);
      let scope=documentScopes.get(binding);
      if(scope&&scope.documentId!==base.documentId)refuse('This checkpoint is already scoped to another document session.');
      if(!scope){scope={documentId:base.documentId,preparations:new WeakSet()};documentScopes.set(binding,scope);}
      const journal=createIconJournal(base.project,id,icon,{id:binding.checkpointId,masterHash:binding.masterHash,backupHash:binding.backupHash});
      stage='prepare';retained={journal,binding,preparation:null};
      const preparedState=await persist(base,journal.base,journal);
      stage='verify-prepared';const preparation=evidence(await adapters.verifyPrepared(cloneIconData(preparedState),validateIconJournal(journal),binding),preparedState,'preparation');
      if(!matchesCheckpointPreparation(binding,preparation,journal))refuse('The exact prepared capture was not verified.');retained.preparation=preparation;await exact(preparedState);scope.preparations.add(preparation);verified={binding,preparation,txId:journal.txId};
      stage='inspect-before';const context=await inspect(preparedState,journal,binding,preparation,true);await exact(preparedState);
      stage='physical-attempted';let outcome;
      // No retries anywhere. Retained state blocks another start after an
      // uncertain attempt; recover() has no path to this adapter method.
      try{outcome=cloneIconData(await adapters.mutatePhysicalOnce(cloneIconData(preparedState),validateIconJournal(journal),context));}catch{outcome={status:'uncertain'};}
      const after=await read();
      if(!outcome||!same(Object.keys(outcome),['status'])||outcome.status!=='confirmed')refuse('Physical mutation failed or is uncertain; it must never be replayed.');
      if(after.documentId!==preparedState.documentId||!same(after.project,journal.base)||!same(after.journal,journal))refuse('Stored state changed during the physical attempt.');
      stage='inspect-after';await inspect(after,journal,binding,preparation);return commit(after,journal,binding,preparation);
    });
  }
  async function recover({binding,preparation}={}){
    return exclusive(async()=>{
      stage='recovery-read';const state=await read(),journal=validateIconJournal(state.journal);
      if(retained&&retained.journal.txId!==journal.txId)refuse('A different unresolved transaction was found.');
      const session=retained||(verified?.txId===journal.txId?verified:null);
      binding=binding||session?.binding;preparation=preparation||session?.preparation;
      if(!matchesCheckpointPreparation(binding,preparation,{...journal,phase:'prepared'}))refuse('Reverify the complete saved checkpoint and preparation in this session.');
      const scope=documentScopes.get(binding);
      if(!scope||scope.documentId!==state.documentId||!scope.preparations.has(preparation))refuse('The prepared evidence is not scoped to this document session.');
      verified={binding,preparation,txId:journal.txId};
      stage='recovery-inspect';await inspect(state,journal,binding,preparation);
      if(journal.phase==='committed'){retained=null;return answer('recognized','Exact committed after-state recognized with no saves.',journal);}
      return commit(state,journal,binding,preparation);
    });
  }
  return Object.freeze({start,recover});
}
