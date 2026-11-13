// Isolated declarations/counters only; no Office, UI or live document actions.
import JSZip from 'jszip';
import {ownedFixture} from './owned-fixture.mjs';
import {assets,installParser} from './checkpoint-fixture.mjs';
import {createCheckpoint,bindSavedCheckpoint,verifyPreparedCheckpointCapture} from '../../src/checkpoint.mjs';
import {createIconController} from '../../src/icon-controller.mjs';
import {xml} from '../../src/worksheet.mjs';
installParser();
const copy=x=>structuredClone(x),same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
export async function controllerFixture(){
  const f=await ownedFixture(),master=await f.zip.generateAsync({type:'uint8array'}),backup=await createCheckpoint(master,f.project,JSZip,assets),binding=await bindSavedCheckpoint(backup.data,master,f.project,JSZip);
  const h={snapshot:{documentId:crypto.randomUUID(),captureId:crypto.randomUUID(),project:copy(f.project),journal:null,persistence:'confirmed'},manual:'KEEP MANUAL CONTENT',imageHash:f.item.cueImage.assetHash,label:`${f.item.icon} · ${f.item.minutes} min`,ownership:'verified',placement:'unchanged',dimensions:'unchanged',format:'unchanged',unrelated:'unchanged',calls:{reads:0,checkpoints:0,proofs:0,inspections:0,saves:0,mutations:0},saveFailures:new Map(),hooks:{},binding,preparation:null};
  h.touch=()=>{h.snapshot.captureId=crypto.randomUUID();};
  h.changeManual=()=>{h.manual+=' CHANGED';h.unrelated='changed';h.touch();};
  h.makePreparation=async journal=>{
    const z=await JSZip.loadAsync(master),path='word/webextensions/webextension1.xml';z.file(path,(await z.file(path).async('string')).replace('</we:properties>',`<we:property name="openstage.guide.pending.v1" value="${xml(JSON.stringify(journal))}"/></we:properties>`));return verifyPreparedCheckpointCapture(binding,await z.generateAsync({type:'uint8array'}),journal,JSZip);
  };
  h.adapters={
    readState:async()=>{const n=++h.calls.reads;await h.hooks.read?.(n,h);return copy(h.snapshot);},
    captureCheckpoint:async state=>{h.calls.checkpoints++;await h.hooks.checkpoint?.(state,h);return {documentId:state.documentId,captureId:state.captureId,binding:h.checkpointOverride||binding};},
    verifyPrepared:async(state,journal)=>{h.calls.proofs++;const preparation=await h.makePreparation(journal);h.preparation=preparation;await h.hooks.proof?.(state,journal,h);return {documentId:state.documentId,captureId:state.captureId,preparation:h.proofOverride||preparation};},
    inspectPhysical:async(context,journal)=>{
      h.calls.inspections++;await h.hooks.inspect?.(context,journal,h);return {...context,ownerTag:journal.icon.ownerTag,imageHash:h.imageHash,displayedLabel:h.label,ownership:h.ownership,placement:h.placement,dimensions:h.dimensions,format:h.format,unrelated:h.unrelated};
    },
    persistExact:async(expected,next)=>{
      const n=++h.calls.saves;await h.hooks.persist?.(n,expected,next,h);
      if(!same(h.snapshot,expected))return {status:'failed'};
      const failure=h.saveFailures.get(n),mode=failure?.persist||'full';
      if(mode==='full'||mode==='project')h.snapshot.project=copy(next.project);
      if(mode==='full'||mode==='journal')h.snapshot.journal=copy(next.journal);
      if(mode!=='none')h.touch();
      await h.hooks.afterPersist?.(n,h);
      if(failure?.status==='throw')throw new Error('Synthetic persistence callback failure');return {status:failure?.status||'confirmed'};
    },
    mutatePhysicalOnce:async(expected,journal,context)=>{
      h.calls.mutations++;await h.hooks.mutate?.(expected,journal,context,h);if(!same(h.snapshot,expected))return {status:'failed'};
      const failure=h.physicalFailure,mode=failure?.change||'full';
      if(mode==='full'||mode==='image')h.imageHash=journal.icon.after.assetHash;
      if(mode==='full'||mode==='label')h.label=`${journal.icon.after.icon} · ${journal.project.modules[0].activities[0].minutes} min`;
      if(mode!=='none')h.touch();
      if(failure?.status==='throw')throw new Error('Synthetic physical callback failure');return {status:failure?.status||'confirmed'};
    }
  };
  h.controller=createIconController(h.adapters);h.request={project:copy(f.project),id:f.item.id,icon:'Assessment'};
  h.capabilities=()=>({binding,preparation:h.preparation});return h;
}
