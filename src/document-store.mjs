import {hasFrontMatter,frontMatterEvidence,validFrontJournal} from './frontmatter.mjs';
import {validate,reconcile,editActivitySettings,editCourseDetails,editActivityMaterials} from './model.mjs';
import {editObjectiveCoverage} from './objective-coverage.mjs';
import {editDeliveryPlan} from './delivery-plans.mjs';

export const KEY='openstage.guide.project.v1',PENDING='openstage.guide.pending.v1';
const clone=x=>x==null?null:structuredClone(x);
const stable=x=>Array.isArray(x)?x.map(stable):x&&typeof x==='object'?Object.fromEntries(Object.keys(x).sort().map(k=>[k,stable(x[k])])):x;
const same=(a,b)=>JSON.stringify(stable(a))===JSON.stringify(stable(b));
function problem(message,code='OSG_INTEGRITY'){const e=new Error(message);e.code=code;return e;}
const pendingError=()=>problem('Course operations are paused. Preserve this document and use Check and recover before continuing.');
const staleError=()=>problem('The document changed. Refresh the course before trying again.','OSG_STALE');

// One store per document runtime. Dependencies are lazy Office adapters in Word,
// or isolated, non-native hosts in the fault tests.
export function createDocumentStore({settings,scan,write,verifyPlacement=async()=>{},checkActivitySettings,writeActivitySettings,verifyFrontMatter,verifyBlank}){
  let fault=null,running=false;
  const api=()=>settings();
  const hostCall=method=>new Promise((resolve,reject)=>{
    try{api()[method](r=>r.status==='succeeded'?resolve():reject(problem(`Word could not ${method==='saveAsync'?'save':'refresh'} the course record. Its persistence is uncertain. Preserve the document. ${r.error?.message||''}`)));}
    catch(e){reject(e);}
  });
  const refreshSettings=()=>hostCall('refreshAsync'),save=()=>hostCall('saveAsync');
  function quarantine(e,journal){
    fault={journal:clone(journal||fault?.journal),cause:e};
    // A failed callback may have persisted everything or nothing. Keep an
    // unresolved local journal; never save a guessed rollback over host state.
    if(fault.journal)api().set(PENDING,clone(fault.journal));
    return problem(`${e.message} Course operations remain paused; use Check and recover.`);
  }
  async function exclusive(action,{recovery=false}={}){
    if(running)throw problem('A course operation is already running. Please wait.','OSG_BUSY');
    if(fault&&!recovery)throw pendingError();
    running=true;
    try{return await action();}
    catch(e){if(['OSG_STALE','OSG_BUSY','OSG_CHECKPOINT','OSG_INPUT'].includes(e.code))throw e;throw quarantine(e);}
    finally{running=false;}
  }
  function stored(){const p=api().get(KEY);return p==null?null:clone(validate(p));}
  const kinds=['course','module','activity','metadata','activity-settings'];
  const cueChecks=j=>j?.cueChecks||(j?.schema===3&&j.settings?.id?[j.settings.id]:[]);
  const timing=j=>j?.schema===3||j?.schema===5&&j.cueChecks?.length;
  async function profile(p,initial=false){if(!hasFrontMatter(p))return;if(!verifyFrontMatter)throw problem('Front-matter package verification is unavailable.');await verifyPlacement(p);await verifyFrontMatter(p,{initial});}
  function receiptMatches(journal,p){return (hasFrontMatter(p)?validFrontJournal(journal,p):(journal?.schema===2||journal?.schema===3&&kinds.includes(journal.kind)))&&journal.phase==='committed'&&typeof journal.txId==='string'&&same(journal.project,p);}
  async function current(){
    await refreshSettings();const p=stored(),journal=api().get(PENDING);
    if((journal||hasFrontMatter(p))&&!receiptMatches(journal,p)){fault={journal:clone(journal),cause:pendingError()};throw pendingError();}
    if(journal?.placement===true)await verifyPlacement(p);
    return p;
  }
  async function canonical(){
    const result=reconcile(await current(),await scan()),receipt=api().get(PENDING);
    if(timing(receipt))await verifyEvidence(receipt,result);
    await profile(result);
    return result;
  }
  function prepared(base,project,kind){
    const checks=cueChecks(api().get(PENDING));
    return {schema:hasFrontMatter(project)?5:checks.length?3:2,txId:crypto.randomUUID(),phase:'prepared',kind,base:clone(base),project:clone(project),...(checks.length?{cueChecks:clone(checks)}:{}),...(hasFrontMatter(project)?{frontMatter:frontMatterEvidence(project),placement:true,cueChecks:project.modules.flatMap(m=>m.activities.map(a=>a.id))}:{})};
  }
  async function prepare(journal){
    fault={journal:clone(journal),cause:null};
    api().set(PENDING,clone(journal));
    await save(); // No content writes until this acknowledgement succeeds.
    await refreshSettings();
    if(!same(api().get(PENDING),journal)||!same(stored(),journal.base))throw problem('Word did not retain this operation journal, or the course changed during preparation.');
  }
  async function finish(journal){
    const result=reconcile(validate(journal.project),await scan());
    if(hasFrontMatter(result)&&(!validFrontJournal(journal,journal.project)||journal.kind==='course'&&!same(result,journal.project)))throw problem('Initial front-matter completion requires every exact proposed field; divergent partial text cannot be adopted.');
    if(journal.schema===5&&!hasFrontMatter(result))throw problem('The governed operation lost its profile.');
    if(journal.placement===true&&!hasFrontMatter(result))await verifyPlacement(result);
    await profile(result,journal.kind==='course');
    if(timing(journal))await verifyEvidence(journal,result);
    await refreshSettings();
    if(!same(api().get(PENDING),journal)||!same(stored(),journal.base))throw problem('The operation journal or course changed while the operation was running.');
    const receipt={schema:journal.schema,txId:journal.txId,phase:'committed',kind:journal.kind,project:clone(result)};
    if(journal.placement===true)receipt.placement=true;
    if(journal.settings)receipt.settings=clone(journal.settings);
    if(journal.cueChecks)receipt.cueChecks=clone(journal.cueChecks);
    if(journal.schema===5)receipt.frontMatter=clone(journal.frontMatter);
    api().set(KEY,clone(result));api().set(PENDING,receipt);
    await save();
    await refreshSettings();
    if(!receiptMatches(api().get(PENDING),stored())||!same(api().get(PENDING),receipt))throw problem('Word did not confirm the completed operation record.');
    // Retain the last confirmed receipt instead of deleting the only evidence
    // of completion in another fallible settings save.
    if(hasFrontMatter(result)){await profile(result,journal.kind==='course');if(!same(reconcile(result,await scan()),result))throw problem('Front-matter content changed during committed verification.');}
    fault=null;return clone(result);
  }
  async function verifyEvidence(journal,project){
    const checks=cueChecks(journal);
    if(!Array.isArray(checks)||!checks.length||new Set(checks).size!==checks.length||checks.some(id=>typeof id!=='string')||!checkActivitySettings)throw problem('Activity timing validation evidence is incomplete. Preserve this copy.');
    if(journal.kind==='activity-settings')await verifySettingsJournal(journal,project);
    for(const id of checks)await checkActivitySettings(project,id);
  }
  async function verifySettingsJournal(journal,project){
    const s=journal.settings;
    if(!s||!checkActivitySettings||!writeActivitySettings)throw problem('Activity settings recovery is unsupported or incomplete. Preserve this copy.');
    const expected=editActivitySettings(journal.base||project,s.id,s.after).next;
    // Prepared settings edits may change only duration and writing lines.
    if(journal.phase==='prepared'&&!same(expected,journal.project))throw problem('The activity settings journal is inconsistent. Preserve this copy.');
    const item=project.modules.flatMap(m=>m.activities).find(a=>a.id===s.id);
    if(!item||item.minutes!==s.after.minutes||item.lines!==s.after.lines)throw problem('Activity settings do not match the completed operation.');
    if(!cueChecks(journal).includes(s.id))throw problem('Activity settings validation evidence is missing.');
  }
  async function metadata(base,result){
    const receipt=api().get(PENDING);
    if(timing(receipt))await verifyEvidence(receipt,result);
    await profile(result);
    if(same(base,result))return clone(result);
    const journal=prepared(base,result,'metadata');
    if(api().get(PENDING)?.placement===true)journal.placement=true;
    try{await prepare(journal);return await finish(journal);}catch(e){throw quarantine(e,journal);}
  }
  function transition(base,next,kind,item){
    validate(next);let expected=clone(base);
    if(kind==='course'){
      if(base||next.revision!==0||next.modules.length)throw staleError();
      return;
    }
    if(!base||next.revision!==base.revision+1||!item)throw staleError();
    if(kind==='module'){
      if(!Array.isArray(item.activities)||item.activities.length)throw problem('A new module must start without activities.');
      expected.modules.push(clone(item));
    }else if(kind==='activity'){
      const targets=next.modules.filter(m=>m.activities.some(a=>a.id===item.id));
      if(targets.length!==1)throw staleError();
      const target=expected.modules.find(m=>m.id===targets[0].id);
      if(!target)throw staleError();
      target.activities.push(clone(item));
    }
    else throw problem('Unsupported course operation.');
    expected.revision++;
    if(!same(expected,next))throw staleError();
  }
  return {
    isBlocked:()=>!!fault,
    // Acceptance-only read diagnostics share the transaction lock, including
    // while quarantined. They must never refresh the settings bag concurrently.
    inspect:action=>exclusive(action,{recovery:true}),
    loadProject:()=>exclusive(canonical),
    readProject:()=>exclusive(canonical),
    captureMaster:reader=>exclusive(async()=>{
      const base=await current(),fresh=reconcile(base,await scan());
      if(!fresh)throw problem('Create an OpenStage course before saving a backup.','OSG_CHECKPOINT');
      const project=await metadata(base,fresh),master=await reader();
      const after=await canonical();
      if(!same(project,after))throw staleError();
      return {project,master};
    }),
    persist:project=>exclusive(async()=>{
      const base=await current(),result=reconcile(base,await scan());
      if(!same(project,result))throw staleError();
      return metadata(base,result);
    }),
    updateCourseDetails:(project,fields)=>exclusive(async()=>{
      const base=await current(),fresh=reconcile(base,await scan());
      if(!base||!same(project,base)||!same(base,fresh))throw staleError();
      // Same journaled metadata path used for direct text reconciliation.
      // No document body, image, title or instructional field is rewritten.
      let next;try{next=editCourseDetails(base,fields);}catch(e){e.code='OSG_INPUT';throw e;}
      return metadata(base,next);
    }),
    updateActivityMaterials:(project,id,materials)=>exclusive(async()=>{
      const base=await current(),fresh=reconcile(base,await scan());
      if(!base||!same(project,base)||!same(base,fresh))throw staleError();
      let next;try{next=editActivityMaterials(base,id,materials);}catch(e){e.code='OSG_INPUT';throw e;}
      return metadata(base,next);
    }),
    updateObjectiveCoverage:(project,change)=>exclusive(async()=>{
      const base=await current(),fresh=reconcile(base,await scan());
      if(!base||!same(project,base)||!same(base,fresh))throw staleError();
      let next;try{next=editObjectiveCoverage(base,change);}catch(e){e.code='OSG_INPUT';throw e;}
      await metadata(base,next);
      // A Word edit arriving during the final legacy metadata save must not
      // leave this feature presenting the old objective as currently approved.
      // Reconcile it using the same metadata protocol; never update annotations'
      // reviewed snapshots from that newer text. This is a bounded reread, not CAS.
      const after=await canonical(),settled=await metadata(stored(),after);
      // A second save introduces its own final callback boundary. Check it
      // once without another persistence loop: keep completed receipts and
      // raw author drafts, and require a fresh user operation if Word changed.
      const final=await canonical();if(!same(final,settled))throw staleError();
      return settled;
    }),
    updateDeliveryPlan:(project,change)=>exclusive(async()=>{
      const base=await current(),fresh=reconcile(base,await scan());
      if(!base||!same(project,base)||!same(base,fresh))throw staleError();
      let next;try{next=editDeliveryPlan(base,change).next;}catch(e){e.code='OSG_INPUT';throw e;}
      await metadata(base,next);
      // Word remains authoritative for content and duration changes arriving
      // during a settings save. One bounded correction and final read prevent
      // a plan from claiming a fresh preview on a second callback race.
      const after=await canonical(),settled=await metadata(stored(),after);
      const final=await canonical();if(!same(final,settled))throw staleError();
      return settled;
    }),
    updateActivitySettings:(project,id,fields)=>exclusive(async()=>{
      const base=await current(),fresh=reconcile(base,await scan());
      if(!same(project,base)||!same(base,fresh))throw staleError();
      if(timing(api().get(PENDING)))await verifyEvidence(api().get(PENDING),fresh);
      await profile(fresh);
      const {next}=editActivitySettings(base,id,fields);
      if(same(base,next))return clone(base);
      if(!checkActivitySettings||!writeActivitySettings)throw problem('Activity settings editing is unavailable.');
      await verifyPlacement(base);
      await checkActivitySettings(base,id);
      const old=base.modules.flatMap(m=>m.activities).find(a=>a.id===id),item=next.modules.flatMap(m=>m.activities).find(a=>a.id===id);
      const journal=prepared(base,next,'activity-settings');journal.placement=true;
      if(!hasFrontMatter(next))journal.schema=3; // Governed guides must retain schema5.
      journal.cueChecks=[...new Set([...cueChecks(journal),id])];
      journal.settings={id,before:{minutes:old.minutes,lines:old.lines},after:{minutes:item.minutes,lines:item.lines}};
      try{
        await prepare(journal);
        if(!same(reconcile(base,await scan()),base))throw problem('Word text changed during preparation. Settings editing was stopped.');
        await verifyPlacement(base);
        await profile(base);
        await writeActivitySettings(base,next,id);
        return await finish(journal);
      }catch(e){throw quarantine(e,journal);}
    }),
    insert:(project,next,kind,item)=>exclusive(async()=>{
      const base=await current(),fresh=reconcile(base,await scan());
      if(!same(project,base)||!same(base,fresh))throw staleError();
      if(timing(api().get(PENDING)))await verifyEvidence(api().get(PENDING),fresh);
      await profile(fresh);
      transition(base,next,kind,item);
      if(kind==='course'&&hasFrontMatter(next)){if(!verifyBlank)throw problem('Blank-document verification is unavailable.','OSG_INPUT');await verifyBlank();}
      const journal=prepared(base,next,kind);
      if(api().get(PENDING)?.placement===true||kind==='activity')journal.placement=true;
      try{
        await prepare(journal);
        // Re-read before writing so manual edits made during preparation cause
        // a safe stop. Recovery never replays this content mutation.
        if(!same(reconcile(base,await scan()),base))throw problem('Word text changed during preparation. Content insertion was stopped.');
        if(journal.placement===true)await verifyPlacement(base);
        if(kind==='course'&&hasFrontMatter(next))await verifyBlank(journal);else await profile(base);
        await write(next,kind,item);
        return await finish(journal);
      }catch(e){throw quarantine(e,journal);}
    }),
    recover:()=>exclusive(async()=>{
      const retained=clone(fault?.journal);
      await refreshSettings();const base=stored(),durable=api().get(PENDING);
      if(receiptMatches(durable,base)){
        // An uncertain save may actually have committed. Do not advance its
        // revision or replay insertion; only reconcile subsequent manual edits.
        const result=reconcile(base,await scan());
        if(durable.placement===true)await verifyPlacement(result);
        if(timing(durable))await verifyEvidence(durable,result);
        await profile(result);
        fault=null;return metadata(base,result);
      }
      // A failed upgrade save can leave the original legacy journal durable
      // while this runtime retains its prepared replacement. Retry that exact
      // conversion only when the recorded legacy source still matches.
      const sameLegacy=durable?.schema===undefined&&retained?.legacy&&same(durable,retained.legacy);
      const journal=clone(sameLegacy?retained:durable||retained);
      if(!journal){if(hasFrontMatter(base))throw pendingError();const result=reconcile(base,await scan());fault=null;return metadata(base,result);}
      if(!journal.project)throw problem('The operation journal is incomplete. Preserve this copy; automatic recovery cannot continue.');
      if([2,3,5].includes(journal.schema)&&(journal.phase!=='prepared'||typeof journal.txId!=='string'))throw problem('The operation receipt does not match the saved course. Preserve this copy; automatic recovery cannot continue.');
      if(journal.schema!==undefined&&journal.schema!==2&&!(journal.schema===3&&kinds.includes(journal.kind))&&!(journal.schema===5&&validFrontJournal(journal,journal.project)))throw problem('This operation journal version is unsupported. Preserve this copy.');
      if(journal.kind==='activity-settings'&&![3,5].includes(journal.schema))throw problem('This activity settings journal version is unsupported. Preserve this copy.');
      if(durable&&retained&&durable.txId!==retained.txId&&!sameLegacy)throw problem('A different unresolved operation was found. Preserve the document before further recovery.');
      // Check all proposed fields before any settings write. Missing, orphaned
      // or duplicate controls leave both the document and journal untouched.
      const checked=reconcile(validate(journal.project),await scan());
      if(hasFrontMatter(journal.project)&&(!validFrontJournal(journal,journal.project)||journal.kind==='course'&&!same(checked,journal.project)))throw problem('Incomplete initial front-matter text cannot be recovered as a manual edit.');
      await profile(checked,journal.kind==='course');
      if(journal.schema===undefined){
        const item=journal.kind==='module'?journal.project.modules.at(-1):journal.kind==='activity'?journal.project.modules.at(-1)?.activities.at(-1):undefined;
        transition(base,journal.project,journal.kind,item);
      }
      const recovery=[2,3,5].includes(journal.schema)?journal:{...prepared(base,journal.project,'legacy-recovery'),legacy:clone(journal)};
      if([2,3,5].includes(recovery.schema)&&!same(base,recovery.base))throw problem('The saved course does not match this unresolved operation. Preserve the document before recovery.');
      try{
        if(!durable||!same(durable,recovery))await prepare(recovery);
        return await finish(recovery);
      }catch(e){throw quarantine(e,recovery);}
    },{recovery:true})
  };
}
