// Pure prospective contract only. No Office adapter/store imports or writes.
// Metadata validation is not evidence about a picture in a live document.
import {validate} from './model.mjs';
import {CUE_ASSETS} from './cue-assets.mjs';
import {cueOwnerTag} from './cue-ownership.mjs';
const uuid=/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/;
const hash=/^[a-f0-9]{64}$/;
const scalar=(value,pattern)=>typeof value==='string'&&pattern.test(value);
const knownIcon=value=>typeof value==='string'&&Object.hasOwn(CUE_ASSETS,value);
const stable=x=>Array.isArray(x)?x.map(stable):x&&typeof x==='object'?Object.fromEntries(Object.keys(x).sort().map(k=>[k,stable(x[k])])):x;
const same=(a,b)=>JSON.stringify(stable(a))===JSON.stringify(stable(b));
const fail=()=>{throw Object.assign(new Error('The icon-only proposal is unsupported or inconsistent. Preserve this guide; no edit is authorized.'),{code:'OSG_ICON_PROPOSAL'});};
function keys(value,names){if(!value||typeof value!=='object'||Array.isArray(value)||!same(Object.keys(value).sort(),[...names].sort()))fail();}
function jsonData(value,seen=new Set()){
  if(value===null||typeof value==='string'||typeof value==='boolean')return;
  if(typeof value==='number'){if(!Number.isFinite(value))fail();return;}
  if(!value||typeof value!=='object'||seen.has(value))fail();seen.add(value);
  const own=Reflect.ownKeys(value),descriptors=Object.getOwnPropertyDescriptors(value);
  if(Array.isArray(value)){
    if(Object.getPrototypeOf(value)!==Array.prototype||own.length!==value.length+1)fail();
    for(let i=0;i<value.length;i++)if(!Object.hasOwn(descriptors,String(i)))fail();
  }else if(![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail();
  for(const key of own){
    if(Array.isArray(value)&&key==='length')continue;
    const property=descriptors[key];
    if(typeof key!=='string'||!property.enumerable||!Object.hasOwn(property,'value'))fail();
    jsonData(property.value,seen);
  }
  seen.delete(value);
}
function ownedProject(project){
  jsonData(project);validate(project);keys(project.cueOwnership,['version']);if(project.cueOwnership.version!==1)fail();
  if(!scalar(project.id,uuid))fail();const seen=new Set();
  for(const module of project.modules){if(!scalar(module.id,uuid))fail();for(const item of module.activities){
    if(!scalar(item.id,uuid)||typeof item.icon!=='string')fail();
    if(!Object.hasOwn(item,'cueImage'))continue;
    keys(item.cueImage,['version','nonce','activityId','assetHash']);const owner=item.cueImage;
    if(!scalar(owner.activityId,uuid)||owner.activityId!==item.id||!scalar(owner.nonce,/^[a-f0-9]{32}$/)||!scalar(owner.assetHash,hash)||!knownIcon(item.icon)||owner.assetHash!==CUE_ASSETS[item.icon].sha256)fail();
    const tag=cueOwnerTag(owner);if(seen.has(tag))fail();seen.add(tag);
  }}
}
export function validateIconProject(project){
  try{ownedProject(project);return structuredClone(project);}catch(error){if(error.code==='OSG_ICON_PROPOSAL')throw error;fail();}
}
export function cloneIconData(value){
  try{jsonData(value);return structuredClone(value);}catch(error){if(error.code==='OSG_ICON_PROPOSAL')throw error;fail();}
}
export function proposeIconTransition(project,id,icon){
  try{
    ownedProject(project);if(!scalar(id,uuid)||!knownIcon(icon))fail();
    const base=structuredClone(project),next=structuredClone(project),item=next.modules.flatMap(m=>m.activities).find(a=>a.id===id);
    if(!item?.cueImage)fail();if(item.icon===icon)return null;
    if(project.revision===Number.MAX_SAFE_INTEGER)fail();
    const before={icon:item.icon,assetHash:item.cueImage.assetHash};
    item.icon=icon;item.cueImage.assetHash=CUE_ASSETS[icon].sha256;next.revision++;ownedProject(next);
    return {base,project:next,icon:{id,ownerTag:cueOwnerTag(item.cueImage),before,after:{icon,assetHash:item.cueImage.assetHash}}};
  }catch(error){if(error.code==='OSG_ICON_PROPOSAL')throw error;fail();}
}
export function validateIconJournal(journal){
  try{
    jsonData(journal);
    keys(journal,['schema','kind','phase','txId','base','project','checkpoint','icon']);
    if(journal.schema!==4||journal.kind!=='activity-icon'||!['prepared','committed'].includes(journal.phase)||!scalar(journal.txId,uuid))fail();
    keys(journal.checkpoint,['id','masterHash','backupHash']);if(!scalar(journal.checkpoint.id,uuid)||!scalar(journal.checkpoint.masterHash,hash)||!scalar(journal.checkpoint.backupHash,hash))fail();
    keys(journal.icon,['id','ownerTag','before','after']);keys(journal.icon.before,['icon','assetHash']);keys(journal.icon.after,['icon','assetHash']);
    if(typeof journal.icon.ownerTag!=='string'||!knownIcon(journal.icon.before.icon)||!knownIcon(journal.icon.after.icon)||!scalar(journal.icon.before.assetHash,hash)||!scalar(journal.icon.after.assetHash,hash))fail();
    ownedProject(journal.project);const expected=proposeIconTransition(journal.base,journal.icon.id,journal.icon.after.icon);
    if(!expected||!same(expected.project,journal.project)||!same(expected.icon,journal.icon))fail();
    return structuredClone(journal);
  }catch(error){if(error.code==='OSG_ICON_PROPOSAL')throw error;fail();}
}
// Checkpoint descriptors here are untrusted data. Live binding and physical
// before/after ownership proof are separate mandatory controller gates.
export function createIconJournal(project,id,icon,checkpoint){
  jsonData(checkpoint);
  const proposal=proposeIconTransition(project,id,icon);if(!proposal)return null;
  return validateIconJournal({schema:4,kind:'activity-icon',phase:'prepared',txId:crypto.randomUUID(),...proposal,checkpoint:structuredClone(checkpoint)});
}
