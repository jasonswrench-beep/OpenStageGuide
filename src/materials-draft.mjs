import {validate,activityMaterials} from './model.mjs';
import {normalizeMaterials} from './materials.mjs';
export const MAX_MATERIAL_DRAFT_BYTES=2*1024*1024;
const bad=()=>{throw new Error('Unsupported materials draft. Keep the original file and choose a supported private materials record.');};
function keys(o,names){if(!o||typeof o!=='object'||Array.isArray(o)||Object.keys(o).length!==names.length||names.some(k=>!Object.hasOwn(o,k)))bad();}
function rawEntry(entry){
  keys(entry,['item','quantity','unit','use','notes']);const result={};
  for(const [key,max] of [['item',200],['quantity',128],['unit',80],['use',20],['notes',2000]]){const s=entry[key];if(typeof s!=='string'||s.length>max||key!=='notes'&&/[\r\n]/.test(s))bad();result[key]=s;}
  if(!['','consumable','reusable'].includes(result.use))bad();return result;
}
export function serializeMaterialsDraft(project,activityId,items,unsavedEntry){
  validate(project);activityMaterials(project,activityId);
  const record={format:'openstage-materials-draft',version:1,projectId:project.id,revision:project.revision,activityId,items:normalizeMaterials(items),unsavedEntry:rawEntry(unsavedEntry)};
  const text=JSON.stringify(record);if(new TextEncoder().encode(text).length>MAX_MATERIAL_DRAFT_BYTES)bad();readMaterialsDraft(text);return text;
}
export function readMaterialsDraft(text){
  if(typeof text!=='string'||new TextEncoder().encode(text).length>MAX_MATERIAL_DRAFT_BYTES)bad();let r;try{r=JSON.parse(text);}catch{bad();}
  keys(r,['format','version','projectId','revision','activityId','items','unsavedEntry']);
  if(r.format!=='openstage-materials-draft'||r.version!==1||!Number.isSafeInteger(r.revision)||r.revision<0)bad();
  for(const id of [r.projectId,r.activityId])if(typeof id!=='string'||!id||id.length>256)bad();
  return {format:r.format,version:1,projectId:r.projectId,revision:r.revision,activityId:r.activityId,items:normalizeMaterials(r.items),unsavedEntry:rawEntry(r.unsavedEntry)};
}
export function restoreMaterialsDraft(text,project,selectedActivity){
  validate(project);const record=readMaterialsDraft(text);
  if(record.projectId!==project.id||record.activityId!==selectedActivity)throw new Error('Choose the matching course and destination activity for this materials draft.');
  const saved=activityMaterials(project,selectedActivity);
  return {draft:{id:selectedActivity,project:structuredClone(project),items:record.items,restored:{recordedRevision:record.revision,savedItems:saved}},unsavedEntry:record.unsavedEntry};
}
