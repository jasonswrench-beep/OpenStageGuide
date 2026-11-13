import {legacyDocxBlob} from './legacy-docx-package.mjs';
import {resolveSelectedPlanScope,selectedScopeLines,SELECTED_PLAN_LIMITS,SELECTED_WORKBOOK_OBJECTIVES} from './selected-plan-scope.mjs';
import {workbookPackage} from './worksheet.mjs';
import {materialsPackage,MATERIALS_WARNING} from './materials-output.mjs';
import {normalizeMaterials,materialKey} from './materials.mjs';
import {publicOutputText as text} from './output-text.mjs';

export const SELECTED_OUTPUT_SCOPE='Selected delivery plan · exact reviewed activity scope';
const stamp=value=>{
 if(typeof value!=='string'||!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value)||!Number.isFinite(Date.parse(value))||new Date(value).toISOString()!==value)throw new Error('Use a valid selected-output generation stamp.');
 return value;
};
const publish=(value,label,required=false)=>{
 if(typeof value!=='string')throw new Error(`The selected ${label} is unsupported. Preserve the master and repair that field before creating this output.`);
 let clean;try{clean=text(value);}catch{throw new Error(`The selected ${label} contains an unsupported character. Preserve the master and repair that field before creating this output.`);}
 if(required&&!clean.trim())throw new Error(`Enter selected ${label} text before creating this output.`);
 return clean;
};
const base=resolved=>({projectId:resolved.projectId,revision:resolved.revision,planId:resolved.planId,planName:resolved.planName,scope:resolved.scope,minutes:resolved.minutes,scopeLabel:SELECTED_OUTPUT_SCOPE,policy:SELECTED_PLAN_LIMITS});
function selectedGroups(project,resolved){
 const modules=new Map(project.modules.map(module=>[module.id,module]));
 return resolved.groups.map(group=>{
  const module=modules.get(group.moduleId);if(!module)throw new Error('A selected module is missing. Preserve and repair this plan.');
  const activities=new Map(module.activities.map(activity=>[activity.id,activity]));
  return {module,activities:group.activityIds.map(id=>{const activity=activities.get(id);if(!activity)throw new Error('A selected activity is missing or belongs to another module. Preserve and repair this plan.');return activity;})};
 });
}

export function planWorkbookPreview(project,id,generation,start='09:00'){
 const resolved=resolveSelectedPlanScope(project,id,start),when=stamp(generation);
 const modules=selectedGroups(project,resolved).map(({module,activities})=>({title:publish(module.title,'module title',true),objective:module.objective===undefined?'':publish(module.objective,'learning objective'),activities:activities.map(activity=>{
  const title=publish(activity.title,'activity title',true),participant=publish(activity.participant,'participant instructions',true);
  if(!Number.isInteger(activity.lines)||activity.lines<0||activity.lines>20)throw new Error('Fix the selected participant writing-line setting before creating a workbook.');
  return {title,participant,lines:activity.lines};
 })}));
 return {...base(resolved),title:publish(resolved.title,'course title',true),stamp:when,objectivePolicy:SELECTED_WORKBOOK_OBJECTIVES,identityLines:selectedScopeLines(resolved,{workbook:true}),modules};
}
export function planWorkbookFiles(project,id,generation,start='09:00'){
 const preview=planWorkbookPreview(project,id,generation,start);
 return workbookPackage(preview,preview.stamp,preview.scopeLabel,preview.identityLines);
}
export async function planWorkbookBlob(project,id,generation,Zip,start='09:00'){
 return legacyDocxBlob(planWorkbookFiles(project,id,generation,start),Zip);
}

export function planMaterialsPreview(project,id,generation,start='09:00'){
 const resolved=resolveSelectedPlanScope(project,id,start),when=stamp(generation),grouped=new Map();let configured=0,total=0;
 selectedGroups(project,resolved).forEach((group,gi)=>group.activities.forEach((activity,ai)=>{
  total++;const items=normalizeMaterials(activity.materials===undefined?[]:activity.materials);if(items.length)configured++;
  for(const item of items){const key=materialKey(item);if(!grouped.has(key))grouped.set(key,{item:publish(item.item,'material name',true),unit:publish(item.unit,'material unit',true),use:item.use,scaled:0,activities:[]});
   const row=grouped.get(key),scaled=Math.round(item.quantity*1000);row.scaled=item.use==='consumable'?row.scaled+scaled:Math.max(row.scaled,scaled);if(!Number.isSafeInteger(row.scaled))throw new Error('The selected materials total is too large.');row.activities.push(`${gi+1}.${ai+1} ${publish(activity.title,'activity title',true)}`);
  }
 }));
 const rows=[...grouped.values()].map(({scaled,...row})=>({...row,quantity:scaled/1000}));
 return {...base(resolved),stamp:when,identityLines:selectedScopeLines(resolved),configuredActivities:configured,totalActivities:total,rows,warning:MATERIALS_WARNING,assumption:'Selected activities in saved plan order. Consumable or retained requirements add. Reusable means a shared pool available again for the next selected activity; use the largest single-activity requirement. Quantities are recorded activity totals, not per-learner counts. Unit case stays distinct; no conversion, parallel-work allowance or stock subtraction. Missing entries do not prove supplies are unnecessary.'};
}
export function planMaterialsFiles(project,id,generation,start='09:00'){
 return materialsPackage(planMaterialsPreview(project,id,generation,start));
}
export async function planMaterialsBlob(project,id,generation,Zip,start='09:00'){
 return legacyDocxBlob(planMaterialsFiles(project,id,generation,start),Zip);
}
const safe=value=>text(value).replace(/[^\p{L}\p{N} _-]/gu,'').slice(0,55)||'Plan';
export const planWorkbookName=(title,planName)=>`${safe(title)} - ${safe(planName)} - Participant Workbook - ${globalThis.crypto.randomUUID()}.docx`;
export const planMaterialsName=planName=>`OpenStage - ${safe(planName)} - Materials Checklist - ${globalThis.crypto.randomUUID()}.docx`;
