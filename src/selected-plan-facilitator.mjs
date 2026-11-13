import {legacyDocxBlob} from './legacy-docx-package.mjs';
import {resolveSelectedPlanScope,selectedScopeLines,SELECTED_WORKBOOK_OBJECTIVES} from './selected-plan-scope.mjs';
import {FACILITATOR_PRIVATE,facilitatorOptions,facilitatorPackage} from './facilitator-output.mjs';
import {normalizeMaterials} from './materials.mjs';
import {ICONS} from './model.mjs';
import {publicOutputText} from './output-text.mjs';

export const SELECTED_FACILITATOR_SCOPE='Selected delivery plan · exact reviewed activity scope';
export const SELECTED_FACILITATOR_LIMITS='Structured text only. Course details and front matter, unmanaged Word additions, images, formatting/layout, headers/footers, attachments and full-master fidelity are omitted. Keep the original master and full backup. Edits to this copy do not update the master. This selection does not approve objective coverage, audience, prerequisites or assessment requirements. Coordinated packs and existing-output regeneration are unavailable.';
const stamp=value=>{
 if(typeof value!=='string'||!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value)||!Number.isFinite(Date.parse(value))||new Date(value).toISOString()!==value)throw new Error('Use a valid selected private-copy generation stamp.');
 return value;
};
const publish=(value,label,required=false)=>{
 if(value===undefined&&!required)return '';
 if(typeof value!=='string'||value.length>50000)throw new Error(`Enter supported selected ${label} text before creating this private copy.`);
 let result;try{result=publicOutputText(value);}catch{throw new Error(`The selected ${label} contains an unsupported character. Preserve your work and correct that field before creating this private copy.`);}
 if(required&&!result.trim())throw new Error(`Enter selected ${label} text before creating this private copy.`);
 return result;
};
const unfinished=(value,label,issues)=>{if(/\b(?:TODO|TBD)\b|\[(?:insert|add|replace)\b[^\]]*\]/i.test(value))issues.push(`${label}: contains an unfinished prompt.`);};

export function planFacilitatorPreview(project,id,options,generation,start='09:00'){
 const resolved=resolveSelectedPlanScope(project,id,start),choice=facilitatorOptions(options),when=stamp(generation),issues=[];
 const modules=new Map(project.modules.map(module=>[module.id,module]));
 const v={title:publish(resolved.title,'course title',true),projectId:resolved.projectId,revision:resolved.revision,stamp:when,options:choice,
  planId:resolved.planId,planName:resolved.planName,selectedScope:resolved.scope,minutes:resolved.minutes,scope:SELECTED_FACILITATOR_SCOPE,
  privacy:FACILITATOR_PRIVATE,limits:SELECTED_FACILITATOR_LIMITS,
  identityLines:[...selectedScopeLines(resolved),SELECTED_WORKBOOK_OBJECTIVES,'Course details and front matter are omitted from this selected copy.'],
  details:[],frontMatter:[],modules:[],issues};
 unfinished(v.title,'Course title',issues);
 v.modules=resolved.groups.map((group,mi)=>{
  const source=modules.get(group.moduleId);if(!source)throw new Error('A selected module is missing. Preserve and repair this plan.');
  const byActivity=new Map(source.activities.map(activity=>[activity.id,activity]));
  const module={title:publish(source.title,'module title',true),objective:publish(source.objective,'learning objective'),activities:[]};
  unfinished(module.title,`Module ${mi+1}`,issues);
  if(!module.objective.trim())issues.push(`Module ${mi+1}: inherited learning objective is blank.`);else unfinished(module.objective,`Module ${mi+1} inherited objective`,issues);
  module.activities=group.activityIds.map((activityId,ai)=>{
   const sourceActivity=byActivity.get(activityId),label=`${mi+1}.${ai+1}`;
   if(!sourceActivity)throw new Error('A selected activity is missing or belongs to another module. Preserve and repair this plan.');
   if(!ICONS.some(([name])=>name===sourceActivity.icon))throw new Error('Choose a supported selected instructional cue label before creating this private copy.');
   if(!Number.isFinite(sourceActivity.minutes)||sourceActivity.minutes<0||sourceActivity.minutes>1440)throw new Error('Fix selected activity duration before creating this private copy.');
   const activity={title:publish(sourceActivity.title,'activity title',true),cue:sourceActivity.icon,minutes:sourceActivity.minutes,
    facilitator:publish(sourceActivity.facilitator,'facilitator instructions'),
    materials:normalizeMaterials(sourceActivity.materials===undefined?[]:sourceActivity.materials).map(item=>({item:publish(item.item,'material name',true),quantity:item.quantity,unit:publish(item.unit,'material unit',true),use:item.use,notes:publish(item.notes,'preparation notes')}))};
   unfinished(activity.title,`Activity ${label} title`,issues);
   if(!activity.facilitator.trim())issues.push(`Activity ${label}: facilitator instructions are blank.`);else unfinished(activity.facilitator,`Activity ${label} facilitator instructions`,issues);
   if(!activity.minutes)issues.push(`Activity ${label}: no teaching time is allocated.`);
   if(choice.answers){activity.answers=publish(sourceActivity.answers,'answers/debrief');if(activity.answers.trim())unfinished(activity.answers,`Activity ${label} answers/debrief`,issues);}
   if(choice.participant){if(!Number.isInteger(sourceActivity.lines)||sourceActivity.lines<0||sourceActivity.lines>20)throw new Error('Fix selected participant writing-line settings before creating this private copy.');activity.participant=publish(sourceActivity.participant,'participant instructions');activity.lines=sourceActivity.lines;if(activity.participant.trim())unfinished(activity.participant,`Activity ${label} participant instructions`,issues);else issues.push(`Activity ${label}: selected participant instructions are blank.`);}
   for(const item of activity.materials){unfinished(item.item,`Activity ${label} material`,issues);unfinished(item.notes,`Activity ${label} preparation notes`,issues);}
   return activity;
  });return module;
 });
 return v;
}
export const planFacilitatorFiles=(project,id,options,generation,start='09:00')=>facilitatorPackage(planFacilitatorPreview(project,id,options,generation,start));
export async function planFacilitatorBlob(project,id,options,generation,Zip,start='09:00'){
 return legacyDocxBlob(planFacilitatorFiles(project,id,options,generation,start),Zip);
}
export const planFacilitatorName=planName=>`OpenStage Private Facilitator Guide - ${publicOutputText(planName).replace(/[^\p{L}\p{N} _-]/gu,'').slice(0,55)||'Plan'} - ${globalThis.crypto.randomUUID()}.docx`;
