import {planSchedule} from './delivery-plans.mjs';
import {publicOutputText as text} from './output-text.mjs';

export const SELECTED_PLAN_LIMITS='This saved plan selects activities and their saved order. Required objectives, audience, prerequisites and assessment policy are unavailable; selection does not prove learning or measurement validity.';
export const SELECTED_WORKBOOK_OBJECTIVES='Inherited objectives from selected modules; selecting activities does not confirm objective coverage or approve plan requirements.';
const exact=(value,label)=>{
 if(typeof value!=='string'||text(value)!==value)throw new Error(`The selected ${label} cannot be shown exactly. Preserve the master and repair the saved identifier explicitly.`);
 return value;
};
const freeze=value=>{if(value&&typeof value==='object'){for(const child of Object.values(value))freeze(child);Object.freeze(value);}return value;};

export function resolveSelectedPlanScope(project,id,start='09:00'){
 const schedule=planSchedule(project,id,start);
 const planName=text(schedule.planName);
 if(!planName.trim())throw new Error('The selected plan name is empty after public text processing. Preserve the master and repair it explicitly.');
 const scope=schedule.scope.map(group=>({moduleId:exact(group.moduleId,'module ID'),activityIds:group.activityIds.map(item=>exact(item,'activity ID'))}));
 // Content is joined only by the output that publishes it. This shared object
 // contains no learner, objective, material, private or opaque source payload.
 const groups=scope.filter(group=>group.activityIds.length).map(group=>({moduleId:group.moduleId,activityIds:[...group.activityIds]}));
 const resolved={projectId:exact(schedule.projectId,'project ID'),revision:schedule.revision,planId:exact(schedule.planId,'plan ID'),planName,title:text(schedule.title),start:schedule.start,end:schedule.end,minutes:schedule.minutes,scope,groups};
 return freeze(resolved);
}

export function selectedScopeLines(resolved,{workbook=false}={}){
 return [`Delivery plan: ${resolved.planName} · Plan ID: ${resolved.planId}`,`Exact selected scope: ${resolved.scope.map(g=>`${g.moduleId}: ${g.activityIds.join(', ')}`).join(' | ')}`,`Selected planned time: ${resolved.minutes} minutes in saved plan order; absolute start times appear only in the agenda.`,...(workbook?[SELECTED_WORKBOOK_OBJECTIVES]:[]),SELECTED_PLAN_LIMITS];
}
