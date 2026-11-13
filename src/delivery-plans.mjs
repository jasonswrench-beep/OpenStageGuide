import {validate,value} from './model.mjs';
import {clockLabel} from './planning.mjs';
import {objectiveCoverage} from './objective-coverage.mjs';

const LIMIT=1000,MAX_GROUPS=200,MAX_SELECTED=2000;
const PLAN_ID=/^[A-Za-z0-9][A-Za-z0-9._:-]{0,199}$/;
const own=(o,k)=>Object.prototype.hasOwnProperty.call(o,k);
const object=o=>!!o&&typeof o==='object'&&!Array.isArray(o);
const fail=message=>{throw new Error(message);};
const safeText=(text,label,max=200)=>{
  if(typeof text!=='string')fail(`Enter ${label} as text.`);
  const v=value(text,label,true);if(v.length>max)fail(`${label} is too long.`);return v;
};
const target=value=>{
  if(value===null||value===undefined)return null;
  if(typeof value!=='number'||!Number.isFinite(value)||value<1||value>10080)fail('Enter a target from 1 to 10080 minutes, up to three decimal places.');
  const scaled=value*1000,rounded=Math.round(scaled);
  if(Math.abs(scaled-rounded)>1e-7)fail('Enter a target from 1 to 10080 minutes, up to three decimal places.');
  return rounded/1000;
};
function root(project){
  validate(project);
  const data=project.deliveryPlans;
  if(data===undefined)return {supported:true,present:false,data:{version:1,projectId:project.id,plans:[]}};
  if(!object(data)||data.version!==1||data.projectId!==project.id||!Array.isArray(data.plans)||data.plans.length>LIMIT)return {supported:false,present:true,reason:'Saved delivery plans use an unsupported version, foreign project, or oversized collection. Preserve the master and use the matching add-in version.'};
  return {supported:true,present:true,data};
}
function check(project,record,duplicates){
  const reasons=[];
  if(!object(record))return ['This plan entry is malformed.'];
  if(typeof record.id!=='string'||!PLAN_ID.test(record.id))reasons.push('The plan ID is missing or contains unsupported characters. Preserve this saved entry; its ID cannot be changed silently.');
  else if(duplicates.has(record.id))reasons.push('The plan ID is duplicated and ambiguous.');
  if(typeof record.name!=='string'||!record.name.trim()||record.name.length>200)reasons.push('The plan name is missing or malformed.');
  try{target(record.targetMinutes);}catch(e){reasons.push(e.message);}
  if(!Array.isArray(record.groups)||record.groups.length>MAX_GROUPS)return [...reasons,'The ordered module groups are malformed or too many.'];
  const modules=new Map(project.modules.map(m=>[m.id,m])),seenModules=new Set(),seenActivities=new Set();let selected=0;
  for(const group of record.groups){
    if(!object(group)||typeof group.moduleId!=='string'||!Array.isArray(group.activityIds)){reasons.push('A module group is malformed.');continue;}
    if(seenModules.has(group.moduleId))reasons.push('A module group is repeated.');seenModules.add(group.moduleId);
    const module=modules.get(group.moduleId);if(!module)reasons.push('A selected module is missing.');
    selected+=group.activityIds.length;if(selected>MAX_SELECTED){reasons.push('The plan has too many selected activities.');continue;}
    const owned=new Set(module?.activities.map(a=>a.id));
    for(const id of group.activityIds){
      if(typeof id!=='string'||!id||id.length>200){reasons.push('A selected activity ID is malformed.');continue;}
      if(seenActivities.has(id))reasons.push('A selected activity is repeated.');seenActivities.add(id);
      if(!owned.has(id))reasons.push('A selected activity is missing or belongs to another module.');
    }
  }
  return [...new Set(reasons)];
}
export function deliveryPlans(project){
  const state=root(project);if(!state.supported)return state;
  const ids=new Map();for(const p of state.data.plans)if(object(p)&&typeof p.id==='string')ids.set(p.id,(ids.get(p.id)||0)+1);
  const duplicates=new Set([...ids].filter(([,count])=>count>1).map(([id])=>id));
  return {...state,plans:state.data.plans.map((p,index)=>({index,id:object(p)&&typeof p.id==='string'?p.id:'',name:object(p)&&typeof p.name==='string'?p.name:`Plan ${index+1}`,record:structuredClone(p),reasons:check(project,p,duplicates)}))};
}
function groupsInput(project,groups,existing){
  if(!Array.isArray(groups)||groups.length>MAX_GROUPS)fail('Choose no more than 200 ordered module groups.');
  if(existing){
    if(!Array.isArray(existing.groups)||existing.groups.length>MAX_GROUPS||existing.groups.some(g=>!object(g)||typeof g.moduleId!=='string'||!Array.isArray(g.activityIds)||g.activityIds.length>MAX_SELECTED)||existing.groups.reduce((n,g)=>n+g.activityIds.length,0)>MAX_SELECTED||new Set(existing.groups.map(g=>g.moduleId)).size!==existing.groups.length)fail('The saved plan has ambiguous or oversized module groups. Preserve it and create a separate plan instead.');
  }
  const old=new Map(Array.isArray(existing?.groups)?existing.groups.filter(object).map(g=>[g.moduleId,g]):[]);
  const result=groups.map(g=>{
    if(!object(g)||Object.keys(g).some(k=>!['moduleId','activityIds'].includes(k))||!Array.isArray(g.activityIds))fail('The selected module group is malformed.');
    return {...structuredClone(old.get(g.moduleId)||{}),moduleId:g.moduleId,activityIds:[...g.activityIds]};
  });
  const reasons=check(project,{id:'temporary',name:'temporary',groups:result},new Set());
  if(reasons.length)fail(reasons.join(' '));return result;
}
export function editDeliveryPlan(project,change){
  const state=deliveryPlans(project);if(!state.supported)fail(state.reason);
  if(!object(change)||!['create','update'].includes(change.kind)||Object.keys(change).some(k=>!['kind','id','name','targetMinutes','groups'].includes(k)))fail('The delivery-plan edit is unsupported.');
  const list=state.data.plans;
  if(list.length>=LIMIT&&change.kind==='create')fail('The delivery-plan collection is full.');
  let current=null,index=-1,id;
  if(change.kind==='update'){
    if(typeof change.id!=='string'||!change.id)fail('Load a saved plan before editing.');
    const matches=state.plans.filter(p=>p.id===change.id);
    if(matches.length!==1)fail('The selected plan is missing or its ID is ambiguous.');
    ({index}=matches[0]);current=list[index];if(!object(current))fail('The selected plan is malformed.');id=change.id;
  }else{
    if(own(change,'id'))fail('New plan IDs are assigned automatically.');
    do{id=crypto.randomUUID();}while(state.plans.some(p=>p.id===id));
  }
  const name=safeText(change.name,'a plan name'),targetMinutes=target(change.targetMinutes);
  const groups=groupsInput(project,change.groups,current);
  const record={...structuredClone(current||{}),id,name,groups};
  if(targetMinutes===null)delete record.targetMinutes;else record.targetMinutes=targetMinutes;
  const proposed=structuredClone(state.data),next=structuredClone(project);
  if(index<0)proposed.plans.push(record);else proposed.plans[index]=record;
  next.deliveryPlans=proposed;
  const assessed=deliveryPlans(next),chosen=assessed.plans.filter(p=>p.id===id);
  if(!assessed.supported||chosen.length!==1||chosen[0].reasons.length)fail('The proposed plan is invalid or ambiguous.');
  if(JSON.stringify(state.data)===JSON.stringify(proposed))return {next:project,id,changed:false};
  next.revision++;return {next,id,changed:true};
}
export function selectedObjectiveLinks(project,scope){
  const annotations=objectiveCoverage(project);
  if(!annotations.supported)return {state:'unsupported',rows:[],summary:null,notice:'Selected-plan objective-link comparison unavailable: saved annotations are unsupported. Audience, prerequisite and plan-required-objective selection are unavailable.'};
  if(project.modules.some(m=>m.objective!==undefined&&typeof m.objective!=='string'))return {state:'unsupported-objective-field',rows:[],summary:null,notice:'Selected-plan objective-link comparison unavailable: one or more course objectives have an unsupported text field. Preserve the saved source; the agenda schedule remains available. Audience, prerequisite and plan-required-objective selection are unavailable.'};
  const selected=new Set(scope.flatMap(g=>g.activityIds));
  const rows=project.modules.map(module=>{
    const objective=module.objective||'',blank=!objective.trim();
    const roles=Object.fromEntries(['practice','assessment'].map(role=>{
      const links=annotations.data.links.filter(l=>l.objectiveModuleId===module.id&&l.roles.includes(role));
      const current=blank?[]:links.filter(l=>l.reviewedObjective===objective),stale=links.filter(l=>blank||l.reviewedObjective!==objective);
      const selectedCurrent=current.filter(l=>selected.has(l.activityId)),selectedStale=stale.filter(l=>selected.has(l.activityId));
      const note=annotations.data.exceptions.find(e=>e.objectiveModuleId===module.id&&e.role===role);
      return [role,{courseCurrent:current.length,selectedCurrent:selectedCurrent.length,courseStale:stale.length,selectedStale:selectedStale.length,lostBySelection:current.length>0&&!selectedCurrent.length,preexistingGap:!current.length,currentCourseGapNote:!!note&&!blank&&note.reviewedObjective===objective,noteConflict:!!note&&!blank&&note.reviewedObjective===objective&&!!current.length}];
    }));
    return {moduleId:module.id,moduleTitle:module.title,objective,blank,roles};
  });
  const nonblank=rows.filter(r=>!r.blank),summary={courseObjectives:rows.length,nonblankObjectives:nonblank.length,blankObjectives:rows.length-nonblank.length,selectedActivities:selected.size,
    lostPractice:nonblank.filter(r=>r.roles.practice.lostBySelection).length,lostAssessment:nonblank.filter(r=>r.roles.assessment.lostBySelection).length,
    preexistingPracticeGaps:nonblank.filter(r=>r.roles.practice.preexistingGap).length,preexistingAssessmentGaps:nonblank.filter(r=>r.roles.assessment.preexistingGap).length,
    staleSelectedLinks:annotations.data.links.filter(l=>selected.has(l.activityId)&&(!project.modules.find(m=>m.id===l.objectiveModuleId)?.objective?.trim()||project.modules.find(m=>m.id===l.objectiveModuleId)?.objective!==l.reviewedObjective)).length};
  const state=!rows.length?'no-objectives':!nonblank.length?'no-objective-text':!annotations.present?'no-annotations':'ready';
  const counts=`Objectives losing all current declared coverage in this selection: ${summary.lostPractice} for practice, ${summary.lostAssessment} for assessment. Pre-existing gaps: ${summary.preexistingPracticeGaps} practice, ${summary.preexistingAssessmentGaps} assessment. Blank objectives: ${summary.blankObjectives}. Selected stale links uncounted: ${summary.staleSelectedLinks}.`;
  const notice=state==='no-objectives'?'Selected-plan objective-link comparison unavailable: this course has no module objectives.':
    state==='no-objective-text'?`Selected-plan objective-link comparison unavailable: all ${summary.courseObjectives} course module objectives are blank.`:
    state==='no-annotations'?`No saved objective annotations exist. Across all ${summary.courseObjectives} course module objectives, ${summary.blankObjectives} are blank and no current practice or assessment links are verified.`:
    `Selected-plan objective-link comparison against all ${summary.courseObjectives} course module objectives: ${counts}`;
  return {state,rows,summary,notice:`${notice} Selected coverage counts only selected activities; the whole-course baseline counts all activities. Both require exact current objective wording. Full-course gap notes do not approve plan omissions. Audience, prerequisite and plan-required-objective selection are unavailable; this is an author-declared design map, not proof of learning or measurement validity.`};
}
export function planSchedule(project,id,start='09:00'){
  const state=deliveryPlans(project);if(!state.supported)fail(state.reason);
  const matches=state.plans.filter(p=>p.id===id);if(matches.length!==1)fail('Choose a unique saved delivery plan.');
  const plan=matches[0];if(plan.reasons.length)fail(`This plan needs repair: ${plan.reasons.join(' ')}`);
  if(!/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(start))fail('Choose a valid start time.');
  if(!plan.record.groups.some(g=>g.activityIds.length))fail('No activities selected. Choose at least one activity before downloading a plan agenda.');
  if(typeof project.title!=='string')fail('Enter a supported course title before creating a selected-plan output.');
  const title=value(project.title,'a course title',true),byModule=new Map(project.modules.map(m=>[m.id,m]));
  let cursor=(Number(start.slice(0,2))*60+Number(start.slice(3)))*60;const beginning=cursor,rows=[],scope=[];
  for(const group of plan.record.groups){const module=byModule.get(group.moduleId);if(typeof module.title!=='string'||!module.title.trim())fail('Give selected modules titles before creating a plan agenda.');
    const byActivity=new Map(module.activities.map(a=>[a.id,a]));const ids=[];
    for(const activityId of group.activityIds){const a=byActivity.get(activityId);if(typeof a.title!=='string'||!a.title.trim())fail('Give selected activities titles before creating a plan agenda.');if(!Number.isFinite(a.minutes)||a.minutes<0||a.minutes>1440)fail('Fix selected activity durations before creating a plan agenda.');
      const end=cursor+Math.round(a.minutes*60);rows.push({moduleId:module.id,moduleTitle:module.title,activityId:a.id,title:a.title,start:clockLabel(cursor),end:clockLabel(end),minutes:a.minutes});ids.push(a.id);cursor=end;}
    scope.push({moduleId:module.id,activityIds:ids});
  }
  const minutes=(cursor-beginning)/60,targetMinutes=target(plan.record.targetMinutes);
  // Targets are authored to thousandths of a minute; scheduled endpoints are
  // whole seconds. Calculate their difference with bounded integer millis,
  // then round only the displayed minute figure to avoid binary tails.
  const varianceMinutes=targetMinutes===null?null:Math.round((Math.round(targetMinutes*60000)-(cursor-beginning)*1000)/60000*1e6)/1e6;
  const report=selectedObjectiveLinks(project,scope),s=report.summary;
  // Public preview/export receives only known aggregate fields. Detailed
  // objective text, link records and full-course gap notes stay in the pane.
  const coverage={state:report.state,summary:s?{
    courseObjectives:s.courseObjectives,nonblankObjectives:s.nonblankObjectives,blankObjectives:s.blankObjectives,selectedActivities:s.selectedActivities,
    lostPractice:s.lostPractice,lostAssessment:s.lostAssessment,preexistingPracticeGaps:s.preexistingPracticeGaps,preexistingAssessmentGaps:s.preexistingAssessmentGaps,staleSelectedLinks:s.staleSelectedLinks
  }:null};
  return {title,projectId:project.id,revision:project.revision,planId:plan.id,planName:plan.record.name,start:clockLabel(beginning),end:clockLabel(cursor),minutes,targetMinutes,varianceMinutes,scope,rows,coverage,coverageNotice:report.notice};
}
export function selectedPlanObjectiveReport(project,id,start='09:00'){
  return selectedObjectiveLinks(project,planSchedule(project,id,start).scope);
}
