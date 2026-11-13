import {validate} from './model.mjs';
import {normalizeMaterials,materialKey} from './materials.mjs';
import {unsupportedObjective,objectiveText} from './objective-text.mjs';
export function planMaterials(project){
  validate(project);const grouped=new Map();let configured=0,total=0;
  for(const m of project.modules)for(const a of m.activities){total++;const items=normalizeMaterials(a.materials===undefined?[]:a.materials);if(items.length)configured++;
    for(const item of items){const key=materialKey(item);if(!grouped.has(key))grouped.set(key,{item:item.item,unit:item.unit,use:item.use,scaled:0,notes:[],activities:[]});const row=grouped.get(key),n=Math.round(item.quantity*1000);row.scaled=item.use==='consumable'?row.scaled+n:Math.max(row.scaled,n);if(!Number.isSafeInteger(row.scaled))throw new Error('The materials total is too large.');if(item.notes&&!row.notes.includes(item.notes))row.notes.push(item.notes);row.activities.push({id:a.id,title:a.title,moduleId:m.id,moduleTitle:m.title});}
  }
  return {projectId:project.id,revision:project.revision,title:project.title,configuredActivities:configured,totalActivities:total,rows:[...grouped.values()].map(({scaled,...row})=>({...row,quantity:scaled/1000})),assumption:'Full course in sequential planned order. Consumable or retained supplies add. Reusable means a shared pool available again for the next activity; use the largest single-activity requirement. Quantities are explicit activity totals, not per-learner counts. Unit case stays distinct; no conversion, parallel-work allowance or stock subtraction.'};
}

export function reviewCourse(project){
  validate(project);const issues=[];
  const add=(severity,type,id,message)=>issues.push({severity,type,id,message});
  if(!project.modules.length)add('warning','course',project.id,'Add at least one module.');
  for(const m of project.modules){
    if(!m.title?.trim())add('error','module',m.id,'The module title is blank.');
    if(unsupportedObjective(m))add('error','module',m.id,'The learning objective has an unsupported value. Preserve this guide and repair this field before creating objective-based outputs.');
    else if(!objectiveText(m).trim())add('warning','module',m.id,'Add a learning objective.');
    if(!m.activities.length)add('warning','module',m.id,'This module has no activities.');
    for(const a of m.activities){
      if(!a.title?.trim())add('error','activity',a.id,'The activity title is blank.');
      if(!a.participant?.trim())add('error','activity',a.id,'Add participant instructions before generating a workbook.');
      if(!a.facilitator?.trim())add('warning','activity',a.id,'Consider adding facilitator instructions.');
      if(!Number.isFinite(a.minutes)||a.minutes<0||a.minutes>1440)add('error','activity',a.id,'The duration must be a number between 0 and 1440 minutes.');
      else if(a.minutes===0)add('warning','activity',a.id,'No teaching time is allocated to this activity.');
    }
    for(const [type,id,fields] of [['module',m.id,[['Module title',m.title],['Learning objective',m.objective]]],...m.activities.map(a=>['activity',a.id,[['Activity title',a.title],['Facilitator instructions',a.facilitator],['Participant instructions',a.participant],['Debrief notes',a.answers]]])]){
      for(const [label,text] of fields)if(typeof text==='string'&&/\b(?:TODO|TBD)\b|\[(?:insert|add|replace)\b[^\]]*\]/i.test(text))add('warning',type,id,`${label} contains an unfinished prompt.`);
    }
  }
  return issues;
}

export function clockLabel(seconds){
  const day=Math.floor(seconds/86400)+1,within=seconds%86400;
  const h=String(Math.floor(within/3600)).padStart(2,'0'),m=String(Math.floor(within%3600/60)).padStart(2,'0'),s=within%60;
  return `${day>1?`Day ${day} · `:''}${h}:${m}${s?':'+String(s).padStart(2,'0'):''}`;
}
export function buildAgenda(project,start='09:00'){
  validate(project);
  if(!/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(start))throw new Error('Choose a valid start time.');
  const [h,m]=start.split(':').map(Number);let cursor=(h*60+m)*60;
  const beginning=cursor,rows=[];
  for(const module of project.modules)for(const a of module.activities){
    if(!Number.isFinite(a.minutes)||a.minutes<0||a.minutes>1440)throw new Error('Fix invalid activity durations before building the agenda.');
    const end=cursor+Math.round(a.minutes*60);
    // Agenda data is an allowlist, never a copy of facilitator/answer records.
    rows.push({moduleId:module.id,moduleTitle:module.title,activityId:a.id,title:a.title,start:clockLabel(cursor),end:clockLabel(end),minutes:a.minutes});
    cursor=end;
  }
  return {title:project.title,revision:project.revision,start:clockLabel(beginning),end:clockLabel(cursor),minutes:(cursor-beginning)/60,rows};
}
