import {legacyDocxBlob} from './legacy-docx-package.mjs';
import {planSchedule} from './delivery-plans.mjs';
import {agendaPackage} from './agenda.mjs';
import {publicOutputText as text} from './output-text.mjs';

export function planAgendaPreview(project,id,start='09:00'){
  const a=planSchedule(project,id,start);
  if(text(a.planId)!==a.planId)throw new Error('The saved plan ID cannot be shown exactly in an agenda. Preserve the master and repair the plan record explicitly.');
  const scope=a.scope.map(g=>({moduleId:text(g.moduleId),activityIds:g.activityIds.map(text)}));
  const status=a.targetMinutes===null?'No target duration set.':a.varianceMinutes>=0?`${a.varianceMinutes} minutes reserved below the target.`:`${-a.varianceMinutes} minutes over the target.`;
  const clean={...a,title:text(a.title),projectId:text(a.projectId),planId:text(a.planId),planName:text(a.planName),rows:a.rows.map(r=>({...r,moduleId:text(r.moduleId),moduleTitle:text(r.moduleTitle),activityId:text(r.activityId),title:text(r.title)})),scope,coverageNotice:text(a.coverageNotice),targetStatus:status};
  clean.subtitle=`${clean.planName} · delivery-plan agenda`;
  clean.scopeLabel='Selected delivery plan · exact reviewed activity scope';
  clean.identityLines=[`Delivery plan: ${clean.planName} · Plan ID: ${clean.planId}`,`Exact selected scope: ${scope.map(g=>`${g.moduleId}: ${g.activityIds.join(', ')}`).join(' | ')}`,status,clean.coverageNotice];
  return clean;
}
export function planAgendaFiles(project,id,start='09:00',stamp=new Date().toISOString()){
  return agendaPackage(planAgendaPreview(project,id,start),stamp);
}
export async function planAgendaBlob(project,id,start,stamp,Zip){
  return legacyDocxBlob(planAgendaFiles(project,id,start,stamp),Zip);
}
export function planAgendaName(title,planName){
  const safe=s=>text(s).replace(/[^\p{L}\p{N} _-]/gu,'').slice(0,55)||'Plan';
  return `${safe(title)} - ${safe(planName)} - Agenda - ${globalThis.crypto.randomUUID()}.docx`;
}
