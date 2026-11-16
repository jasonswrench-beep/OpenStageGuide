// The course map: each objective with the activities that practice it and the assessments that check it. It is Guide's learning alignment, extended in two ways:
//  * the PLAN's map, from the design desk's notes, before the guide has the blocks: which activities and assessments the author has said are for each objective, the level
//    of the objective and of the assessment, and the lesson it is to be taught in;
//  * the GUIDE's map, from the guide itself (guide-alignment.mjs reads what is in it), with each objective's level and where it is, and its minutes.
// Either one is a table a syllabus or an accreditation report can use (Word table in the guide, or CSV). Nothing here refuses anything: a gap is said in words and the
// author decides.
import {checkObjective,levelById,levelRank} from './design-bloom.mjs';
import {alignment} from './guide-alignment.mjs';
import {readGuide,lessonMinutes} from './guide-model.mjs';

const text=t=>String(t??'').replace(/\s+/g,' ').trim();

/** The map of the plan. @param {ReturnType<typeof import('./design-model.mjs').cleanDesign>} design */
export function planMap(design){
  const lessonOf=new Map(),moduleOf=new Map();
  for(const m of design.storyboard)for(const l of m.lessons)for(const o of l.objectives){if(!lessonOf.has(o)){lessonOf.set(o,l);moduleOf.set(o,m);}}
  const rows=design.objectives.map(o=>{
    const read=checkObjective(o.text);
    const level=o.level||read.level||'';
    const activities=design.activities.filter(a=>a.objectives.includes(o.id));
    const assessments=design.assessments.filter(a=>a.objectives.includes(o.id));
    const issues=[];
    if(!text(o.text))issues.push({rule:'MAP-EMPTY',message:'This objective has no words yet.'});
    else if(!level)issues.push({rule:'MAP-NO-LEVEL',message:'No level: choose one, or start the objective with an action verb.'});
    if(read.weak)issues.push({rule:'MAP-WEAK-VERB',message:`"${read.verb}" cannot be seen or measured, so an assessment cannot check it.`});
    if(!activities.length)issues.push({rule:'MAP-NO-ACTIVITY',message:'No activity gives practice at this.'});
    if(!assessments.length)issues.push({rule:'MAP-NO-ASSESSMENT',message:'No assessment checks this.'});
    const below=assessments.filter(a=>a.level&&level&&levelRank(a.level)<levelRank(level));
    for(const a of below)issues.push({rule:'MAP-LEVEL-LOW',message:`"${a.title||'An assessment'}" checks at ${levelById(a.level).label}, below this objective's ${levelById(level).label}.`});
    if(!lessonOf.has(o.id)&&design.storyboard.length)issues.push({rule:'MAP-NO-LESSON',message:'It is not in any lesson of the storyboard.'});
    return {id:o.id,text:o.text,level,levelLabel:level?levelById(level).label:'',module:moduleOf.get(o.id)?.title||'',lesson:lessonOf.get(o.id)?.title||'',
      activities:activities.map(a=>({id:a.id,title:a.title,minutes:a.minutes,model:a.model,event:a.event})),
      assessments:assessments.map(a=>({id:a.id,title:a.title,kind:a.kind,level:a.level,slots:a.slots})),issues};
  });
  const loose={
    activities:design.activities.filter(a=>!a.objectives.length).map(a=>({id:a.id,title:a.title,message:'This activity is not for any objective.'})),
    assessments:design.assessments.filter(a=>!a.objectives.length).map(a=>({id:a.id,title:a.title,message:'This assessment checks no objective.'})),
  };
  return {kind:'plan',rows,loose,summary:summarize(rows)};
}

/** The map of what is in the guide. @param {{text:string,style?:string}[]} paras */
export function guideMap(paras){
  const a=alignment(paras),guide=readGuide(paras);
  const where=new Map();
  for(const m of guide.modules)for(const l of m.lessons)where.set(l.text,{module:m.text||m.title,minutes:lessonMinutes(l)});
  const rows=a.objectives.map((o,i)=>{
    const read=checkObjective(o.text),level=read.level||'';
    const issues=[];
    if(read.weak)issues.push({rule:'MAP-WEAK-VERB',message:`"${read.verb}" cannot be seen or measured, so an assessment cannot check it.`});
    else if(!level)issues.push({rule:'MAP-NO-LEVEL',message:'No level: start the objective with an action verb.'});
    if(!o.activities.length)issues.push({rule:'MAP-NO-ACTIVITY',message:'No activity gives practice at this.'});
    if(!o.items.length)issues.push({rule:'MAP-NO-ASSESSMENT',message:'No assessment item checks this.'});
    return {id:`g${i+1}`,index:o.index,text:o.text,level,levelLabel:level?levelById(level).label:'',module:o.module||'',lesson:o.lesson||'',minutes:where.get(o.lesson)?.minutes??0,
      activities:o.activities.map(x=>({title:x.title||'An activity',similar:x.similar,index:x.index})),
      assessments:o.items.map(x=>({title:x.label,similar:x.similar,index:x.index})),issues};
  });
  return {kind:'guide',rows,loose:{activities:a.loose.filter(x=>x.kind==='activity').map(x=>({title:x.text,message:x.why,index:x.index})),assessments:a.loose.filter(x=>x.kind==='item').map(x=>({title:x.text,message:x.why,index:x.index}))},summary:summarize(rows)};
}

function summarize(rows){
  return {objectives:rows.length,withActivity:rows.filter(r=>r.activities.length).length,withAssessment:rows.filter(r=>r.assessments.length).length,
    withBoth:rows.filter(r=>r.activities.length&&r.assessments.length).length,withIssues:rows.filter(r=>r.issues.length).length};
}
export function mapSummaryText(map){
  const s=map.summary,n=(c,w)=>`${c} ${w}${c===1?'':'s'}`;
  if(!s.objectives)return map.kind==='plan'?'There are no objectives yet. Write them on the Needs and objectives screen.':'The guide has no Objective blocks yet. Add one to a lesson, then write the same words in an Activity\'s Objective part and in an assessment item\'s Learning objective.';
  return `${n(s.objectives,'objective')}: ${s.withActivity} with an activity, ${s.withAssessment} with an assessment, ${s.withBoth} with both.`;
}

// ---- as a table, for a syllabus or an accreditation report

export const MAP_COLUMNS=Object.freeze(['Module','Lesson','Objective','Level','Activities','Assessments']);
const list=items=>items.map(x=>x.title||'(untitled)').join('; ');
/** The map as rows of text: a heading row and a row for each objective. */
export function mapTable(map){
  return {head:[...MAP_COLUMNS],body:map.rows.map(r=>[r.module,r.lesson,text(r.text),r.levelLabel,list(r.activities),list(r.assessments)])};
}
const cell=v=>{const t=String(v??'');return /[",\r\n]/.test(t)?`"${t.replace(/"/g,'""')}"`:t;};
/** The map as a CSV file's text (UTF-8, with a byte-order mark so a spreadsheet opens it as UTF-8, and line ends a spreadsheet expects). */
export function mapCsv(map){
  const t=mapTable(map);
  return `﻿${[t.head,...t.body].map(r=>r.map(cell).join(',')).join('\r\n')}\r\n`;
}
