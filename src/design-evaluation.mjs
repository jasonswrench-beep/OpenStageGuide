// The evaluation plan by Kirkpatrick's four levels, and the file PrismStatistics opens for it. Each measure the author plans (what is measured, of whom, when) becomes variables
// in a .pstat file with empty rows: reaction items for level 1, a score before and after for level 2, behavior ratings for level 3, a measure of the work before and after for
// level 4. When the answers are typed in or pasted, PrismStatistics has them ready, with the level in each variable's label and the plan in the notes the file carries.
// What the file cannot hold as data (the plan's wording) is in `source.report.notes`; there is no profile for it in the core yet (CORE-REQUESTS G1).
import {KIRKPATRICK} from './design-model.mjs';
import {datasetProject,MAX_ROWS} from './guide-pstat.mjs';

export const MEASURE_KINDS=Object.freeze([
  {id:'rating',label:'Statements rated on a scale',for:[1,3]},
  {id:'pre-post',label:'A score before and after',for:[2]},
  {id:'score',label:'One score',for:[2,3]},
  {id:'metric',label:'A measure of the work, before and after',for:[4]},
]);
export const kindsForLevel=level=>MEASURE_KINDS.filter(k=>k.for.includes(Number(level)));
const clean=t=>String(t??'').replace(/\s+/g,' ').trim();
const slug=t=>clean(t).normalize('NFKD').replace(/[̀-ͯ]/g,'').replace(/[^A-Za-z0-9]+/g,'_').replace(/^_+|_+$/g,'').toLowerCase().slice(0,24).replace(/_+$/g,'');
const plural=(n,w)=>`${n} ${w}${n===1?'':'s'}`;

/** The points of a scale as the author typed them ("1, 2, 3, 4, 5" or "Disagree, Neutral, Agree"): value labels when they are words, nothing when they are numbers. */
export function scalePoints(text){
  const points=String(text||'1, 2, 3, 4, 5').split(/[,;\n]/).map(clean).filter(Boolean);
  if(points.length<2||points.length>7)throw new Error('A scale has 2 to 7 points, separated by commas.');
  const nums=points.map(p=>/^-?\d+(\.\d+)?$/.test(p)?Number(p):null);
  if(nums.every(n=>n!==null))return {labels:[],low:Math.min(...nums),high:Math.max(...nums),points};
  return {labels:points.map((p,i)=>({value:i+1,label:p})),low:1,high:points.length,points};
}

/** What is wrong with the plan for a file, in words, or an empty list. The author can still save what is there. */
export function evaluationIssues(design){
  const out=[];
  const ms=design.evaluation.measures;
  if(!ms.length)out.push('There are no measures yet. Add one for each level you will evaluate.');
  for(const m of ms){
    const name=m.name||'A measure';
    if(!m.name)out.push(`A level ${m.level} measure has no name.`);
    if(m.kind==='rating'&&!m.statements.length)out.push(`"${name}" is statements rated on a scale, and has no statements.`);
    if(!m.when)out.push(`"${name}" does not say when it is measured.`);
  }
  for(const k of KIRKPATRICK)if(ms.length&&!ms.some(m=>m.level===k.id))out.push(`Nothing is planned for ${k.label.replace(/^Level \d: /,'level '+k.id+', ')}.`);
  return out;
}

/** The project.json of the file for the plan. @param {ReturnType<typeof import('./design-model.mjs').cleanDesign>} design */
export function evaluationProject(design,{guideTitle='',rows=0,appVersion='OpenStage Guide'}={}){
  const ms=design.evaluation.measures;
  if(!ms.length)throw new Error('The plan has no measures yet. Add a measure first.');
  const n=Number(rows);
  if(!Number.isInteger(n)||n<0||n>MAX_ROWS)throw new Error(`Give the number of people as a whole number from 0 to ${MAX_ROWS}.`);
  const objectives=new Map(design.objectives.map(o=>[o.id,o]));
  const variables=[],columns=[],names=new Set(['respondent']),notes=[];
  let id=1;
  const add=(wanted,label,measure,labels,type='numeric')=>{
    let name=wanted.slice(0,60),k=2;
    while(names.has(name.toLowerCase())){name=`${wanted.slice(0,56)}_${k++}`;}
    names.add(name.toLowerCase());
    variables.push({id:id++,name,label,measure,value_labels:labels,missing_values:[]});
    columns.push({type,values:Array.from({length:n},()=>null)});
    return name;
  };
  variables.push({id:id++,name:'respondent',label:'Respondent (a number for each person)',measure:'scale',value_labels:[],missing_values:[]});
  columns.push({type:'numeric',values:Array.from({length:n},(_,i)=>i+1)});
  notes.push(`The evaluation plan of the guide${guideTitle?` "${clean(guideTitle)}"`:''}, by Kirkpatrick's four levels. The answers are empty: type or paste them in, one row for each person (${n?plural(n,'row'):'no rows yet'}).`);
  for(const m of design.evaluation.measures.slice().sort((a,b)=>a.level-b.level)){
    const level=KIRKPATRICK.find(k=>k.id===m.level),tag=`L${m.level}`,base=`l${m.level}_${slug(m.name)||'measure'}`;
    const where=[m.when?`when: ${m.when}`:'',m.who?`of: ${m.who}`:''].filter(Boolean).join('; ');
    const goals=m.objectives.map(o=>objectives.get(o)).filter(Boolean).map(o=>clean(o.text)).filter(Boolean);
    let made=[];
    if(m.kind==='rating'){
      const sc=scalePoints(m.scale);
      made=m.statements.map((s,i)=>add(`${base}_q${String(i+1).padStart(2,'0')}`,`${tag}: ${s}`,'ordinal',sc.labels));
      notes.push(`${level.label}, "${m.name||'Measure'}": ${plural(m.statements.length,'statement')} rated ${sc.points.join(', ')}${made.length?` (${made[0]}${made.length>1?` to ${made.at(-1)}`:''})`:''}${where?`; ${where}`:''}.`);
    }else if(m.kind==='pre-post'){
      made=[add(`${base}_pre`,`${tag}: ${m.name||'Score'}, before`,'scale',[]),add(`${base}_post`,`${tag}: ${m.name||'Score'}, after`,'scale',[])];
      notes.push(`${level.label}, "${m.name||'Measure'}": a score before (${made[0]}) and after (${made[1]}); the gain is the second minus the first, which PrismStatistics works out${where?`; ${where}`:''}.`);
    }else if(m.kind==='metric'){
      made=[add(`${base}_before`,`${tag}: ${m.name||'Measure'}, before`,'scale',[]),add(`${base}_after`,`${tag}: ${m.name||'Measure'}, after`,'scale',[])];
      notes.push(`${level.label}, "${m.name||'Measure'}": a measure of the work before (${made[0]}) and after (${made[1]})${where?`; ${where}`:''}.`);
    }else{
      made=[add(base,`${tag}: ${m.name||'Score'}`,'scale',[])];
      notes.push(`${level.label}, "${m.name||'Measure'}": one score (${made[0]})${where?`; ${where}`:''}.`);
    }
    if(goals.length)notes.push(`"${m.name||'Measure'}" is for: ${goals.join(' / ')}`);
    if(clean(m.note))notes.push(`"${m.name||'Measure'}": ${clean(m.note)}`);
  }
  if(clean(design.evaluation.notes))notes.push(clean(design.evaluation.notes));
  return datasetProject({datasetName:clean(guideTitle)||'Evaluation',fileName:`${clean(guideTitle)||'Guide'} evaluation plan (OpenStage Guide)`,appVersion,rows:n,variables,columns,notes});
}

// ---- the plan as a table, for the guide

export const EVALUATION_COLUMNS=Object.freeze(['Level','Measure','What is measured','When','Of whom']);
export function evaluationTable(design){
  const kindLabel=k=>MEASURE_KINDS.find(x=>x.id===k)?.label||clean(k);
  return {head:[...EVALUATION_COLUMNS],body:design.evaluation.measures.slice().sort((a,b)=>a.level-b.level).map(m=>[
    KIRKPATRICK.find(k=>k.id===m.level).label,clean(m.name),kindLabel(m.kind),clean(m.when),clean(m.who)])};
}
