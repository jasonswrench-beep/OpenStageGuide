// The design desk's notes, kept in the guide's own project file (the `design` member of a guide): the needs analysis, the objectives, the activities and
// assessments planned for them, the storyboard, and the evaluation plan. They are the author's words and choices; nothing here is worked out by a program
// that the author did not see, and nothing needs a model. The guide's pages are not changed by them: the desk puts them into the guide, a new guide or another
// tool's file only when the author asks (storyboard.mjs, map.mjs, evaluation.mjs).
import {BLOOM_LEVELS} from './design-bloom.mjs';

export const DESIGN_VERSION=1;
export const LIMITS=Object.freeze({objectives:80,activities:300,assessments:150,modules:30,lessons:200,measures:80,statements:200,models:40,text:4000,title:200,short:80,chars:400000});
export const NEED_CAUSES=Object.freeze([
  {id:'knowledge',label:'They do not know how, or what',note:'Training can help.',training:true},
  {id:'practice',label:'They know, and have had too little practice',note:'Training with practice can help.',training:true},
  {id:'tools',label:'The tools, process or information are in the way',note:'Training will not fix this by itself.',training:false},
  {id:'motivation',label:'There is little reason to do it, or it is not rewarded',note:'Training will not fix this by itself.',training:false},
  {id:'clarity',label:'They are not sure what is expected',note:'Say it first; training may not be needed.',training:false},
]);
export const ASSESSMENT_KINDS=Object.freeze([
  {id:'quiz',label:'Quiz or test'},{id:'demo',label:'Demonstration or skills check'},{id:'project',label:'Project or work product'},
  {id:'discussion',label:'Discussion or reflection'},{id:'observation',label:'Observation'},{id:'self',label:'Self-check'},
]);
export const KIRKPATRICK=Object.freeze([
  {id:1,label:'Level 1: Reaction',asks:'How did they find it?',kinds:['rating items','open comments']},
  {id:2,label:'Level 2: Learning',asks:'What did they learn?',kinds:['test before and after','skills check','self-scored instrument']},
  {id:3,label:'Level 3: Behavior',asks:'Are they doing it at work?',kinds:['observation','manager rating','self-report','work data']},
  {id:4,label:'Level 4: Results',asks:'Did it change what the work is for?',kinds:['measure of the work','cost or time','quality']},
]);

const clean=t=>String(t??'').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').trim();
const isObj=v=>v!==null&&typeof v==='object'&&!Array.isArray(v);
const str=(v,max,what)=>{
  const t=clean(v);
  if(t.length>max)throw new Error(`${what} can be ${max.toLocaleString('en-US')} characters at most.`);
  return t;
};
const arr=(v,max,what)=>{
  if(v===undefined||v===null)return [];
  if(!Array.isArray(v))throw new Error(`The design notes are not in the right form (${what}).`);
  if(v.length>max)throw new Error(`There can be ${max} ${what} at most.`);
  return v;
};

export function emptyDesign(){
  return {
    version:DESIGN_VERSION,
    needs:{problem:'',audience:'',current:'',desired:'',causes:[],constraints:'',success:'',delivery:'',notes:''},
    objectives:[],models:[],activities:[],assessments:[],storyboard:[],
    evaluation:{measures:[],notes:''},
  };
}

/** The ids a list already uses: a new one is the next that is free (o1, o2, o3...). */
export function nextId(prefix,...lists){
  const used=new Set(lists.flat().map(x=>x?.id));
  for(let n=1;;n++)if(!used.has(`${prefix}${n}`))return `${prefix}${n}`;
}

const idList=(v,known,what)=>[...new Set(arr(v,LIMITS.objectives,what).map(String).filter(id=>known.has(id)))];

/**
 * A design checked and tidied: fields it does not know are dropped, a missing field takes its empty value, text is trimmed, every id a list points at must exist
 * (one that does not is taken out), and every list has its limit. Throws a message fit to show when something is beyond what a design can hold.
 */
export function cleanDesign(raw){
  const d=emptyDesign();
  if(raw===null||raw===undefined)return d;
  if(!isObj(raw))throw new Error('The design notes are not in the right form, so they were not opened.');
  if(raw.version!==undefined&&(!Number.isInteger(raw.version)||raw.version<1))throw new Error('The design notes have no usable version, so they were not opened.');
  if(raw.version>DESIGN_VERSION)throw new Error('The design notes were saved by a newer OpenStage Guide. Update the program, then open the guide again.');
  const n=isObj(raw.needs)?raw.needs:{};
  d.needs={
    problem:str(n.problem,LIMITS.text,'The problem'),audience:str(n.audience,LIMITS.text,'The audience'),current:str(n.current,LIMITS.text,'What they do now'),
    desired:str(n.desired,LIMITS.text,'What they should do'),causes:[...new Set(arr(n.causes,NEED_CAUSES.length,'causes').map(String).filter(c=>NEED_CAUSES.some(x=>x.id===c)))],
    constraints:str(n.constraints,LIMITS.text,'The constraints'),success:str(n.success,LIMITS.text,'How success is measured'),delivery:str(n.delivery,LIMITS.text,'How it is delivered'),notes:str(n.notes,LIMITS.text,'The notes'),
  };
  const objectives=arr(raw.objectives,LIMITS.objectives,'objectives').map((o,i)=>{
    if(!isObj(o))throw new Error(`Objective ${i+1} is not in the right form.`);
    const level=BLOOM_LEVELS.some(l=>l.id===o.level)?o.level:'';
    return {id:str(o.id,LIMITS.short,'An id')||`o${i+1}`,text:str(o.text,LIMITS.text,'An objective'),level};
  });
  if(new Set(objectives.map(o=>o.id)).size!==objectives.length)throw new Error('Two objectives have the same id.');
  d.objectives=objectives;
  const oids=new Set(objectives.map(o=>o.id));
  const modelIds=arr(raw.models,LIMITS.models,'models').map(String);
  d.models=[...new Set(modelIds)];
  d.activities=arr(raw.activities,LIMITS.activities,'activities').map((a,i)=>{
    if(!isObj(a))throw new Error(`Activity ${i+1} is not in the right form.`);
    const minutes=Number(a.minutes);
    return {id:str(a.id,LIMITS.short,'An id')||`a${i+1}`,title:str(a.title,LIMITS.title,'An activity title'),model:str(a.model,LIMITS.short,'A model'),event:str(a.event,LIMITS.title,'An event'),
      minutes:Number.isInteger(minutes)&&minutes>=0&&minutes<=999?minutes:0,objectives:idList(a.objectives,oids,'objectives'),note:str(a.note,LIMITS.text,'A note')};
  });
  if(new Set(d.activities.map(a=>a.id)).size!==d.activities.length)throw new Error('Two activities have the same id.');
  d.assessments=arr(raw.assessments,LIMITS.assessments,'assessments').map((a,i)=>{
    if(!isObj(a))throw new Error(`Assessment ${i+1} is not in the right form.`);
    const slots=Number(a.slots);
    return {id:str(a.id,LIMITS.short,'An id')||`s${i+1}`,title:str(a.title,LIMITS.title,'An assessment title'),
      kind:ASSESSMENT_KINDS.some(k=>k.id===a.kind)?a.kind:'quiz',level:BLOOM_LEVELS.some(l=>l.id===a.level)?a.level:'',
      slots:Number.isInteger(slots)&&slots>=0&&slots<=50?slots:0,objectives:idList(a.objectives,oids,'objectives'),note:str(a.note,LIMITS.text,'A note')};
  });
  if(new Set(d.assessments.map(a=>a.id)).size!==d.assessments.length)throw new Error('Two assessments have the same id.');
  const aids=new Set(d.activities.map(a=>a.id)),sids=new Set(d.assessments.map(a=>a.id));
  let lessons=0;
  d.storyboard=arr(raw.storyboard,LIMITS.modules,'modules').map((m,i)=>{
    if(!isObj(m))throw new Error(`Module ${i+1} is not in the right form.`);
    const ls=arr(m.lessons,LIMITS.lessons,'lessons');
    lessons+=ls.length;
    if(lessons>LIMITS.lessons)throw new Error(`There can be ${LIMITS.lessons} lessons at most.`);
    return {id:str(m.id,LIMITS.short,'An id')||`m${i+1}`,title:str(m.title,LIMITS.title,'A module title'),lessons:ls.map((l,j)=>{
      if(!isObj(l))throw new Error(`Lesson ${j+1} of module ${i+1} is not in the right form.`);
      const minutes=Number(l.minutes);
      return {id:str(l.id,LIMITS.short,'An id')||`l${lessons}-${j+1}`,title:str(l.title,LIMITS.title,'A lesson title'),minutes:Number.isInteger(minutes)&&minutes>=0&&minutes<=999?minutes:0,
        objectives:idList(l.objectives,oids,'objectives'),activities:idList(l.activities,aids,'activities'),assessments:idList(l.assessments,sids,'assessments'),note:str(l.note,LIMITS.text,'A note')};
    })};
  });
  const ids=d.storyboard.flatMap(m=>[m.id,...m.lessons.map(l=>l.id)]);
  if(new Set(ids).size!==ids.length)throw new Error('Two parts of the storyboard have the same id.');
  const ev=isObj(raw.evaluation)?raw.evaluation:{};
  d.evaluation={measures:arr(ev.measures,LIMITS.measures,'measures').map((m,i)=>{
    if(!isObj(m))throw new Error(`Measure ${i+1} is not in the right form.`);
    const level=Number(m.level);
    return {id:str(m.id,LIMITS.short,'An id')||`e${i+1}`,level:KIRKPATRICK.some(k=>k.id===level)?level:1,name:str(m.name,LIMITS.title,'A measure name'),kind:str(m.kind,LIMITS.short,'A kind of measure'),
      when:str(m.when,LIMITS.title,'When'),who:str(m.who,LIMITS.title,'Who'),scale:str(m.scale,LIMITS.title,'A scale'),
      statements:arr(m.statements,LIMITS.statements,'statements').map(s=>str(s,LIMITS.text,'A statement')).filter(Boolean),objectives:idList(m.objectives,oids,'objectives'),note:str(m.note,LIMITS.text,'A note')};
  }),notes:str(ev.notes,LIMITS.text,'The evaluation notes')};
  if(new Set(d.evaluation.measures.map(m=>m.id)).size!==d.evaluation.measures.length)throw new Error('Two measures have the same id.');
  if(JSON.stringify(d).length>LIMITS.chars)throw new Error('The design notes are larger than a guide can keep.');
  return d;
}

/** The notes as the text the project file keeps, or null for notes that are still empty. */
export function designText(design){
  const d=cleanDesign(design);
  return JSON.stringify(d)===JSON.stringify(emptyDesign())?null:JSON.stringify(d);
}
/** The notes from the text the project file keeps (null or unreadable text gives empty notes, and an unreadable text is said so by `readDesign`). */
export function readDesign(text){
  if(text===null||text===undefined||text==='')return emptyDesign();
  let raw;
  try{raw=JSON.parse(text);}catch{throw new Error('The design notes in this guide are not readable, so the design desk was not opened. Nothing was changed.');}
  return cleanDesign(raw);
}

// ---- changes that touch more than one list

/** An objective taken out of the notes, and out of everything that pointed at it. */
export function removeObjective(design,id){
  const d=structuredClone(design);
  d.objectives=d.objectives.filter(o=>o.id!==id);
  const drop=list=>list.forEach(x=>{x.objectives=x.objectives.filter(o=>o!==id);});
  drop(d.activities);drop(d.assessments);drop(d.evaluation.measures);
  d.storyboard.forEach(m=>drop(m.lessons));
  return d;
}
export function removeActivity(design,id){
  const d=structuredClone(design);
  d.activities=d.activities.filter(a=>a.id!==id);
  d.storyboard.forEach(m=>m.lessons.forEach(l=>{l.activities=l.activities.filter(a=>a!==id);}));
  return d;
}
export function removeAssessment(design,id){
  const d=structuredClone(design);
  d.assessments=d.assessments.filter(a=>a.id!==id);
  d.storyboard.forEach(m=>m.lessons.forEach(l=>{l.assessments=l.assessments.filter(a=>a!==id);}));
  return d;
}

/** The minutes the storyboard adds up to, and by module. */
export function storyboardMinutes(design){
  const modules=design.storyboard.map(m=>({id:m.id,title:m.title,minutes:m.lessons.reduce((a,l)=>a+(l.minutes||0),0)}));
  return {total:modules.reduce((a,m)=>a+m.minutes,0),modules};
}
