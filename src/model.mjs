import {normalizeMaterials} from './materials.mjs';
import {hasFrontMatter,validateFrontMatter,createFrontMatter,frontMatterFields} from './frontmatter.mjs';
export const SCHEMA = 1;
export function activityMaterials(project,id){validate(project);const a=project.modules.flatMap(m=>m.activities).find(a=>a.id===id);if(!a)throw new Error('Select an existing activity.');return normalizeMaterials(a.materials===undefined?[]:a.materials);}
export function editActivityMaterials(project,id,materials){
  validate(project);const next=structuredClone(project),a=next.modules.flatMap(m=>m.activities).find(a=>a.id===id);if(!a)throw new Error('Select an existing activity.');
  const prior=normalizeMaterials(a.materials===undefined?[]:a.materials),items=normalizeMaterials(materials);
  if(JSON.stringify(prior)!==JSON.stringify(items)){a.materials=items;next.revision++;}return next;
}
export const COURSE_DETAIL_FIELDS=Object.freeze(['audience','purpose','authors','attribution']);
export function courseDetails(project){
  const details=project?.details;
  if(details!==undefined&&(!details||typeof details!=='object'||Array.isArray(details)))throw new Error('The course details are unsupported. Preserve this document before continuing.');
  return Object.fromEntries(COURSE_DETAIL_FIELDS.map(key=>{
    const text=details?.[key]??'';
    if(typeof text!=='string'||text.length>50000)throw new Error('The course details are unsupported. Preserve this document before continuing.');
    return [key,text];
  }));
}
export function normalizeCourseDetails(fields){
  if(!fields||typeof fields!=='object'||Array.isArray(fields))throw new Error('Enter the course details.');
  const details={};
  for(const key of COURSE_DETAIL_FIELDS){
    if(typeof fields[key]!=='string')throw new Error(`Enter ${key} as text.`);
    details[key]=value(fields[key],key);
  }
  return details;
}
export function editCourseDetails(project,fields){
  validate(project);
  const prior=courseDetails(project),details=normalizeCourseDetails(fields);
  const next=structuredClone(project);
  if(COURSE_DETAIL_FIELDS.some(key=>details[key]!==prior[key])){
    next.details={...next.details,...details};next.revision++;
  }
  return next;
}
export const ICONS = [
  ['Objective','Objective.png'],['Key point','Key Point.png'],['Lecturette','Lecture.png'],
  ['Slide','PowerPoint.png'],['Media','Media.png'],['Training instrument','Training Instrument.png'],
  ['Learning activity','Interaction.png'],['Participant workbook','Workbook.png'],['Handout','Hand-Out.png'],
  ['Discussion','Questions.png'],['Time on task','Time on Task.png'],['Trainer notes','Trainer Notes.png'],
  ['Transition','Transitions.png'],['Assessment','Assessment.png']
];
export function value(text, name, required = false) {
  const v = String(text ?? '').replace(/\r\n?/g, '\n').trim();
  if (required && !v) throw new Error(`Enter ${name}.`);
  if (v.length > 50000) throw new Error(`${name} is too long for this prototype.`);
  return v;
}
export function course(title,details) {
  const project={ schema: SCHEMA, id: crypto.randomUUID(), title: value(title,'a course title',true), revision: 0, modules: [] };
  if(details!==undefined)project.details=normalizeCourseDetails(details);
  return project;
}
export function wrenchCourse(title,details,selected=[]){const p=course(title,details);p.schema=2;p.frontMatter=createFrontMatter(selected);return validate(p);}
export function validate(project) {
  if (!project || ![SCHEMA,2].includes(project.schema) || !Array.isArray(project.modules) || !project.id || !project.title) throw new Error('This document has unsupported OpenStage Guide data. Keep a copy and use the matching add-in version.');
  if(!Number.isSafeInteger(project.revision)||project.revision<0)throw new Error('The course revision is invalid. Preserve this document before repairing it.');
  courseDetails(project);
  const ids = new Set([project.id]);
  for (const m of project.modules) {
    if (!m.id || ids.has(m.id) || !Array.isArray(m.activities)) throw new Error('The course has missing or duplicate module IDs. Keep a copy before repairing it.');
    ids.add(m.id);
    for (const a of m.activities) {
      if (!a.id || ids.has(a.id)) throw new Error('The course has missing or duplicate activity IDs.');
      normalizeMaterials(a.materials===undefined?[]:a.materials);
      ids.add(a.id);
    }
  }
  validateFrontMatter(project,ids);return project;
}
export function addModule(project, title, objective) {
  validate(project);
  const next = structuredClone(project);
  const m = { id: crypto.randomUUID(), title: value(title,'a module title',true), objective: value(objective,'an objective'), activities: [] };
  next.modules.push(m); next.revision++;
  return { next, item: m };
}
export function addActivity(project, moduleId, fields) {
  validate(project);
  const next = structuredClone(project), m = next.modules.find(x=>x.id===moduleId);
  if (!m) throw new Error('Select a module first.');
  const minutes = Number(fields.minutes);
  if (!Number.isFinite(minutes) || minutes < 0 || minutes > 1440) throw new Error('Enter a duration between 0 and 1440 minutes.');
  const a = { id: crypto.randomUUID(), title: value(fields.title,'an activity title',true),
    icon: ICONS.some(([label])=>label===fields.icon) ? fields.icon : 'Learning activity', minutes,
    facilitator: value(fields.facilitator,'facilitator instructions'),
    participant: value(fields.participant,'participant instructions',true),
    answers: value(fields.answers,'facilitator answers'), lines: Math.max(0,Math.min(20,Number(fields.lines)||0)) };
  m.activities.push(a); next.revision++;
  if(Object.hasOwn(fields,'materials'))a.materials=normalizeMaterials(fields.materials);
  return { next, item: a };
}
export function total(project) { return project.modules.reduce((n,m)=>n+m.activities.reduce((s,a)=>s+a.minutes,0),0); }
export function editActivitySettings(project,id,fields){
  validate(project);
  const next=structuredClone(project),item=next.modules.flatMap(m=>m.activities).find(a=>a.id===id);
  if(!item)throw new Error('Select an existing activity.');
  const minutes=Number(fields.minutes),lines=Number(fields.lines);
  if(String(fields.minutes??'').trim()===''||!Number.isFinite(minutes)||minutes<0||minutes>1440)throw new Error('Enter a duration between 0 and 1440 minutes.');
  if(String(fields.lines??'').trim()===''||!Number.isInteger(lines)||lines<0||lines>20)throw new Error('Enter a whole number of writing lines between 0 and 20.');
  if(item.minutes!==minutes||item.lines!==lines){item.minutes=minutes;item.lines=lines;next.revision++;}
  return {next,item};
}
export function participantView(project) {
  validate(project);
  // Deliberately allowlist learner fields. Never serialize the whole authoring record into an output.
  return { title: project.title, revision: project.revision, modules: project.modules.map(m=>({title:m.title, objective:m.objective,
    activities:m.activities.map(a=>({title:a.title, participant:a.participant, lines:a.lines}))})) };
}
export function reconcile(project, records) {
  if(project)validate(project);
  const next = structuredClone(project), map = new Map(),expected=managedTags(project);
  for (const {tag,text} of records) {
    if (typeof tag!=='string'||!tag.startsWith('osg:')) continue;
    if (map.has(tag)) throw new Error('A managed field was copied in Word. Duplicate tags must be resolved before export. Keep a backup of the document.');
    if(!expected.has(tag))throw new Error('An unexpected managed field exists in Word. Course records are incomplete or stale. Preserve the document and use recovery before continuing.');
    if(typeof text!=='string')throw new Error('A managed field could not be read. Preserve the document before continuing.');
    map.set(tag, text.replace(/\r/g,'\n').replace(/\u200b/g,'').trim());
  }
  if(!project)return null;
  const read=(type,id)=>{
    const tag=`osg:${type}:${id}`;
    if (!map.has(tag)) throw new Error('A managed field was deleted from Word. Export is paused so content is not silently omitted. Use Undo or reopen your saved copy.');
    return map.get(tag);
  };
  next.title = read('course',next.id);
  if (!next.title) throw new Error('The course title cannot be blank.');
  if(hasFrontMatter(next))for(const s of next.frontMatter.sections){s.heading=read('section-heading',s.id);s.body=read('section-body',s.id);if(!s.heading)throw new Error('A front-matter heading cannot be blank. Preserve this guide.');}
  for (const m of next.modules) {
    m.title=read('module',m.id); m.objective=read('objective',m.id);
    for (const a of m.activities) {
      a.title=read('activity',a.id);
      a.facilitator=read('facilitator',a.id);
      a.participant=read('participant',a.id);
      a.answers=read('answers',a.id);
    }
  }
  if (JSON.stringify(next)!==JSON.stringify(project)) next.revision++;
  return next;
}
export function managedTags(project){
  if(!project)return new Set();
  validate(project);
  return new Set([`osg:course:${project.id}`,...frontMatterFields(project).map(f=>f.tag),...project.modules.flatMap(m=>[
    `osg:module:${m.id}`,`osg:objective:${m.id}`,
    ...m.activities.flatMap(a=>['activity','facilitator','participant','answers'].map(type=>`osg:${type}:${a.id}`))
  ])]);
}
