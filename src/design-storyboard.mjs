// What the storyboard becomes. The storyboard is the author's plan: modules, each with lessons, each lesson with its minutes, the objectives it teaches, the activities
// in it and the assessments that check it (design-model.mjs). From it, when the author asks:
//   * the guide's own course plan: modules, lessons with their time lines, an Objective block for each objective, an Activity block for each activity (the model and its
//     event in the facilitator's notes), an Assessment block for each assessment, and a Transition at the end of each module;
//   * a skeleton for OpenStage Author (a project with a scene for each lesson and a slide for each screen), through the writer the guide's e-learning storyboard uses;
//   * a skeleton for OpenStage Scroll (an .openscroll package: course.json, in the format Scroll documents), a section for each module and a lesson for each lesson;
//   * empty item slots for OpenAssess, each tagged with the objective, its level, the chapter and the lesson, as a question table OpenAssess's own import reads.
// Author and Scroll have no field for an objective or its Bloom level, so each lesson says them in words (a block of the lesson on Scroll, a note in Author's hand-off notes).
// Nothing here is sent anywhere: each is a file or a guide the author keeps.
import {moduleHeading,lessonHeading,timeBlock,cueBlock,standardPages} from './guide-ooxml.mjs';
import {checkObjective,levelById} from './design-bloom.mjs';
import {modelById} from './design-models.mjs';

const plain=t=>String(t??'').replace(/\s+/g,' ').trim();
const lessonCount=d=>d.storyboard.reduce((a,m)=>a+m.lessons.length,0);
export function needStoryboard(design){
  if(!design.storyboard.length||!lessonCount(design))throw new Error('The storyboard has no lessons yet. Add a module and a lesson to it first.');
}
const objectiveById=d=>new Map(d.objectives.map(o=>[o.id,o]));
const levelOf=o=>o.level||checkObjective(o.text).level||'';
const minutesText=n=>(n>0?` (${n} min)`:'');

// What is said about a model, in the facilitator's words, for the note of an activity.
function modelNote(a){
  const m=a.model&&modelById(a.model);
  return [m?`Planned with ${m.name}${a.event?`, event: ${a.event}`:''}.`:'',plain(a.note)].filter(Boolean).join(' ');
}

// ---- the guide's own course plan

/**
 * The title page and standard pages, then each module, lesson and block of the storyboard, as the Word content a guide is made of.
 * @param {ReturnType<typeof import('./design-model.mjs').cleanDesign>} design @param {{title:string}} options @param {any} ctx
 */
export function planBody(design,{title},ctx){
  needStoryboard(design);
  const objectives=objectiveById(design),acts=new Map(design.activities.map(a=>[a.id,a])),asm=new Map(design.assessments.map(a=>[a.id,a]));
  let out=standardPages(plain(title)||'Course plan'),cumulative=0;
  design.storyboard.forEach((m,mi)=>{
    out+=moduleHeading(mi+1,plain(m.title)||`Module ${mi+1}`);
    m.lessons.forEach((l,li)=>{
      out+=lessonHeading(mi+1,li+1,plain(l.title)||`Lesson ${mi+1}.${li+1}`);
      out+=timeBlock({minutes:l.minutes||0,cumulative},ctx);
      cumulative+=l.minutes||0;
      for(const id of l.objectives){const o=objectives.get(id);if(o&&plain(o.text))out+=cueBlock({cue:'objective',text:plain(o.text)},ctx);}
      // A lesson's time is the sum of the minutes of its blocks when they have any (Update guide works it out so), so the blocks say their minutes only when they add up to the
      // minutes the author planned for the lesson; otherwise the lesson keeps its planned time, and each activity's minutes are in its notes.
      const sum=l.activities.map(id=>acts.get(id)).filter(Boolean).reduce((n,x)=>n+(x.minutes||0),0);
      const inBlocks=sum>0&&sum===(l.minutes||0);
      for(const id of l.activities){
        const a=acts.get(id);
        if(!a)continue;
        const goals=a.objectives.map(x=>objectives.get(x)).filter(Boolean).map(o=>plain(o.text)).filter(Boolean);
        out+=cueBlock({cue:'activity',text:plain(a.title)||'Activity',minutes:inBlocks?(a.minutes||undefined):undefined,
          notes:[modelNote(a),a.minutes&&!inBlocks?`Planned time: ${a.minutes} minutes.`:'',goals.length?`For: ${goals.join(' / ')}`:''].filter(Boolean).join('\n')},ctx);
      }
      for(const id of l.assessments){
        const a=asm.get(id);
        if(!a)continue;
        const checks=a.objectives.map(x=>objectives.get(x)).filter(Boolean).map(o=>plain(o.text)).filter(Boolean);
        out+=cueBlock({cue:'assessment',text:plain(a.title)||'Assessment',notes:[checks.length?`Checks: ${checks.join(' / ')}`:'',a.slots?`Item slots for each objective: ${a.slots}.`:'',plain(a.note)].filter(Boolean).join('\n')},ctx);
      }
    });
    out+=cueBlock({cue:'transition',text:'Say how this module leads into the next one.'},ctx);
  });
  return out;
}

// ---- OpenStage Author

/**
 * The screens of a skeleton for Author, in the shape the e-learning storyboard's screens have (guide-elearning.mjs), so guide-author.mjs writes them as it writes those:
 * a screen for each lesson (its objectives as on-screen text), a screen for each activity and a screen for each assessment. Everything Author has no field for is in the
 * notes of the screen: the level of each objective, the model an activity comes from, and the objectives each assessment checks.
 */
export function authorScreens(design){
  needStoryboard(design);
  const objectives=objectiveById(design),acts=new Map(design.activities.map(a=>[a.id,a])),asm=new Map(design.assessments.map(a=>[a.id,a]));
  const screens=[];
  design.storyboard.forEach((m,mi)=>m.lessons.forEach((l,li)=>{
    const lesson=`Lesson ${mi+1}.${li+1}: ${plain(l.title)||'Lesson'}`;
    let n=0;
    const add=(title,onScreen,notes)=>{
      n++;
      screens.push({title,lesson,number:n,onScreen:onScreen.join('\n'),narration:'',narrationWords:0,blocks:notes.map(([cue,text])=>({cue,text}))});
    };
    const goals=l.objectives.map(id=>objectives.get(id)).filter(o=>o&&plain(o.text));
    add(plain(l.title)||'Lesson',goals.map(o=>plain(o.text)),[
      ...goals.map(o=>['developer-note',`Objective (${levelById(levelOf(o))?.label||'no level'}): ${plain(o.text)}`]),
      ...(l.minutes?[['developer-note',`Planned time: ${l.minutes} minutes. Module: ${plain(m.title)}.`]]:[]),
      ...(plain(l.note)?[['developer-note',plain(l.note)]]:[]),
    ]);
    for(const id of l.activities){
      const a=acts.get(id);
      if(!a)continue;
      add(plain(a.title)||'Activity',[],[['interaction',`${plain(a.title)||'Activity'}${a.minutes?` (${a.minutes} min)`:''}`],...(modelNote(a)?[['developer-note',modelNote(a)]]:[])]);
    }
    for(const id of l.assessments){
      const a=asm.get(id);
      if(!a)continue;
      const checks=a.objectives.map(x=>objectives.get(x)).filter(o=>o&&plain(o.text));
      add(plain(a.title)||'Check',[],[['knowledge-check',`${plain(a.title)||'Knowledge check'}${a.slots?`: ${a.slots} ${a.slots===1?'question':'questions'} for each objective`:''}`],
        ...checks.map(o=>['developer-note',`Checks (${levelById(levelOf(o))?.label||'no level'}): ${plain(o.text)}`])]);
    }
  }));
  return screens;
}

// ---- OpenStage Scroll

const SCROLL_APPEARANCE=Object.freeze({width:'reading',padding:'medium',background:'page'});
const sid=(prefix,n)=>`${prefix}-${n}`;
/** A Scroll id: a lower-case letter, then lower-case letters, digits or hyphens (up to 64). */
export const scrollId=(s,fallback)=>{
  const t=plain(s).toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g,'').replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'').slice(0,40).replace(/-+$/g,'');
  return /^[a-z]/.test(t)?t:fallback;
};

/**
 * The course.json of an .openscroll package (OpenScroll's FORMAT-SPEC.md, schema version 3): a section for each module and a lesson for each lesson, each with a heading,
 * its objectives as a list, and its activities and assessments as lists. A skeleton: the author writes the content in Scroll.
 * @param {ReturnType<typeof import('./design-model.mjs').cleanDesign>} design @param {{title:string,author?:string,language?:string,now?:string}} options
 */
export function scrollCourse(design,{title,author='',language='en',now=new Date().toISOString().replace(/\.\d+Z$/,'Z')}){
  needStoryboard(design);
  const objectives=objectiveById(design),acts=new Map(design.activities.map(a=>[a.id,a])),asm=new Map(design.assessments.map(a=>[a.id,a]));
  const used=new Set();
  const unique=(base,fallback)=>{let id=scrollId(base,fallback),n=2;const root=id;while(used.has(id))id=`${root.slice(0,60)}-${n++}`;used.add(id);return id;};
  const block=(prefix,type,variant,content)=>({id:unique(prefix,prefix),type,variant,appearance:{...SCROLL_APPEARANCE},content});
  const sections=design.storyboard.map((m,mi)=>({id:unique(`${m.title}`,sid('module',mi+1)),title:plain(m.title)||`Module ${mi+1}`,lessons:m.lessons.map((l,li)=>{
    const blocks=[block(`heading-${mi+1}-${li+1}`,'heading','heading',{text:plain(l.title)||'Lesson'})];
    const goals=l.objectives.map(id=>objectives.get(id)).filter(o=>o&&plain(o.text));
    if(goals.length){
      blocks.push(block(`objectives-${mi+1}-${li+1}`,'statement','note',{text:'By the end of this lesson'}));
      blocks.push(block(`objective-list-${mi+1}-${li+1}`,'list','bulleted',{items:goals.map(o=>plain(o.text).slice(0,2000))}));
    }
    const planned=l.activities.map(id=>acts.get(id)).filter(Boolean).map(a=>`${plain(a.title)||'Activity'}${a.minutes?` (${a.minutes} min)`:''}`);
    if(planned.length)blocks.push(block(`activities-${mi+1}-${li+1}`,'list','bulleted',{items:planned}));
    const checks=l.assessments.map(id=>asm.get(id)).filter(Boolean).map(a=>plain(a.title)||'Check');
    if(checks.length)blocks.push(block(`checks-${mi+1}-${li+1}`,'list','checklist',{items:checks}));
    return {id:unique(`${m.title}-${l.title}`,sid('lesson',`${mi+1}-${li+1}`)),title:plain(l.title)||`Lesson ${mi+1}.${li+1}`,blocks};
  })}));
  return {
    format:'openscroll',schemaVersion:3,
    course:{id:unique(title,'course'),title:plain(title)||'Course',description:plain(design.needs.problem).slice(0,500),author:plain(author),language,createdAt:now,updatedAt:now,cover:null,
      theme:{accentColor:'#0b3fb3',font:'sans',pageWidth:'standard'},settings:{navigation:'free',completion:'all-lessons',passingScore:80},sections},
    assets:[],
  };
}
export const scrollFileName=title=>`${plain(title).replace(/[\\/:*?"<>|\u0000-\u001f]/g,'-').slice(0,80)||'Course'}.openscroll`;
/** The package: a zip holding exactly course.json. @param {object} course @param {new()=>any} Zip */
export async function scrollBytes(course,Zip){
  const zip=new Zip();
  zip.file('course.json',JSON.stringify(course,null,2),{date:new Date(Date.UTC(2026,0,1)),compression:'DEFLATE'});
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}

// ---- OpenAssess: empty item slots

export const SLOT_COLUMNS=Object.freeze(['Type','Question','Points','Chapter','Section','Objective','Bloom','Tags','ID']);
const csvCell=v=>{const t=String(v??'');return /[",\r\n]/.test(t)?`"${t.replace(/"/g,'""')}"`:t;};
/**
 * The item slots, as rows of OpenAssess's question table: an essay row with a placeholder stem (OpenAssess holds no question without words), the objective's words and
 * level, the module as the chapter and the lesson as the section, and the tag "slot" so they can be found and written. Each assessment of the plan gives `slots` rows for
 * each objective it checks.
 */
export function itemSlots(design){
  const objectives=objectiveById(design),module=new Map(),lesson=new Map();
  for(const m of design.storyboard)for(const l of m.lessons)for(const id of l.assessments){if(!module.has(`${id}`)){module.set(id,m.title);lesson.set(id,l.title);}}
  const rows=[];
  for(const a of design.assessments){
    if(!a.slots)continue;
    for(const oid of a.objectives){
      const o=objectives.get(oid);
      if(!o||!plain(o.text))continue;
      const level=levelOf(o)||a.level||'';
      for(let n=1;n<=a.slots;n++){
        const level_=levelById(level)?.label||'';
        rows.push({Type:'ES',Question:`[Slot] Write a question that checks: ${plain(o.text)}`,Points:'1',Chapter:plain(module.get(a.id)||''),Section:plain(lesson.get(a.id)||''),Objective:plain(o.text),
          Bloom:level,Tags:`slot; ${a.kind}${level_?`; ${level_}`:''}`,ID:`guide-${o.id}-${a.id}-${n}`});
      }
    }
  }
  return rows;
}
/** The CSV text of the slots: UTF-8 with a byte-order mark. Throws when the plan has none. */
export function itemSlotsCsv(design){
  const rows=itemSlots(design);
  if(!rows.length)throw new Error('There are no item slots yet. On the Course map, give an assessment the number of item slots for each objective it checks, and link it to its objectives.');
  return `﻿${[SLOT_COLUMNS.join(','),...rows.map(r=>SLOT_COLUMNS.map(c=>csvCell(r[c])).join(','))].join('\r\n')}\r\n`;
}
