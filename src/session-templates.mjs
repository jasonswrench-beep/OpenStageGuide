// Session templates (milestone 1 of the product spec of 2026-11-22): browse the ready-made sessions, preview one without changing anything, apply a COPY of it to the guide,
// and save a lesson of the guide as a template of the author's own.
//
// A session is a lesson. A template is applied as ordinary paragraphs (a lesson with its time line, its facts and one Activity block for each row of the template), alone
// in the module the cursor is in or in a module of its own, so everything that reads a guide reads the copy, and an edit to one copy touches neither another copy nor the
// catalog (src/session-catalog.mjs, which is read-only: a bundled template can only be copied). The copy says where it came from in words: a "Template:" fact under the
// lesson's time line and a "Source:" part in each activity, so there is no hidden record to fall out of step.
import {TEMPLATES} from './session-catalog.mjs';
import {ACTIVITY_PARTS} from './guide-cues.mjs';
import {cueBlock,lessonHeading,moduleHeading,timeBlock} from './guide-ooxml.mjs';
import {factsBlock,readFacts,CLOCK} from './guide-facts.mjs';
import {readGuide,readBlock,blockSpan,blockAnchor,kindOf,cueFromText} from './guide-model.mjs';
import {rowEntry,activityById} from './session-text.mjs';

export const MAX_MINE=30,MAX_ROWS=40;
const FIELD_LIMIT=4000,TITLE_LIMIT=120;
const PART_IDS=ACTIVITY_PARTS.map(p=>p.id);
const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();
const lines=t=>String(t??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);

// ---- the templates as the tool lists them
export const TIME_BANDS=Object.freeze([
  {id:'short',label:'Up to 60 minutes',test:m=>m<=60},
  {id:'medium',label:'61 to 75 minutes',test:m=>m>60&&m<=75},
  {id:'long',label:'76 minutes or more',test:m=>m>75}
]);
export const PACK_LABELS=Object.freeze({Launch:'Launch templates',Expansion:'Expansion pack',Yours:'Your templates'});
export const groupText=g=>g&&g.max?`${g.min} to ${g.max} people`:'';
const sumMinutes=rows=>rows.reduce((n,r)=>n+(Number(r.minutes)||0),0);

/** A bundled template with its rows written out (its activities by ID, from the catalog). */
function activityOf(id){const a=activityById(id);if(!a)throw new Error(`The catalog has no activity ${id}.`);return a;}
const bundled=t=>Object.freeze({
  id:t.id,version:t.version,mine:false,title:t.title,pack:t.pack,family:t.family,minutes:t.minutes,audience:t.audience,group:t.group,
  purpose:t.purpose,output:t.output,materials:t.materials,prerequisites:t.prerequisites,rationale:t.rationale,
  rows:Object.freeze(t.rows.map(id=>Object.freeze(rowEntry(activityOf(id)))))
});
export const BUNDLED=Object.freeze(TEMPLATES.map(bundled));

/** The author's own templates as they are kept (cleanMine reads them back). */
const MINE_FIELDS=['id','title','purpose','output','audience','prerequisites','rationale'];
function cleanRow(r){
  if(!r||typeof r!=='object')return null;
  const title=oneLine(r.title).slice(0,TITLE_LIMIT);
  if(!title)return null;
  const parts={};
  for(const id of PART_IDS)parts[id]=String(r.parts?.[id]??'').slice(0,FIELD_LIMIT);
  if(!lines(parts.instructions).length)return null;
  const minutes=Number(r.minutes)>0?Math.min(Math.round(Number(r.minutes)),999):0;
  return {title,minutes,parts,notes:String(r.notes??'').slice(0,FIELD_LIMIT),activityId:typeof r.activityId==='string'?r.activityId.slice(0,10):''};
}
export function cleanMine(raw){
  if(!Array.isArray(raw))return [];
  const out=[],seen=new Set();
  for(const r of raw){
    if(!r||typeof r!=='object')continue;
    const id=typeof r.id==='string'?r.id.slice(0,60):'',title=oneLine(r.title).slice(0,TITLE_LIMIT);
    const rows=(Array.isArray(r.rows)?r.rows:[]).slice(0,MAX_ROWS).map(cleanRow).filter(Boolean);
    if(!id||!title||seen.has(id)||!rows.length)continue;
    const g=r.group&&Number.isInteger(r.group.min)&&Number.isInteger(r.group.max)&&r.group.min>=1&&r.group.max>r.group.min&&r.group.max<=999?{min:r.group.min,max:r.group.max}:null;
    const t={id,version:1,mine:true,title,pack:'Yours',family:'Yours',minutes:sumMinutes(rows),group:g,rows:Object.freeze(rows)};
    for(const f of MINE_FIELDS.slice(2))t[f]=oneLine(r[f]).slice(0,400);
    t.materials=(Array.isArray(r.materials)?r.materials:[]).map(m=>oneLine(m).slice(0,120)).filter(Boolean).slice(0,20);
    seen.add(id);out.push(Object.freeze(t));
    if(out.length>=MAX_MINE)break;
  }
  return out;
}
export const serializeMine=list=>JSON.stringify(list.map(({mine,pack,family,minutes,version,...t})=>t));
export const allTemplates=(mine=[])=>[...BUNDLED,...mine];
export const templateById=(id,mine=[])=>allTemplates(mine).find(t=>t.id===id)||null;

/** What a browse card shows: the title, what it is for, what it makes, who it is for, the group and the minutes. */
export const templateCard=t=>({id:t.id,title:t.title,pack:t.pack,family:t.family,minutes:t.minutes,audience:t.audience||'',group:groupText(t.group),purpose:t.purpose||'',output:t.output||'',mine:!!t.mine});

// ---- search and filters: words across the title, the goal, the output, the audience and the tags (family, pack, materials, activity titles)
export function matchTemplates(list,{text='',family='',band='',pack=''}={}){
  const words=oneLine(text).toLowerCase().split(' ').filter(Boolean);
  const wanted=TIME_BANDS.find(b=>b.id===band);
  return list.filter(t=>{
    if(family&&t.family!==family)return false;
    if(pack&&t.pack!==pack)return false;
    if(wanted&&!wanted.test(t.minutes))return false;
    if(!words.length)return true;
    const hay=[t.title,t.purpose,t.output,t.audience,t.family,PACK_LABELS[t.pack]||t.pack,...(t.materials||[]),...t.rows.map(r=>r.title)].join(' ').toLowerCase();
    return words.every(w=>hay.includes(w));
  });
}

// ---- preview: read-only. Nothing here changes the guide or the template.
const CLOCK_TEXT=(base,add)=>{const t=base+add;return `${String(Math.floor(t/60)%24).padStart(2,'0')}:${String(t%60).padStart(2,'0')}`;};
export function checkStart(start){
  const s=oneLine(start);
  if(!s)return '';
  const c=CLOCK.exec(s);
  if(!c)throw new Error('Give the start time as a clock time such as 9:00, or leave it out.');
  return `${Number(c[1])}:${c[2]}`;
}
export function templatePreview(t,{start=''}={}){
  const s=checkStart(start),base=s?Number(s.split(':')[0])*60+Number(s.split(':')[1]):null;
  let offset=0;
  const agenda=t.rows.map((r,i)=>{
    const row={order:i+1,title:r.title,minutes:r.minutes,phase:r.parts.phase||'',grouping:r.parts.grouping||'',output:r.parts.output||'',offset,
      from:base===null?`${offset} min`:CLOCK_TEXT(base,offset),to:base===null?`${offset+r.minutes} min`:CLOCK_TEXT(base,offset+r.minutes)};
    offset+=r.minutes;
    return row;
  });
  const access=[...new Set(t.rows.flatMap(r=>lines(r.parts.access)))];
  return {
    ...templateCard(t),
    outcomes:t.output?[t.output]:[],prerequisites:t.prerequisites||'',rationale:t.rationale||'',materials:[...(t.materials||[])],
    agenda,total:offset,adjustments:access,
    note:t.mine?'This is one of your own templates. Applying it writes a copy into the guide.':'Applying it writes a copy into the guide. The template itself is never changed.'
  };
}

// ---- apply: a lesson written as paragraphs, from a COPY of the template's rows
const partsFor=r=>Object.fromEntries(PART_IDS.map(id=>[id,String(r.parts?.[id]??'')]));
const templateFact=t=>t.mine?`Template: ${t.title} (your template)`:`Template: ${t.title} (${t.id}, version ${t.version})`;

/**
 * The paragraphs of one session: its heading, time line, facts (what it is for, its materials, the group, the template it came from, and "New day" when a start time is given),
 * an Objective block for what the group leaves with, and an Activity block for each row. `m` and `l` are the module and lesson numbers it goes under, `before` the minutes taught before it.
 */
export function sessionXml(t,{title,audience='',start='',m,l,before=0},ctx){
  const name=oneLine(title);
  if(!name)throw new Error('Type a title for the session first.');
  if(!t.rows.length)throw new Error('This template has no activities.');
  const s=checkStart(start),who=oneLine(audience);
  const facts={
    purpose:[oneLine(t.purpose).replace(/\.$/,''),who?`For ${who}`:''].filter(Boolean).join('. ')+(t.purpose||who?'.':''),
    materials:(t.materials||[]).join(', '),group:groupText(t.group),newDay:s,extra:[templateFact(t)]
  };
  let out=lessonHeading(m,l,name)+timeBlock({minutes:sumMinutes(t.rows),cumulative:before},ctx)+factsBlock(facts);
  if(t.output)out+=cueBlock({cue:'objective',text:`Participants leave with: ${oneLine(t.output).replace(/\.$/,'')}.`},ctx);
  for(const r of t.rows)out+=cueBlock({cue:'activity',text:r.title,minutes:r.minutes||undefined,parts:partsFor(r),notes:r.notes||''},ctx);
  return out;
}
/** The same session in a module of its own, with the Transition every module ends with. */
export function moduleXml(t,opts,ctx){
  return moduleHeading(opts.m,oneLine(opts.moduleTitle||opts.title))+sessionXml(t,{...opts,l:1},ctx)+cueBlock({cue:'transition',text:'Say how this module leads into the next one.'},ctx);
}

/** Where an activity put in from the library goes: after the block the cursor is in, named in words. The picker shows it before anything is inserted. */
export function destinationText(paras,at){
  if(!paras.length)return 'Put your cursor in a lesson first.';
  try{
    const anchor=blockAnchor(paras,Math.max(at,0)),span=blockSpan(paras,anchor),head=paras[span?span.start:anchor]||paras[anchor];
    const lesson=readGuide(paras).modules.flatMap(m=>m.lessons).filter(l=>l.index<=anchor).at(-1);
    const cue=head&&kindOf(head)==='cue'?cueFromText(head.text):null,first=oneLine(String(head?.text??'').replace(/^[^A-Za-z0-9]+/,'')).slice(0,40);
    const where=lesson?` in ${lesson.text}`:'';
    return cue?`It goes after the ${cue.label} block "${first.replace(new RegExp(`^${cue.label}\\s*`,'i'),'')||cue.label}"${where}.`:`It goes after "${first}"${where}.`;
  }catch{return 'Put your cursor in a lesson first.';}
}

// ---- save as a template: the lesson the cursor is in, with a deliberate choice about what to keep
/**
 * `keepExamples` keeps each activity's case study; `keepNotes` keeps its facilitator notes. Without them a saved template holds the instructions, materials, steps and prompts
 * and nothing that may name a person. Returns {template, skipped} (skipped: activities with no instructions, which a template cannot hold).
 */
export function templateFromLesson(paras,at,{name,keepExamples=false,keepNotes=false}={}){
  const guide=readGuide(paras),lessons=guide.modules.flatMap(m=>m.lessons);
  const lesson=lessons.filter(l=>l.index<=at).at(-1);
  if(!lesson||at>lesson.end)throw new Error('Put your cursor inside a lesson first.');
  const title=oneLine(name||lesson.title).slice(0,TITLE_LIMIT);
  if(!title)throw new Error('Type a name for the template first.');
  const facts=readFacts(paras,lesson).facts;
  const rows=[];let skipped=0;
  for(let i=lesson.index+1;i<=lesson.end;i++){
    if(kindOf(paras[i])!=='cue')continue;
    const cue=cueFromText(paras[i].text);
    if(!cue||cue.id!=='activity')continue;
    const span=blockSpan(paras,i);
    if(!span||span.start!==i)continue;
    i=span.end;
    const b=readBlock(paras,span.start);
    if(!b){skipped++;continue;}
    let parts,rowTitle,notes=b.notes||'';
    if(b.parts){parts=partsFor({parts:b.parts});rowTitle=oneLine(b.text);}
    else{
      const all=lines(b.text);rowTitle=oneLine(all[0]);
      parts=partsFor({parts:{instructions:all.slice(1).join('\n')}});
    }
    if(!rowTitle||!lines(parts.instructions).length){skipped++;continue;}
    if(!keepExamples)parts.caseStudy='';
    if(!keepNotes)notes='';
    rows.push({title:rowTitle,minutes:Number(b.minutes)||0,parts,notes,activityId:''});
    if(rows.length>=MAX_ROWS)break;
  }
  if(!rows.length)throw new Error('This lesson has no activities to save. Add Activity blocks with instructions first.');
  const m=/^(\d+)\s*to\s*(\d+)\s*people/i.exec(facts.group||'');
  const template={
    title,purpose:oneLine(facts.purpose).slice(0,400),output:'',audience:'',prerequisites:'',rationale:'',
    materials:oneLine(facts.materials).split(/\s*,\s*/).filter(Boolean).slice(0,20),
    group:m?{min:Number(m[1]),max:Number(m[2])}:null,rows
  };
  return {template,skipped};
}
/** Adds a saved template among the author's own: one with the same title is replaced (saving it again updates it), a new one goes first. */
export function addMine(mine,saved,{now=()=>Date.now()}={}){
  const key=oneLine(saved.title).toLowerCase(),same=mine.find(t=>oneLine(t.title).toLowerCase()===key);
  if(!same&&mine.length>=MAX_MINE)throw new Error(`You have saved ${MAX_MINE} templates, which is the most. Remove one first.`);
  const entry={...saved,id:same?same.id:`mine-${now().toString(36)}-${mine.length}`};
  const list=cleanMine([entry,...mine.filter(t=>t!==same)]);
  if(!list.some(t=>t.id===entry.id))throw new Error('That lesson could not be saved as a template.');
  return {list,replaced:!!same,id:entry.id};
}
export const removeMine=(mine,id)=>mine.filter(t=>t.id!==id);
