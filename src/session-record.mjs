// What the guide remembers about the materials and decks made from it (milestones 2 and 3 of the product spec of 2026-11-22), kept in the project file as one more member, `sessions`,
// like the design desk's notes and the form settings. The guide itself stays ordinary paragraphs with no hidden record: this holds only what cannot be read from them, which is
// which outputs were made, from what, and when, and the card decks the author is editing.
//
// An output (a run sheet, a preparation checklist, a participant workbook, a card deck, a slide deck) remembers the revision it came from: a fingerprint of each activity's words
// that matter to that output (fingerprints below). When the guide is read again and an activity's words have changed, the output is "Review needed" and the changes are named, so
// the author can refresh just that output. Nothing is ever rewritten on its own, and a file the program made is a copy: an earlier copy that was edited by hand is never touched.
import {readGuide,readBlock,blockSpan,kindOf,cueFromText} from './guide-model.mjs';
import {ACTIVITY_PARTS} from './guide-cues.mjs';

export const SESSIONS_VERSION=1;
export const STATUSES=Object.freeze(['Suggested','Ready','Review needed','Excluded']);
export const OUTPUT_KINDS=Object.freeze({runsheet:'Run sheet',checklist:'Preparation checklist',workbook:'Participant workbook',cards:'Card deck',slides:'Slide deck'});
export const LIMITS=Object.freeze({outputs:60,decks:30,cards:120,sources:400,name:120,scope:120});

const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();
const norm=t=>oneLine(t).toLowerCase();
const isObject=v=>v!==null&&typeof v==='object'&&!Array.isArray(v);
const iso=v=>{const t=Date.parse(v);return Number.isFinite(t)?new Date(t).toISOString():'';};

// ---- the fingerprint of an activity: FNV-1a over its words, for each part and for the parts together
export function hashText(text){
  let h=0x811c9dc5;
  const s=String(text??'');
  for(let i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,0x01000193)>>>0;}
  return h.toString(16).padStart(8,'0');
}
// Which words of an activity matter to which output. A change to anything else (the minutes of a block the output does not show) does not flag it.
const PART_SETS=Object.freeze({
  runsheet:'all',
  workbook:['title','objective','output','grouping','needs','instructions','caseStudy'],
  checklist:['title','materials','grouping'],
  slides:['title','minutes','output','grouping','instructions','steps','debrief','materials'],
  cards:['title','output','instructions','caseStudy']
});
const LABELS={title:'Title',minutes:'Minutes',notes:'Facilitator notes',...Object.fromEntries(ACTIVITY_PARTS.map(p=>[p.id,p.label]))};

/** Every Activity block of the guide for learners or not, in order, with its lesson and its words. `scope` may be {lesson:index} to keep one lesson's. */
export function activityBlocks(paras,{lessonIndex=null}={}){
  const out=[],seen=new Map();
  for(const mod of readGuide(paras).modules)for(const lesson of mod.lessons){
    for(let i=lesson.index+1;i<=lesson.end;i++){
      if(kindOf(paras[i])!=='cue')continue;
      const cue=cueFromText(paras[i].text);
      if(!cue||cue.id!=='activity')continue;
      const span=blockSpan(paras,i);
      if(!span||span.start!==i)continue;
      i=span.end;
      const b=readBlock(paras,span.start);
      if(!b)continue;
      const title=oneLine(String(b.text).split('\n')[0]);
      const n=(seen.get(norm(title))||0)+1;seen.set(norm(title),n);
      if(lessonIndex!==null&&lesson.index!==lessonIndex)continue;
      out.push({key:`${norm(title)}#${n}`,title,lesson:lesson.title,lessonIndex:lesson.index,index:span.start,minutes:Number(b.minutes)||0,audience:b.audience,parts:b.parts||{instructions:String(b.text).split('\n').slice(1).join('\n')},notes:b.notes||''});
    }
  }
  return out;
}
/** The fingerprints of the activities for one kind of output: [{key,title,hash,parts:{id:hash}}], in the order the guide has them. */
export function fingerprints(paras,kind,options){
  const set=PART_SETS[kind];
  if(!set)throw new Error('That kind of output is not known.');
  return activityBlocks(paras,options).map(a=>{
    const words={title:a.title,minutes:String(a.minutes||''),notes:a.notes,...Object.fromEntries(Object.entries(a.parts).map(([k,v])=>[k,String(v??'')]))};
    const ids=set==='all'?Object.keys(words):set;
    const parts={};
    for(const id of ids)if(words[id]!==undefined&&words[id]!=='')parts[id]=hashText(words[id]);
    return {key:a.key,title:a.title,hash:hashText(Object.entries(parts).map(([k,v])=>`${k}=${v}`).join(';')),parts};
  });
}

/** What changed between the fingerprints an output was made from and the guide now: named in words, never in the activities' own text. */
export function diffSources(before,now){
  const old=new Map(before.map(s=>[s.key,s])),cur=new Map(now.map(s=>[s.key,s]));
  const changed=[],added=[],removed=[];
  for(const s of now){
    const o=old.get(s.key);
    if(!o){added.push(s.title);continue;}
    if(o.hash!==s.hash){
      const ids=[...new Set([...Object.keys(o.parts),...Object.keys(s.parts)])].filter(id=>o.parts[id]!==s.parts[id]);
      changed.push({title:s.title,parts:ids.map(id=>LABELS[id]||id)});
    }
  }
  for(const o of before)if(!cur.has(o.key))removed.push(o.title);
  const common=now.map(s=>s.key).filter(k=>old.has(k)),was=before.map(s=>s.key).filter(k=>cur.has(k));
  const reordered=common.length>1&&common.join('|')!==was.join('|');
  return {changed,added,removed,reordered,same:!changed.length&&!added.length&&!removed.length&&!reordered};
}
export function describeDiff(d){
  if(d.same)return ['Nothing in it has changed since it was made.'];
  return [
    ...d.changed.map(c=>`${c.title}: ${c.parts.join(', ')} changed.`),
    ...d.added.map(t=>`${t} is new in the guide.`),
    ...d.removed.map(t=>`${t} is not in the guide any more.`),
    ...(d.reordered?['The activities are in a different order.']:[])
  ];
}

// ---- the member
const id=(v,fallback)=>typeof v==='string'&&/^[a-z0-9-]{1,40}$/i.test(v)?v:fallback;
function cleanSource(s){
  if(!isObject(s)||typeof s.key!=='string'||typeof s.title!=='string'||typeof s.hash!=='string')return null;
  const parts={};
  if(isObject(s.parts))for(const [k,v] of Object.entries(s.parts))if(/^[A-Za-z]{1,20}$/.test(k)&&typeof v==='string'&&/^[0-9a-f]{8}$/.test(v))parts[k]=v;
  return {key:s.key.slice(0,200),title:oneLine(s.title).slice(0,200),hash:s.hash.slice(0,16),parts};
}
function cleanOutput(o,i){
  if(!isObject(o)||!OUTPUT_KINDS[o.kind])return null;
  return {
    id:id(o.id,`o${i+1}`),kind:o.kind,name:oneLine(o.name).slice(0,LIMITS.name)||OUTPUT_KINDS[o.kind],scope:oneLine(o.scope).slice(0,LIMITS.scope)||'The whole guide',
    lessonIndex:Number.isInteger(o.lessonIndex)&&o.lessonIndex>=0?o.lessonIndex:null,
    builtAt:iso(o.builtAt),excluded:o.excluded===true,
    sources:(Array.isArray(o.sources)?o.sources:[]).slice(0,LIMITS.sources).map(cleanSource).filter(Boolean)
  };
}
function cleanDeck(d,i){
  if(!isObject(d))return null;
  const cards=(Array.isArray(d.cards)?d.cards:[]).slice(0,LIMITS.cards).map(c=>{
    if(!isObject(c))return null;
    const side=v=>(Array.isArray(v)?v:[]).slice(0,8).map(l=>oneLine(l).slice(0,300)).filter(Boolean);
    const front=side(c.front);
    return front.length?{front,back:side(c.back)}:null;
  }).filter(Boolean);
  return {
    id:id(d.id,`d${i+1}`),title:oneLine(d.title).slice(0,200)||'Cards',type:['activity','toolkit','facilitator'].includes(d.type)?d.type:'activity',
    paper:d.paper==='A4'?'A4':'Letter',perPage:[1,2,4,6,8].includes(d.perPage)?d.perPage:4,cropMarks:d.cropMarks===true,flip:d.flip==='short'?'short':'long',
    sequence:d.sequence==='separate'?'separate':'alternate',
    activity:oneLine(d.activity).slice(0,200),cards
  };
}
export const emptySessions=()=>({version:SESSIONS_VERSION,outputs:[],decks:[]});
/** The member as text -> the clean member. Anything not whole is left out; a text that is not ours is the same as an empty member (never throws). */
export function readSessions(text){
  let data;
  try{data=JSON.parse(String(text??''));}catch{return emptySessions();}
  if(!isObject(data)||data.version!==SESSIONS_VERSION)return emptySessions();
  return {
    version:SESSIONS_VERSION,
    outputs:(Array.isArray(data.outputs)?data.outputs:[]).slice(0,LIMITS.outputs).map(cleanOutput).filter(Boolean),
    decks:(Array.isArray(data.decks)?data.decks:[]).slice(0,LIMITS.decks).map(cleanDeck).filter(Boolean)
  };
}
export const sessionsText=s=>JSON.stringify(readSessions(JSON.stringify(s)));

/** The status shown for an output: Excluded when the author left it out, Suggested when it was never made, Review needed when the guide has changed since, else Ready. */
export function outputStatus(output,paras){
  if(output.excluded)return {status:'Excluded',diff:null};
  if(!output.builtAt)return {status:'Suggested',diff:null};
  const diff=diffSources(output.sources,fingerprints(paras,output.kind,output.lessonIndex===null?{}:{lessonIndex:output.lessonIndex}));
  return {status:diff.same?'Ready':'Review needed',diff};
}
const nextId=(list,prefix)=>{let n=list.length+1;while(list.some(x=>x.id===`${prefix}${n}`))n++;return `${prefix}${n}`;};
/** Records that an output was made now from the guide as it is: a new record, or the same one made again (same kind and name), so refreshing keeps its place. */
export function recordOutput(sessions,{kind,name,scope='The whole guide',lessonIndex=null,paras,now=()=>new Date().toISOString()}){
  const next=readSessions(JSON.stringify(sessions)),same=next.outputs.find(o=>o.kind===kind&&o.name===name&&o.scope===scope);
  const sources=fingerprints(paras,kind,lessonIndex===null?{}:{lessonIndex});
  const rec={id:same?same.id:nextId(next.outputs,'o'),kind,name,scope,lessonIndex,builtAt:now(),excluded:false,sources};
  if(same)next.outputs=next.outputs.map(o=>o===same?rec:o);
  else{
    if(next.outputs.length>=LIMITS.outputs)next.outputs=next.outputs.slice(0,LIMITS.outputs-1);
    next.outputs=[rec,...next.outputs];
  }
  return readSessions(JSON.stringify(next));
}
export const setExcluded=(sessions,outputId,excluded)=>({...sessions,outputs:sessions.outputs.map(o=>o.id===outputId?{...o,excluded:!!excluded}:o)});
export const removeOutput=(sessions,outputId)=>({...sessions,outputs:sessions.outputs.filter(o=>o.id!==outputId)});
/** A deck saved among the guide's own: replaced when it has the same id, else added. */
export function saveDeck(sessions,deck){
  const next=readSessions(JSON.stringify(sessions)),clean=cleanDeck(deck,next.decks.length);
  if(!clean||!clean.cards.length)throw new Error('A deck needs at least one card with words on its front.');
  const same=next.decks.find(d=>d.id===clean.id);
  if(same)next.decks=next.decks.map(d=>d===same?clean:d);
  else{
    if(next.decks.length>=LIMITS.decks)throw new Error(`A guide holds up to ${LIMITS.decks} decks. Remove one first.`);
    clean.id=nextId(next.decks,'d');next.decks=[...next.decks,clean];
  }
  return {sessions:readSessions(JSON.stringify(next)),id:clean.id};
}
export const removeDeck=(sessions,deckId)=>({...sessions,decks:sessions.decks.filter(d=>d.id!==deckId)});
