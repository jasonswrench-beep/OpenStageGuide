// Delivery versions on the guide: a shorter or longer session planned from the same guide. A version is a name, a target length, a start
// time and an ordered list of the lessons to teach. It is kept in the document's settings (adapter.versionMemory), so it travels with
// the file. Lessons have no ids in a guide, so a version names each lesson by its title (without its number, which changes when lessons
// are moved), and a title that is no longer in the guide is reported, not guessed at. The older tool kept plans in a course record.
import {readGuide,lessonMinutes} from './guide-model.mjs';
import {blocksOf,START_TIME} from './guide-agenda.mjs';
import {clockLabel} from './planning.mjs';

export const VERSIONS_KEY='osg.versions.v1';
export const MAX_VERSIONS=20,MAX_LESSONS=200,MAX_TARGET=10080;
const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();
export const lessonKey=t=>oneLine(t).toLowerCase();

// The versions as read back from the document: anything that is not whole is left out rather than half-used.
export function readVersions(text){
  let raw=null;
  try{raw=JSON.parse(String(text||'')||'null');}catch{raw=null;}
  const out=[];
  if(!raw||raw.v!==1||!Array.isArray(raw.versions))return out;
  for(const v of raw.versions){
    if(!v||typeof v!=='object')continue;
    const name=oneLine(v.name).slice(0,80),id=typeof v.id==='string'?v.id.slice(0,40):'';
    if(!name||!id||out.some(x=>x.id===id))continue;
    out.push({id,name,
      target:Number.isInteger(v.target)&&v.target>0&&v.target<=MAX_TARGET?v.target:0,
      start:START_TIME.test(String(v.start||''))?String(v.start):'09:00',
      lessons:Array.isArray(v.lessons)?v.lessons.filter(s=>typeof s==='string'&&oneLine(s)).map(s=>oneLine(s).slice(0,200)).slice(0,MAX_LESSONS):[]});
    if(out.length>=MAX_VERSIONS)break;
  }
  return out;
}
export const writeVersions=versions=>JSON.stringify({v:1,versions});

// Every lesson of the guide, in order.
export function lessonList(paras){
  return readGuide(paras).modules.flatMap(m=>m.lessons.map(l=>({key:lessonKey(l.title),title:l.title,text:l.text,minutes:lessonMinutes(l),module:m.text||m.title,moduleIndex:m.index,index:l.index,lesson:l})));
}

// A new version holds every lesson, in the guide's order, so the author takes lessons out rather than building from nothing.
export function newVersion(all,existing=[],{now=()=>Date.now()}={}){
  if(existing.length>=MAX_VERSIONS)throw new Error(`You have ${MAX_VERSIONS} versions, which is the most. Delete one first.`);
  let name='New version',n=1;
  while(existing.some(v=>lessonKey(v.name)===lessonKey(name)))name=`New version ${++n}`;
  return {id:`v${now().toString(36)}${existing.length}`,name,target:0,start:'09:00',lessons:all.slice(0,MAX_LESSONS).map(l=>l.title)};
}

// Takes a lesson in or out. A lesson put back goes where the guide has it, among those already in the version.
export function toggleLesson(version,all,title){
  const key=lessonKey(title),has=version.lessons.some(t=>lessonKey(t)===key);
  if(has)return {...version,lessons:version.lessons.filter(t=>lessonKey(t)!==key)};
  const order=all.map(l=>l.key),mine=order.indexOf(key);
  let at=version.lessons.length;
  for(let i=0;i<version.lessons.length;i++){const o=order.indexOf(lessonKey(version.lessons[i]));if(o>mine){at=i;break;}}
  const lessons=[...version.lessons];lessons.splice(at,0,oneLine(title));
  return {...version,lessons};
}
export function moveLesson(version,title,delta){
  const key=lessonKey(title),i=version.lessons.findIndex(t=>lessonKey(t)===key),j=i+delta;
  if(i<0||j<0||j>=version.lessons.length)return version;
  const lessons=[...version.lessons];[lessons[i],lessons[j]]=[lessons[j],lessons[i]];
  return {...version,lessons};
}

// The version's agenda: its lessons in its order, with clock times from its start, the total and how far it is from the target.
export function versionAgenda(paras,version,{title='Guide'}={}){
  const m=START_TIME.exec(version.start||'09:00');
  if(!m)throw new Error('Choose a valid start time.');
  const begin=Number(m[1])*3600+Number(m[2])*60,all=lessonList(paras);
  const rows=[],missing=[];
  let cursor=begin,lastModule=null,lessons=0;
  for(const t of version.lessons){
    const l=all.find(x=>x.key===lessonKey(t));
    if(!l){missing.push(t);continue;}
    if(l.moduleIndex!==lastModule){rows.push({kind:'module',text:l.module,index:l.moduleIndex});lastModule=l.moduleIndex;}
    const first=cursor,last=cursor+l.minutes*60;
    rows.push({kind:'lesson',text:l.text,index:l.index,start:clockLabel(first),end:clockLabel(last),minutes:l.minutes});
    let at=first;
    for(const b of blocksOf(paras,l.lesson)){
      if(b.minutes<=0)continue;
      rows.push({kind:'block',text:`${b.label}${b.title?`: ${b.title}`:''}`,cue:b.cue,index:b.index,start:clockLabel(at),end:clockLabel(at+b.minutes*60),minutes:b.minutes});
      at+=b.minutes*60;
    }
    lessons++;cursor=last;
  }
  const minutes=(cursor-begin)/60;
  return {title:`${title} - ${version.name}`,start:clockLabel(begin),end:clockLabel(cursor),minutes,lessons,rows,notes:missing.map(t=>({text:`"${t}" is no longer in the guide, so it is left out of this version.`})),target:version.target||0,variance:version.target?minutes-version.target:null,missing};
}

const n=(c,w)=>`${c} ${c===1?w:w+'s'}`;
export function versionSummary(a){
  const base=a.lessons?`${n(a.lessons,'lesson')}, ${n(a.minutes,'minute')}, ${a.start} to ${a.end}.`:'No lessons are in this version yet. Tick the lessons to teach.';
  if(!a.target)return `${base} No target set.`;
  const v=a.variance;
  return `${base} ${v===0?`Exactly on the target of ${a.target} minutes.`:v>0?`${n(v,'minute')} over the target of ${a.target}.`:`${n(-v,'minute')} under the target of ${a.target}.`}`;
}
// A target as typed: nothing, or a whole number of minutes from 1 to 10080.
export function parseTarget(text){
  const t=oneLine(text);
  if(!t)return 0;
  if(!/^\d{1,5}$/.test(t)||Number(t)<1||Number(t)>MAX_TARGET)throw new Error(`Enter a target from 1 to ${MAX_TARGET} minutes, or leave it empty.`);
  return Number(t);
}
