// The facts about a lesson (slices K1 and K2 of docs/KIT-SPEC-2026-11-15.md): its purpose, materials and method, the range of minutes the facilitator may take, how
// the people work together, the size of the group, whether it may be left out, and whether the facilitator chooses some of a set. Each is one paragraph in the style
// OSG Fact ("Purpose: ..."), directly under the lesson's Time line, in a fixed order. They are for the facilitator: the participant workbook never has them.
// Pure functions over the guide's paragraphs and the facts the tool holds; nothing here touches a document.
import {STYLE_NAMES,STYLE_IDS} from './guide-styles.mjs';
import {para,run} from './guide-ooxml.mjs';

export const METHODS=Object.freeze(['Lecture','Discussion','Small-group work','Pair work','Individual work','Demonstration','Role play','Game','Case study','Video','Break']);
export const MODES=Object.freeze(['Whole group','Small groups','Pairs','Individual','Break']);
export const FACT_LIMITS=Object.freeze({words:400,group:60,minutes:1440,choose:99});

// The order the facts are written in, with the label each has.
const ORDER=Object.freeze(['purpose','materials','method','range','mode','group','newDay','optional','choose']);
const LABELS=Object.freeze({purpose:'Purpose',materials:'Materials',method:'Method',range:'Range',mode:'Interaction',group:'Group size',newDay:'New day',optional:'Optional',choose:'Choose'});
const KEY_OF=new Map(Object.entries(LABELS).map(([k,v])=>[v.toLowerCase(),k]));
const OPTIONAL_WORDS='This lesson may be left out.';
const clean=t=>String(t??'').replace(/[\r\n\u0007]+$/,'').replace(/\s+/g,' ').trim();

export const emptyFacts=()=>({purpose:'',materials:'',method:'',range:null,mode:'',group:'',newDay:'',optional:false,choose:null,chooseNote:'',extra:[]});
export const hasFacts=f=>!!(f&&(f.purpose||f.materials||f.method||f.range||f.mode||f.group||f.newDay||f.optional||f.choose||(f.extra&&f.extra.length)));
// "New day: 9:00" says a lesson starts a new day at that time. The clock of the agenda goes back to it there.
export const CLOCK=/^([01]?\d|2[0-3]):([0-5]\d)$/;

// "15 to 20 minutes" -> {min:15,max:20}; "2 of 5. Pick the ones that fit." -> {n:2,m:5}
const RANGE=/^(\d{1,4})\s*(?:to|-|–)\s*(\d{1,4})\s*(?:min(?:ute)?s?)?\.?$/i;
const CHOOSE=/^(\d{1,2})\s*of\s*(\d{1,2})\b[.:;,]?\s*(.*)$/i;

// The facts a lesson holds: from its Time line to the end of its header. `lesson` is readGuide's record (timeIndex, headerEnd).
// Returns {facts, span:{start,end}|null}; a label the guide does not know is kept, word for word, in `extra`, so saving the facts does not lose it.
export function readFacts(paras,lesson){
  const facts=emptyFacts();
  if(!lesson||lesson.timeIndex<0||lesson.headerEnd<=lesson.timeIndex)return {facts,span:null};
  const start=lesson.timeIndex+1,end=lesson.headerEnd;
  for(let i=start;i<=end;i++){
    const p=paras[i];
    if(String(p?.style||'')!==STYLE_NAMES.fact)continue;
    const text=clean(p.text),m=/^([A-Za-z][A-Za-z ]{0,24}):\s*(.*)$/.exec(text),key=m?KEY_OF.get(m[1].toLowerCase()):null;
    if(!key){facts.extra.push(text);continue;}
    const value=m[2].trim();
    if(key==='range'){const r=RANGE.exec(value);if(r)facts.range={min:Number(r[1]),max:Number(r[2])};else facts.extra.push(text);}
    else if(key==='newDay'){const c=CLOCK.exec(value);if(c)facts.newDay=`${Number(c[1])}:${c[2]}`;else facts.extra.push(text);}
    else if(key==='optional')facts.optional=true;
    else if(key==='choose'){const c=CHOOSE.exec(value);if(c){facts.choose={n:Number(c[1]),m:Number(c[2])};facts.chooseNote=c[3].trim();}else facts.extra.push(text);}
    else facts[key]=value;
  }
  return {facts,span:{start,end}};
}

// Throws, in plain words, what cannot be written; gives the clean facts back.
export function checkFacts(input){
  const f={...emptyFacts(),...input,extra:[...(input?.extra||[])]};
  for(const key of ['purpose','materials','chooseNote']){
    f[key]=clean(f[key]);
    if(f[key].length>FACT_LIMITS.words)throw new Error(`The ${key==='chooseNote'?'note about choosing':key} is too long: it holds up to ${FACT_LIMITS.words} letters.`);
  }
  f.method=clean(f.method);f.mode=clean(f.mode);f.group=clean(f.group);
  if(f.method&&!METHODS.includes(f.method)&&!f.methodKept)throw new Error(`Choose a method from the list (${METHODS.join(', ')}), or none.`);
  if(f.mode&&!MODES.includes(f.mode)&&!f.modeKept)throw new Error(`Choose how people work together from the list (${MODES.join(', ')}), or none.`);
  if(f.group.length>FACT_LIMITS.group)throw new Error(`The size of the group is too long: it holds up to ${FACT_LIMITS.group} letters.`);
  f.newDay=clean(f.newDay);
  if(f.newDay){const c=CLOCK.exec(f.newDay);if(!c)throw new Error('The time a new day starts is a clock time, such as 9:00 or 13:30.');f.newDay=`${Number(c[1])}:${c[2]}`;}
  if(f.range){
    const {min,max}=f.range;
    if(![min,max].every(n=>Number.isInteger(n)&&n>=1&&n<=FACT_LIMITS.minutes))throw new Error(`A range of minutes is two whole numbers from 1 to ${FACT_LIMITS.minutes}.`);
    if(max<=min)throw new Error('The most minutes must be more than the least.');
  }
  if(f.choose){
    const {n,m}=f.choose;
    if(![n,m].every(x=>Number.isInteger(x)&&x>=1&&x<=FACT_LIMITS.choose))throw new Error(`Choose n of m: two whole numbers from 1 to ${FACT_LIMITS.choose}.`);
    if(n>m)throw new Error('The facilitator cannot choose more than there are: n must not be more than m.');
  }
  return f;
}

const textOf=(key,f)=>{
  switch(key){
    case 'purpose':case 'materials':case 'method':case 'mode':case 'group':case 'newDay':return f[key];
    case 'range':return f.range?`${f.range.min} to ${f.range.max} minutes`:'';
    case 'optional':return f.optional?OPTIONAL_WORDS:'';
    case 'choose':return f.choose?`${f.choose.n} of ${f.choose.m}${f.chooseNote?`. ${f.chooseNote}`:'.'}`:'';
    default:return '';
  }
};
// The paragraphs for these facts, as Word-format XML: one for each fact that has words, in the fixed order, then any it did not know.
export function factsBlock(input){
  const f=checkFacts(input);
  const lines=ORDER.map(key=>[key,textOf(key,f)]).filter(([,t])=>t);
  return lines.map(([key,t])=>para(STYLE_IDS.fact,run(`${LABELS[key]}:`,STYLE_IDS.cueLabel)+run(` ${t}`))).join('')
    +f.extra.map(t=>para(STYLE_IDS.fact,run(t))).join('');
}
// The same lines as plain text, for tests and messages.
export function factsLines(input){
  const f=checkFacts(input);
  return [...ORDER.map(key=>[key,textOf(key,f)]).filter(([,t])=>t).map(([key,t])=>`${LABELS[key]}: ${t}`),...f.extra];
}

// What the Review says about the facts of the guide's lessons. `lessons` are readGuide's lessons with their minutes from lessonMinutes.
export function factIssues(paras,guide,minutesOf){
  const out=[];
  for(const mod of guide.modules)for(const lesson of mod.lessons){
    const {facts,span}=readFacts(paras,lesson);
    if(!span)continue;
    const name=lesson.text||'A lesson',at=span.start;
    if(facts.method&&!METHODS.includes(facts.method))out.push({level:'warn',rule:'LESSON-METHOD',message:`${name} has a method the guide does not know: "${facts.method}". The methods are ${METHODS.join(', ')}.`,paragraph:at});
    if(facts.mode&&!MODES.includes(facts.mode))out.push({level:'warn',rule:'LESSON-MODE',message:`${name} says people work together in a way the guide does not know: "${facts.mode}". The ways are ${MODES.join(', ')}.`,paragraph:at});
    if(facts.choose&&facts.choose.n>facts.choose.m)out.push({level:'warn',rule:'LESSON-CHOOSE',message:`${name} asks the facilitator to choose ${facts.choose.n} of ${facts.choose.m}, which cannot be done.`,paragraph:at});
    if(facts.range){
      const planned=minutesOf(lesson);
      if(planned<facts.range.min||planned>facts.range.max)out.push({level:'warn',rule:'LESSON-RANGE',message:`${name} is planned for ${planned} minutes, outside its range of ${facts.range.min} to ${facts.range.max}.`,paragraph:at});
    }
  }
  return out;
}

// ---- days, breaks and the stated length (kit slice K2)
// The days of the guide: a day starts at the first lesson, and again at each lesson whose facts say "New day". Each day has its lessons, its minutes and whether any block in it is a Break
// (or a lesson whose method is Break). `cueIdsOf(lesson)` gives the ids of the cues of the lesson's blocks.
export function dayPlan(paras,guide,minutesOf,cueIdsOf){
  const days=[];
  let current=null;
  for(const mod of guide.modules)for(const lesson of mod.lessons){
    const {facts}=readFacts(paras,lesson);
    if(!current||facts.newDay){current={start:facts.newDay||'',lessons:[],minutes:0,hasBreak:false,first:lesson};days.push(current);}
    current.lessons.push(lesson);
    current.minutes+=minutesOf(lesson);
    if(facts.method==='Break'||cueIdsOf(lesson).includes('break'))current.hasBreak=true;
  }
  return days;
}
export const LONG_DAY_MINUTES=180;
export function dayIssues(paras,guide,minutesOf,cueIdsOf){
  const out=[];
  for(const [n,day] of dayPlan(paras,guide,minutesOf,cueIdsOf).entries()){
    if(day.minutes>LONG_DAY_MINUTES&&!day.hasBreak)out.push({level:'warn',rule:'DAY-BREAK',message:`${days(n,day)} is ${hoursWords(day.minutes)} long with no break in it. Add a Break block with its minutes.`,paragraph:day.first.index});
  }
  return out;
}
const days=(n,day)=>day.start||n>0?`Day ${n+1}`:'The guide';
export function hoursWords(minutes){
  const h=Math.floor(minutes/60),m=minutes%60;
  return `${h?`${h} ${h===1?'hour':'hours'}`:''}${h&&m?' ':''}${m?`${m} ${m===1?'minute':'minutes'}`:''}`||'0 minutes';
}
// The guide says how long the programme is (in its settings) and its lessons add up to something else.
export function lengthIssue(totalMinutes,stated){
  if(!stated||!Number.isInteger(stated)||stated<1||stated===totalMinutes)return null;
  return {level:'warn',rule:'PROGRAM-LENGTH',message:`The guide says the programme is ${hoursWords(stated)}, but its lessons add up to ${hoursWords(totalMinutes)}.`};
}
