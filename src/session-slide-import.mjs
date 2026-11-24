// Bringing changes back from a presentation to the activities they came from (milestone 4 of the product spec of 2026-11-22: "Import slides, map proposed changes to existing
// activities, and accept a subset"). The slides the program makes (session-decks.mjs, pptx-export.mjs) hold an activity's task, grouping, time and what it makes on the slide, and the
// steps, supplies, set-up, debrief prompts, what goes wrong and other ways to take part in the speaker notes. This reads a presentation the same way and lists, for each slide it can
// match to an activity, only the parts that differ: the author ticks the ones to take, and nothing is written until then.
//
// Matching is by the activity's name, which each slide's title carries ("Scenario rehearsal (09:22, 18 min)"); a slide whose title was changed, and a slide that is not an
// activity's, are listed as not matched and are never applied. A change to a part that the guide has also changed since the slides were made is a conflict: it is listed,
// not ticked, and says so. Lines taken out of the notes of a slide are not taken out of the guide; only what the slide says is proposed.
import {activityBlocks,hashText} from './session-record.mjs';
import {readSteps,stepLine} from './session-text.mjs';
import {ACTIVITY_PARTS} from './guide-cues.mjs';
import {readBlock,blockSpan} from './guide-model.mjs';

const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();
const norm=t=>oneLine(t).toLowerCase();
const lines=t=>String(t??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
const LABELS=Object.fromEntries(ACTIVITY_PARTS.map(p=>[p.id,p.label]));
const TIME_SUFFIX=/\s*\((?:Day \d+ · )?\d{1,2}:\d{2}, \d+ min\)\s*$/;
const trimStop=t=>oneLine(t).replace(/[.\s]+$/,'');
// A line the slide writer had to shorten ends in an ellipsis and still matches the full line it was cut from.
const sameLine=(slide,guide)=>{const a=oneLine(slide),b=oneLine(guide);return a===b||(a.endsWith('…')&&b.startsWith(a.slice(0,-1).trimEnd()));};
const sameLines=(a,b)=>a.length===b.length&&a.every((l,i)=>sameLine(l,b[i]));

export const slideName=title=>oneLine(title).replace(TIME_SUFFIX,'');

/** What one slide says about its activity: the lines of the body and of the notes sorted into the parts they came from. */
export function readSlide(slide){
  const all=String(slide.text||'').split('\n').map(oneLine).filter(Boolean);
  const title=oneLine(slide.title)||all[0]||'';
  const body=all[0]===title?all.slice(1):all;
  const out={name:slideName(title),title,task:[],grouping:null,output:null,steps:[],materials:null,preparation:null,debrief:[],challenges:[],access:[]};
  for(const l of body){
    if(/^Grouping:\s*/i.test(l))out.grouping=trimStop(l.replace(/^Grouping:\s*/i,''));
    else if(/^Make:\s*/i.test(l))out.output=trimStop(l.replace(/^Make:\s*/i,''));
    else if(/^Time:\s/i.test(l))continue;
    else out.task.push(l);
  }
  for(const n of lines(slide.notes).map(oneLine)){
    let m;
    if((m=/^\d{1,2}:\d{2}\s+(\d+) min:\s*(.*)$/.exec(n)))out.steps.push(stepLine({min:Number(m[1]),text:m[2]}));
    else if((m=/^Supplies:\s*(.*)$/.exec(n)))out.materials=trimStop(m[1]);
    else if((m=/^Set up:\s*(.*)$/.exec(n)))out.preparation=oneLine(m[1]);
    else if((m=/^Debrief:\s*(.*)$/.exec(n)))out.debrief.push(m[1]);
    else if((m=/^If something goes wrong:\s*(.*)$/.exec(n)))out.challenges.push(m[1]);
    else if((m=/^Other ways to take part:\s*(.*)$/.exec(n)))out.access.push(m[1]);
  }
  return out;
}

/** The change a slide proposes for each part it speaks of, against the activity's parts now. */
function proposals(read,parts){
  const out=[],cur=id=>String(parts[id]??'');
  const propose=(id,to)=>out.push({part:id,label:LABELS[id],from:cur(id),to});
  if(read.task.length){
    const have=lines(cur('instructions'));
    if(!sameLines(read.task,have.slice(0,read.task.length))||have.length<read.task.length)propose('instructions',[...read.task,...have.slice(read.task.length)].join('\n'));
  }
  if(read.grouping!==null&&trimStop(cur('grouping'))!==read.grouping)propose('grouping',read.grouping);
  if(read.output!==null&&trimStop(cur('output'))!==read.output)propose('output',read.output);
  if(read.steps.length){
    const have=readSteps(cur('steps')).steps.map(stepLine);
    if(!sameLines(read.steps,have))propose('steps',read.steps.join('\n'));
  }
  if(read.materials!==null&&!sameLine(read.materials,trimStop(lines(cur('materials')).join(', '))))propose('materials',read.materials.split(/,\s*/).filter(Boolean).join('\n'));
  if(read.preparation!==null&&!sameLine(read.preparation,lines(cur('preparation')).join(' ')))propose('preparation',read.preparation);
  for(const [id,found] of [['debrief',read.debrief],['challenges',read.challenges],['access',read.access]])
    if(found.length&&!sameLines(found,lines(cur(id))))propose(id,found.join('\n'));
  return out;
}

/**
 * The plan for a presentation: `slides` is the deck as the program reads it ({order,title,text,notes}), `previous` the fingerprints the slides were made from (the "slides" output of
 * session-record.mjs), if the program made them. Returns {matched:[{slide,activity,changes}], unchanged, unmatched, changes}; each change has an id, the part, its words now and its
 * proposed words, and `conflict` when the guide's part has changed since the slides were made.
 */
export function planSlideImport(paras,slides,{previous=null}={}){
  const acts=activityBlocks(paras),used=new Set(),matched=[],unchanged=[],unmatched=[];
  for(const slide of slides){
    const read=readSlide(slide);
    const act=acts.find(a=>!used.has(a.key)&&norm(a.title)===norm(read.name));
    if(!act){unmatched.push({order:slide.order,title:oneLine(slide.title)||read.title});continue;}
    used.add(act.key);
    const was=previous?previous.find(s=>s.key===act.key):null;
    const changes=proposals(read,act.parts).map(c=>({...c,id:`${act.key}:${c.part}`,conflict:!!(was&&was.parts[c.part]!==undefined&&was.parts[c.part]!==hashText(c.from))}));
    if(changes.length)matched.push({slide:{order:slide.order,title:oneLine(slide.title)},activity:{key:act.key,title:act.title,lesson:act.lesson},changes});
    else unchanged.push(act.title);
  }
  return {matched,unchanged,unmatched,changes:matched.flatMap(m=>m.changes)};
}
export const planSummary=plan=>{
  const n=plan.changes.length,c=plan.changes.filter(x=>x.conflict).length;
  return [n?`${n} ${n===1?'change':'changes'} in ${plan.matched.length} ${plan.matched.length===1?'activity':'activities'}`:'Nothing in the presentation differs from the guide',
    c?`${c} of them also changed in the guide since the slides were made`:'',plan.unmatched.length?`${plan.unmatched.length} ${plan.unmatched.length===1?'slide was':'slides were'} not matched to an activity`:''].filter(Boolean).join('. ')+'.';
};

/**
 * What to write for one activity: where its block is now, the text of its first line (to be sure it is the same block), and the arguments the block is written from again, with the
 * accepted parts in place of the guide's. `overrides` is {part: words}. Throws, in words, when the activity is not there any more or is not written as parts.
 */
export function activityRewrite(paras,key,overrides){
  const a=activityBlocks(paras).find(x=>x.key===key);
  if(!a)throw new Error('That activity is not in the guide any more, so nothing was changed in it.');
  const span=blockSpan(paras,a.index),b=span&&readBlock(paras,span.start);
  if(!b||!b.parts)throw new Error(`${a.title} is not written as labelled parts, so the presentation's changes cannot be put into it. Edit it first.`);
  return {span,guard:b.firstText,args:{cue:'activity',text:b.text.split('\n')[0],minutes:b.minutes||undefined,audience:b.audience,parts:{...b.parts,...overrides},notes:b.notes}};
}
