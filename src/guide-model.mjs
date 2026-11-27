// Reads a guide from the document and works out what "Update guide" should change.
// There is no hidden record: the guide is whatever the document's styles and text say it is, so a hand edit,
// a copy and paste or Word's Undo can never put the pane out of step. All functions here are pure.
import {STYLE_NAMES} from './guide-styles.mjs';
import {ACK_HEADING,PERSPECTIVE_HEADING,APPENDIX_HEADING} from './guide-ooxml.mjs';
import {allCues,ACTIVITY_PARTS} from './guide-cues.mjs';
import {ITEM_TYPES,RATING_SCALES,itemTypeByKeyLabel,readNumericKey} from './guide-items.mjs';
import {formPlan} from './guide-forms.mjs';
import {factIssues,dayIssues,lengthIssue} from './guide-facts.mjs';
import {captionEdits} from './guide-reference.mjs';
import {readSteps,readSource,readNeedFrom,activityById} from './session-text.mjs';

const MODULE=/^Module\s+(\d+)\s*:\s*(.*)$/s;
const LESSON=/^Lesson\s+(\d+)\.(\d+)\s*:\s*(.*)$/s;
// A picture can leave a placeholder character at the start of a paragraph's text, so allow junk before the words.
const TIME=/^[^A-Za-z]*Time:\s*(\d+)\s+minutes?\.\s*Cumulative time:\s*(\d+)\s+minutes?\.?\s*$/;
const BLOCK_MINUTES=/\((\d+) min\)\s*$/;
const TRANSITION=/^[^A-Za-z]*Transition\b/;

export const kindOf=p=>{
  const s=p.style||'';
  if(s===STYLE_NAMES.guideTitle)return 'guide-title';
  if(s===STYLE_NAMES.module)return 'module';
  if(s===STYLE_NAMES.lesson)return 'lesson';
  if(s===STYLE_NAMES.time)return 'time';
  if(s===STYLE_NAMES.cue||s===STYLE_NAMES.cueFacilitator)return 'cue';
  if(s===STYLE_NAMES.body)return 'body';
  if(s===STYLE_NAMES.notes)return 'notes';
  if(s===STYLE_NAMES.assessment)return 'assessment';
  if(s===STYLE_NAMES.question)return 'question';
  if(s===STYLE_NAMES.option)return 'option';
  if(s===STYLE_NAMES.answerLine)return 'answer-line';
  if(s===STYLE_NAMES.answerKey)return 'answer-key';
  if(s===STYLE_NAMES.formTitle||s===STYLE_NAMES.formTitleFacilitator)return 'form-title';
  if(s===STYLE_NAMES.formText)return 'form-text';
  if(s===STYLE_NAMES.formLine)return 'form-line';
  if(s===STYLE_NAMES.formField)return 'form-field';
  if(s===STYLE_NAMES.formBox)return 'form-box';
  if(s===STYLE_NAMES.formSource)return 'form-source';
  if(s===STYLE_NAMES.formPrint)return 'form-print';
  if(s===STYLE_NAMES.formCard)return 'form-card';
  if(s===STYLE_NAMES.fact)return 'fact';
  if(s===STYLE_NAMES.term)return 'term';
  if(s===STYLE_NAMES.reference)return 'reference';
  if(s===STYLE_NAMES.caption)return 'caption';
  if(s===STYLE_NAMES.quote)return 'quote';
  if(s===STYLE_NAMES.quoteBy)return 'quote-by';
  if(s===STYLE_NAMES.tableHead)return 'table-head';
  if(s===STYLE_NAMES.tableCell)return 'table-cell';
  if(s===STYLE_NAMES.ratingBox)return 'rating-box';
  if(s===STYLE_NAMES.itemNote)return 'item-note';
  if(s===STYLE_NAMES.overview)return 'overview';
  if(s===STYLE_NAMES.keyTableHead)return 'key-table-head';
  if(s===STYLE_NAMES.keyTableCell)return 'key-table-cell';
  if(p.builtIn==='Title'||/^title$/i.test(s))return 'title';
  return 'other';
};
const clean=t=>String(t??'').replace(/[\r\n\u0007]+$/,'').trim();
export const lessonMinutes=lesson=>lesson.blockMinutes>0?lesson.blockMinutes:(lesson.timeMinutes??0);

// paras: [{text,style,builtIn}] in document order. Indexes in the result are positions in that list.
export function readGuide(paras){
  const guide={isGuide:false,title:'',modules:[]};
  let mod=null,lesson=null;
  const ensureModule=()=>{
    if(!mod){mod={index:-1,text:'',title:'',numbered:false,lessons:[],hasTransition:false,transitionIndex:-1,end:-1};guide.modules.push(mod);}
    return mod;
  };
  paras.forEach((p,index)=>{
    const text=clean(p.text);
    switch(kindOf(p)){
      // Only the guide's own title style marks a document as a guide; an ordinary Title paragraph is just a fallback name.
      case 'guide-title':guide.isGuide=true;guide.title=text;break;
      case 'title':if(!guide.title)guide.title=text;break;
      case 'module':{
        guide.isGuide=true;
        const m=MODULE.exec(text);
        mod={index,text,title:m?m[2].trim():text,numbered:!!m,lessons:[],hasTransition:false,transitionIndex:-1,end:index};
        guide.modules.push(mod);lesson=null;break;
      }
      case 'lesson':{
        guide.isGuide=true;ensureModule();
        const m=LESSON.exec(text);
        lesson={index,text,title:m?m[3].trim():text,numbered:!!m,timeIndex:-1,headerEnd:-1,timeMinutes:null,timeCumulative:null,blockMinutes:0,cues:0,end:index};
        mod.lessons.push(lesson);break;
      }
      case 'time':{
        if(lesson&&lesson.timeIndex<0){
          lesson.timeIndex=index;lesson.headerEnd=index;
          const m=TIME.exec(text);
          if(m){lesson.timeMinutes=Number(m[1]);lesson.timeCumulative=Number(m[2]);}
        }
        break;
      }
      // The facts about a lesson (purpose, materials, method, ...) sit straight under its time line; where they end is where its blocks begin.
      case 'fact':{
        if(lesson&&lesson.headerEnd>=0&&index===lesson.headerEnd+1)lesson.headerEnd=index;
        break;
      }
      case 'cue':{
        // A Transition leads out of the module, so it is the module's block and not part of any one lesson.
        if(mod&&TRANSITION.test(text)){mod.hasTransition=true;if(mod.transitionIndex<0)mod.transitionIndex=index;break;}
        if(lesson){lesson.cues++;const m=BLOCK_MINUTES.exec(text);if(m)lesson.blockMinutes+=Number(m[1]);}
        break;
      }
      default:break;
    }
  });
  // Where each lesson and module ends, so new content can go after a whole lesson, never into the middle of one.
  // A lesson runs to its last block before the next heading or Transition; text the author wrote after that is theirs.
  const isStop=i=>{
    const p=paras[i],k=kindOf(p);
    return k==='module'||k==='lesson'||k==='assessment'||k==='guide-title'||k==='title'||/^heading\b/i.test(String(p.style||''))||/^Heading/.test(String(p.builtIn||''))||(k==='cue'&&TRANSITION.test(clean(p.text)));
  };
  for(const m of guide.modules){
    let end=m.index;
    for(const l of m.lessons){
      let last=l.index;
      for(let i=l.index+1;i<paras.length&&!isStop(i);i++)if(['time','fact','cue','body','notes'].includes(kindOf(paras[i])))last=i;
      l.end=last;end=Math.max(end,last);
    }
    if(m.transitionIndex>=0)end=Math.max(end,blockSpan(paras,m.transitionIndex).end);
    m.end=end;
  }
  return guide;
}

// The paragraphs that make up the block the cursor is in: the cue paragraph, any further lines and its notes.
export function blockSpan(paras,at){
  const k=i=>kindOf(paras[i]||{}),part=i=>k(i)==='body'||k(i)==='notes';
  if(k(at)!=='cue'&&!part(at))return null;
  let start=at;
  while(k(start)!=='cue'&&start>0&&(part(start-1)||k(start-1)==='cue'))start--;
  let end=at;
  while(part(end+1))end++;
  return {start,end};
}

// Where new content goes, so it never lands in the middle of another block. All return a paragraph index to insert after.
export function blockAnchor(paras,at){
  const span=blockSpan(paras,at);
  if(span){
    // On a module's Transition, a new block still belongs inside the module, before the Transition: in the last lesson, or, when the
    // module has no lessons, straight after what comes before the Transition. After it the block would belong to no module at all,
    // and the next module the author adds would then land above it (seen in the author's own recording, 2026-11-08).
    const mod=readGuide(paras).modules.find(m=>m.index>=0&&m.transitionIndex===span.start);
    if(mod)return mod.lessons.length?mod.lessons.at(-1).end:Math.max(span.start-1,mod.index);
    return span.end;
  }
  if(kindOf(paras[at]||{})==='lesson'){
    const lesson=readGuide(paras).modules.flatMap(m=>m.lessons).find(l=>l.index===at);
    if(lesson&&lesson.timeIndex===at+1)return lesson.headerEnd;
  }
  const header=headerEndAt(paras,at);
  if(header>=0)return header;
  return Math.max(at,0);
}
// On a lesson's Time line, or on one of the facts under it: the last fact (or the Time line itself when there are no facts), where whatever is added to the lesson goes.
// Between the time line and the facts it would hide the facts from Lesson details and from the lists made from them (found in the alpha 7 manual, 2026-11-15). -1 elsewhere.
export function headerEndAt(paras,at){
  const lesson=readGuide(paras).modules.flatMap(m=>m.lessons).find(l=>l.timeIndex>=0&&at>=l.timeIndex&&at<=l.headerEnd);
  return lesson?lesson.headerEnd:-1;
}
export function transitionAnchor(paras,at){
  const modules=readGuide(paras).modules.filter(m=>m.index>=0),inside=modules.filter(m=>m.index<=at);
  if(!inside.length)throw new Error('Put your cursor inside a module first.');
  const mod=inside.at(-1);
  if(mod.hasTransition)throw new Error(`${mod.text||'This module'} already has a Transition block. Select it in the document, then choose Edit this block.`);
  return {anchor:mod.end,module:mod};
}
// The last paragraph of what a module's overview put under its heading (its labelled lines and its lessons table), or the heading
// itself when there is none. Text the author typed among them does not move it, and it stops at the first lesson, block or heading.
export function moduleIntroEnd(paras,modIndex){
  let last=modIndex;
  for(let i=modIndex+1;i<paras.length;i++){
    const p=paras[i],k=kindOf(p);
    if(k==='overview'||k==='table-head'||k==='table-cell')last=i;
    else if(k!=='other'||/^heading\b/i.test(String(p.style||''))||/^Heading/.test(String(p.builtIn||'')))break;
  }
  return last;
}
export function lessonPlacement(paras,at){
  const guide=readGuide(paras),modules=guide.modules.filter(m=>m.index>=0),inside=modules.filter(m=>m.index<=at);
  if(!inside.length)throw new Error('Put your cursor inside a module first, or add a module.');
  const mod=inside.at(-1),m=modules.indexOf(mod)+1;
  const started=mod.lessons.filter(l=>l.index<=at),after=started.length?started.at(-1):null;
  // Cursor in or after a lesson: the new one follows it. Cursor on the module heading: the new one joins the end of the module.
  const prev=after||mod.lessons.at(-1)||null,l=(after?started.length:mod.lessons.length)+1;
  const anchor=prev?prev.end:moduleIntroEnd(paras,mod.index);
  const before=guide.modules.flatMap(x=>x.lessons).filter(x=>x.index<=anchor).reduce((s,x)=>s+lessonMinutes(x),0);
  return {anchor,m,l,before,after:prev?prev.text:mod.text};
}
export function modulePlacement(paras,at){
  const guide=readGuide(paras),modules=guide.modules.filter(m=>m.index>=0);
  if(!modules.length){
    const gi=paras.findIndex(p=>/^heading 2$/i.test(String(p.style||''))&&/^training at a glance\s*$/i.test(clean(p.text)));
    const prompt=gi>=0&&kindOf(paras[gi+1]||{})==='other'&&!/^heading/i.test(String(paras[gi+1].style||''))?gi+1:gi;
    return {anchor:prompt>=0?prompt:Math.max(at,0),n:1};
  }
  const inside=modules.filter(m=>m.index<=at),mod=inside.length?inside.at(-1):modules.at(-1);
  return {anchor:moduleTailEnd(paras,mod),n:modules.indexOf(mod)+2};
}
// A module's own last paragraph is its Transition (or last lesson). Guide blocks that sit straight after it, such as slides put in while the
// cursor was on the Transition, still belong to that module: the next module goes after them, not above them. Plain text the author wrote
// there is theirs and does not move the module's end.
function moduleTailEnd(paras,mod){
  let end=mod.end;
  while(end+1<paras.length&&['cue','body','notes'].includes(kindOf(paras[end+1])))end++;
  return end;
}

// Which cue a cue paragraph is: its text starts with the cue's label (after the placeholder a picture can leave), and the
// longest label wins, so "Time on task" is not read as something shorter.
export const cueFromText=text=>{
  const first=clean(text).replace(/^[^A-Za-z]+/,'');
  return [...allCues()].sort((a,b)=>b.label.length-a.label.length).find(c=>first.toLowerCase().startsWith(c.label.toLowerCase()))||null;
};

// Reads a block back into the form's fields so it can be edited.
export function readBlock(paras,at){
  const span=blockSpan(paras,at);
  if(!span)return null;
  const first=clean(paras[span.start].text).replace(/^[^A-Za-z]+/,'');
  const cue=cueFromText(paras[span.start].text);
  if(!cue)return null;
  let rest=first.slice(cue.label.length).replace(/^\s*\t?\s*/,'');
  let minutes='';
  const m=BLOCK_MINUTES.exec(rest);
  if(m){minutes=m[1];rest=rest.slice(0,m.index).trim();}
  const audience=paras[span.start].style===STYLE_NAMES.cueFacilitator?'facilitator':'everyone',firstText=clean(paras[span.start].text);
  // A picture is a paragraph of the block's body style with no words in it. The pane cannot rewrite it, so it says there is one.
  const hasPicture=paras.slice(span.start+1,span.end+1).some(isPictureParagraph);
  // An activity written as labelled parts reads back as its title, its parts and its own facilitator notes.
  const labelled=cue.id==='activity'?readActivityParts(paras,span):null;
  if(labelled)return {span,cue:cue.id,text:rest,notes:labelled.notes,parts:labelled.parts,minutes,audience,firstText,hasPicture};
  // A slide's own text sits under its picture in the shaded notes style, labelled "Slide text:", before its "Facilitator notes:".
  // The two are kept apart so editing a block's notes never swallows the slide's text (and the other way round).
  const lines=[rest],notes=[],slideText=[];
  let into=notes;
  for(let i=span.start+1;i<=span.end;i++){
    const text=clean(paras[i].text);
    if(kindOf(paras[i])==='notes'){
      if(/^Slide text:\s*/i.test(text)){into=slideText;into.push(text.replace(/^Slide text:\s*/i,''));}
      else if(/^Facilitator notes:\s*/i.test(text)){into=notes;into.push(text.replace(/^Facilitator notes:\s*/i,''));}
      else into.push(text);
    }
    else if(!isPictureParagraph(paras[i]))lines.push(text);
  }
  return {span,cue:cue.id,text:lines.join('\n'),notes:notes.join('\n'),slideText:slideText.join('\n'),minutes,audience,firstText,hasPicture};
}
// Word leaves a placeholder character, or nothing, where a picture is; there are no letters or digits in it.
export const isPictureParagraph=p=>kindOf(p)==='body'&&!/[\p{L}\p{N}]/u.test(String(p.text??''));

// The paragraphs after an activity's first line, read as labelled parts: each starts with its label (in its own style: the
// shaded facilitator-notes style for materials and preparation, the plain block style for the rest) and may run on in further
// paragraphs of that style. Anything else makes it an ordinary block (null), as does having no instructions.
function readActivityParts(paras,span){
  const parts={},notes=[];
  let current=null;
  const facilitatorPart=id=>ACTIVITY_PARTS.find(p=>p.id===id).for==='facilitator';
  for(let i=span.start+1;i<=span.end;i++){
    const text=clean(paras[i].text),isNotes=kindOf(paras[i])==='notes',m=/^([^:]{1,30}):\s*(.*)$/s.exec(text);
    const def=m&&ACTIVITY_PARTS.find(p=>p.label.toLowerCase()===m[1].trim().toLowerCase()&&(p.for==='facilitator')===isNotes);
    // A part's label may stand alone, with its lines (bulleted in Word) in the paragraphs under it.
    if(def){current=def.id;parts[current]=m[2]?[m[2]]:[];}
    else if(isNotes&&/^Facilitator notes:/i.test(text)){current='notes';notes.push(text.replace(/^Facilitator notes:\s*/i,''));}
    else if(current==='notes'&&isNotes)notes.push(text);
    else if(current&&current!=='notes'&&facilitatorPart(current)===isNotes)parts[current].push(text);
    else return null;
  }
  if(!parts.instructions||!parts.instructions.length)return null;
  return {parts:Object.fromEntries(Object.entries(parts).map(([k,v])=>[k,v.join('\n')])),notes:notes.join('\n')};
}

// ---- a block with an icon and a label and nothing else: added in place (Insert > Block...) and not yet written, or written and then emptied.
const PART_LABEL=/^(Objective|Materials|Preparation|Instructions|Case study|Facilitator notes|Slide text):\s*/i;
const hasWords=text=>/[\p{L}\p{N}]/u.test(String(text??''));
/**
 * Every block with no words in it: nothing after its label (the minutes it was given do not count as words, but a block that is only its
 * minutes, a Break of 15 minutes, says something and is left alone), no lines under it, no picture, and no words in any labelled part.
 * Each is `{index, cue, label}`, the paragraph of its first line.
 */
export function emptyBlocks(paras){
  const out=[];
  for(let i=0;i<paras.length;i++){
    if(kindOf(paras[i])!=='cue')continue;
    const cue=cueFromText(paras[i].text);
    if(!cue)continue;
    const span=blockSpan(paras,i);
    if(!span||span.start!==i)continue;
    const line=clean(paras[i].text),first=line.replace(/^[^A-Za-z]+/,'').slice(cue.label.length);
    const minutes=BLOCK_MINUTES.test(first);
    if(hasWords(first.replace(BLOCK_MINUTES,'')))continue;
    let anything=minutes;
    for(let j=span.start+1;j<=span.end&&!anything;j++)if(isPictureParagraph(paras[j])||hasWords(clean(paras[j].text).replace(PART_LABEL,'')))anything=true;
    if(!anything)out.push({index:i,cue:cue.id,label:cue.label});
  }
  return out;
}

// What the cursor is in, for the "selected" line on the Build screen.
export function describeAt(paras,at){
  const block=readBlock(paras,at);
  if(block)return {type:'block',block,label:`${cueLabel(block.cue)} block`};
  const item=readItem(paras,at);
  if(item)return {type:'item',item,label:item.label};
  const guide=readGuide(paras);
  for(const mod of guide.modules){
    if(mod.index===at)return {type:'module',module:mod,label:mod.text};
    for(const l of mod.lessons)if(l.index===at||l.timeIndex===at)return {type:'lesson',lesson:l,module:mod,label:l.text};
  }
  return null;
}
const cueLabel=id=>(allCues().find(c=>c.id===id)||{label:'Cue'}).label;

// What "Update guide" would change: heading numbers, each lesson's time block, question numbers, and the numbers and citations of forms.
// `options.forms` is the guide's form settings (guide-forms.mjs readSettings); without it forms are numbered the automatic way.
// Cumulative time is the minutes taught before the lesson starts, so a facilitator
// can see at the top of a lesson where the clock should stand.
export function updatePlan(paras,options={}){
  const guide=readGuide(paras),edits=[];
  let cumulative=0,m=0;
  for(const mod of guide.modules){
    if(mod.index>=0){
      m++;
      const want=`Module ${m}: ${mod.title}`;
      if(mod.numbered&&clean(paras[mod.index].text)!==want)edits.push({kind:'heading',index:mod.index,text:want});
    }
    let l=0;
    for(const lesson of mod.lessons){
      l++;
      const want=`Lesson ${m}.${l}: ${lesson.title}`;
      if(lesson.numbered&&clean(paras[lesson.index].text)!==want)edits.push({kind:'heading',index:lesson.index,text:want});
      const minutes=lessonMinutes(lesson);
      if(lesson.timeIndex>=0&&(lesson.timeMinutes!==minutes||lesson.timeCumulative!==cumulative))edits.push({kind:'time',index:lesson.timeIndex,minutes,cumulative});
      cumulative+=minutes;
    }
  }
  // Question numbers restart in each assessment section. Items without a "Question N." label (open comments) are not counted.
  const assess=readAssessments(paras);
  for(const section of assess.sections){
    let n=0;
    for(const item of section.items){
      if(item.number===null)continue;
      n++;
      if(item.number!==n)edits.push({kind:'question',index:item.index,from:`Question ${item.number}.`,to:`Question ${n}.`});
    }
  }
  const forms=formPlan(paras,options.forms||{});
  edits.push(...forms.edits);
  edits.push(...captionEdits(paras));
  // The ids of the cues of a lesson's blocks, so a day with no break in it can be found.
  const cueIdsOf=l=>{const ids=[];for(let i=l.index+1;i<=l.end;i++)if(kindOf(paras[i])==='cue'){const c=cueFromText(paras[i].text);if(c)ids.push(c.id);}return ids;};
  const length=lengthIssue(cumulative,options.forms?.time?.minutes);
  // An Expected answer is what to listen for in the answers to an Ask: with no Ask before it in its lesson (the nearest block before it that is not another Expected answer),
  // it listens for nothing (the spec's note of kit slice K3).
  const expectedIssues=[];
  for(const mod of guide.modules)for(const l of mod.lessons){
    let before=null;
    for(let i=l.index+1;i<=l.end;i++){
      if(kindOf(paras[i])!=='cue')continue;
      const c=cueFromText(paras[i].text);
      if(!c)continue;
      if(c.id==='expected-answer'&&before!=='ask')expectedIssues.push({level:'warn',rule:'CUE-EXPECTED',message:`An Expected answer in ${l.text||'a lesson'} has no Ask before it, so it has no question to listen for answers to. Put it right after the Ask it belongs to.`,paragraph:i});
      if(c.id!=='expected-answer')before=c.id;
    }
  }
  // A block that was added (Insert > Block...) and never written: its icon and label, and nothing else.
  const lessonAt=i=>guide.modules.flatMap(m=>m.lessons).find(l=>i>l.index&&i<=l.end);
  const emptyIssues=emptyBlocks(paras).map(b=>({level:'warn',rule:'BLOCK-EMPTY',message:`A ${b.label} block${lessonAt(b.index)?` in ${lessonAt(b.index).text||'a lesson'}`:''} has no words yet. Type its words in the guide, or delete the block.`,paragraph:b.index}));
  const facts={issues:[...factIssues(paras,guide,lessonMinutes),...dayIssues(paras,guide,lessonMinutes,cueIdsOf),...(length?[length]:[]),...expectedIssues,...emptyIssues,...activityIssues(paras,guide)]};
  return {guide,assess,forms,facts,edits,totalMinutes:cumulative};
}

// What the Review says about Activity blocks written as parts (the session templates, milestone 1 of the product spec): timed steps that do not add up to the activity's minutes
// (a shorter block with unchanged steps is found the same way), a step with no minutes, a need that names an earlier activity that is not there any more, and an activity made from a
// version of the catalog that has since been replaced.
const normTitle=t=>String(t??'').replace(/\s+/g,' ').trim().toLowerCase();
export function activityIssues(paras,guide){
  const out=[],earlier=new Set();
  for(const mod of guide.modules)for(const lesson of mod.lessons){
    for(let i=lesson.index+1;i<=lesson.end;i++){
      if(kindOf(paras[i])!=='cue')continue;
      const cue=cueFromText(paras[i].text);
      if(!cue||cue.id!=='activity')continue;
      const span=blockSpan(paras,i);
      if(!span||span.start!==i)continue;
      i=span.end;
      const b=readBlock(paras,span.start);
      if(!b)continue;
      const title=String(b.text).split('\n')[0].trim()||'An activity',at=span.start;
      if(b.parts){
        if(b.parts.steps){
          const s=readSteps(b.parts.steps),planned=Number(b.minutes)||0;
          if(s.without.length)out.push({level:'warn',rule:'ACTIVITY-STEP-MINUTES',message:`A timed step of "${title}" gives no minutes, so its steps cannot be added up. Write each one as "3 min: what happens".`,paragraph:at});
          else if(planned&&s.steps.length&&s.minutes!==planned)out.push({level:'warn',rule:'ACTIVITY-STEPS',message:`The timed steps of "${title}" add up to ${s.minutes} minutes, but the activity is ${planned} minutes. Change a step or the activity's minutes so that they agree.`,paragraph:at});
        }
        const src=b.parts.source?readSource(b.parts.source):null,made=src&&activityById(src.id);
        if(made&&made.version>src.version)out.push({level:'warn',rule:'ACTIVITY-SOURCE',message:`"${title}" was made from version ${src.version} of the catalog activity "${made.title}", and the catalog now has version ${made.version}. Put the activity in again to see what changed.`,paragraph:at});
        for(const line of String(b.parts.needs||'').split(/\r?\n/)){
          const from=readNeedFrom(line);
          if(from&&!earlier.has(normTitle(from)))out.push({level:'warn',rule:'ACTIVITY-NEEDS',message:`"${title}" needs the output of "${from}", but no activity with that name comes before it. It may have been renamed, moved after this one or removed. Choose another earlier activity, or take the line out.`,paragraph:at});
        }
      }
      earlier.add(normTitle(title));
    }
  }
  return out;
}

// Whether the contents table lists the headings the guide has. Page numbers cannot be told from here, so a table
// that lists every heading can still have old page numbers; Update guide refreshes it either way.
const alnum=s=>String(s??'').replace(/[^\p{L}\p{N}]+/gu,'').toLowerCase();
// When it is out of date, `why` says in what way (how many lines against how many headings, or the first line that differs), so a
// table that the pane calls out of date while it looks right in Word can be told apart from one that really is.
const shown=t=>{const x=String(t??'').replace(/[\r\n\u0007]+$/,'').replace(/\s+/g,' ').trim();return x.length>40?`${x.slice(0,40)}...`:x;};
export function contentsState(paras){
  const entryParas=paras.filter(p=>/^toc [1-9]$/i.test(String(p.style||'')));
  // Word's pane API reports the table's first line, which holds the start of the field, with no style name at all (seen 2026-11-03), so a line
  // with no style that ends in a page number and sits just above the first line that does have a table style is counted as the first line.
  const firstAt=entryParas.length?paras.indexOf(entryParas[0]):-1,before=firstAt>0?paras[firstAt-1]:null;
  if(before&&!String(before.style||'').trim()&&/[	 ]\d+\s*$/.test(String(before.text||'')))entryParas.unshift(before);
  const entries=entryParas.map(p=>alnum(p.text));
  const headParas=paras.filter(p=>/^(heading [12]|osg module|osg lesson|osg assessment)$/i.test(String(p.style||''))&&alnum(p.text)),heads=headParas.map(p=>alnum(p.text));
  if(!entries.length){
    const placeholder=paras.some(p=>/update this table to list the modules and lessons/i.test(String(p.text||'')));
    return {present:placeholder,stale:placeholder};
  }
  const at=heads.findIndex((h,i)=>!(entries[i]||'').startsWith(h));
  // The two paragraphs just above the table's first line, with their styles: a line Word does not report in a table style shows up here.
  const first=paras.indexOf(entryParas[0]),above=[first-2,first-1].filter(i=>i>=0).map(i=>`"${shown(paras[i].text)}" (${paras[i].style})`).join(' and ');
  if(entries.length!==heads.length)return {present:true,stale:true,why:`The table has ${entries.length} lines and the guide has ${heads.length} headings. Just above its first line are ${above}.${at>=0?` The first heading without a matching line is "${shown(headParas[at].text)}"${entryParas[at]?`, where the table has "${shown(entryParas[at].text)}"`:''}.`:''}`};
  if(at>=0)return {present:true,stale:true,why:`Line ${at+1} of the table reads "${shown(entryParas[at].text)}" where the heading is "${shown(headParas[at].text)}".`};
  return {present:true,stale:false};
}

// What to tell the author about the guide itself. Each issue knows its paragraph so the pane can go there.
export function guideIssues(plan){
  const {guide,edits}=plan,out=[];
  if(!guide.isGuide)return out;
  if(!guide.title)out.push({level:'warn',rule:'GUIDE-TITLE',message:'The guide has no title. Use the Title style for the first line.'});
  for(const mod of guide.modules){
    const where=mod.index>=0?mod.index:undefined,name=mod.text||'Lessons outside a module';
    if(mod.index<0)out.push({level:'warn',rule:'LESSON-NO-MODULE',message:'Some lessons come before any module.',paragraph:guide.modules[0]===mod&&mod.lessons[0]?mod.lessons[0].index:undefined});
    if(mod.index>=0&&!mod.numbered)out.push({level:'warn',rule:'MODULE-NAME',message:`"${name}" is not written as "Module N: Title", so it is not renumbered.`,paragraph:where});
    if(mod.index>=0&&!mod.lessons.length)out.push({level:'warn',rule:'MODULE-EMPTY',message:`${name} has no lessons yet.`,paragraph:where});
    if(mod.index>=0&&!mod.hasTransition)out.push({level:'warn',rule:'MODULE-TRANSITION',message:`${name} has no Transition block to lead into the next module.`,paragraph:where});
    for(const lesson of mod.lessons){
      if(lesson.timeIndex<0)out.push({level:'fail',rule:'LESSON-TIME',message:`${lesson.text||'A lesson'} has no time block.`,paragraph:lesson.index});
      else if(lesson.timeMinutes===null)out.push({level:'warn',rule:'LESSON-TIME-TEXT',message:`The time block under ${lesson.text} does not read "Time: N minutes. Cumulative time: M minutes."`,paragraph:lesson.timeIndex});
      if(!lesson.numbered)out.push({level:'warn',rule:'LESSON-NAME',message:`"${lesson.text}" is not written as "Lesson M.N: Title", so it is not renumbered.`,paragraph:lesson.index});
      if(!lesson.cues)out.push({level:'warn',rule:'LESSON-EMPTY',message:`${lesson.text||'A lesson'} has no blocks yet.`,paragraph:lesson.index});
    }
  }
  for(const section of plan.assess.sections){
    if(!section.items.length)out.push({level:'warn',rule:'ASSESS-EMPTY',message:`${section.title} has no items yet.`,paragraph:section.index});
    for(const item of section.items)
      if(item.options&&!item.hasKey)out.push({level:'warn',rule:'ITEM-KEY',message:`${item.label} has no answer key.`,paragraph:item.index});
  }
  out.push(...(plan.forms?.issues||[]));
  out.push(...(plan.facts?.issues||[]));
  if(edits.length)out.push({level:'warn',rule:'OUT-OF-DATE',message:`${edits.length} number${edits.length===1?'':'s'} or time${edits.length===1?' is':'s are'} out of date. Choose Update guide.`,paragraph:edits[0].index});
  return out;
}

// A flat list for the Contents on the Build screen.
export function outline(plan){
  const rows=[];
  let cumulative=0;
  for(const mod of plan.guide.modules){
    if(mod.index>=0)rows.push({kind:'module',index:mod.index,text:mod.text});
    for(const lesson of mod.lessons){
      const minutes=lessonMinutes(lesson);
      rows.push({kind:'lesson',index:lesson.index,text:lesson.text,minutes,cumulative});
      cumulative+=minutes;
    }
  }
  for(const section of plan.assess.sections)rows.push({kind:'assessment',index:section.index,text:section.title,items:section.items.length});
  return rows.sort((a,b)=>a.index-b.index);
}

// ---- Assessments: sections of items, read back from the document's styles like everything else.
// The paragraphs that belong to an item, the first being its question. A table's cells are paragraphs too.
const ITEM_KINDS=['question','option','answer-line','answer-key','table-head','table-cell','rating-box','item-note','key-table-head','key-table-cell'];
const KEY_LABEL=/^[^A-Za-z]*Answer key \(([^)]+)\):\s*(.*)$/s;
const QUESTION_LABEL=/^[^A-Za-z]*Question\s+(\d+)\.\s*(.*)$/s;
const softBreaks=t=>String(t??'').replace(/[\u000b\r]+/g,'\n');
const endsSection=(paras,i)=>{
  const p=paras[i],k=kindOf(p);
  return k==='module'||k==='lesson'||k==='assessment'||k==='guide-title'||k==='title'||/^heading\b/i.test(String(p.style||''))||/^Heading/.test(String(p.builtIn||''));
};

// The paragraphs of the item that starts at `start` (a question): its options, answer lines and key.
export function itemSpan(paras,at){
  const k=i=>kindOf(paras[i]||{});
  if(!ITEM_KINDS.includes(k(at)))return null;
  let start=at;
  while(k(start)!=='question'){if(start>0&&ITEM_KINDS.includes(k(start-1)))start--;else return null;}
  let end=start;
  while(ITEM_KINDS.slice(1).includes(k(end+1)))end++;
  return {start,end};
}

function summarizeItem(paras,span){
  const k=i=>kindOf(paras[i]);
  const q=QUESTION_LABEL.exec(clean(paras[span.start].text));
  const number=q?Number(q[1]):null;
  let key=null,options=0,boxes=0,notes=0,firstHead='';
  for(let i=span.start+1;i<=span.end;i++){
    if(k(i)==='option')options++;
    if(k(i)==='rating-box')boxes++;
    if(k(i)==='item-note')notes++;
    if(k(i)==='table-head'&&!firstHead)firstHead=clean(paras[i].text);
    if(!key&&k(i)==='answer-key')key=KEY_LABEL.exec(clean(paras[i].text));
  }
  // No key and no options is an open comment; check boxes with no key are rating statements, or a semantic differential when
  // the table's first column is headed "Left word"; labelled actions with no key are an action plan. Options with no key are read
  // as multiple choice, so they can still be fixed.
  const guess=boxes?(/^Left word/i.test(firstHead)?'semantic-differential':'rating-statements'):notes?'action-plan':options?'multiple-choice':'open-comment';
  const type=key?itemTypeByKeyLabel(key[1]):ITEM_TYPES.find(t=>t.id===guess);
  const label=number!==null?`Question ${number}${type&&type.keyLabel?` (${type.keyLabel})`:''}`:(type?type.label:'Item');
  return {index:span.start,end:span.end,number,type:type?type.id:null,hasKey:!!key,options:options>0,label};
}

export function readAssessments(paras){
  const sections=[];
  paras.forEach((p,index)=>{if(kindOf(p)==='assessment')sections.push({index,title:clean(p.text),items:[],end:index});});
  for(const section of sections){
    let last=section.index;
    for(let i=section.index+1;i<paras.length&&!endsSection(paras,i);i++)
      if(clean(paras[i].text)||ITEM_KINDS.includes(kindOf(paras[i])))last=i;
    section.end=last;
    for(let i=section.index+1;i<=last;i++)
      if(kindOf(paras[i])==='question'){const span=itemSpan(paras,i);section.items.push(summarizeItem(paras,span));i=span.end;}
  }
  return {sections};
}

// An item read back into the fields of its form, so it can be edited.
export function readItem(paras,at){
  const span=itemSpan(paras,at);
  if(!span)return null;
  const k=i=>kindOf(paras[i]),summary=summarizeItem(paras,span);
  if(!summary.type)return null;
  const type=ITEM_TYPES.find(t=>t.id===summary.type);
  let stem=softBreaks(clean(paras[span.start].text)),number=summary.number;
  const q=QUESTION_LABEL.exec(stem);
  if(q)stem=q[2];
  if(type.instruction&&stem.endsWith(` ${type.instruction}`))stem=stem.slice(0,-type.instruction.length-1);
  const item={span,type:type.id,number,stem,options:[],correct:[],truth:null,answer:'',why:'',points:'',objective:'',lines:0,firstText:clean(paras[span.start].text),label:summary.label};
  let keyText='',lineCount=0,notes=0;
  const heads=[],cells=[],rubricHeads=[],rubricCells=[];
  item.criteria=[];
  for(let i=span.start+1;i<=span.end;i++){
    const text=softBreaks(clean(paras[i].text));
    if(k(i)==='option')item.options.push(text);
    else if(k(i)==='answer-line')lineCount++;
    else if(k(i)==='item-note')notes++;
    else if(k(i)==='table-head')heads.push(text);
    else if(k(i)==='table-cell')cells.push(text);
    else if(k(i)==='key-table-head')rubricHeads.push(text);
    else if(k(i)==='key-table-cell')rubricCells.push(text);
    else if(k(i)==='answer-key'){
      const key=KEY_LABEL.exec(text),field=/^[^A-Za-z]*(Why|Points|Objective):\s*(.*)$/s.exec(text),criterion=/^[^A-Za-z]*Criterion\s+\d+:\s*(.*)$/s.exec(text);
      if(key)keyText=key[2];
      else if(criterion){const m=/^(.*?)\s*\((\d{1,3}(?:\.\d{1,2})?) points?\)$/s.exec(criterion[1].trim());item.criteria.push(m?{text:m[1].trim(),points:m[2]}:{text:criterion[1].trim(),points:''});}
      else if(field)item[field[1].toLowerCase()]=field[2].trim();
    }
  }
  if(type.options)item.correct=keyText.replace(/\.$/,'').split(/[,\s]+/).filter(Boolean).map(l=>'abcdefghijklmnopqrstuvwxyz'.indexOf(l.toLowerCase())).filter(i=>i>=0&&i<item.options.length);
  else if(type.truefalse)item.truth=keyText?/^true/i.test(keyText):null;
  else if(type.rubric){
    // The key's table: a header row (the criterion column, then each level with its points) and a row for each criterion.
    // In the key's own table styles it is for the facilitator; in the ordinary table styles, under the task, learners see it too.
    item.learnerRubric=!rubricHeads.length&&heads.length>0;
    const hs=item.learnerRubric?heads:rubricHeads,cs=item.learnerRubric?cells:rubricCells;
    item.levels=hs.slice(1).map(h=>{const m=/^(.*?)\s*\((\d{1,3}(?:\.\d{1,2})?) points?\)$/s.exec(h);return m?{text:m[1].trim(),points:m[2]}:{text:h,points:''};});
    const cols=hs.length||1;
    item.rubricRows=Array.from({length:Math.floor(cs.length/cols)},(_,r)=>cs.slice(r*cols,(r+1)*cols));
  }
  else if(type.numeric)Object.assign(item,readNumericKey(keyText));
  else if(type.answer)item.answer=keyText;
  else if(type.matching){
    // Rows of three cells (number, item, a space for the letter); the key says which choice each item matches.
    const matches=new Map([...keyText.matchAll(/(\d+)\s*=\s*([a-z])/gi)].map(m=>[Number(m[1]),m[2].toLowerCase()]));
    item.pairs=cells.filter((_,i)=>i%3===1).map((text,i)=>({item:text,match:item.options['abcdefghijklmnopqrstuvwxyz'.indexOf(matches.get(i+1)??'')]??''}));
    const used=new Set(item.pairs.map(p=>p.match.toLowerCase()));
    item.extras=item.options.filter(o=>!used.has(o.toLowerCase()));
  }else if(type.ordering){
    // The key lists the letters in the right order; the items themselves are the mixed-up list.
    item.orderItems=keyText.replace(/\.$/,'').split(/[,\s]+/).filter(Boolean).map(l=>item.options['abcdefghijklmnopqrstuvwxyz'.indexOf(l.toLowerCase())]).filter(s=>s!==undefined);
  }else if(type.semantic){
    item.sdPoints=Math.max(0,heads.length-2);
    item.sdPairs=cells.reduce((acc,text,i)=>{if(i%2===0)acc.push({item:text,match:''});else acc.at(-1).match=text;return acc;},[]);
  }else if(type.action){
    item.actions=notes;
  }else if(type.rating){
    item.statements=cells;
    item.labels=heads.slice(1);
    item.scale=RATING_SCALES.find(s=>s.labels.length===item.labels.length&&s.labels.every((l,i)=>l.toLowerCase()===item.labels[i].toLowerCase()))?.id??'custom';
  }
  item.lines=type.lines?lineCount:0;
  return item;
}

function sectionAt(paras,at){
  const {sections}=readAssessments(paras);
  return sections.filter(s=>s.index<=at&&at<=s.end).at(-1)||null;
}

// What the pane tells the author before they type: the section a new item would go into and the item it would follow,
// or null when the cursor is not in a section.
export function itemTarget(paras,at){
  const section=sectionAt(paras,at);
  if(!section)return null;
  const span=itemSpan(paras,at);
  return {title:section.title,after:span?summarizeItem(paras,span).label:null,number:itemPlacement(paras,at).number};
}

// Where a new item goes: inside the assessment section the cursor is in, after the item the cursor is in (or after
// the section's last item), numbered to follow the items before it.
export function itemPlacement(paras,at){
  const section=sectionAt(paras,at);
  if(!section)throw new Error('Items go inside an assessment section, and your cursor is not in one.');
  const inItem=itemSpan(paras,at);
  // The new item follows the item the cursor is in; from anywhere else in the section it follows the last item.
  const before=section.items.filter(i=>i.number!==null&&(!inItem||i.index<=inItem.start)).length;
  let anchor=inItem?inItem.end:section.items.length?section.items.at(-1).end:section.index;
  if(!inItem&&!section.items.length)
    for(let i=section.index+1;i<=section.end;i++)if(clean(paras[i].text))anchor=i; // after the note and the information lines
  return {anchor,number:before+1,section};
}

// Where a new assessment section goes: after the section the cursor is in, else after the last module or section, else at the end.
export function sectionPlacement(paras,at){
  const {sections}=readAssessments(paras),inside=sections.filter(s=>s.index<=at&&at<=s.end).at(-1);
  if(inside)return {anchor:inside.end};
  const guide=readGuide(paras);
  const ends=[...guide.modules.map(m=>m.end),...sections.map(s=>s.end)].filter(i=>i>=0);
  if(ends.length)return {anchor:Math.max(...ends)};
  let last=paras.length-1;
  while(last>0&&!clean(paras[last].text))last--;
  return {anchor:Math.max(last,0)};
}

// ---- the standard pages: which are in the document, and where a missing one goes
const isHeadingLevel=(p,level)=>new RegExp(`^heading ${level}$`,'i').test(String(p.style||''))||p.builtIn===`Heading${level}`;
export function standardPageState(paras){
  const find=(level,name)=>paras.findIndex(p=>isHeadingLevel(p,level)&&alnum(p.text)===alnum(name));
  return {
    acknowledgements:find(1,ACK_HEADING),contents:find(1,'Contents'),gettingStarted:find(1,'Getting Started'),
    about:find(2,'About This Guide'),perspective:find(2,PERSPECTIVE_HEADING),preparation:find(2,'Program Preparation'),glance:find(2,'Training at a Glance'),
    appendix:find(1,APPENDIX_HEADING)
  };
}
// The last paragraph of the page that starts at heading `at`: it runs to the next heading, module, lesson or assessment.
function pageEnd(paras,at){
  let i=at+1;
  while(i<paras.length&&!(['module','lesson','assessment','guide-title','title'].includes(kindOf(paras[i]))||/^heading\b/i.test(String(paras[i].style||''))||/^Heading/.test(String(paras[i].builtIn||''))))i++;
  return i-1;
}
// The standard pages a guide can still be given: Acknowledgements (before the Contents) and The Program in Perspective (between
// Graphic Cues and Program Preparation). Returns the pages to add, last one in the document first so that adding them in order
// never moves the place of the next, and the names of any that could not be placed. The Appendix is separate: it is optional.
export function missingStandardPages(paras){
  const s=standardPageState(paras),pages=[],cannot=[];
  const titleAt=paras.findIndex(p=>kindOf(p)==='guide-title');
  if(s.acknowledgements<0){
    const next=[s.contents,s.gettingStarted].filter(i=>i>=0).sort((a,b)=>a-b)[0];
    if(next!==undefined)pages.push({page:'acknowledgements',name:ACK_HEADING,after:next-1});
    else if(titleAt>=0)pages.push({page:'acknowledgements',name:ACK_HEADING,after:titleAt});
    else cannot.push(ACK_HEADING);
  }
  if(s.perspective<0){
    const next=[s.preparation,s.glance].filter(i=>i>=0).sort((a,b)=>a-b)[0];
    if(next!==undefined)pages.push({page:'perspective',name:PERSPECTIVE_HEADING,after:next-1});
    else if(s.about>=0)pages.push({page:'perspective',name:PERSPECTIVE_HEADING,after:pageEnd(paras,s.about)});
    else if(s.gettingStarted>=0)pages.push({page:'perspective',name:PERSPECTIVE_HEADING,after:s.gettingStarted});
    else cannot.push(PERSPECTIVE_HEADING);
  }
  return {pages:pages.sort((a,b)=>b.after-a.after),cannot};
}
// Where the Appendix goes: at the end of the guide, after the last paragraph that has any words. `exists` says it is there already.
export function appendixPlacement(paras){
  const at=standardPageState(paras).appendix;
  if(at>=0)return {exists:true,at};
  let last=paras.length-1;
  while(last>0&&!clean(paras[last].text))last--;
  return {exists:false,anchor:Math.max(last,0)};
}

// ---- tidying the guide: moving and deleting a whole module or lesson
// A lesson runs from its heading to the next lesson, module, assessment section or top-level heading, or to the module's Transition.
// A module (or an assessment section) runs to the next module, assessment section, title or Heading 1. Whatever the author typed
// inside a span is part of it, so it moves or goes with the heading. Headings of lower levels the author added stay inside.
const topBoundary=(paras,i)=>{const p=paras[i],k=kindOf(p);return k==='module'||k==='assessment'||k==='guide-title'||k==='title'||isHeadingLevel(p,1);};
// A span that runs to the end of the document stops before the empty paragraphs that end it: Word always keeps one there, so it is
// the document's, not the module's, and moving it would leave a stray empty paragraph between modules (seen in real Word 2026-11-03).
const spanFrom=(paras,start,stop)=>{
  let i=start+1;
  while(i<paras.length&&!stop(i))i++;
  let end=i-1;
  if(i>=paras.length)while(end>start&&kindOf(paras[end])==='other'&&!clean(paras[end].text))end--;
  return {start,end};
};
export const lessonSpan=(paras,at)=>spanFrom(paras,at,i=>topBoundary(paras,i)||kindOf(paras[i])==='lesson'||(kindOf(paras[i])==='cue'&&TRANSITION.test(clean(paras[i].text))));
export const topSpan=(paras,at)=>spanFrom(paras,at,i=>topBoundary(paras,i));

// What the tidy buttons act on, from the cursor: the lesson it is in, else the module it is in, else nothing.
export function tidyTarget(paras,at){
  if(at<0)return null;
  const guide=readGuide(paras);
  for(const mod of guide.modules)for(const lesson of mod.lessons){
    const span=lessonSpan(paras,lesson.index);
    if(at>=span.start&&at<=span.end)return {kind:'lesson',lesson,module:mod,span,text:clean(paras[lesson.index].text)};
  }
  for(const mod of guide.modules){
    if(mod.index<0)continue;
    const span=topSpan(paras,mod.index);
    if(at>=span.start&&at<=span.end)return {kind:'module',module:mod,span,text:clean(paras[mod.index].text)};
  }
  return null;
}
// Where the target would go if it moved one place up (dir -1) or down (dir 1): the paragraph to put it after (-1 is the top of the
// document), or a message saying why it cannot. A lesson swaps places with the lesson beside it in its own module; a module swaps
// with the module or assessment section beside it.
export function tidyMove(paras,target,dir){
  const noun=target.kind;
  if(target.kind==='lesson'){
    const lessons=target.module.lessons,j=lessons.indexOf(target.lesson)+dir;
    if(!lessons[j])return {message:`${target.text} is already the ${dir<0?'first':'last'} lesson in its module. A lesson moves only within its own module.`};
    const other=lessonSpan(paras,lessons[j].index);
    return {span:target.span,after:dir<0?other.start-1:other.end,over:clean(paras[lessons[j].index].text)};
  }
  const {span}=target;
  if(dir<0){
    let i=span.start-1;
    while(i>=0&&!['module','assessment'].includes(kindOf(paras[i]))&&!topBoundary(paras,i))i--;
    if(i<0||!['module','assessment'].includes(kindOf(paras[i])))return {message:`${target.text} is already the first ${noun}.`};
    return {span,after:i-1,over:clean(paras[i].text)};
  }
  const n=span.end+1;
  if(n>=paras.length||!['module','assessment'].includes(kindOf(paras[n])))return {message:`${target.text} is already the last ${noun}.`};
  return {span,after:topSpan(paras,n).end,over:clean(paras[n].text)};
}
// What deleting the target would take away, so the pane can say so before it does.
export function tidyDeletion(paras,target){
  const {span}=target;
  return {span,paragraphs:span.end-span.start+1,lessons:target.kind==='module'?target.module.lessons.length:0,text:target.text};
}

// ---- moving about the guide from the keyboard. Each action names where to go; the pane's shortcuts file gives each a key.
export const NAVIGATION_ACTIONS=Object.freeze({
  nextModule:'next-module',previousModule:'previous-module',nextLesson:'next-lesson',previousLesson:'previous-lesson',
  goToContents:'contents',goToGlance:'glance'
});
// Where one of those actions goes from the cursor at paragraph `at`: {index,text} for the paragraph to move to, or {message}
// saying why there is nowhere to go. Next is the first heading after the cursor, previous the nearest one before it.
export function navigationTarget(paras,at,kind){
  const guide=readGuide(paras),words=s=>String(s??'').replace(/[^\p{L}\p{N}]+/gu,'').toLowerCase();
  const isHeading2=p=>/^heading 2$/i.test(String(p.style||''))||/^Heading2$/.test(String(p.builtIn||''));
  const isHeading1=p=>/^heading 1$/i.test(String(p.style||''))||/^Heading1$/.test(String(p.builtIn||''));
  const go=(list,dir,noun)=>{
    const hit=dir>0?list.find(h=>h.index>at):[...list].reverse().find(h=>h.index<at);
    return hit?{index:hit.index,text:hit.text}:{message:`There is no ${dir>0?'next':'previous'} ${noun} in this guide.`};
  };
  const modules=guide.modules.filter(m=>m.index>=0).map(m=>({index:m.index,text:m.text}));
  const lessons=guide.modules.flatMap(m=>m.lessons).map(l=>({index:l.index,text:l.text}));
  switch(kind){
    case 'next-module':return go(modules,1,'module');
    case 'previous-module':return go(modules,-1,'module');
    case 'next-lesson':return go(lessons,1,'lesson');
    case 'previous-lesson':return go(lessons,-1,'lesson');
    case 'contents':{
      const i=paras.findIndex(p=>isHeading1(p)&&words(p.text)==='contents');
      return i>=0?{index:i,text:clean(paras[i].text)}:{message:'This guide has no Contents page. Start a guide with the standard pages to get one.'};
    }
    case 'glance':{
      const i=paras.findIndex(p=>isHeading2(p)&&words(p.text)==='trainingataglance');
      return i>=0?{index:i,text:clean(paras[i].text)}:{message:'This guide has no Training at a Glance heading. Choose Build Training at a Glance to make one.'};
    }
    default:return {message:'That is not a place this can go to.'};
  }
}
