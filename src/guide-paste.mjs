// Pasting many questions at once, in the plain-text layout that test banks commonly use (the author's OpenAssess calls it the
// Standard Format). This reads the common part of that layout: a numbered question, lettered options with a * before the correct
// one (or an Answer: line, or an Answers: list at the end), and a Type: line before a question that is not multiple choice.
// It does not guess: a question it cannot read is named, with its line, and nothing is inserted until every question reads.
// Pure functions: text in, items (the shape itemBlock takes) and problems out.
import {LIMITS,MAX_OPTIONS,itemTypeById} from './guide-items.mjs';
import {itemBlock} from './guide-assess-ooxml.mjs';
import {makeContext} from './guide-ooxml.mjs';

export const MAX_PASTE=50;
// What each Type: line means here. MC (or no line) is multiple choice; the others are named as test banks name them.
const TYPES=Object.freeze({MC:'multiple-choice',MA:'multiple-response',MR:'multiple-response',TF:'true-false',E:'reflection',S:'short-answer',F:'fill-blank',MT:'matching'});
const QUESTION=/^\s*(\d{1,3})\s*[.)]\s+(\S.*)$/;
const OPTION=/^\s*(\*)?\s*\(?([A-Za-z])[.)]\s+(\S.*)$/;
const FEEDBACK=/^\s*(?:[~@#]|Feedback\s*:)/i;
const META=/^\s*(?:Title|DIF|REF|OBJ|TOP|BLOOM|MSC|NAT|LOC|PTS)\s*:/i;
const METADATA_TAIL=/\s{2,}(?:DIF|REF|OBJ|TOP|BLOOM|MSC|NAT|LOC|PTS)\s*:.*$/i;
const LETTERS='abcdefgh';
const PLACEHOLDER={reflection:'Add a model answer, or what to look for.','short-answer':'Add what a good answer includes.'};

// Splits the text into questions, each with its stem, options, the answers written next to it and its own line number.
export function parseStandardFormat(text){
  const lines=String(text??'').replace(/\r\n?/g,'\n').split('\n');
  const questions=[],block=new Map();
  let q=null,pending={},inAnswers=false,entries=0,ignored=0,prevBlank=true;
  const finish=()=>{if(q){questions.push(q);q=null;}};
  lines.forEach((raw,i)=>{
    const line=raw.trim(),no=i+1;
    // An Answers: list is the lines after its heading up to the first blank line, so questions can follow it.
    if(!line){prevBlank=true;if(inAnswers&&entries>0)inAnswers=false;return;}
    const wasBlank=prevBlank;prevBlank=false;
    // A Type:, Points: or Title: line starts the next question only between questions: after a blank line, or once the open
    // question has its options. In the middle of a question's own text (a stem quoting a label, say) it is part of the text.
    const between=!q||wasBlank||q.options.length>0||q.inline.length>0;
    if(/^Answers\s*:\s*$/i.test(line)){finish();inAnswers=true;entries=0;return;}
    let m;
    if(inAnswers){
      if((m=/^(\d{1,3})\s*[.):]\s*(.+)$/.exec(line))){block.set(Number(m[1]),[...(block.get(Number(m[1]))||[]),m[2].trim()]);entries++;}
      else ignored++;
      return;
    }
    if(between&&(m=/^Type\s*:\s*([A-Za-z]+)\s*$/i.exec(line))){finish();pending.type=m[1].toUpperCase();return;}
    if(between&&(m=/^Points\s*:\s*(\d{1,3}(?:\.\d{1,2})?)\s*$/i.exec(line))){finish();pending.points=m[1];return;}
    if(FEEDBACK.test(line)||(META.test(line)&&between)){ignored++;return;}
    if(q&&(m=/^(?:Answer|Ans)\s*[:.]\s*(.+)$/i.exec(line))){q.inline.push(m[1].replace(METADATA_TAIL,'').trim());return;}
    if((m=QUESTION.exec(line))){
      finish();
      q={number:Number(m[1]),line:no,type:pending.type,points:pending.points,stem:[m[2].trim()],options:[],inline:[]};
      pending={};return;
    }
    if(!q){ignored++;return;}
    if((m=OPTION.exec(line))){q.options.push({letter:m[2].toLowerCase(),text:m[3].trim(),star:!!m[1]});return;}
    if(q.options.length)q.options.at(-1).text+=` ${line}`;else q.stem.push(line);
  });
  finish();
  return {questions,block,ignored};
}

const tokens=raws=>raws.flatMap(r=>String(r).split(/[\s,;]+|\band\b/i)).map(t=>t.replace(/[().]/g,'').trim()).filter(Boolean);
const dedupe=list=>{const seen=new Set();return list.filter(t=>{const k=t.toLowerCase();if(seen.has(k))return false;seen.add(k);return true;});};

// One question into the item the guide builds from, or a problem that says what to fix.
function toItem(q,block){
  const fail=message=>({problem:{number:q.number,line:q.line,message}});
  const stem=q.stem.join(' ').trim(),opts=q.options,answers=[...q.inline,...(block.get(q.number)||[])];
  if(!stem)return fail('The question has no text.');
  let id=q.type?TYPES[q.type]:null;
  if(q.type&&!id)return fail(`Type: ${q.type} is not read here. Add this question with Add an item.`);
  const isTF=opts.length===2&&opts.every(o=>/^(true|false)$/i.test(o.text))&&new Set(opts.map(o=>o.text.toLowerCase())).size===2;
  if(!id)id=isTF?'true-false':opts.length?'multiple-choice':answers.length?'short-answer':null;
  if(!id)return fail('The question has no options and no Type: line, so it cannot be told from a heading. Add Type: E, S or F before it, or an Answer: line after it.');
  const item={type:id,stem,points:q.points};
  const note=[];
  const letterOk=opts.every((o,i)=>o.letter===LETTERS[i]);
  if(['multiple-choice','multiple-response','true-false'].includes(id)){
    if(opts.length<2)return fail('A question with options needs at least two lettered options (a. b. c.).');
    if(!letterOk)return fail('The option letters skip or do not start at a. Put them in order: a. b. c.');
    if(opts.length>MAX_OPTIONS&&id!=='true-false')return fail(`It has ${opts.length} options and a guide question can have up to ${MAX_OPTIONS}.`);
    const starred=opts.map((o,i)=>o.star?i:-1).filter(i=>i>=0);
    let correct=starred;
    if(id==='true-false'){
      if(starred.length>1)return fail('Both True and False are marked.');
      let truth;
      if(starred.length)truth=/^true$/i.test(opts[starred[0]].text);
      else{
        // The key can say T or F, True or False, or the letter of the True or False option.
        const t=tokens(answers)[0];
        if(t&&/^(t|true)$/i.test(t))truth=true;
        else if(t&&/^(f|false)$/i.test(t))truth=false;
        else if(t&&t.length===1&&LETTERS.indexOf(t.toLowerCase())>=0&&LETTERS.indexOf(t.toLowerCase())<opts.length)truth=/^true$/i.test(opts[LETTERS.indexOf(t.toLowerCase())].text);
        else if(t)return fail(`The answer "${t}" is not True, False, T, F or the letter of one of them.`);
      }
      if(truth===undefined)return fail('No answer is marked. Put * before True or False, or add an Answer: line.');
      return {item:{...item,truth},note};
    }
    if(!correct.length){
      const toks=tokens(answers);
      const idx=toks.map(t=>t.length===1?LETTERS.indexOf(t.toLowerCase()):opts.findIndex(o=>o.text.toLowerCase()===t.toLowerCase()));
      const bad=toks.find((t,i)=>idx[i]<0||idx[i]>=opts.length);
      if(toks.length&&bad!==undefined)return fail(`The answer "${bad}" is not one of the options.`);
      correct=[...new Set(idx)];
    }
    if(!correct.length)return fail('No correct answer is marked. Put * before the letter of the correct option, or add an Answer: line.');
    if(id==='multiple-choice'&&correct.length>1)return fail('More than one correct answer is marked. If more than one is meant, put Type: MA before the question.');
    return {item:{...item,options:opts.map(o=>o.text),correct},note};
  }
  if(id==='matching'){
    const pairs=opts.map(o=>{const at=o.text.indexOf('=');return at<0?null:{item:o.text.slice(0,at).trim(),match:o.text.slice(at+1).trim()};});
    if(pairs.some(p=>!p||!p.item||!p.match))return fail('Each option of a matching question needs an item, an equals sign and its match, like: a. Marconi = Radio.');
    if(pairs.length<2)return fail('A matching question needs at least two pairs.');
    if(pairs.length>LIMITS.pairs)return fail(`A matching question can have up to ${LIMITS.pairs} pairs.`);
    return {item:{...item,pairs,extras:[]},note};
  }
  // Short answer, fill in the blank and essay: the words after the options or in the answers are what the key says.
  const accepted=dedupe([...opts.map(o=>o.text),...answers.map(a=>a.trim()).filter(Boolean)]);
  if(id==='fill-blank'){
    if(!accepted.length)return fail('A fill-in-the-blank question needs its accepted answers: as options, or in an Answer: line.');
    return {item:{...item,answer:accepted.join('; ')},note};
  }
  if(!accepted.length){note.push(`${itemTypeById(id).label} question ${q.number} has a placeholder in its key.`);return {item:{...item,answer:PLACEHOLDER[id]},note};}
  return {item:{...item,answer:accepted.join(id==='reflection'?'\n':'; ')},note};
}

// Everything the pane needs: the items in order, the problems with their lines, and counts for the message.
export function readPastedQuestions(text){
  const {questions,block,ignored}=parseStandardFormat(text);
  const items=[],problems=[],notes=[];
  // A number used twice is only a problem when an Answers: list is keyed by it, for the questions are numbered again anyway.
  // Such questions are not read at all, for whatever is read of them would be matched with the wrong answers.
  const seen=new Map();
  for(const q of questions)seen.set(q.number,(seen.get(q.number)||0)+1);
  const reported=new Set();
  for(const q of questions){
    if(seen.get(q.number)>1&&block.has(q.number)){
      if(!reported.has(q.number)){
        reported.add(q.number);
        problems.push({number:q.number,line:q.line,message:`Number ${q.number} is used ${seen.get(q.number)} times and the Answers: list is keyed by it, so the answers cannot be matched. Number the questions one after another, or mark each correct option with a *.`});
      }
      continue;
    }
    const r=toItem(q,block);
    if(r.problem){problems.push(r.problem);continue;}
    // The builder has the last word: a question it would refuse is a problem now, with its number, not a surprise on insert.
    try{itemBlock({...r.item,number:1},makeContext());}
    catch(e){problems.push({number:q.number,line:q.line,message:e.message});continue;}
    items.push(r.item);notes.push(...r.note);
  }
  const counts=new Map();
  for(const it of items)counts.set(it.type,(counts.get(it.type)||0)+1);
  return {found:questions.length,items,problems,notes,ignored,counts};
}
