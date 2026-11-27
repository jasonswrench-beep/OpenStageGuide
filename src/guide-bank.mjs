// Importing questions from an OpenAssess bank file (.openassess), the author's separate question-bank project.
// The file is read by the shared core's bank reader, @openstage/bank (BANK-1; carried in ./openstage-bank, see scripts/vendor-bank.mjs),
// which gives each question with its key and its details, the folders in order, and what it could not read. This file keeps what is the
// guide's own: which questions a guide item can hold, and how. A question the guide cannot hold (a picture or equation, a type it does not
// have, no key) is not guessed at: it is skipped and named with the reason, since the author cannot fix it here. Retired questions are left
// out. Pure apart from the reader.
import {readBank,BankError,BANK_FORMAT_VERSION} from './openstage-bank/index.mjs';
import {itemBlock} from './guide-assess-ooxml.mjs';
import {makeContext} from './guide-ooxml.mjs';
import {MAX_OPTIONS,LIMITS} from './guide-items.mjs';

export const MAX_BANK_ADD=50;
const MAX_BANK_BYTES=200*1024*1024;

// The reader's text ({html,text}) as plain lines, or as one line.
const linesOf=t=>String(t?.text??'').trim();
const plainOf=t=>linesOf(t).replace(/\s*\n\s*/g,' ').trim();
const num=n=>String(Number(Number(n).toFixed(10)));
const hasMediaOrMath=q=>q.pictures?.length||q.equations?.length||/<oa-(?:media|math)\b/.test(JSON.stringify([q.stem,q.options,q.key,q.feedback]));

// One question as the bank reader gives it into the item the guide builds from ({item,notes}), or {problem} saying why it cannot be.
export function questionToItem(q){
  const title=String(q?.title||q?.id||'Untitled question');
  const fail=message=>({problem:{id:q?.id??'',title,message}});
  if(!q||typeof q!=='object'||typeof q.type!=='string'||!q.key||typeof q.key!=='object')return fail('The question could not be read.');
  if(hasMediaOrMath(q))return fail('It has a picture or an equation, which the guide cannot bring across yet. Add it with Add an item.');
  const stem=linesOf(q.stem);
  if(!stem)return fail('The question has no text.');
  const k=q.key,base={stem,points:q.points>0?num(q.points):'',objective:q.metadata?.objective||'',why:plainOf(q.feedback?.correct)||plainOf(q.feedback?.general)};
  const notes=[];
  let item;
  switch(q.type){
    case 'multipleChoice':case 'multipleResponse':{
      const opts=(q.options||[]).map(o=>({text:plainOf(o.text),correct:!!o.correct}));
      if(opts.length<2)return fail('It has fewer than two options.');
      if(opts.length>MAX_OPTIONS)return fail(`It has ${opts.length} options and a guide question can have up to ${MAX_OPTIONS}.`);
      if(opts.some(o=>!o.text))return fail('An option has no text.');
      const correct=opts.map((o,i)=>o.correct?i:-1).filter(i=>i>=0);
      if(!correct.length)return fail('No correct option is marked.');
      if(q.type==='multipleChoice'&&correct.length>1)return fail('More than one option is marked correct, and a multiple choice question has one.');
      item={...base,type:q.type==='multipleChoice'?'multiple-choice':'multiple-response',options:opts.map(o=>o.text),correct};
      break;
    }
    case 'trueFalse':
      if(typeof k.correct!=='boolean')return fail('No answer is marked: it is not said whether the statement is true or false.');
      item={...base,type:'true-false',truth:k.correct};break;
    case 'fillInBlank':{
      const blanks=k.blanks||[];
      if(blanks.length!==1)return fail(`It has ${blanks.length} blanks, and the guide's fill in the blank has one.`);
      const accepted=(blanks[0].accepted||[]).map(a=>String(a).trim()).filter(Boolean);
      if(!accepted.length)return fail('The blank has no accepted answer.');
      item={...base,type:'fill-blank',answer:accepted.join('; ')};break;
    }
    case 'shortAnswer':{
      const accepted=(k.accepted||[]).map(a=>String(a).trim()).filter(Boolean);
      if(!accepted.length)return fail('It has no accepted answer.');
      item={...base,type:'short-answer',answer:accepted.join('; ')};break;
    }
    case 'matching':{
      const pairs=(k.pairs||[]).map(p=>({item:plainOf(p.left),match:String(p.right??'').trim()}));
      if(pairs.length<2)return fail('It has fewer than two pairs.');
      if(pairs.length>LIMITS.pairs)return fail(`It has ${pairs.length} pairs and a guide matching question can have up to ${LIMITS.pairs}.`);
      if(pairs.some(p=>!p.item||!p.match))return fail('A pair has an empty side.');
      const extras=(k.distractors||[]).map(d=>String(d).trim()).filter(Boolean);
      if(extras.length>LIMITS.extras)return fail(`It has ${extras.length} extra choices and the guide takes up to ${LIMITS.extras}.`);
      item={...base,type:'matching',pairs,extras};break;
    }
    case 'ordering':{
      // The reader gives the items in their right order.
      const items=(k.items||[]).map(i=>plainOf(i.text));
      if(items.length<LIMITS.minOrderItems||items.length>LIMITS.orderItems)return fail(`It has ${items.length} items and a guide ordering question has ${LIMITS.minOrderItems} to ${LIMITS.orderItems}.`);
      item={...base,type:'ordering',orderItems:items};break;
    }
    case 'numeric':{
      let value,tolerance='';
      if(k.mode==='range'&&Number.isFinite(k.min)&&Number.isFinite(k.max)&&k.max>=k.min){
        // The guide's key is a number and how far off still counts, so a range becomes its middle and half its width.
        value=(k.min+k.max)/2;tolerance=k.max>k.min?num((k.max-k.min)/2):'';
        if(tolerance)notes.push(`${title}: the range ${num(k.min)} to ${num(k.max)} is written as ${num(value)} plus or minus ${tolerance}.`);
      }else if(Number.isFinite(k.value)){value=k.value;if(k.mode==='tolerance'&&Number.isFinite(k.tolerance)&&k.tolerance>0)tolerance=num(k.tolerance);}
      else return fail('It has no correct number.');
      item={...base,type:'numeric',answer:num(value),tolerance,units:String(k.units||'').trim()};break;
    }
    case 'essay':{
      const model=linesOf(k.modelAnswer);
      const rubric=(k.rubric||[]).map(r=>`${r.criterion}${r.description?`: ${r.description}`:''} (${num(r.points)} ${Number(r.points)===1?'point':'points'})`);
      const key=[model,...(rubric.length?['Rubric:',...rubric]:[])].filter(Boolean).join('\n');
      if(!key)notes.push(`${title}: an essay with no model answer gets a placeholder in its key.`);
      item={...base,type:'reflection',answer:key||'Add a model answer, or what to look for.'};break;
    }
    default:return fail(`The type "${q.type}" is not read here. Add this question with Add an item.`);
  }
  try{itemBlock({...item,number:1},makeContext());}
  catch(e){return fail(e.message);}
  return {item,notes};
}

// Why a file is not read at all (BANK-1 section 7.1), in the guide's words.
function refusal(e){
  const said=/manifest says it is ("[^"]*")/.exec(e.message)?.[1];
  switch(e.code){
    case 'no-manifest':return 'This file has no manifest.json, so it is not an OpenAssess bank.';
    case 'damaged-manifest':return 'The bank\'s manifest could not be read, so the file may be damaged.';
    case 'not-a-bank':return `This file says it is ${said||'something else'}, not an OpenAssess bank.`;
    case 'no-version':return 'This bank does not say which version of the bank format it is written in, so it is not read. Save it again in OpenAssess.';
    case 'newer-format':return `This bank was written by a newer OpenAssess than this guide can read (it reads bank format ${BANK_FORMAT_VERSION}). Update OpenStage Guide, or save the bank in the older OpenAssess.`;
    default:return 'This file could not be opened as an OpenAssess bank. It may be damaged, or it may not be a bank.';
  }
}

// Opens a bank file's bytes. Returns the bank's title, its questions in folder order (each readable one as {id,title,item,notes}, others
// as {id,title,problem}), then the question files the reader could not read, and the folders for choosing from (each with the questions in
// it and in the folders under it).
export async function readBankFile(bytes){
  if(bytes.byteLength>MAX_BANK_BYTES)throw new Error('This file is over 200 MB, which is more than the pane can read. Export the part you need from OpenAssess.');
  let bank;
  try{bank=await readBank(bytes instanceof Uint8Array?bytes:new Uint8Array(bytes));}
  catch(e){throw new Error(refusal(e instanceof BankError?e:{code:'',message:''}));}
  const all=[];
  for(const q of bank.questions){
    if(q.metadata.status==='retired')continue;
    const r=questionToItem(q);
    all.push({id:q.id,title:q.title||q.id,...(r.problem?{problem:r.problem}:{item:r.item,notes:r.notes})});
  }
  // A question file that was not read is named with the reader's reason, so the author can see it is there.
  for(const p of bank.problems)if(p.skipped&&p.questionId&&!all.some(q=>q.id===p.questionId))all.push({id:p.questionId,title:p.questionId,problem:{id:p.questionId,title:p.questionId,message:p.message}});
  const listed=new Set(all.map(q=>q.id));
  const under=f=>[...f.questionIds.filter(id=>listed.has(id)),...bank.folders.filter(c=>c.parentId===f.id).flatMap(under)];
  const folders=bank.folders.map(f=>({id:f.id,name:f.name,depth:f.depth,ids:[...new Set(under(f))]})).filter(f=>f.ids.length);
  return {title:bank.title,all,folders};
}
