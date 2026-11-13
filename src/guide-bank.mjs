// Importing questions from an OpenAssess bank file (.openassess), the author's separate question-bank project.
// A bank is a zip: manifest.json, one questions/q_*.json for each question and folders.json for the folder tree. This reads
// the questions and turns each into the item the guide builds from. A question the guide cannot hold (a picture or equation,
// a type it does not have, no key) is not guessed at: it is skipped and named with the reason, since the author cannot fix it
// here. The question model is documented in OpenAssess's packages/assess/src/model/question.ts. Pure apart from reading the zip.
import {itemBlock} from './guide-assess-ooxml.mjs';
import {makeContext} from './guide-ooxml.mjs';
import {MAX_OPTIONS,LIMITS} from './guide-items.mjs';

export const BANK_FORMAT='openassess-bank';
export const BANK_VERSION=1;
export const MAX_BANK_ADD=50;
const MAX_BANK_BYTES=200*1024*1024;

const ENTITIES={amp:'&',lt:'<',gt:'>',quot:'"',apos:"'",nbsp:' '};
const decode=s=>s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi,(m,e)=>{
  if(e[0]==='#'){const n=e[1]==='x'||e[1]==='X'?parseInt(e.slice(2),16):parseInt(e.slice(1),10);return Number.isFinite(n)&&n>0&&n<0x110000?String.fromCodePoint(n):m;}
  return ENTITIES[e.toLowerCase()]??m;
});
// A rich-text field ({format,content}) as plain lines: paragraphs, line breaks and list items each on a line, cells apart by a space.
export function richToText(rt){
  if(rt===undefined||rt===null)return '';
  const raw=typeof rt==='string'?rt:String(rt.content??'');
  const html=typeof rt==='string'?false:rt.format==='rich';
  const text=html?decode(raw.replace(/<br\s*\/?>/gi,'\n').replace(/<\/(?:p|li|tr|blockquote|pre|h[1-6]|ul|ol|table)>/gi,'\n').replace(/<\/(?:td|th)>/gi,' ').replace(/<[^>]*>/g,'')):raw;
  return text.split(/\r?\n/).map(l=>l.replace(/[ \t ]+/g,' ').trim()).filter(Boolean).join('\n');
}
const plainOf=rt=>richToText(rt).replace(/\s*\n\s*/g,' ').trim();
const num=n=>String(Number(Number(n).toFixed(10)));
const hasMediaOrMath=q=>/<oa-(?:media|math)\b/.test(JSON.stringify([q.stem,q.body,q.feedback]));

// One bank question into the item the guide builds from ({item,notes}), or {problem} saying why it cannot be.
export function questionToItem(q){
  const title=String(q?.title||q?.id||'Untitled question');
  const fail=message=>({problem:{id:q?.id??'',title,message}});
  if(!q||typeof q!=='object'||typeof q.type!=='string'||typeof q.body!=='object'||!q.body)return fail('The question file could not be read.');
  if(hasMediaOrMath(q)||(q.media?.length||q.math?.length))return fail('It has a picture or an equation, which the guide cannot bring across yet. Add it with Add an item.');
  const stem=richToText(q.stem);
  if(!stem)return fail('The question has no text.');
  const b=q.body,base={stem,points:q.points>0?num(q.points):'',objective:q.metadata?.objective||'',why:plainOf(q.feedback?.correct)||plainOf(q.feedback?.general)};
  const notes=[];
  let item;
  switch(q.type){
    case 'multipleChoice':case 'multipleResponse':{
      const opts=(b.options||[]).map(o=>({text:plainOf(o.content),correct:!!o.correct}));
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
      if(typeof b.correct!=='boolean')return fail('No answer is marked: it is not said whether the statement is true or false.');
      item={...base,type:'true-false',truth:b.correct};break;
    case 'fillInBlank':{
      if((b.blanks||[]).length!==1)return fail(`It has ${(b.blanks||[]).length} blanks, and the guide's fill in the blank has one.`);
      const accepted=(b.blanks[0].accepted||[]).map(a=>String(a).trim()).filter(Boolean);
      if(!accepted.length)return fail('The blank has no accepted answer.');
      item={...base,type:'fill-blank',answer:accepted.join('; ')};break;
    }
    case 'shortAnswer':{
      const accepted=(b.accepted||[]).map(a=>String(a).trim()).filter(Boolean);
      if(!accepted.length)return fail('It has no accepted answer.');
      item={...base,type:'short-answer',answer:accepted.join('; ')};break;
    }
    case 'matching':{
      const pairs=(b.pairs||[]).map(p=>({item:plainOf(p.left),match:String(p.right??'').trim()}));
      if(pairs.length<2)return fail('It has fewer than two pairs.');
      if(pairs.length>LIMITS.pairs)return fail(`It has ${pairs.length} pairs and a guide matching question can have up to ${LIMITS.pairs}.`);
      if(pairs.some(p=>!p.item||!p.match))return fail('A pair has an empty side.');
      const extras=(b.distractors||[]).map(d=>String(d).trim()).filter(Boolean);
      if(extras.length>LIMITS.extras)return fail(`It has ${extras.length} extra choices and the guide takes up to ${LIMITS.extras}.`);
      item={...base,type:'matching',pairs,extras};break;
    }
    case 'ordering':{
      const items=[...(b.items||[])].sort((x,y)=>x.position-y.position).map(i=>plainOf(i.content));
      if(items.length<LIMITS.minOrderItems||items.length>LIMITS.orderItems)return fail(`It has ${items.length} items and a guide ordering question has ${LIMITS.minOrderItems} to ${LIMITS.orderItems}.`);
      item={...base,type:'ordering',orderItems:items};break;
    }
    case 'numeric':{
      let value,tolerance='';
      if(b.mode==='range'&&Number.isFinite(b.min)&&Number.isFinite(b.max)&&b.max>=b.min){
        // The guide's key is a number and how far off still counts, so a range becomes its middle and half its width.
        value=(b.min+b.max)/2;tolerance=b.max>b.min?num((b.max-b.min)/2):'';
        if(tolerance)notes.push(`${title}: the range ${num(b.min)} to ${num(b.max)} is written as ${num(value)} plus or minus ${tolerance}.`);
      }else if(Number.isFinite(b.value)){value=b.value;if(b.mode==='tolerance'&&Number.isFinite(b.tolerance)&&b.tolerance>0)tolerance=num(b.tolerance);}
      else return fail('It has no correct number.');
      item={...base,type:'numeric',answer:num(value),tolerance,units:String(b.units||'').trim()};break;
    }
    case 'essay':{
      const model=richToText(b.modelAnswer);
      const rubric=(b.rubric||[]).map(r=>`${r.criterion}${r.description?`: ${r.description}`:''} (${num(r.points)} ${Number(r.points)===1?'point':'points'})`);
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

// Opens a bank file's bytes. Zip is the JSZip class. Returns the bank's title, its questions in folder order (each readable
// ones as {id,title,item}, others as {id,title,problem}) and its folders for choosing from.
export async function readBankFile(bytes,{Zip}){
  if(!Zip)throw new Error('The pane could not load what it needs to read a bank. Close it, open it again from the ribbon, then try again.');
  if(bytes.byteLength>MAX_BANK_BYTES)throw new Error('This file is over 200 MB, which is more than the pane can read. Export the part you need from OpenAssess.');
  let zip;
  try{zip=await Zip.loadAsync(bytes);}catch{throw new Error('This file could not be opened as an OpenAssess bank. It may be damaged, or it may not be a bank.');}
  const manifestFile=zip.file('manifest.json');
  if(!manifestFile)throw new Error('This file has no manifest.json, so it is not an OpenAssess bank.');
  let manifest;
  try{manifest=JSON.parse(await manifestFile.async('string'));}catch{throw new Error('The bank\'s manifest could not be read, so the file may be damaged.');}
  if(manifest?.format!==BANK_FORMAT)throw new Error(`This file says it is "${String(manifest?.format)}", not an OpenAssess bank.`);
  if(Number(manifest.formatVersion??0)>BANK_VERSION)throw new Error(`This bank was written by a newer OpenAssess (format ${manifest.formatVersion}); this guide reads format ${BANK_VERSION}. Update the guide, or export the bank again from the older version.`);
  const files=zip.file(/^questions\/[^/]+\.json$/);
  const byId=new Map();
  for(const f of files){
    const id=f.name.slice('questions/'.length,-'.json'.length);
    let q;
    try{q=JSON.parse(await f.async('string'));}catch{byId.set(id,{id,title:id,problem:{id,title:id,message:'The question file could not be read.'}});continue;}
    if(q?.metadata?.status==='retired')continue;
    const r=questionToItem({...q,id:q?.id||id});
    byId.set(id,{id,title:String(q?.title||id),...(r.problem?{problem:r.problem}:{item:r.item,notes:r.notes})});
  }
  // Folder order first, as the author arranged it, then anything not in a folder.
  let folderFile=null;
  const ff=zip.file('folders.json');
  if(ff){try{folderFile=JSON.parse(await ff.async('string'));}catch{folderFile=null;}}
  const folders=[],seen=new Set(),ordered=[];
  const take=ids=>{const out=[];for(const id of ids||[])if(byId.has(id)&&!out.includes(id))out.push(id);return out;};
  // A folder's questions include those in the folders under it.
  const idsUnder=n=>[...take(n.questionIds),...(Array.isArray(n.children)?n.children.flatMap(idsUnder):[])];
  const walk=(nodes,depth)=>{
    for(const n of Array.isArray(nodes)?nodes:[]){
      folders.push({id:String(n.id??''),name:String(n.name??'Folder'),depth,ids:[...new Set(idsUnder(n))]});
      for(const id of take(n.questionIds))if(!seen.has(id)){seen.add(id);ordered.push(id);}
      walk(n.children,depth+1);
    }
  };
  walk(folderFile?.root,0);
  const rest=[...byId.keys()].filter(id=>!seen.has(id)).sort();
  const all=[...ordered,...rest].map(id=>byId.get(id));
  return {title:String(manifest.title||'Untitled bank'),all,folders:folders.filter(f=>f.ids.length)};
}
