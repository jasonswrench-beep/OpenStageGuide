// Backups of the guide. The guide is one Word document, and everything it keeps (its look, its slide memory, its delivery versions) lives
// inside that file, so a backup is a complete copy of the document saved as an ordinary .docx that opens in Word on any computer.
// This file names the copy and reads a file back to say what is in it: before a copy is saved (so one that does not match the document
// is never offered) and when the author chooses a saved backup to open again. Nothing here talks to Word.
import {readStyles,textOf} from './guide-check.mjs';
import {readGuide,kindOf,cueFromText} from './guide-model.mjs';

export const MAX_BACKUP_BYTES=64*1024*1024;
const MAX_PART_BYTES=200*1024*1024; // the unpacked document part: far more than a 64 MB file of real text can hold
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const plural=(n,word)=>`${n} ${word}${n===1?'':'s'}`;
const two=n=>String(n).padStart(2,'0');

// "Title - Backup 2026-11-08 0915.docx", in the computer's own time, with nothing in the title a file name cannot hold.
export function backupFileName(title,when=new Date()){
  const clean=String(title??'').replace(/[^\p{L}\p{N} _-]/gu,'').replace(/\s+/g,' ').trim().slice(0,60)||'OpenStage Guide';
  const day=`${when.getFullYear()}-${two(when.getMonth()+1)}-${two(when.getDate())}`,time=`${two(when.getHours())}${two(when.getMinutes())}`;
  return `${clean} - Backup ${day} ${time}.docx`;
}

export const sizeText=bytes=>bytes>=1024*1024?`${(bytes/1024/1024).toFixed(1)} MB`:`${Math.max(1,Math.round(bytes/1024))} KB`;

async function fingerprintOf(bytes){
  try{
    const digest=await globalThis.crypto.subtle.digest('SHA-256',bytes);
    return Array.from(new Uint8Array(digest),n=>n.toString(16).padStart(2,'0')).join('');
  }catch{return '';}
}

const NOT_WORD='This file is not a Word document that can be opened here. Choose a backup saved from OpenStage Guide, a .docx file.';

// What is in a .docx (as bytes), or an Error in plain words. `Zip` is JSZip; `DOMParserClass` reads the XML.
export async function inspectBackup(bytes,Zip,{DOMParserClass=globalThis.DOMParser}={}){
  if(!bytes||!bytes.length)throw new Error('That file is empty.');
  if(bytes.length>MAX_BACKUP_BYTES)throw new Error(`That file is ${sizeText(bytes.length)}. Backups up to ${sizeText(MAX_BACKUP_BYTES)} can be read.`);
  if(!Zip)throw new Error('The pane could not load what it needs to read a backup. Close it, open it again from the ribbon, then try again.');
  let zip;
  try{zip=await Zip.loadAsync(bytes);}catch{throw new Error(NOT_WORD);}
  const documentPart=zip.file('word/document.xml');
  if(!documentPart||!zip.file('[Content_Types].xml'))throw new Error(NOT_WORD);
  if((documentPart._data?.uncompressedSize||0)>MAX_PART_BYTES)throw new Error('The text of that document is too large to read here.');
  const parse=text=>new DOMParserClass().parseFromString(text,'application/xml');
  const doc=parse(await documentPart.async('string'));
  const body=doc.getElementsByTagNameNS(W,'body')[0];
  if(!body)throw new Error(NOT_WORD);
  const stylesPart=zip.file('word/styles.xml');
  const styles=stylesPart?readStyles(parse(await stylesPart.async('string')).documentElement):new Map();
  // Every paragraph in document order, a table's cells included, as the guide's own reading of a document takes them.
  const paras=Array.from(body.getElementsByTagNameNS(W,'p')).map((p,i)=>{
    const ppr=Array.from(p.childNodes).find(n=>n.nodeType===1&&n.localName==='pPr');
    const id=ppr&&Array.from(ppr.childNodes).find(n=>n.nodeType===1&&n.localName==='pStyle')?.getAttributeNS(W,'val');
    return {text:textOf(p),style:id?(styles.get(id)?.name||id):'Normal',builtIn:i===0&&id==='Title'?'Title':''};
  });
  if(!paras.length)throw new Error('That document has no paragraphs in it.');
  const guide=readGuide(paras),modules=guide.modules.filter(m=>m.index>=0);
  let core='';
  try{core=await zip.file('docProps/core.xml')?.async('string')||'';}catch{core='';}
  const tag=name=>(new RegExp(`<${name}[^>]*>([^<]*)</${name}>`).exec(core)||[])[1]?.trim()||'';
  const titled=paras.find(p=>kindOf(p)==='guide-title'&&p.text.trim());
  return {
    bytes:bytes.length,
    fingerprint:await fingerprintOf(bytes),
    paragraphs:paras.length,
    images:Object.keys(zip.files).filter(n=>/^word\/media\/[^/]+$/.test(n)).length,
    title:titled?titled.text.trim():(tag('dc:title')||paras.find(p=>p.text.trim())?.text.trim().slice(0,80)||''),
    created:tag('dcterms:created'),
    modified:tag('dcterms:modified'),
    isGuide:!!guide.isGuide,
    modules:modules.length,
    lessons:modules.reduce((n,m)=>n+m.lessons.length,0),
    slides:paras.filter(p=>kindOf(p)==='cue'&&cueFromText(p.text)?.id==='slide').length
  };
}

// Whether a copy just taken says the same about the guide as the document does ('' when it does). The copy is read back before it is
// offered, so a copy Word made badly is caught here and not found on the day it is needed.
export function copyProblem(info,live){
  if(!info.isGuide&&live.isGuide)return 'The copy Word made does not read as a guide, so it was not saved. Try again; if it happens again, close the pane, open it again from the ribbon and try once more.';
  if(info.modules!==live.modules||info.lessons!==live.lessons)
    return `The copy Word made has ${plural(info.modules,'module')} and ${plural(info.lessons,'lesson')}, but the document has ${plural(live.modules,'module')} and ${plural(live.lessons,'lesson')}, so the copy was not saved. Try again.`;
  return '';
}

// What to tell the author about a file, one fact to a line.
export function backupFacts(info){
  const when=iso=>{const t=Date.parse(iso);return Number.isFinite(t)?new Date(t).toLocaleString('en-US',{dateStyle:'long',timeStyle:'short'}):'';};
  const lines=[info.title?`Title: ${info.title}`:'No title'];
  lines.push(info.isGuide
    ?`${plural(info.modules,'module')}, ${plural(info.lessons,'lesson')}, ${plural(info.slides,'slide')}.`
    :'This is a Word document, but it was not made with the guide tools: it has no modules or lessons.');
  lines.push(`${plural(info.paragraphs,'paragraph')}, ${plural(info.images,'image')}, ${sizeText(info.bytes)}.`);
  const saved=when(info.modified);
  if(saved)lines.push(`Last saved in Word: ${saved}.`);
  if(info.fingerprint)lines.push(`Fingerprint: ${info.fingerprint.slice(0,12)} (the same file always gives the same twelve characters).`);
  return lines;
}
