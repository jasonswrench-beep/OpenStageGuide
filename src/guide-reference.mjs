// The reference pages of a guide (kit slice K4 of docs/KIT-SPEC-2026-11-15.md): a glossary, a list of references, figure captions with their source, pull quotes, and a
// licence and permissions page. Each is paragraphs in styles of its own, found again by their shape, so the guide stays a list of styled paragraphs with no hidden record.
// The glossary and the references are Heading 1 sections at the end of the guide that the tool adds to in alphabetical order; the licence page is built from what the
// author says and built again over itself. Learners have all of it. Pure functions: the guide's paragraphs in, Word-format XML and positions out.
import {STYLE_NAMES,STYLE_IDS} from './guide-styles.mjs';
import {para,run,heading} from './guide-ooxml.mjs';

export const GLOSSARY_HEADING='Glossary',REFERENCES_HEADING='References',LICENCE_HEADING='Licence and permissions';
const clean=t=>String(t??'').replace(/[\r\n\u0007]+$/,'').replace(/\s+/g,' ').trim();
const alnum=t=>clean(t).normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[^\p{L}\p{N}]+/gu,'').toLowerCase();
const isHeading1=p=>/^heading 1$/i.test(String(p?.style||''))||/^Heading1$/.test(String(p?.builtIn||''));
export const LIMITS=Object.freeze({term:80,meaning:600,reference:600,caption:300,source:300,quote:500,by:120,holder:120,words:1200,place:120});
const need=(value,what,max)=>{
  const v=clean(value);
  if(!v)throw new Error(`Type ${what} first.`);
  if(v.length>max)throw new Error(`${what[0].toUpperCase()}${what.slice(1)} is too long: it holds up to ${max} letters.`);
  return v;
};
const sentence=t=>/[.!?]$/.test(t)?t:`${t}.`;

// A section of the guide: its Heading 1, and the paragraphs of the one style that follow it. null when there is no such heading.
function sectionOf(paras,title,style){
  const heading=paras.findIndex(p=>isHeading1(p)&&alnum(p.text)===alnum(title));
  if(heading<0)return null;
  let end=heading;
  while(end+1<paras.length&&String(paras[end+1].style||'')===style)end++;
  return {heading,end,entries:paras.slice(heading+1,end+1).map((p,i)=>({index:heading+1+i,text:clean(p.text)}))};
}
// Where an entry goes to keep a section in order: after the last entry that sorts at or before it, or straight after the heading.
const placeIn=(entries,heading,key,keyOf)=>{
  let after=heading;
  for(const e of entries)if(keyOf(e.text).localeCompare(key,undefined,{sensitivity:'base',numeric:true})<=0)after=e.index;
  return after;
};
const termOf=text=>(/^(.*?)\s*:\s*(.*)$/.exec(text)||[,text,''])[1];

// ---- the glossary
export const termLine=(term,meaning)=>para(STYLE_IDS.term,run(`${term}:`,STYLE_IDS.cueLabel)+run(` ${meaning}`));
export function readGlossary(paras){
  const s=sectionOf(paras,GLOSSARY_HEADING,STYLE_NAMES.term);
  return s?{...s,terms:s.entries.map(e=>({index:e.index,term:termOf(e.text),text:e.text}))}:null;
}
// What adding a term does: where it goes, and the XML (with the Glossary heading when the guide has none yet, at the end of the guide).
export function planTerm(paras,term,meaning){
  const t=need(term,'the term',LIMITS.term),m=need(meaning,'what it means',LIMITS.meaning);
  if(t.includes(':'))throw new Error('A term cannot hold a colon: the colon is what separates it from its meaning.');
  const g=readGlossary(paras);
  if(!g)return {after:paras.length-1,body:heading(1,GLOSSARY_HEADING)+termLine(t,m),created:true,term:t};
  const clash=g.terms.find(x=>alnum(x.term)===alnum(t));
  if(clash)throw new Error(`The glossary already has "${clash.term}". Change its words on the page.`);
  return {after:placeIn(g.terms,g.heading,t,termOf),body:termLine(t,m),created:false,term:t};
}

// ---- the references
export const referenceLine=text=>para(STYLE_IDS.reference,run(text));
export function readReferences(paras){return sectionOf(paras,REFERENCES_HEADING,STYLE_NAMES.reference);}
export function planReference(paras,text){
  const r=need(text,'the reference',LIMITS.reference);
  const s=readReferences(paras);
  if(!s)return {after:paras.length-1,body:heading(1,REFERENCES_HEADING)+referenceLine(r),created:true};
  if(s.entries.some(e=>alnum(e.text)===alnum(r)))throw new Error('The references already hold that one. Change its words on the page.');
  return {after:placeIn(s.entries,s.heading,r,t=>t),body:referenceLine(r),created:false};
}

// ---- figure captions, numbered by Update guide
export const captionLine=({caption,source})=>{
  const c=need(caption,'the caption',LIMITS.caption),src=clean(source);
  if(src.length>LIMITS.source)throw new Error(`The source is too long: it holds up to ${LIMITS.source} letters.`);
  return para(STYLE_IDS.caption,run(`Figure: ${sentence(c)}${src?` Source: ${sentence(src)}`:''}`));
};
const CAPTION=/^Figure\s*[\d.]*\s*:\s*(.*)$/s;
// The number each caption should have: by module (2.1, 2.2) inside a module, by order (1, 2) outside one. Returns the edits Update guide makes.
export function captionEdits(paras){
  const edits=[];
  let module=0,appendix=false,inModule=0,loose=0;
  paras.forEach((p,index)=>{
    const style=String(p.style||'');
    if(style===STYLE_NAMES.module){module++;inModule=0;appendix=false;}
    else if(/^heading 1$/i.test(style)&&/^appendix\b/i.test(clean(p.text)))appendix=true;
    if(style!==STYLE_NAMES.caption)return;
    const m=CAPTION.exec(clean(p.text));
    if(!m)return;
    const number=module>0&&!appendix?`${module}.${++inModule}`:String(++loose);
    const want=`Figure ${number}: ${m[1]}`;
    if(clean(p.text)!==want)edits.push({kind:'caption',index,text:want});
  });
  return edits;
}

// ---- a pull quote with its attribution
export const quoteBlock=({quote,by})=>{
  const q=need(quote,'the quotation',LIMITS.quote),who=need(by,'who said it',LIMITS.by);
  return para(STYLE_IDS.quote,run(`“${q.replace(/^["“]+|["”]+$/g,'')}”`))+para(STYLE_IDS.quoteBy,run(who));
};

// ---- the licence and permissions page
const CC=(name,url,words)=>({id:name.toLowerCase().replace(/\s+/g,'-'),label:name,url,words});
export const LICENCES=Object.freeze([
  {id:'all-rights',label:'All rights reserved',words:({holder,year})=>`© ${year} ${holder}. All rights reserved. No part of this guide may be copied, shared or changed without the written permission of ${holder}.`},
  ...[
    CC('CC BY 4.0','https://creativecommons.org/licenses/by/4.0/','You may share and adapt this guide for any purpose, including commercially, if you give credit.'),
    CC('CC BY-NC 4.0','https://creativecommons.org/licenses/by-nc/4.0/','You may share and adapt this guide if you give credit and do not use it commercially.'),
    CC('CC BY-SA 4.0','https://creativecommons.org/licenses/by-sa/4.0/','You may share and adapt this guide for any purpose, including commercially, if you give credit and share what you make from it under the same licence.'),
    CC('CC BY-NC-SA 4.0','https://creativecommons.org/licenses/by-nc-sa/4.0/','You may share and adapt this guide if you give credit, do not use it commercially and share what you make from it under the same licence.')
  ].map(c=>({id:c.id,label:c.label,words:({holder,year})=>`© ${year} ${holder}. This guide is licensed under ${c.label} (${c.url}). ${c.words}`})),
  {id:'custom',label:'My own words',words:({words})=>words}
]);
export const licenceById=id=>LICENCES.find(l=>l.id===id)||null;
export function licenceBody({licence,holder,year,words,reviewed,place}){
  const l=licenceById(licence);
  if(!l)throw new Error('Choose a licence, or choose My own words.');
  const y=clean(year)||String(new Date().getFullYear());
  if(!/^\d{4}$/.test(y))throw new Error('The year is four digits, such as 2026.');
  const h=l.id==='custom'?clean(holder):need(holder,'who holds the rights',LIMITS.holder);
  const w=l.id==='custom'?need(words,'the words of the licence or permission',LIMITS.words):'';
  const lines=[l.words({holder:h,year:y,words:w})];
  const r=clean(reviewed),p=clean(place);
  if(r)lines.push(`Content that dates (figures, laws, guidance) was last reviewed on ${r}.`);
  if(p)lines.push(`Where the law applies: ${sentence(p)}`);
  if(clean(reviewed).length>60||p.length>LIMITS.place)throw new Error(`The review date and the place are short: up to 60 and ${LIMITS.place} letters.`);
  return heading(1,LICENCE_HEADING)+lines.map(t=>para(null,run(t))).join('');
}
// Where the page is, and how far it runs: its heading and the plain paragraphs after it up to the next heading or a styled block.
export function findLicence(paras){
  const gi=paras.findIndex(p=>isHeading1(p)&&alnum(p.text)===alnum(LICENCE_HEADING));
  if(gi<0)return null;
  let end=gi;
  while(end+1<paras.length&&!isHeading1(paras[end+1])&&['Normal',''].includes(String(paras[end+1].style||'Normal'))&&String(paras[end+1].builtIn||'')!=='Title')end++;
  return {start:gi,end};
}
