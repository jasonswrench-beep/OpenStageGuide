// A handout pack: the forms to give out, from one lesson, one module or the whole guide, in one Word file for copying. It is made from the guide's own
// Word package and styles (src/guide-docx.mjs), so the forms look as they do in the guide; each one starts on a new page, a form printed landscape is a section
// of its own, and nothing meant for the facilitator goes in: not the print line, and not a form marked for the facilitator only.
import {STYLE_NAMES} from './guide-styles.mjs';
import {formSpans,readForms} from './guide-forms.mjs';

export const PACK_SCOPES=Object.freeze([
  {id:'lesson',label:'The lesson the cursor is in'},
  {id:'module',label:'The module the cursor is in'},
  {id:'guide',label:'The whole guide'}
]);

// The paragraphs a scope covers, and what to call it. A lesson runs from its heading to the next lesson, module, assessment or heading, and a module to the next
// module, assessment or heading, so the forms written under a lesson's blocks are inside it (the guide model's own `end` stops at the last block).
const isHeading=style=>/^(heading \d|title)$/i.test(style);
export function packRange(paras,at,scope){
  const style=i=>String(paras[i]?.style||'');
  if(scope==='guide')return {start:0,end:paras.length-1,name:'the guide'};
  const kinds=scope==='lesson'?[STYLE_NAMES.lesson]:scope==='module'?[STYLE_NAMES.module]:null;
  if(!kinds)throw new Error('Choose a lesson, a module or the whole guide.');
  const stops=scope==='lesson'?[STYLE_NAMES.module,STYLE_NAMES.lesson,STYLE_NAMES.assessment,STYLE_NAMES.guideTitle]:[STYLE_NAMES.module,STYLE_NAMES.assessment,STYLE_NAMES.guideTitle];
  const stopsAt=i=>stops.includes(style(i))||isHeading(style(i));
  let start=-1;
  for(let i=Math.min(at,paras.length-1);i>=0;i--){
    if(kinds.includes(style(i))){start=i;break;}
    if(stopsAt(i))break;
  }
  if(start<0)throw new Error(scope==='lesson'?'Put your cursor in a lesson first, or choose the module or the whole guide.':'Put your cursor in a module first, or choose the whole guide.');
  let end=start;
  while(end+1<paras.length&&!stopsAt(end+1))end++;
  return {start,end,name:String(paras[start].text||'').trim()||(scope==='lesson'?'the lesson':'the module')};
}

// Puts a page break in front of a paragraph, in the place the page setup of a paragraph expects it (after its style and keep settings).
export function withPageBreakBefore(xml){
  if(/<w:pageBreakBefore\b/.test(xml))return xml;
  if(/<w:pPr>/.test(xml))return xml.replace(/(<w:pPr>)((?:<w:(?:pStyle|keepNext|keepLines)\b[^>]*\/>)*)/,'$1$2<w:pageBreakBefore/>');
  return xml.replace(/^<w:p(\s[^>]*)?>/,(m,attrs)=>`<w:p${attrs||''}><w:pPr><w:pageBreakBefore/></w:pPr>`);
}

/**
 * The paragraphs of a pack. `entries` are the guide's paragraphs with their XML (exportState().paragraphs), `paras` the same paragraphs with their marks
 * (paragraphs()), and `at` the paragraph the cursor is in. A form is in the pack when it is for everyone and it is inside the scope or cited from a block
 * inside it. Returns {entries, forms, scopeName}; throws, saying what to do, when no form qualifies.
 */
export function packEntries(entries,paras,{at=0,scope='guide'}={}){
  if(entries.length!==paras.length)throw new Error('The guide changed while the pack was being made. Try again.');
  const range=packRange(paras,at,scope);
  const spans=formSpans(paras),read=readForms(paras);
  const citedIds=new Set(read.cites.filter(c=>c.index>=range.start&&c.index<=range.end).map(c=>c.id));
  const chosen=spans.filter(s=>s.audience==='everyone'&&((s.index>=range.start&&s.index<=range.end)||(s.id&&citedIds.has(s.id))));
  if(!chosen.length){
    const any=spans.some(s=>s.audience==='everyone');
    throw new Error(`There is no form for everyone in ${range.name}${scope==='guide'?'':' or cited from it'}. ${any?'Choose a wider part of the guide, or':'Start a form in the Forms tool, or'} change a form for the facilitator only into one for everyone.`);
  }
  const out=[];
  let previousLandscape=false;
  chosen.forEach((s,n)=>{
    const landscape=s.orientation==='landscape';
    for(let i=s.index;i<=s.end;i++){
      if(i===s.printIndex)continue; // the print line is for the facilitator
      const e=entries[i];
      // A form starts on a page of its own, unless it starts a section (a change of orientation does that already).
      const startsPage=i===s.index&&n>0&&landscape===previousLandscape;
      out.push({...e,...(startsPage&&e.xml?{xml:withPageBreakBefore(e.xml)}:{}),...(landscape?{landscape:true}:{landscape:false})});
    }
    previousLandscape=landscape;
  });
  return {entries:out,forms:chosen.map(s=>s.text),scopeName:range.name};
}
