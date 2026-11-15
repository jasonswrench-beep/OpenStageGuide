// Layout choices for the Word file (kit slice K6 of docs/KIT-SPEC-2026-11-15.md). They belong to the output, not to the content: they are kept with the guide (in its
// settings) and applied when the guide is saved as a Word file, so the guide on the page does not change. The choices: each module starting a new page, a running header
// (the guide's title, or the name of the module), and a footer with the page number (on the right, or on the outside of two-sided pages), with words of the author's.
// Pure functions: the body of the Word document in, the body, the header and footer parts and the relationships they need out.
import {STYLE_NAMES,STYLE_IDS} from './guide-styles.mjs';

export const HEADERS=Object.freeze(['none','title','module']);
export const FOOTERS=Object.freeze(['none','plain','mirrored']);
export const HEADER_LABELS=Object.freeze({none:'No running header',title:'The title of the guide',module:'The name of the module'});
export const FOOTER_LABELS=Object.freeze({none:'No footer',plain:'The page number on the right',mirrored:'The page number on the outside, for two-sided printing'});
export const FOOTER_WORDS_MAX=120;
/** The heading of the one-page summary: the Word file puts the page on a sheet of its own. */
export const SUMMARY_HEADING='One-page summary';
export const LAYOUT_DEFAULTS=Object.freeze({dividers:false,header:'none',footer:'none',footerText:'',marginIcons:false,slideBeside:false});

const oneLine=(t,max)=>String(t??'').replace(/[\u0000-\u001f\u007f]+/g,' ').replace(/\s+/g,' ').trim().slice(0,max);
/** Any stored value as a complete, safe layout: unknown words fall back to the defaults. */
export function cleanLayout(v){
  const d=v&&typeof v==='object'?v:{};
  return {
    dividers:d.dividers===true,
    header:HEADERS.includes(d.header)?d.header:'none',
    footer:FOOTERS.includes(d.footer)?d.footer:'none',
    footerText:oneLine(d.footerText,FOOTER_WORDS_MAX),
    // Each block's icon in a column of its own at the left, and a slide's picture beside its script (src/guide-margin-layout.mjs)
    marginIcons:d.marginIcons===true,
    slideBeside:d.slideBeside===true
  };
}
export const isDefaultLayout=l=>{const c=cleanLayout(l);return !c.dividers&&c.header==='none'&&c.footer==='none'&&!c.footerText&&!c.marginIcons&&!c.slideBeside;};
/** One sentence on what the layout does, for the message after the Word file is made; '' for the default layout. */
export function describeLayout(layout){
  const l=cleanLayout(layout),said=[];
  if(l.dividers)said.push('each module starts on a new page');
  if(l.header==='title')said.push('the title of the guide runs along the top');
  if(l.header==='module')said.push('the name of the module runs along the top');
  if(l.footer==='plain')said.push('the page number is at the foot of each page');
  if(l.footer==='mirrored')said.push('the page number is on the outside of each page, for printing on both sides');
  if(l.footerText&&l.footer!=='none')said.push('your footer words are in the footer');
  if(l.marginIcons)said.push('the icon of each block is in a column at the left');
  if(l.slideBeside)said.push('the picture of each slide is beside its script');
  return said.length?`The layout you chose is in it: ${said.join('; ')}.`:'';
}

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_HEADER='application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const CT_FOOTER='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const SMALL='<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>';
const text=t=>t?`<w:r>${SMALL}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`:'';
const field=(instr,shown)=>`<w:r>${SMALL}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${SMALL}<w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r>${SMALL}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${SMALL}<w:t>${esc(shown)}</w:t></w:r><w:r>${SMALL}<w:fldChar w:fldCharType="end"/></w:r>`;
// A tab to the right margin, whatever the width of the page: the same footer serves portrait and landscape sections.
const RIGHT='<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>';
const part=(root,body)=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${root} xmlns:w="${W}" xmlns:r="${REL}">${body}</w:${root}>`;

/**
 * The header and footer parts a layout needs, how a section refers to them, and whether the settings must turn on separate odd and even pages.
 * `title` is the guide's title. A "front" header (empty) is made when the header names the module, so that the pages before the first module do not name one.
 * @returns {{parts:{name:string,type:string,xml:string}[], rels:{id:string,type:string,target:string}[], body:string, front:string, evenOdd:boolean, replaces:{header:boolean,footer:boolean}, hasFrontHeader:boolean}}
 */
export function layoutParts(layout,{title=''}={}){
  const l=cleanLayout(layout),heading=oneLine(title,200),parts=[],rels=[];
  const evenOdd=l.footer==='mirrored';
  const add=(kind,id,xml)=>{
    const name=`word/${kind}Layout${id}.xml`;
    parts.push({name:`/${name}`,type:kind==='header'?CT_HEADER:CT_FOOTER,xml});
    const rid=`rIdOsgLayout${kind==='header'?'H':'F'}${id}`;
    rels.push({id:rid,type:`${REL}/${kind}`,target:`${kind}Layout${id}.xml`});
    return rid;
  };
  const header=(jc)=>l.header==='module'
    ?`<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr>${field(`STYLEREF "${STYLE_NAMES.module}"`,heading)}</w:p>`
    :`<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr>${text(heading)}</w:p>`;
  const footer=(first)=>{
    const words=text(l.footerText),page=`${text('Page ')}${field('PAGE','1')}`;
    // On the page that has its outside on the right (odd pages, and every page when the footer is plain) the words are at the left and the number at the right;
    // on the other pages of a two-sided document the number is at the left and the words are at the right.
    return first?`<w:p>${words}${RIGHT}${page}</w:p>`:`<w:p>${page}${RIGHT}${words}</w:p>`;
  };
  const refs={body:'',front:''};
  const ref=(kind,type,rid)=>`<w:${kind}Reference w:type="${type}" r:id="${rid}"/>`;
  if(l.header!=='none'){
    const odd=add('header',1,part('hdr',header('right')));
    refs.body+=ref('header','default',odd);
    if(evenOdd)refs.body+=ref('header','even',add('header',2,part('hdr',header('left'))));
    if(l.header==='module'){
      const empty=part('hdr','<w:p/>');
      const f1=add('header',3,empty);
      refs.front+=ref('header','default',f1);
      if(evenOdd)refs.front+=ref('header','even',add('header',4,empty));
    }else refs.front=refs.body.split(/(?=<w:headerReference)/).filter(Boolean).join('');
  }
  if(l.footer!=='none'){
    const odd=add('footer',1,part('ftr',footer(true)));
    const f=ref('footer','default',odd)+(evenOdd?ref('footer','even',add('footer',2,part('ftr',footer(false)))):'');
    refs.body+=f;refs.front+=f;
  }
  return {parts,rels,body:refs.body,front:refs.front,evenOdd,replaces:{header:l.header!=='none',footer:l.footer!=='none'},hasFrontHeader:l.header==='module'};
}

// ---- applying a layout to the body of the Word document
const PARAGRAPH=/<w:p[ >][\s\S]*?<\/w:p>/g;
const styleOf=xml=>/<w:pStyle w:val="([^"]+)"/.exec(xml)?.[1]||'';
const plainText=xml=>Array.from(xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g),m=>m[1]).join('').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/\s+/g,' ').trim();
/** A paragraph that starts a new page: Word's own page-break-before, put where the paragraph's properties keep it (after the style, keep with next and keep lines). */
export function withPageBreakBefore(paragraphXml){
  if(/<w:pageBreakBefore\b/.test(paragraphXml))return paragraphXml;
  const m=/<w:pPr>\s*<w:pStyle [^>]*\/>(?:\s*<w:keepNext\/>)?(?:\s*<w:keepLines\/>)?/.exec(paragraphXml);
  if(m)return paragraphXml.replace(m[0],`${m[0]}<w:pageBreakBefore/>`);
  if(/<w:pPr>/.test(paragraphXml))return paragraphXml.replace('<w:pPr>','<w:pPr><w:pageBreakBefore/>');
  return paragraphXml.replace(/^<w:p(\s[^>]*)?>/,m2=>`${m2}<w:pPr><w:pageBreakBefore/></w:pPr>`);
}
const MODULE_IDS=new Set([STYLE_IDS.module,STYLE_IDS.assessment]);
/**
 * The body with the page breaks the layout asks for: each module and assessment starting a page when `dividers` is on, and the one-page summary (when the guide has
 * one) on a page of its own, so that it prints on one side. `summaryHeading` is the words of its heading.
 */
export function layoutBody(body,layout,{summaryHeading=''}={}){
  const l=cleanLayout(layout);
  let out=body;
  if(l.dividers)out=out.replace(PARAGRAPH,p=>MODULE_IDS.has(styleOf(p))?withPageBreakBefore(p):p);
  if(summaryHeading){
    let at=-1,len=0;
    for(const m of out.matchAll(PARAGRAPH))if(/^Heading\d$/.test(styleOf(m[0]))&&plainText(m[0])===summaryHeading){at=m.index;len=m[0].length;break;}
    if(at>=0){
      const headEnd=at+len,tableEnd=out.indexOf('</w:tbl>',headEnd);
      let tail=out.slice(headEnd);
      // The paragraph after the page's table starts the next page.
      if(tableEnd>=0){
        const from=tableEnd+'</w:tbl>'.length,m=/<w:p[ >][\s\S]*?<\/w:p>/.exec(out.slice(from));
        if(m)tail=out.slice(headEnd,from)+out.slice(from,from+m.index)+withPageBreakBefore(m[0])+out.slice(from+m.index+m[0].length);
      }
      out=out.slice(0,at)+withPageBreakBefore(out.slice(at,headEnd))+tail;
    }
  }
  return out;
}

// A section's page setup, written for the break before the first module when the header names the module.
const FRONT_SECTION=`<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p>`;
function withRefs(chunk,refs,replaces){
  return chunk.replace(/<w:sectPr(\s[^>]*)?>([\s\S]*?)<\/w:sectPr>/g,(all,attrs,inner)=>{
    let rest=inner;
    if(replaces.header)rest=rest.replace(/<w:headerReference\b[^>]*\/>/g,'');
    if(replaces.footer)rest=rest.replace(/<w:footerReference\b[^>]*\/>/g,'');
    return `<w:sectPr${attrs||''}>${refs}${rest}</w:sectPr>`;
  });
}
/**
 * The body with every section pointing at the layout's header and footer parts. When the header names the module, the sections before the first module point at an empty
 * header (and a section break is made there if the body has none), so that the front pages do not carry the name of a module that has not begun.
 */
export function bodyWithLayoutRefs(body,spec){
  if(!spec.parts.length)return body;
  if(!spec.hasFrontHeader)return withRefs(body,spec.body,spec.replaces);
  let first=-1;
  for(const m of body.matchAll(PARAGRAPH))if(MODULE_IDS.has(styleOf(m[0]))){first=m.index;break;}
  // No module at all: nothing begins, so nothing is named.
  if(first<0)return withRefs(body,spec.front,spec.replaces);
  let front=body.slice(0,first);
  const back=body.slice(first),closes=/<\/w:sectPr><\/w:pPr><\/w:p>\s*$/.test(front);
  // The front pages end in a section break of their own, unless they are empty or already end in one.
  if(!closes&&/<w:(p|tbl)[ >]/.test(front))front+=FRONT_SECTION;
  return withRefs(front,spec.front,spec.replaces)+withRefs(back,spec.body,spec.replaces);
}
/** The settings part with separate odd and even pages turned on (what mirrored headers and footers need), put where the schema keeps it. */
export function withEvenAndOdd(settingsXml){
  if(/<w:evenAndOddHeaders\b/.test(settingsXml))return settingsXml;
  const after=/<w:(defaultTabStop|autoHyphenation|consecutiveHyphenLimit|hyphenationZone|doNotHyphenateCaps|showEnvelope|summaryLength|clickAndTypeStyle|defaultTableStyle)\b[^>]*\/>/g;
  let last=null;for(const m of settingsXml.matchAll(after))last=m;
  if(last)return settingsXml.slice(0,last.index+last[0].length)+'<w:evenAndOddHeaders/>'+settingsXml.slice(last.index+last[0].length);
  if(/<w:compat>/.test(settingsXml))return settingsXml.replace('<w:compat>','<w:evenAndOddHeaders/><w:compat>');
  return settingsXml.replace('</w:settings>','<w:evenAndOddHeaders/></w:settings>');
}
