// Changing the words of one paragraph of the guide, without disturbing anything else in it. Pure functions: Word-format XML in, Word-format
// XML out, so they are tested without a page. The program's editor (program/ui/editor.js) is the one that calls them.
//
// A paragraph the editor can change reads as a lead that is shown but not typed, then the words that are typed:
//
//   <icon> <label> <tab>  words          a cue line:      the cue's picture, its name, a tab, then the author's words
//   <label>               words          notes:           "Facilitator notes: " in front of the words
//   "Module 1: "          title          a module or lesson heading: the number is written by Update guide
//                         words          everything else: just the words
//
// Anything else in a paragraph (a picture of the author's, a check box, a citation, a bookmark, a field, a line break, a hyperlink) is not
// words. A paragraph that holds any of it is not changed by the editor in this version, and is changed with the tools instead; it is never
// rewritten by the editor, so nothing in it can be lost. Paragraphs the program writes itself (the time line, form lines and boxes, the
// contents) are locked for the same reason.
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
export const MAX_WORDS=100000;
const LABEL_STYLE='OSGCueLabel';

// Styles whose words the program writes. Their paragraphs are shown, not typed in.
const LOCKED=[/^OSG Time$/,/^OSG Form Line$/,/^OSG Answer Line$/,/^OSG Form Box$/,/^OSG Form Field$/,/^TOC \d$/i];
// Headings: one line, a number the program writes, and no new paragraph made by pressing Enter in them.
const HEADINGS=new Set(['OSG Guide Title','OSG Module','OSG Lesson','OSG Assessment','OSG Form Title','OSG Form Title Facilitator','heading 1','heading 2','heading 3','Title']);
const PREFIX={'OSG Module':/^Module \d+: /,'OSG Lesson':/^Lesson \d+\.\d+: /};
// Where a new paragraph goes on after a cue line: the cue's own words are one line, the rest of a block is in the body style (and goes on as it is).
const AFTER_CUE=new Set(['OSG Cue','OSG Cue Facilitator']);

export const isHeading=style=>HEADINGS.has(style);
export const isLocked=style=>LOCKED.some(r=>r.test(style||''));

const kids=el=>Array.from(el.childNodes).filter(n=>n.nodeType===1);
const named=(el,name)=>kids(el).find(n=>n.localName===name&&n.namespaceURI===W);
const attr=(el,name)=>el?.getAttributeNS(W,name)||'';

// Words that go into XML text: no characters XML cannot hold, no line ends (a paragraph is one line of words; Enter makes the next).
export function cleanWords(words){
  if(typeof words!=='string')throw new Error('Type the words as text.');
  const text=words.replace(/\r\n|\r|\n|\u2028|\u2029/g,' ').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g,'');
  if(text.length>MAX_WORDS)throw new Error(`A paragraph can hold up to ${MAX_WORDS.toLocaleString('en-US')} characters.`);
  return text;
}

const escText=s=>s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

// What kind of inline piece is this child of a paragraph?
function pieceOf(child){
  if(child.namespaceURI!==W||child.localName!=='r')return {kind:'other'};
  const parts=kids(child).filter(c=>c.localName!=='rPr');
  const rPr=named(child,'rPr');
  if(parts.length===1&&parts[0].localName==='drawing'){
    // The cue's own picture sits in line with the text, lowered to its middle and named "Icon n"; the author's pictures are named "Picture n".
    const docPr=Array.from(parts[0].getElementsByTagNameNS(WP,'docPr'))[0];
    return /^Icon\b/.test(docPr?.getAttribute('name')||'')?{kind:'icon'}:{kind:'other'};
  }
  if(parts.length===1&&parts[0].localName==='tab')return {kind:'tab'};
  if(parts.length&&parts.every(p=>p.localName==='t')){
    const text=parts.map(p=>p.textContent??'').join('');
    return attr(rPr&&named(rPr,'rStyle'),'val')===LABEL_STYLE?{kind:'label',text}:{kind:'text',text,rPr};
  }
  return {kind:'other'};
}

/**
 * Read a paragraph into its lead and its words.
 * @param {string} xml the paragraph's Word-format XML
 * @param {string} style the paragraph's style name
 * @param {{parse:(xml:string)=>Document}} io
 * @returns {{editable:boolean, reason?:string, lead:{kind:string,text?:string}[], text:string, prefix:string}}
 */
export function readParagraph(xml,style,{parse}){
  const no=reason=>({editable:false,reason,lead:[],text:'',prefix:''});
  if(isLocked(style))return no('The program writes this part. Update guide keeps it right.');
  if(!xml||!xml.trim())return no('This part is made by the program.');
  let root;
  try{root=parse(xml).documentElement;}catch{return no('This part could not be read.');}
  if(!root||root.namespaceURI!==W||root.localName!=='p')return no('This part is a table. Its cells are changed with the tools beside the guide.');
  const lead=[];
  const words=[];
  let sawText=false;
  for(const child of kids(root)){
    if(child.namespaceURI===W&&child.localName==='pPr')continue;
    const piece=pieceOf(child);
    if(piece.kind==='icon'||piece.kind==='tab'||piece.kind==='label'){
      if(sawText)return no('This part holds more than words, so it is changed with the tools beside the guide.');
      lead.push(piece.kind==='label'?{kind:'label',text:piece.text}:{kind:piece.kind});
    }else if(piece.kind==='text'){
      sawText=true;
      words.push(piece.text);
    }else return no('This part holds more than words (a picture, a check box, a citation, a bookmark or a field), so it is changed with the tools beside the guide.');
  }
  let text=words.join('');
  let prefix='';
  const rx=PREFIX[style];
  const m=rx&&rx.exec(text);
  if(m){prefix=m[0];text=text.slice(prefix.length);}
  return {editable:true,lead,text,prefix};
}

const serialize=(root,io)=>io.serialize(root);

// The root of a paragraph and its parts, ready to be changed.
function open(xml,style,io){
  const info=readParagraph(xml,style,io);
  if(!info.editable)throw new Error(info.reason);
  const root=io.parse(xml).documentElement;
  const textRuns=kids(root).filter(c=>pieceOf(c).kind==='text');
  return {info,root,textRuns};
}

function newRun(doc,text,rPr){
  const run=doc.createElementNS(W,'w:r');
  if(rPr)run.appendChild(rPr.cloneNode(true));
  const t=doc.createElementNS(W,'w:t');
  t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');
  t.appendChild(doc.createTextNode(text));
  run.appendChild(t);
  return run;
}

/**
 * The paragraph with other words. The lead, the paragraph's properties and the look of its first run stay; nothing else is touched. When
 * the words are the same, the very same XML comes back.
 * @param {string} xml @param {string} style @param {string} words @param {{parse:(xml:string)=>Document, serialize:(el:Element)=>string}} io
 */
export function setWords(xml,style,words,io){
  const clean=cleanWords(words);
  const {info,root,textRuns}=open(xml,style,io);
  if(clean===info.text)return xml;
  const doc=root.ownerDocument;
  const rPr=textRuns.map(r=>named(r,'rPr')).find(Boolean)||null;
  const at=textRuns.length?textRuns[0]:null;
  const full=info.prefix+clean;
  const run=full?newRun(doc,full,rPr):null;
  if(at){
    if(run)root.insertBefore(run,at);
    for(const r of textRuns)root.removeChild(r);
  }else if(run)root.appendChild(run);
  return serialize(root,io);
}

/** The words of an editable paragraph, or null. */
export function wordsOf(xml,style,io){
  const info=readParagraph(xml,style,io);
  return info.editable?info.text:null;
}

// The new paragraph that Enter makes after a paragraph: a block's cue line goes on in the body style; everything else goes on as it was
// (a note in the notes style, a list item in the list). Headings do not split.
export function nextParagraphSpec(xml,style,io){
  if(isHeading(style)||isLocked(style))return null;
  const root=io.parse(xml).documentElement;
  const pPr=named(root,'pPr');
  const pStyle=named(pPr||{childNodes:[]},'pStyle');
  const numPr=pPr&&named(pPr,'numPr');
  if(AFTER_CUE.has(style))return {styleId:'OSGCueBody',numPr:null};
  return {styleId:pStyle?attr(pStyle,'val'):null,numPr:numPr?io.serialize(numPr):null};
}

/** A paragraph that holds only words, in the style (and list) a spec names. */
export function paragraphWith(spec,words,{parse,serialize:ser}){
  const text=cleanWords(words);
  const pPr=spec.styleId||spec.numPr?`<w:pPr>${spec.styleId?`<w:pStyle w:val="${escText(spec.styleId)}"/>`:''}${spec.numPr||''}</w:pPr>`:'';
  const run=text?`<w:r><w:t xml:space="preserve">${escText(text)}</w:t></w:r>`:'';
  const root=parse(`<w:p xmlns:w="${W}">${pPr}${run}</w:p>`).documentElement;
  return ser(root);
}

/**
 * Enter in the middle of the words: the paragraph keeps the words before the caret, a new paragraph takes the rest. Null when this paragraph
 * cannot be split (a heading, a locked or a not-words paragraph).
 * @returns {{first:string, second:string}|null}
 */
export function splitWords(xml,style,offset,io){
  const info=readParagraph(xml,style,io);
  if(!info.editable||isHeading(style))return null;
  const spec=nextParagraphSpec(xml,style,io);
  if(!spec)return null;
  const at=Math.max(0,Math.min(info.text.length,offset));
  return {first:setWords(xml,style,info.text.slice(0,at),io),second:paragraphWith(spec,info.text.slice(at),io)};
}

/**
 * Backspace at the start of a paragraph: its words join the end of the one before. Null when they cannot be joined (either is not just words,
 * the paragraph has a lead of its own, or the one before is a heading).
 * @returns {{xml:string, offset:number}|null}
 */
export function joinWords(previous,current,io){
  const a=readParagraph(previous.xml,previous.style,io),b=readParagraph(current.xml,current.style,io);
  if(!a.editable||!b.editable||isHeading(previous.style)||b.lead.length)return null;
  if(isHeading(current.style))return null;
  return {xml:setWords(previous.xml,previous.style,a.text+b.text,io),offset:a.text.length};
}
