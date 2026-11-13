// Formatting inside a sentence: the words of a paragraph as spans, each a stretch of words with one look. Pure functions on plain data, so they
// are tested without a page.
//
// A span is {text, rpr}: the words, and the properties of the Word run they are in as a self-contained piece of XML (the run's <w:rPr>, or ''
// for a run with none). Keeping the properties as Word wrote them means everything the program does not model (the language, a ligature
// setting, a theme font) goes through an edit untouched; only the settings the author changes are rewritten, and in the order Word's schema gives.
//
// The settings the author can change, which a "change" names (a value of null takes a setting off, so the paragraph's style decides again):
//   bold, italic, underline, strike   true | false | null
//   script                            'superscript' | 'subscript' | 'baseline' | null
//   color                             '#RRGGBB' | null
//   font                              a font name | null
//   size                              points, 6 to 96 | null
//   style                             the id of a character style | null
import {RPR_ORDER,STYLE_LIMITS,specCss} from './guide-styles-edit.mjs';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const COLOR=/^#[0-9A-Fa-f]{6}$/;
const FONT=/^[\p{L}\p{N} .,'()+-]{1,64}$/u;
const ID=/^[A-Za-z][A-Za-z0-9]{0,63}$/;
export const RUN_SETTINGS=Object.freeze(['bold','italic','underline','strike','script','color','font','size','style']);
export const SCRIPTS=Object.freeze(['superscript','subscript','baseline']);

const kids=el=>Array.from(el.childNodes).filter(n=>n.nodeType===1);
const named=(el,name)=>kids(el).find(n=>n.localName===name&&n.namespaceURI===W);
const attr=(el,name)=>el?.getAttributeNS(W,name)||'';
const isOn=el=>!!el&&!['false','0','off'].includes((attr(el,'val')||'1').toLowerCase());
const bad=message=>{throw new Error(message);};
const WRAP=`<w:rPr xmlns:w="${W}"`;

// ---- run properties as text

/** A run's properties as the text a span holds: no namespace declaration for w, and '' when there are none. */
export function rprText(rPr,io){
  if(!rPr||!kids(rPr).length)return '';
  return io.serialize(rPr).replace(` xmlns:w="${W}"`,'');
}
/** The properties element for a span's text (a new empty one for ''), in a document of its own. */
export function parseRpr(rpr,io){
  const xml=rpr?rpr.replace(/^<w:rPr/,WRAP):`${WRAP}/>`;
  const root=io.parse(xml).documentElement;
  if(!root||root.namespaceURI!==W||root.localName!=='rPr')bad('A run\'s properties are not in the right form.');
  return root;
}

/** The settings a run's properties hold, in the form of a change (only what is set). */
export function fmtOfRpr(rpr,io){
  if(!rpr)return {};
  return fmtOfEl(parseRpr(rpr,io));
}
/** The same, from the properties element itself (a run's <w:rPr>, or null). */
export function fmtOfEl(rPr){
  const fmt={};
  if(!rPr)return fmt;
  for(const [key,name] of [['bold','b'],['italic','i'],['strike','strike']]){const e=named(rPr,name);if(e)fmt[key]=isOn(e);}
  const u=named(rPr,'u');
  if(u)fmt.underline=(attr(u,'val')||'single')!=='none';
  const va=named(rPr,'vertAlign');
  if(va&&SCRIPTS.includes(attr(va,'val')))fmt.script=attr(va,'val');
  const color=named(rPr,'color');
  if(color&&/^[0-9A-Fa-f]{6}$/.test(attr(color,'val')))fmt.color=`#${attr(color,'val').toUpperCase()}`;
  const fonts=named(rPr,'rFonts');
  if(fonts&&attr(fonts,'ascii'))fmt.font=attr(fonts,'ascii');
  const sz=named(rPr,'sz');
  if(sz&&/^\d+$/.test(attr(sz,'val')))fmt.size=Number(attr(sz,'val'))/2;
  const style=named(rPr,'rStyle');
  if(style&&attr(style,'val'))fmt.style=attr(style,'val');
  return fmt;
}

/** A change, checked: a message says what is wrong with it. Unknown settings are refused. */
export function cleanChange(change){
  if(change===null||typeof change!=='object'||Array.isArray(change))bad('The formatting is not in the right form.');
  const out={};
  for(const [k,v] of Object.entries(change)){
    if(!RUN_SETTINGS.includes(k))bad(`${k} is not a setting the words can have.`);
    if(v===undefined)continue;
    if(v===null){out[k]=null;continue;}
    if(['bold','italic','underline','strike'].includes(k)){if(typeof v!=='boolean')bad(`${k} is on or off.`);out[k]=v;}
    else if(k==='script'){if(!SCRIPTS.includes(v))bad('Raised or lowered text is superscript, subscript or neither.');out[k]=v;}
    else if(k==='color'){if(typeof v!=='string'||!COLOR.test(v))bad('A color is six hexadecimal digits after a #, for example #1F3A5F.');out[k]=v.toUpperCase();}
    else if(k==='font'){if(typeof v!=='string'||!FONT.test(v))bad('A font name can have letters, digits, spaces and . , \' ( ) + - only, up to 64 characters.');out[k]=v;}
    else if(k==='size'){const [lo,hi]=STYLE_LIMITS.size;if(typeof v!=='number'||!Number.isFinite(v)||v<lo||v>hi)bad(`The font size must be a number from ${lo} to ${hi}.`);out[k]=Math.round(v*2)/2;}
    else if(k==='style'){if(typeof v!=='string'||!ID.test(v))bad('That character style is not known.');out[k]=v;}
  }
  return out;
}

function place(parent,el){
  const rank=n=>{const i=RPR_ORDER.indexOf(n.localName);return i<0?RPR_ORDER.length:i;};
  const mine=rank(el);
  parent.insertBefore(el,kids(parent).find(n=>rank(n)>mine)||null);
}
const drop=(parent,...names)=>{for(const n of names){for(const e of kids(parent).filter(c=>c.localName===n&&c.namespaceURI===W))parent.removeChild(e);}};
function put(rPr,name,attrs={}){
  const doc=rPr.ownerDocument;
  let e=named(rPr,name);
  if(!e){e=doc.createElementNS(W,`w:${name}`);place(rPr,e);}
  for(const a of Array.from(e.attributes||[]).map(x=>x.name))e.removeAttribute(a);
  for(const [k,v] of Object.entries(attrs))e.setAttributeNS(W,`w:${k}`,v);
  return e;
}

/** The run properties with a change laid over them, as text. Nothing else in them is touched. */
export function withFmt(rpr,change,io){
  const c=cleanChange(change);
  if(!Object.keys(c).length)return rpr;
  const rPr=parseRpr(rpr,io);
  for(const [k,v] of Object.entries(c)){
    if(['bold','italic'].includes(k)){
      const names=k==='bold'?['b','bCs']:['i','iCs'];
      if(v===null)drop(rPr,...names);else for(const n of names)put(rPr,n,v?{}:{val:'0'});
    }else if(k==='strike'){
      if(v===null)drop(rPr,'strike');else put(rPr,'strike',v?{}:{val:'0'});
    }else if(k==='underline'){
      if(v===null)drop(rPr,'u');else put(rPr,'u',{val:v?'single':'none'});
    }else if(k==='script'){
      if(v===null)drop(rPr,'vertAlign');else put(rPr,'vertAlign',{val:v});
    }else if(k==='color'){
      if(v===null)drop(rPr,'color');else put(rPr,'color',{val:v.slice(1)});
    }else if(k==='font'){
      if(v===null)drop(rPr,'rFonts');else put(rPr,'rFonts',{ascii:v,hAnsi:v,eastAsia:v,cs:v});
    }else if(k==='size'){
      if(v===null)drop(rPr,'sz','szCs');else{put(rPr,'sz',{val:String(Math.round(v*2))});put(rPr,'szCs',{val:String(Math.round(v*2))});}
    }else if(k==='style'){
      if(v===null)drop(rPr,'rStyle');else put(rPr,'rStyle',{val:v});
    }
  }
  return rprText(rPr,io);
}

/** What a run's settings look like on a page, as CSS properties for an element. @param {Record<string, any>} fmt */
export function runCss(fmt){
  const css=specCss({font:{...(fmt.font?{name:fmt.font}:{}),...(fmt.size?{size:fmt.size}:{}),...(fmt.color?{color:fmt.color}:{}),...(fmt.bold!==undefined?{bold:fmt.bold}:{}),...(fmt.italic!==undefined?{italic:fmt.italic}:{}),...(fmt.underline!==undefined?{underline:fmt.underline}:{}),...(fmt.strike!==undefined?{strike:fmt.strike}:{})}});
  if(fmt.script==='superscript'||fmt.script==='subscript'){css.verticalAlign=fmt.script==='superscript'?'super':'sub';if(!fmt.size)css.fontSize='smaller';}
  return css;
}

// ---- spans

/** @typedef {{text:string, rpr:string}} Span */

export const textOf=spans=>spans.map(s=>s.text).join('');
/** The spans with the empty ones left out and neighbours with the same look made one. */
export function normalize(spans){
  const out=[];
  for(const s of spans){
    if(!s.text)continue;
    const last=out[out.length-1];
    if(last&&last.rpr===s.rpr)last.text+=s.text;else out.push({text:s.text,rpr:s.rpr});
  }
  return out;
}
export const sameSpans=(a,b)=>{const x=normalize(a),y=normalize(b);return x.length===y.length&&x.every((s,i)=>s.text===y[i].text&&s.rpr===y[i].rpr);};

/** The part of the spans from `start` to `end` (offsets in characters). */
export function slice(spans,start,end=Infinity){
  const out=[];
  let at=0;
  for(const s of spans){
    const from=Math.max(start,at),to=Math.min(end,at+s.text.length);
    if(to>from)out.push({text:s.text.slice(from-at,to-at),rpr:s.rpr});
    at+=s.text.length;
  }
  return out;
}

/** The properties of the character before `offset`, or of the first character when there is none before (as typing in a word processor goes on). */
export function rprAt(spans,offset){
  const parts=spans.filter(s=>s.text);
  if(!parts.length)return '';
  if(offset<=0)return parts[0].rpr;
  let at=0;
  for(const s of parts){at+=s.text.length;if(offset<=at)return s.rpr;}
  return parts[parts.length-1].rpr;
}

/**
 * Other words in a stretch of the spans. New words take the look of the character before them (or what `rpr` says); words that go keep no trace.
 * @param {Span[]} spans @param {number} start @param {number} end @param {string} text @param {string} [rpr]
 */
export function replaceText(spans,start,end,text,rpr){
  const look=rpr!==undefined?rpr:rprAt(spans,start);
  return normalize([...slice(spans,0,start),...(text?[{text,rpr:look}]:[]),...slice(spans,end)]);
}

/** The spans with a change laid over the settings of the characters from `start` to `end`. */
export function applyChange(spans,start,end,change,io){
  const c=cleanChange(change);
  if(!Object.keys(c).length||end<=start)return normalize(spans);
  const cache=new Map();
  const changed=rpr=>{if(!cache.has(rpr))cache.set(rpr,withFmt(rpr,c,io));return cache.get(rpr);};
  return normalize([...slice(spans,0,start),...slice(spans,start,end).map(s=>({text:s.text,rpr:changed(s.rpr)})),...slice(spans,end)]);
}

/**
 * What the characters from `start` to `end` are like, setting by setting: the value they all have, or 'mixed' when they differ, or undefined when
 * none of them says. (With no characters, the look of the character before the caret.)
 */
export function summarize(spans,start,end,io){
  const parts=end>start?slice(spans,start,end):[{text:'',rpr:rprAt(spans,start)}];
  const out={};
  const fmts=parts.map(p=>fmtOfRpr(p.rpr,io));
  for(const k of RUN_SETTINGS){
    const values=fmts.map(f=>f[k]);
    out[k]=values.every(v=>v===values[0])?values[0]:'mixed';
  }
  return out;
}
