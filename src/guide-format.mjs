// Formatting a paragraph by hand: its alignment, the space around it, its line spacing and its indents. Pure functions on Word-format XML.
// What a paragraph's style says is not changed here (see guide-styles-edit.mjs for that); this is the direct formatting a paragraph carries
// over its style, which Word shows and keeps. A value left out is left as it is; null takes the paragraph's own setting away, so it follows
// its style again.
//
// The order of the properties in a paragraph's properties is the order Word's schema gives; Word may refuse a file that has them in another.
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PPR_ORDER=['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];
export const ALIGNMENTS=Object.freeze({left:'left',center:'center',right:'right',justify:'both'});
const FROM_WORD=Object.freeze({left:'left',start:'left',center:'center',right:'right',end:'right',both:'justify'});
export const FORMAT_LIMITS=Object.freeze({space:[0,144],line:[0.8,3],indent:[0,216],first:[-72,72]});
// Paragraphs whose spacing the program writes to make them what they are (a line to write on, a box of a set height, a row of fields).
const STRUCTURE=[/^OSG Form Line$/,/^OSG Form Box$/,/^OSG Form Field$/,/^OSG Answer Line$/,/^OSG Time$/,/^TOC \d$/i];

const kids=el=>Array.from(el.childNodes).filter(n=>n.nodeType===1);
const named=(el,name)=>kids(el).find(n=>n.localName===name&&n.namespaceURI===W);
const attr=(el,name)=>el?.getAttributeNS(W,name)||'';
const twips=pt=>Math.round(pt*20);
const points=t=>Math.round(Number(t)/20*100)/100;

/** Whether this paragraph's formatting can be changed by hand, and if not, why. */
export function canFormat(xml,style,io){
  if(STRUCTURE.some(r=>r.test(style||'')))return {ok:false,reason:'The program sets the spacing of this part, so it is left as it is.'};
  if(!xml||!/^<w:p[\s>/]/.test(xml))return {ok:false,reason:'Only paragraphs of words can be formatted by hand here. A table is changed with its own commands.'};
  let root=null;
  try{root=io.parse(xml).documentElement;}catch{root=null;}
  if(!root||root.namespaceURI!==W||root.localName!=='p')return {ok:false,reason:'This part could not be read.'};
  return {ok:true};
}

/** The paragraph's own formatting as an element reads it: each value, or null where the paragraph follows its style. */
export function formatOf(p){
  const pPr=named(p,'pPr');
  const out={alignment:null,spaceBefore:null,spaceAfter:null,lineSpacing:null,leftIndent:null,firstLineIndent:null};
  if(!pPr)return out;
  const jc=named(pPr,'jc');
  if(jc)out.alignment=FROM_WORD[attr(jc,'val')]??null;
  const sp=named(pPr,'spacing');
  if(sp){
    if(attr(sp,'before')!=='')out.spaceBefore=points(attr(sp,'before'));
    if(attr(sp,'after')!=='')out.spaceAfter=points(attr(sp,'after'));
    if(attr(sp,'line')!==''&&(attr(sp,'lineRule')||'auto')==='auto')out.lineSpacing=Math.round(Number(attr(sp,'line'))/240*100)/100;
  }
  const ind=named(pPr,'ind');
  if(ind){
    const left=attr(ind,'left')||attr(ind,'start');
    if(left!=='')out.leftIndent=points(left);
    if(attr(ind,'hanging')!=='')out.firstLineIndent=-points(attr(ind,'hanging'));
    else if(attr(ind,'firstLine')!=='')out.firstLineIndent=points(attr(ind,'firstLine'));
  }
  return out;
}
export const readParagraphFormat=(xml,io)=>formatOf(io.parse(xml).documentElement);

function number(v,[lo,hi],what){
  if(typeof v!=='number'||!Number.isFinite(v)||v<lo||v>hi)throw new Error(`${what} must be a number from ${lo} to ${hi}.`);
  return Math.round(v*100)/100;
}

// Put an element among a paragraph's properties where the schema puts it.
function place(pPr,el){
  const rank=n=>{const i=PPR_ORDER.indexOf(n.localName);return i<0?PPR_ORDER.length:i;};
  const mine=rank(el);
  const after=kids(pPr).find(n=>rank(n)>mine);
  pPr.insertBefore(el,after||null);
}

/**
 * The paragraph with other formatting. fmt: {alignment:'left'|'center'|'right'|'justify'|null, spaceBefore, spaceAfter (points), lineSpacing (a
 * multiple of one line), leftIndent, firstLineIndent (points; negative hangs)}. Nothing else in the paragraph changes.
 */
export function withParagraphFormat(xml,style,fmt,io){
  const can=canFormat(xml,style,io);
  if(!can.ok)throw new Error(can.reason);
  const root=io.parse(xml).documentElement;
  const doc=root.ownerDocument;
  let pPr=named(root,'pPr');
  const need=()=>{if(!pPr){pPr=doc.createElementNS(W,'w:pPr');root.insertBefore(pPr,root.firstChild);}return pPr;};
  const el=name=>{let e=pPr&&named(pPr,name);if(!e){e=doc.createElementNS(W,`w:${name}`);place(need(),e);}return e;};
  const drop=name=>{const e=pPr&&named(pPr,name);if(e)pPr.removeChild(e);};
  const set=(e,name,value)=>{if(value===null)e.removeAttributeNS(W,name);else e.setAttributeNS(W,`w:${name}`,String(value));};

  if('alignment' in fmt){
    if(fmt.alignment===null)drop('jc');
    else{
      if(!ALIGNMENTS[fmt.alignment])throw new Error('Alignment must be left, centered, right or justified.');
      el('jc').setAttributeNS(W,'w:val',ALIGNMENTS[fmt.alignment]);
    }
  }
  const spacing=('spaceBefore' in fmt)||('spaceAfter' in fmt)||('lineSpacing' in fmt);
  if(spacing){
    const sp=el('spacing');
    if('spaceBefore' in fmt)set(sp,'before',fmt.spaceBefore===null?null:twips(number(fmt.spaceBefore,FORMAT_LIMITS.space,'Space before')));
    if('spaceAfter' in fmt)set(sp,'after',fmt.spaceAfter===null?null:twips(number(fmt.spaceAfter,FORMAT_LIMITS.space,'Space after')));
    if('lineSpacing' in fmt){
      const rule=attr(sp,'lineRule');
      if(rule&&rule!=='auto'&&fmt.lineSpacing!==null)throw new Error('This paragraph has a line height of a set size, so its line spacing is left as it is.');
      if(fmt.lineSpacing===null){set(sp,'line',null);set(sp,'lineRule',null);}
      else{set(sp,'line',Math.round(number(fmt.lineSpacing,FORMAT_LIMITS.line,'Line spacing')*240));set(sp,'lineRule','auto');}
    }
    if(!Array.from(sp.attributes||[]).length)pPr.removeChild(sp);
  }
  if(('leftIndent' in fmt)||('firstLineIndent' in fmt)){
    const ind=el('ind');
    if('leftIndent' in fmt){set(ind,'left',fmt.leftIndent===null?null:twips(number(fmt.leftIndent,FORMAT_LIMITS.indent,'The indent')));set(ind,'start',null);}
    if('firstLineIndent' in fmt){
      set(ind,'hanging',null);set(ind,'firstLine',null);
      if(fmt.firstLineIndent!==null){
        const v=number(fmt.firstLineIndent,FORMAT_LIMITS.first,'The first line indent');
        if(v<0)set(ind,'hanging',twips(-v));else if(v>0)set(ind,'firstLine',twips(v));
      }
    }
    if(!Array.from(ind.attributes||[]).length)pPr.removeChild(ind);
  }
  if(pPr&&!kids(pPr).length)root.removeChild(pPr);
  return io.serialize(root);
}
