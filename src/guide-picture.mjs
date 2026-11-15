// Changing a picture that is in the guide: its description (what a screen reader says in place of it), whether it is only decoration, and how wide it is.
// The caption under a picture is a paragraph of its own (src/guide-reference.mjs), so it is changed there. Pure functions: the XML of one paragraph in, the XML out.
// Only a picture of the author's (one named "Picture n", made by Insert > Picture) is changed; the icons at the front of blocks are the program's own and are left alone.
import {PICTURE_WIDTHS} from './image-prepare.mjs';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const DEC_URI='{C183D7F6-B498-43B3-948B-1728B52AA6E4}';
const EMU=914400,MAX_HEIGHT=8*EMU;
export const DESCRIPTION_LIMIT=300;
const clean=t=>String(t??'').replace(/\s+/g,' ').trim();
const first=(root,ns,name)=>Array.from(root.getElementsByTagNameNS(ns,name))[0]||null;
const ofAuthor=d=>/^Picture\b/.test(first(d,WP,'docPr')?.getAttribute('name')||'');

function open(xml,io){
  const root=io.parse(xml).documentElement;
  if(!root||root.namespaceURI!==W||root.localName!=='p')throw new Error('That is not a paragraph.');
  return root;
}
const pictures=root=>Array.from(root.getElementsByTagNameNS(W,'drawing')).filter(ofAuthor);

/** The author's picture in a paragraph, or null when it holds none (or more than one, which is changed in Word). */
export function readPicture(xml,io){
  let root;
  try{root=open(xml||'',io);}catch{return null;}
  const found=pictures(root);
  if(found.length!==1)return null;
  const d=found[0],docPr=first(d,WP,'docPr'),extent=first(d,WP,'extent');
  const cx=Number(extent?.getAttribute('cx'))||0,cy=Number(extent?.getAttribute('cy'))||0;
  const decorative=Array.from(docPr.getElementsByTagNameNS(A,'ext')).some(e=>e.getAttribute('uri')===DEC_URI);
  const inches=cx/EMU;
  const near=PICTURE_WIDTHS.find(w=>Math.abs(w.inches-inches)<0.06);
  return {description:docPr.getAttribute('descr')||'',decorative,width:near?near.id:'',inches:Math.round(inches*100)/100,cx,cy};
}

/**
 * The paragraph with the picture described, marked as decoration or not, and set to one of the widths the program offers (its height keeps the picture's shape).
 * A picture that is not decoration needs a description.
 * With `maxCx` (in EMU) the picture is never made wider than that: a picture in a table cell is kept inside its cell.
 * @param {string} xml @param {{description?:string, decorative?:boolean, width?:string, maxCx?:number}} change
 */
export function setPicture(xml,{description,decorative,width,maxCx},io){
  const root=open(xml,io);
  const found=pictures(root);
  if(found.length!==1)throw new Error('This paragraph does not hold exactly one picture of yours, so it is changed in Word.');
  const d=found[0],docPr=first(d,WP,'docPr'),extent=first(d,WP,'extent');
  const alt=clean(description);
  if(!decorative&&!alt)throw new Error('Describe the picture in a few words for people who cannot see it, or tick "This picture is only decoration".');
  if(alt.length>DESCRIPTION_LIMIT)throw new Error(`The description is too long: it holds up to ${DESCRIPTION_LIMIT} letters.`);
  docPr.setAttribute('descr',decorative?'':alt);
  const marked=Array.from(docPr.getElementsByTagNameNS(A,'extLst'));
  if(decorative&&!Array.from(docPr.getElementsByTagNameNS(A,'ext')).some(e=>e.getAttribute('uri')===DEC_URI)){
    const holder=io.parse(`<root xmlns:a="${A}"><a:extLst><a:ext uri="${DEC_URI}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst></root>`).documentElement.firstChild;
    docPr.appendChild(root.ownerDocument.importNode(holder,true));
  }else if(!decorative)for(const e of marked)if(Array.from(e.getElementsByTagNameNS(A,'ext')).some(x=>x.getAttribute('uri')===DEC_URI))docPr.removeChild(e);
  if(width){
    const w=PICTURE_WIDTHS.find(x=>x.id===width);
    if(!w)throw new Error('Choose one of the widths in the list.');
    const cx0=Number(extent.getAttribute('cx'))||0,cy0=Number(extent.getAttribute('cy'))||0;
    if(cx0>0&&cy0>0){
      let cx=Math.round(w.inches*EMU),cy=Math.round(cx*cy0/cx0);
      if(cy>MAX_HEIGHT){cx=Math.round(cx*MAX_HEIGHT/cy);cy=MAX_HEIGHT;}
      if(maxCx&&cx>maxCx){cy=Math.round(cy*maxCx/cx);cx=maxCx;}
      extent.setAttribute('cx',String(cx));extent.setAttribute('cy',String(cy));
      for(const e of Array.from(d.getElementsByTagNameNS(A,'ext')))if(e.hasAttribute('cx')&&e.hasAttribute('cy')){e.setAttribute('cx',String(cx));e.setAttribute('cy',String(cy));}
    }
  }
  return io.serialize(root);
}
