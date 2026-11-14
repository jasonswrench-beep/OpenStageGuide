// Cleans the Word XML of one guide paragraph (or table) of everything that could make Word reach outside the file when it is opened or when
// its fields are updated: fields other than the program's own contents, page number and page count; links; legacy and embedded objects;
// attributes that carry an address; and drawings that are not plain inline pictures.
//
// A guide file is data a person may be sent. Its paragraphs are kept as Word XML and written into the Word files the program makes, and the program
// tells the person to press Ctrl+A and F9 in Word, which updates every field. A field such as INCLUDEPICTURE, INCLUDETEXT, LINK or a
// HYPERLINK to a file or a network path would then be followed by Word, which can send the person's Windows sign-in to a remote computer. So
// such things never stay in a paragraph: they are taken out when a guide is opened and again when a Word file is made.
//
// The text of what is taken out stays where it can (a field keeps the words of its result, a link keeps its words). The program's own
// paragraphs hold none of these, so they pass through untouched, byte for byte (a test keeps that true).
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC='http://schemas.openxmlformats.org/drawingml/2006/picture';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The only fields a guide may keep: the program's own contents table, and the page number and page count. */
export const ALLOWED_FIELDS=/^\s*(TOC|PAGE|NUMPAGES)\b/i;
/** An instruction that names an address, whatever field it is in, is never kept. */
const ADDRESS=/[a-z][a-z0-9+.-]*:\/\/|\\\\|file:|\bmailto:/i;
export const fieldAllowed=instruction=>ALLOWED_FIELDS.test(String(instruction||''))&&!ADDRESS.test(String(instruction||''));

// Elements that can pull in or run something from outside, wherever they are in the paragraph.
const DROP=new Set(['altChunk','subDoc','object','pict','control','dataBinding','delInstrText','txbxContent','OLEObject','imagedata']);
// Attributes (in any namespace) that hold an address.
const ADDRESS_ATTRIBUTES=new Set(['href','link','src','Target']);

const kids=el=>Array.from(el.childNodes||[]).filter(n=>n.nodeType===1);
const isW=(el,name)=>el.namespaceURI===W&&el.localName===name;

function unwrap(el){
  const parent=el.parentNode;
  while(el.firstChild)parent.insertBefore(el.firstChild,el);
  parent.removeChild(el);
}

/** Take a field's control elements out; the run that held only one of them goes too. */
function removeControl(node){
  const run=node.parentNode;
  run.removeChild(node);
  if(run&&isW(run,'r')&&!kids(run).some(k=>!isW(k,'rPr')))run.parentNode?.removeChild(run);
}

/** Every complex field in document order, with the elements that make it and the instruction it holds. */
function complexFields(root){
  const controls=[];
  const walk=el=>{
    for(const k of kids(el)){
      if(k.namespaceURI===W&&(k.localName==='fldChar'||k.localName==='instrText'))controls.push(k);
      else walk(k);
    }
  };
  walk(root);
  const fields=[],stray=[],stack=[];
  for(const node of controls){
    if(node.localName==='instrText'){
      const top=stack.at(-1);
      if(top&&top.phase==='instruction'){top.instruction+=node.textContent||'';top.nodes.push(node);}else stray.push(node);
      continue;
    }
    const type=node.getAttributeNS(W,'fldCharType')||node.getAttribute('w:fldCharType');
    if(type==='begin')stack.push({nodes:[node],instruction:'',phase:'instruction'});
    else if(type==='separate'){const top=stack.at(-1);if(top){top.phase='result';top.nodes.push(node);}else stray.push(node);}
    else if(type==='end'){const top=stack.pop();if(top){top.nodes.push(node);fields.push(top);}else stray.push(node);}
    else stray.push(node);
  }
  // A field that is never closed cannot be told apart from one made to hide what it does.
  for(const open of stack)fields.push({...open,instruction:'(unfinished field)'});
  return {fields,stray};
}

function isInlinePicture(drawing){
  const inline=kids(drawing).find(k=>k.namespaceURI===WP&&k.localName==='inline');
  if(!inline||kids(drawing).some(k=>k.namespaceURI===WP&&k.localName==='anchor'))return false;
  const data=Array.from(inline.getElementsByTagNameNS(A,'graphicData'))[0];
  return !!data&&data.getAttribute('uri')===PIC&&inline.getElementsByTagNameNS(PIC,'pic').length===1;
}

/**
 * Clean the XML of one paragraph or table. Returns the XML (the very same text when nothing needed to go) and how many things were taken out.
 * @param {string} xml @param {{parse:(x:string)=>Document, serialize:(e:Element)=>string}} io
 * @returns {{xml:string, removed:number, kinds:string[]}}
 */
export function cleanParagraphXml(xml,io){
  const text=String(xml??'');
  if(!text.trim())return {xml:text,removed:0,kinds:[]};
  // Quick look first: most paragraphs hold none of what is cleaned, and are not parsed again.
  if(!/instrText|fldChar|fldSimple|hyperlink|altChunk|subDoc|<[a-z]*:?(object|pict|control|dataBinding|OLEObject|imagedata|txbxContent|delInstrText)\b|\b(href|link|src|Target)=|drawing|hlink/i.test(text))return {xml:text,removed:0,kinds:[]};
  let doc;
  try{doc=io.parse(text);}catch{return {xml:text,removed:0,kinds:[]};}
  const root=doc?.documentElement;
  if(!root||doc.getElementsByTagName('parsererror').length)return {xml:text,removed:0,kinds:[]};
  const kinds=new Set();
  let removed=0;
  const note=kind=>{removed++;kinds.add(kind);};

  // Complex fields first: each one that is not the program's own loses its control elements and keeps the words of its result.
  const {fields,stray}=complexFields(root);
  for(const field of fields){
    if(fieldAllowed(field.instruction))continue;
    for(const node of field.nodes)if(node.parentNode)removeControl(node);
    note('field');
  }
  for(const node of stray){if(node.parentNode)removeControl(node);note('field');}

  const visit=el=>{
    for(const c of kids(el)){
      const name=c.localName;
      if(DROP.has(name)){el.removeChild(c);note(name);continue;}
      if(isW(c,'fldSimple')){
        if(fieldAllowed(c.getAttributeNS(W,'instr')||c.getAttribute('w:instr'))){visit(c);continue;}
        visit(c);unwrap(c);note('field');continue;
      }
      if(isW(c,'hyperlink')){
        const external=Array.from(c.attributes||[]).some(a=>a.namespaceURI===REL||/^r:/.test(a.name));
        const internal=!!(c.getAttributeNS(W,'anchor')||c.getAttribute('w:anchor'));
        if(external||!internal){visit(c);unwrap(c);note('link');continue;}
      }
      if(isW(c,'drawing')&&!isInlinePicture(c)){el.removeChild(c);note('drawing');continue;}
      if(c.namespaceURI===A&&(name==='hlinkClick'||name==='hlinkHover')){el.removeChild(c);note('link');continue;}
      for(const a of Array.from(c.attributes||[])){
        if(ADDRESS_ATTRIBUTES.has(a.localName)||(a.namespaceURI===REL&&(a.localName==='link'||a.localName==='href'))){c.removeAttributeNode(a);note('address');}
      }
      visit(c);
    }
  };
  visit(root);
  if(!removed)return {xml:text,removed:0,kinds:[]};
  return {xml:io.serialize(root),removed,kinds:[...kinds]};
}
