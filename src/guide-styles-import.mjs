// Taking the styles of a Word file into the guide. Pure functions on what listStyles reads: the styles of the file, the styles of the guide, and
// the author's choices, which give the edits to the guide's styles (see guide-styles-edit.mjs) and a few words about what was done.
//
// A style of the file that has the name of a style of the guide changes that style; any other can be added as a style of the author's own.
// What the guide cannot hold (a font size of 200 points, a font name with odd characters) is brought inside what it can, or left out, and said.
import {ALIGNMENTS} from './guide-format.mjs';
import {STYLE_LIMITS,cleanStyleEdits,diffSpec,idForName,mergeSpec,niceName} from './guide-styles-edit.mjs';

export const IMPORT_LIMITS=Object.freeze({styles:150});
const COLOR=/^#[0-9A-Fa-f]{6}$/;
const FONT=/^[\p{L}\p{N} .,'()+-]{1,64}$/u;
const SIDES=['top','left','bottom','right'];
const NEW='@new';
export const AS_NEW=NEW;

const norm=name=>String(name||'').toLowerCase().replace(/\s+/g,' ').trim();
const clamp=(v,[lo,hi])=>Math.min(hi,Math.max(lo,v));
const round=v=>Math.round(v*100)/100;

/**
 * What a style of a Word file comes to, as a spec the guide can hold. Numbers outside the limits are brought to the nearest limit, and what
 * cannot be held is left out; `notes` say each of those once.
 * @param {any} effective
 * @returns {{spec: any, notes: string[]}}
 */
export function fitSpec(effective){
  const notes=new Set();
  const spec={font:{},paragraph:{}};
  const f=effective?.font||{},p=effective?.paragraph||{};
  const number=(v,limits,what)=>{
    if(typeof v!=='number'||!Number.isFinite(v)){notes.add(`${what} was left out.`);return undefined;}
    const c=round(clamp(v,limits));
    if(c!==round(v))notes.add(`${what} was brought to ${c}, the nearest the guide can hold.`);
    return c;
  };
  if(f.name!==undefined){if(typeof f.name==='string'&&FONT.test(f.name))spec.font.name=f.name;else notes.add('A font name the guide cannot hold was left out.');}
  if(f.size!==undefined){const v=number(f.size,STYLE_LIMITS.size,'The font size');if(v!==undefined)spec.font.size=v;}
  if(f.color!==undefined){if(typeof f.color==='string'&&COLOR.test(f.color))spec.font.color=f.color.toUpperCase();else notes.add('A color the guide cannot hold was left out.');}
  for(const k of ['bold','italic','underline','strike'])if(typeof f[k]==='boolean')spec.font[k]=f[k];
  if(p.alignment!==undefined&&ALIGNMENTS[p.alignment])spec.paragraph.alignment=p.alignment;
  for(const [k,limits,what] of [['spaceBefore',STYLE_LIMITS.space,'The space before'],['spaceAfter',STYLE_LIMITS.space,'The space after'],['lineSpacing',STYLE_LIMITS.line,'The line spacing'],['leftIndent',STYLE_LIMITS.indent,'The indent'],['firstLineIndent',STYLE_LIMITS.first,'The first line indent']]){
    if(p[k]!==undefined){const v=number(p[k],limits,what);if(v!==undefined)spec.paragraph[k]=v;}
  }
  for(const k of ['keepWithNext','keepTogether'])if(typeof p[k]==='boolean')spec.paragraph[k]=p[k];
  if(typeof effective?.shading==='string'&&COLOR.test(effective.shading))spec.shading=effective.shading.toUpperCase();
  if(Array.isArray(effective?.borders)){
    const borders=effective.borders.filter(b=>b&&SIDES.includes(b.location)&&COLOR.test(b.color||'')&&Number.isFinite(b.width)).slice(0,4).map(b=>({location:b.location,color:b.color.toUpperCase(),width:round(clamp(b.width,[0.25,6]))}));
    if(borders.length)spec.borders=borders;
  }
  if(!Object.keys(spec.font).length)delete spec.font;
  if(!Object.keys(spec.paragraph).length)delete spec.paragraph;
  return {spec,notes:[...notes]};
}

/** A spec with what Word means by saying nothing filled in (not bold, not italic, left aligned), so two specs can be compared. */
export function plainSpec(spec={}){
  const font={bold:false,italic:false,underline:false,strike:false,...(spec.font||{})};
  const paragraph={alignment:'left',keepWithNext:false,...(spec.paragraph||{})};
  return {font,paragraph,...(spec.shading?{shading:spec.shading}:{}),...(spec.borders?{borders:spec.borders}:{})};
}

/** What must change in `was` to be `want`: the settings that differ, and the borders when they are not the same. */
export function changeBetween(was,want){
  const a=plainSpec(was),b=plainSpec(want);
  const change=diffSpec(a,b);
  if(b.borders&&JSON.stringify(b.borders)!==JSON.stringify(a.borders))change.borders=b.borders;
  return change;
}

/**
 * The paragraph styles of a Word file the author can choose from, with the guide's style each would change (matched by name, without regard
 * to capital letters or the spaces in it), or null when the guide has none of that name.
 * @param {any[]} file  listStyles of the file's styles part   @param {any[]} guide  listStyles of the guide's styles
 * @returns {{fileId:string,name:string,spec:any,notes:string[],to:string|null}[]}
 */
export function candidates(file,guide){
  const byName=new Map();
  for(const g of guide)if(g.type==='paragraph'&&!byName.has(norm(g.name)))byName.set(norm(g.name),g);
  const out=file.filter(s=>s.type==='paragraph').slice(0,IMPORT_LIMITS.styles).map(s=>{
    const {spec,notes}=fitSpec(s.effective);
    return {fileId:s.id,name:String(s.name||s.id),spec,notes,to:byName.get(norm(s.name))?.id??null};
  });
  // The styles that match come first: they are the ones most people came for.
  return [...out.filter(c=>c.to),...out.filter(c=>!c.to)];
}

const uniqueName=(name,taken)=>{
  const used=new Set([...taken].map(norm));
  const base=name.replace(/[<>&"\u0000-\u001f]/g,'').replace(/\s+/g,' ').trim().slice(0,STYLE_LIMITS.name-12)||'Imported style';
  if(!used.has(norm(base)))return base;
  let n=2;
  while(used.has(norm(`${base} (${n})`)))n++;
  return `${base} (${n})`;
};

/**
 * The edits to the guide's styles that bring in the styles the author chose.
 * @param {{file:any[], guide:any[], edits:{overrides:Record<string,any>,custom:any[]}, picks:{fileId:string,to:string}[]}} input
 *   `file` is listStyles of the file, `guide` is listStyles of the guide, `edits` the guide's present edits, and each pick says which style of the file goes to which style of the guide (or AS_NEW)
 * @returns {{edits:any, changed:string[], added:string[], unchanged:string[], notes:string[]}}
 */
export function planImport({file,guide,edits,picks}){
  const next=JSON.parse(JSON.stringify(edits));
  const changed=[],added=[],unchanged=[],notes=[];
  const taken=[...guide.map(s=>s.name),...next.custom.map(c=>c.name)];
  const ids=[...guide.map(s=>s.id),...next.custom.map(c=>c.id)];
  const normal=guide.find(s=>norm(s.name)==='normal')||guide.find(s=>s.type==='paragraph');
  for(const pick of picks){
    const f=file.find(s=>s.id===pick.fileId);
    if(!f)continue;
    const {spec,notes:n}=fitSpec(f.effective);
    for(const m of n)notes.push(`${niceName(f.name)}: ${m}`);
    if(pick.to===NEW){
      if(!normal)throw new Error('The guide has no ordinary text style to start a style from.');
      const name=uniqueName(niceName(f.name||f.id),taken);
      const id=idForName(name,ids);
      taken.push(name);ids.push(id);
      next.custom.push({id,name,type:'paragraph',basedOn:normal.id,spec:changeBetween(normal.effective,spec)});
      added.push(name);
      continue;
    }
    const target=guide.find(s=>s.id===pick.to&&s.type==='paragraph');
    if(!target)throw new Error(`The guide has no style to put ${niceName(f.name)} in.`);
    const change=changeBetween(target.effective,spec);
    if(!Object.keys(change).length){unchanged.push(niceName(target.name));continue;}
    if(target.custom){
      const mine=next.custom.find(c=>c.id===target.id);
      if(!mine)throw new Error(`${niceName(target.name)} is not one of your own styles.`);
      mine.spec=mergeSpec(mine.spec,change);
    }else next.overrides[target.id]=mergeSpec(next.overrides[target.id]??{},change);
    changed.push(niceName(target.name));
  }
  return {edits:cleanStyleEdits(next),changed,added,unchanged,notes};
}

/** What the author is told when the styles are in. */
export function importSummary({changed,added,unchanged}){
  const list=names=>names.length<=3?names.join(', '):`${names.slice(0,3).join(', ')} and ${names.length-3} more`;
  const parts=[];
  if(changed.length)parts.push(`${changed.length===1?'1 style':`${changed.length} styles`} changed (${list(changed)})`);
  if(added.length)parts.push(`${added.length===1?'1 style':`${added.length} styles`} added as your own (${list(added)})`);
  if(!parts.length)return unchanged.length?'Those styles already look the way the file has them, so nothing changed.':'No style was chosen.';
  return `${parts.join(' and ')}. Ctrl+Z takes it back.`;
}
