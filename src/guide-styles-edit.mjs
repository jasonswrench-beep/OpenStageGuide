// Editing the styles of the guide: changing how a built-in style looks, and making styles of the author's own. Pure functions on the styles
// part of a Word document (word/styles.xml).
//
// The edits are kept apart from the styles the program writes: {overrides: {styleId: spec}, custom: [{id, name, type, basedOn, spec}]}. The
// styles part is always made again from the program's own definitions with the edits laid over them, so taking an edit away puts the style
// back exactly as it was, and the same edits on the same definitions give the same styles part.
//
// A spec has only what the author set: {font:{name,size,color,bold,italic,underline,strike}, paragraph:{alignment,spaceBefore,spaceAfter,
// lineSpacing,leftIndent,firstLineIndent,keepWithNext,keepTogether}, shading}. Sizes and spaces are in points, line spacing is a multiple of
// one line, colors are #RRGGBB. Everything is written as real Word style properties, in the order Word's schema gives, so the styles are
// the styles a Word user sees and can use.
import {ALIGNMENTS,FORMAT_LIMITS} from './guide-format.mjs';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const STYLE_LIMITS=Object.freeze({custom:60,name:60,font:64,size:[6,96],space:FORMAT_LIMITS.space,line:FORMAT_LIMITS.line,indent:[0,360],first:[-360,360]});
const PPR_ORDER=['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];
export const RPR_ORDER=['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath'];
const STYLE_ORDER=['name','aliases','basedOn','next','link','autoRedefine','hidden','uiPriority','semiHidden','unhideWhenUsed','qFormat','locked','personal','personalCompose','personalReply','rsid','pPr','rPr','tblPr','trPr','tcPr','tblStylePr'];
const FROM_WORD={left:'left',start:'left',center:'center',right:'right',end:'right',both:'justify'};
const COLOR=/^#[0-9A-Fa-f]{6}$/;
const SIDES=['top','left','bottom','right'];
const FONT=/^[\p{L}\p{N} .,'()+-]{1,64}$/u;

const kids=el=>Array.from(el.childNodes).filter(n=>n.nodeType===1);
const named=(el,name)=>kids(el).find(n=>n.localName===name&&n.namespaceURI===W);
const attr=(el,name)=>el?.getAttributeNS(W,name)||'';
const isOn=el=>!!el&&!['false','0','off'].includes((attr(el,'val')||'1').toLowerCase());
const twips=pt=>Math.round(pt*20);
const points=t=>Math.round(Number(t)/20*100)/100;

// ---- checking

const bad=m=>{throw new Error(m);};
const range=(v,[lo,hi],what)=>{if(typeof v!=='number'||!Number.isFinite(v)||v<lo||v>hi)bad(`${what} must be a number from ${lo} to ${hi}.`);return Math.round(v*100)/100;};

/** A clean copy of a spec, or a message saying what is wrong with it. Unknown keys are dropped. */
export function cleanSpec(spec){
  if(spec===null||typeof spec!=='object'||Array.isArray(spec))bad('A style\'s settings are not in the right form.');
  const out={};
  const f=spec.font;
  if(f!==undefined&&f!==null){
    if(typeof f!=='object'||Array.isArray(f))bad('A style\'s font settings are not in the right form.');
    const o={};
    if(f.name!==undefined){if(typeof f.name!=='string'||!FONT.test(f.name))bad('A font name can have letters, digits, spaces and . , \' ( ) + - only, up to 64 characters.');o.name=f.name;}
    if(f.size!==undefined)o.size=range(f.size,STYLE_LIMITS.size,'The font size');
    if(f.color!==undefined){if(typeof f.color!=='string'||!COLOR.test(f.color))bad('A color is six hexadecimal digits after a #, for example #1F3A5F.');o.color=f.color.toUpperCase();}
    for(const k of ['bold','italic','underline','strike']){if(f[k]!==undefined){if(typeof f[k]!=='boolean')bad(`${k} is on or off.`);o[k]=f[k];}}
    if(Object.keys(o).length)out.font=o;
  }
  const p=spec.paragraph;
  if(p!==undefined&&p!==null){
    if(typeof p!=='object'||Array.isArray(p))bad('A style\'s paragraph settings are not in the right form.');
    const o={};
    if(p.alignment!==undefined){if(!ALIGNMENTS[p.alignment])bad('Alignment must be left, center, right or justify.');o.alignment=p.alignment;}
    if(p.spaceBefore!==undefined)o.spaceBefore=range(p.spaceBefore,STYLE_LIMITS.space,'Space before');
    if(p.spaceAfter!==undefined)o.spaceAfter=range(p.spaceAfter,STYLE_LIMITS.space,'Space after');
    if(p.lineSpacing!==undefined)o.lineSpacing=range(p.lineSpacing,STYLE_LIMITS.line,'Line spacing');
    if(p.leftIndent!==undefined)o.leftIndent=range(p.leftIndent,STYLE_LIMITS.indent,'The indent');
    if(p.firstLineIndent!==undefined)o.firstLineIndent=range(p.firstLineIndent,STYLE_LIMITS.first,'The first line indent');
    for(const k of ['keepWithNext','keepTogether']){if(p[k]!==undefined){if(typeof p[k]!=='boolean')bad(`${k} is on or off.`);o[k]=p[k];}}
    if(Object.keys(o).length)out.paragraph=o;
  }
  if(spec.shading!==undefined&&spec.shading!==null){
    if(typeof spec.shading!=='string'||!COLOR.test(spec.shading))bad('A shading color is six hexadecimal digits after a #.');
    out.shading=spec.shading.toUpperCase();
  }
  if(spec.borders!==undefined&&spec.borders!==null){
    if(!Array.isArray(spec.borders)||spec.borders.length>4)bad('A style has up to four borders: left, right, top and bottom.');
    const seen=new Set();
    out.borders=spec.borders.map(b=>{
      if(!b||typeof b!=='object'||!SIDES.includes(b.location)||seen.has(b.location))bad('A border is on the left, right, top or bottom, once.');
      seen.add(b.location);
      if(typeof b.color!=='string'||!COLOR.test(b.color))bad('A border color is six hexadecimal digits after a #.');
      return {location:b.location,color:b.color.toUpperCase(),width:range(b.width,[0.25,6],'A border\'s width')};
    });
    if(!out.borders.length)delete out.borders;
  }
  return out;
}

const ID=/^[A-Za-z][A-Za-z0-9]{0,63}$/;
/** Edits as a clean copy (or a message). Ids and names are checked here once, however the edits got here. */
export function cleanStyleEdits(raw){
  if(raw===undefined||raw===null)return {overrides:{},custom:[]};
  if(typeof raw!=='object'||Array.isArray(raw))bad('The styles are not in the right form.');
  const overrides={};
  const o=raw.overrides??{};
  if(typeof o!=='object'||Array.isArray(o))bad('The changes to styles are not in the right form.');
  for(const [id,spec] of Object.entries(o)){
    if(!ID.test(id)||['__proto__','constructor','prototype'].includes(id))bad('A style has an id this program cannot use.');
    const clean=cleanSpec(spec);
    if(Object.keys(clean).length)overrides[id]=clean;
  }
  const custom=[];
  if(raw.custom!==undefined&&!Array.isArray(raw.custom))bad('The styles of your own are not in the right form.');
  if((raw.custom||[]).length>STYLE_LIMITS.custom)bad(`There can be up to ${STYLE_LIMITS.custom} styles of your own.`);
  const seen=new Set();
  for(const c of raw.custom||[]){
    if(!c||typeof c!=='object'||!ID.test(String(c.id))||typeof c.name!=='string')bad('A style of your own is not in the right form.');
    const name=c.name.replace(/\s+/g,' ').trim();
    if(!name||name.length>STYLE_LIMITS.name||/[<>&"\u0000-\u001f]/.test(name))bad(`A style's name is up to ${STYLE_LIMITS.name} letters, without < > & or quotation marks.`);
    if(!['paragraph','character'].includes(c.type))bad('A style is for paragraphs or for characters.');
    if(!ID.test(String(c.basedOn)))bad('A style of your own is based on another style.');
    if(seen.has(c.id.toLowerCase())||seen.has('n:'+name.toLowerCase()))bad(`There is already a style called ${name}.`);
    seen.add(c.id.toLowerCase());seen.add('n:'+name.toLowerCase());
    custom.push({id:c.id,name,type:c.type,basedOn:c.basedOn,spec:cleanSpec(c.spec??{})});
  }
  return {overrides,custom};
}

// ---- writing

function placeIn(parent,el,order){
  const rank=n=>{const i=order.indexOf(n.localName);return i<0?order.length:i;};
  const mine=rank(el);
  parent.insertBefore(el,kids(parent).find(n=>rank(n)>mine)||null);
}
function ensure(parent,name,order,doc,{create=true}={}){
  let e=named(parent,name);
  if(!e&&create){e=doc.createElementNS(W,`w:${name}`);placeIn(parent,e,order);}
  return e;
}
const drop=(parent,name)=>{const e=named(parent,name);if(e)parent.removeChild(e);};
const toggle=(rPr,name,on,doc,both=false)=>{
  const names=both?[name,`${name}Cs`]:[name];
  for(const n of names){const e=ensure(rPr,n,RPR_ORDER,doc);if(on)e.removeAttributeNS(W,'val');else e.setAttributeNS(W,'w:val','0');}
};

function writeSpec(style,spec,doc){
  const f=spec.font,p=spec.paragraph;
  if(f&&Object.keys(f).length){
    const rPr=ensure(style,'rPr',STYLE_ORDER,doc);
    if(f.name!==undefined){const e=ensure(rPr,'rFonts',RPR_ORDER,doc);for(const k of ['ascii','hAnsi','cs','eastAsia'])e.setAttributeNS(W,`w:${k}`,f.name);for(const k of ['asciiTheme','hAnsiTheme','cstheme','eastAsiaTheme'])e.removeAttributeNS(W,k);}
    if(f.bold!==undefined)toggle(rPr,'b',f.bold,doc,true);
    if(f.italic!==undefined)toggle(rPr,'i',f.italic,doc,true);
    if(f.strike!==undefined)toggle(rPr,'strike',f.strike,doc);
    if(f.color!==undefined)ensure(rPr,'color',RPR_ORDER,doc).setAttributeNS(W,'w:val',f.color.slice(1));
    if(f.size!==undefined){for(const n of ['sz','szCs'])ensure(rPr,n,RPR_ORDER,doc).setAttributeNS(W,'w:val',String(Math.round(f.size*2)));}
    if(f.underline!==undefined){const e=ensure(rPr,'u',RPR_ORDER,doc);e.setAttributeNS(W,'w:val',f.underline?'single':'none');}
  }
  const shade=spec.shading;
  if((p&&Object.keys(p).length)||shade||spec.borders){
    const pPr=ensure(style,'pPr',STYLE_ORDER,doc);
    if(p){
      if(p.alignment!==undefined)ensure(pPr,'jc',PPR_ORDER,doc).setAttributeNS(W,'w:val',ALIGNMENTS[p.alignment]);
      if(p.spaceBefore!==undefined||p.spaceAfter!==undefined||p.lineSpacing!==undefined){
        const sp=ensure(pPr,'spacing',PPR_ORDER,doc);
        if(p.spaceBefore!==undefined)sp.setAttributeNS(W,'w:before',String(twips(p.spaceBefore)));
        if(p.spaceAfter!==undefined)sp.setAttributeNS(W,'w:after',String(twips(p.spaceAfter)));
        if(p.lineSpacing!==undefined){sp.setAttributeNS(W,'w:line',String(Math.round(p.lineSpacing*240)));sp.setAttributeNS(W,'w:lineRule','auto');}
      }
      if(p.leftIndent!==undefined||p.firstLineIndent!==undefined){
        const ind=ensure(pPr,'ind',PPR_ORDER,doc);
        if(p.leftIndent!==undefined){ind.setAttributeNS(W,'w:left',String(twips(p.leftIndent)));ind.removeAttributeNS(W,'start');}
        if(p.firstLineIndent!==undefined){
          ind.removeAttributeNS(W,'hanging');ind.removeAttributeNS(W,'firstLine');
          if(p.firstLineIndent<0)ind.setAttributeNS(W,'w:hanging',String(twips(-p.firstLineIndent)));else if(p.firstLineIndent>0)ind.setAttributeNS(W,'w:firstLine',String(twips(p.firstLineIndent)));
        }
      }
      for(const [k,n] of [['keepWithNext','keepNext'],['keepTogether','keepLines']])if(p[k]!==undefined){
        if(p[k])ensure(pPr,n,PPR_ORDER,doc);else drop(pPr,n);
      }
    }
    if(shade){const e=ensure(pPr,'shd',PPR_ORDER,doc);e.setAttributeNS(W,'w:val','clear');e.setAttributeNS(W,'w:color','auto');e.setAttributeNS(W,'w:fill',shade.slice(1));}
    if(spec.borders){
      const bdr=ensure(pPr,'pBdr',PPR_ORDER,doc);
      for(const side of SIDES){
        const b=spec.borders.find(x=>x.location===side);
        if(!b)continue;
        const e=ensure(bdr,side,['top','left','bottom','right','between','bar'],doc);
        e.setAttributeNS(W,'w:val','single');e.setAttributeNS(W,'w:sz',String(Math.max(2,Math.round(b.width*8))));e.setAttributeNS(W,'w:space','4');e.setAttributeNS(W,'w:color',b.color.slice(1));
      }
    }
  }
}

/** The id a name of the author's would have, which no style in `taken` (compared without regard to case) has. */
export function idForName(name,taken){
  const used=new Set([...taken].map(t=>String(t).toLowerCase()));
  let base=name.normalize('NFKD').replace(/[^A-Za-z0-9]/g,'');
  if(!/^[A-Za-z]/.test(base))base=`Style${base}`;
  base=base.slice(0,40);
  let id=base,n=2;
  while(used.has(id.toLowerCase()))id=`${base}${n++}`;
  return id;
}

/**
 * The styles part with the edits laid over it. With no edits the same text comes back.
 * @param {string} stylesXml @param {{overrides:Record<string,object>,custom:object[]}} edits @param {{parse:(x:string)=>Document, serialize:(e:Element)=>string}} io
 */
export function applyStyleEdits(stylesXml,edits,io){
  const clean=cleanStyleEdits(edits);
  if(!Object.keys(clean.overrides).length&&!clean.custom.length)return stylesXml;
  const root=io.parse(stylesXml).documentElement;
  if(!root||root.namespaceURI!==W||root.localName!=='styles')bad('That is not a styles part.');
  const doc=root.ownerDocument;
  const byId=new Map(kids(root).filter(s=>s.localName==='style').map(s=>[attr(s,'styleId'),s]));
  for(const [id,spec] of Object.entries(clean.overrides)){
    const style=byId.get(id);
    if(style)writeSpec(style,spec,doc);
  }
  for(const c of clean.custom){
    if(byId.has(c.id))bad(`There is already a style with the id ${c.id}.`);
    if(!byId.has(c.basedOn)&&!clean.custom.some(x=>x.id===c.basedOn))bad(`The style ${c.name} is based on a style that is not there.`);
    const style=doc.createElementNS(W,'w:style');
    style.setAttributeNS(W,'w:type',c.type);
    style.setAttributeNS(W,'w:customStyle','1');
    style.setAttributeNS(W,'w:styleId',c.id);
    for(const [tag,val] of [['name',c.name],['basedOn',c.basedOn],...(c.type==='paragraph'?[['next',c.id]]:[])]){
      const e=doc.createElementNS(W,`w:${tag}`);e.setAttributeNS(W,'w:val',val);placeIn(style,e,STYLE_ORDER);
    }
    placeIn(style,doc.createElementNS(W,'w:qFormat'),STYLE_ORDER);
    writeSpec(style,c.spec,doc);
    root.appendChild(style);
    byId.set(c.id,style);
  }
  return io.serialize(root);
}

// ---- reading

const THEME_FACE=/^(major|minor)/i;
/** The spec a style element holds itself, in the form a spec has (only what it says). `themeFonts` names the theme's two fonts, for a style that asks for "the body font". */
export function readSpec(style,themeFonts={}){
  const spec={};
  const rPr=named(style,'rPr'),pPr=named(style,'pPr');
  if(rPr){
    const font={};
    const fonts=named(rPr,'rFonts');
    if(fonts&&attr(fonts,'ascii'))font.name=attr(fonts,'ascii');
    else if(fonts){
      const theme=attr(fonts,'asciiTheme')||attr(fonts,'hAnsiTheme');
      const face=THEME_FACE.test(theme)?themeFonts[theme.slice(0,5).toLowerCase()]:undefined;
      if(face)font.name=face;
    }
    const sz=named(rPr,'sz');
    if(sz&&attr(sz,'val'))font.size=Number(attr(sz,'val'))/2;
    const color=named(rPr,'color');
    if(color&&/^[0-9A-Fa-f]{6}$/.test(attr(color,'val')))font.color=`#${attr(color,'val').toUpperCase()}`;
    if(named(rPr,'b'))font.bold=isOn(named(rPr,'b'));
    if(named(rPr,'i'))font.italic=isOn(named(rPr,'i'));
    if(named(rPr,'strike'))font.strike=isOn(named(rPr,'strike'));
    const u=named(rPr,'u');
    if(u)font.underline=attr(u,'val')!=='none';
    if(Object.keys(font).length)spec.font=font;
  }
  if(pPr){
    const p={};
    const jc=named(pPr,'jc');
    if(jc&&FROM_WORD[attr(jc,'val')])p.alignment=FROM_WORD[attr(jc,'val')];
    const sp=named(pPr,'spacing');
    if(sp){
      if(attr(sp,'before'))p.spaceBefore=points(attr(sp,'before'));
      if(attr(sp,'after'))p.spaceAfter=points(attr(sp,'after'));
      if(attr(sp,'line')&&(attr(sp,'lineRule')||'auto')==='auto')p.lineSpacing=Math.round(Number(attr(sp,'line'))/240*100)/100;
    }
    const ind=named(pPr,'ind');
    if(ind){
      if(attr(ind,'left'))p.leftIndent=points(attr(ind,'left'));
      if(attr(ind,'hanging'))p.firstLineIndent=-points(attr(ind,'hanging'));
      else if(attr(ind,'firstLine'))p.firstLineIndent=points(attr(ind,'firstLine'));
    }
    if(named(pPr,'keepNext'))p.keepWithNext=isOn(named(pPr,'keepNext'));
    if(named(pPr,'keepLines'))p.keepTogether=isOn(named(pPr,'keepLines'));
    if(Object.keys(p).length)spec.paragraph=p;
    const shd=named(pPr,'shd');
    if(shd&&/^[0-9A-Fa-f]{6}$/.test(attr(shd,'fill')))spec.shading=`#${attr(shd,'fill').toUpperCase()}`;
    const bdr=named(pPr,'pBdr');
    if(bdr){
      const borders=[];
      for(const side of SIDES){
        const b=named(bdr,side);
        if(b&&!['nil','none'].includes(attr(b,'val'))&&/^[0-9A-Fa-f]{6}$/.test(attr(b,'color')))borders.push({location:side,color:`#${attr(b,'color').toUpperCase()}`,width:Math.max(0.25,Math.round(Number(attr(b,'sz')||4)/8*100)/100)});
      }
      if(borders.length)spec.borders=borders;
    }
  }
  return spec;
}

/**
 * Every style the styles part has that an author can use: id, name, kind, the style it is based on, whether it is the author's own (its id is
 * in `authorIds`; the program's own OSG styles are Word custom styles too, but they are not the author's), what it
 * says itself (`own`) and what it comes to once the styles it is based on and the document's defaults are counted (`effective`).
 * `options.themeFonts` ({major, minor}) names the fonts of a Word file's theme; with `options.paragraphDefaults` the document's paragraph
 * defaults (the space after and the line spacing Word gives every paragraph) count too, and borders are carried by `effective`, as a style
 * from a Word file needs.
 */
export function listStyles(stylesXml,io,authorIds=new Set(),options={}){
  const themeFonts=options.themeFonts||{};
  const root=io.parse(stylesXml).documentElement;
  const styles=kids(root).filter(s=>s.localName==='style');
  const byId=new Map(styles.map(s=>[attr(s,'styleId'),s]));
  const defaults=named(root,'docDefaults');
  const defRPr=defaults&&named(named(defaults,'rPrDefault')||{childNodes:[]},'rPr');
  const base={font:{}};
  // The document's defaults are what a style that says nothing of its own font comes to (readSpec reads the properties of whatever holds an rPr).
  if(defRPr)base.font=readSpec({childNodes:[defRPr]},themeFonts).font||{};
  base.paragraph={};
  if(options.paragraphDefaults){
    const defPPr=defaults&&named(named(defaults,'pPrDefault')||{childNodes:[]},'pPr');
    if(defPPr)base.paragraph=readSpec({childNodes:[defPPr]}).paragraph||{};
  }
  const merged=(spec,over)=>({font:{...(spec.font||{}),...(over.font||{})},paragraph:{...(spec.paragraph||{}),...(over.paragraph||{})},...((over.shading??spec.shading)?{shading:over.shading??spec.shading}:{}),...(options.paragraphDefaults&&(over.borders??spec.borders)?{borders:over.borders??spec.borders}:{})});
  const effective=(id,depth=0)=>{
    const s=byId.get(id);
    if(!s||depth>12)return {font:{...base.font},paragraph:{...base.paragraph}};
    const parent=attr(named(s,'basedOn'),'val');
    return merged(parent?effective(parent,depth+1):{font:{...base.font},paragraph:{...base.paragraph}},readSpec(s,themeFonts));
  };
  return styles.filter(s=>['paragraph','character'].includes(attr(s,'type'))&&!named(s,'semiHidden')).map(s=>{
    const id=attr(s,'styleId');
    return {id,name:attr(named(s,'name'),'val'),type:attr(s,'type'),basedOn:attr(named(s,'basedOn'),'val'),custom:authorIds.has(id),own:readSpec(s,themeFonts),effective:effective(id),defaultStyle:attr(s,'default')==='1'};
  });
}

/** A style's name as a person reads it: Word's own names for its built-in styles are lower case ("heading 1", "toc 2"). */
export const niceName=name=>String(name??'').replace(/^(heading|toc) (\d)$/i,(m,k,n)=>`${k.toLowerCase()==='toc'?'TOC':'Heading'} ${n}`);

/** What a spec looks like on a page, as CSS properties for an element (points are 4/3 of a pixel). */
export function specCss(spec){
  const css={};
  const f=spec.font||{},p=spec.paragraph||{};
  const px=pt=>`${Math.round(pt*133.33)/100}px`;
  if(f.name)css.fontFamily=`"${f.name}", system-ui, sans-serif`;
  if(f.size)css.fontSize=px(f.size);
  if(f.color)css.color=f.color;
  if(f.bold!==undefined)css.fontWeight=f.bold?'700':'400';
  if(f.italic!==undefined)css.fontStyle=f.italic?'italic':'normal';
  if(f.underline!==undefined||f.strike!==undefined){
    const d=[f.underline?'underline':'',f.strike?'line-through':''].filter(Boolean).join(' ');
    css.textDecoration=d||'none';
  }
  if(p.alignment)css.textAlign=p.alignment;
  if(p.spaceBefore!==undefined)css.marginTop=px(p.spaceBefore);
  if(p.spaceAfter!==undefined)css.marginBottom=px(p.spaceAfter);
  if(p.lineSpacing!==undefined)css.lineHeight=String(Math.round(p.lineSpacing*120)/100);
  if(p.leftIndent!==undefined)css.marginLeft=px(p.leftIndent);
  if(p.firstLineIndent!==undefined)css.textIndent=px(p.firstLineIndent);
  if(spec.shading)css.backgroundColor=spec.shading;
  for(const b of spec.borders||[])css[`border${b.location[0].toUpperCase()}${b.location.slice(1)}`]=`${Math.max(1,Math.round(b.width*1.33))}px solid ${b.color}`;
  return css;
}

// ---- what the Styles panel needs

/** Both specs laid together, the second over the first, group by group. */
export function mergeSpec(a={},b={}){
  const out={};
  for(const part of ['font','paragraph']){
    const m={...(a[part]||{}),...(b[part]||{})};
    if(Object.keys(m).length)out[part]=m;
  }
  const shading=b.shading??a.shading;
  if(shading)out.shading=shading;
  const borders=b.borders??a.borders;
  if(borders)out.borders=borders;
  return out;
}

/**
 * What the author changed: the settings in `now` that are not what `was` had (a setting that is not in `now` is not a change). The Styles panel
 * shows what a style comes to, and stores only what was changed, so everything else goes on following the style it is based on.
 */
export function diffSpec(was={},now={}){
  const out={};
  for(const part of ['font','paragraph']){
    const d={};
    for(const [k,v] of Object.entries(now[part]||{}))if(v!==undefined&&(was[part]||{})[k]!==v)d[k]=v;
    if(Object.keys(d).length)out[part]=d;
  }
  if(now.shading&&now.shading!==was.shading)out.shading=now.shading;
  return out;
}

const channel=c=>{const v=c/255;return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4);};
const luminance=hex=>{const n=parseInt(hex.slice(1),16);return 0.2126*channel((n>>16)&255)+0.7152*channel((n>>8)&255)+0.0722*channel(n&255);};
/** The contrast of two #RRGGBB colors, from 1 to 21 (WCAG). */
export function contrastRatio(a,b){
  if(!COLOR.test(a)||!COLOR.test(b))return null;
  const x=luminance(a),y=luminance(b);
  return Math.round((Math.max(x,y)+0.05)/(Math.min(x,y)+0.05)*100)/100;
}
/** Text of this size and weight needs 4.5 to 1 against what it is on, or 3 to 1 when it is large (18 points, or 14 and bold). */
export const contrastNeeded=(size,bold)=>((size||11)>=18||((size||11)>=14&&bold))?3:4.5;

/**
 * How the guide's page draws styles that the author changed or made: for a style's name, the CSS to lay over its usual look (cssFor, or null when
 * nothing was changed), and the built-in style a style of the author's is based on, whose usual look it starts from (baseOf, or null).
 * @param {{overrides:Record<string,object>,custom:object[]}} edits @param {ReturnType<typeof listStyles>} list
 */
export function previewStyles(edits,list){
  const byName=new Map(list.map(s=>[s.name,s]));
  const byId=new Map(list.map(s=>[s.id,s]));
  const chain=name=>{
    const out=[];
    let s=byName.get(name),guard=0;
    while(s&&guard++<12){out.unshift(s);s=s.custom?byId.get(s.basedOn):null;}
    return out;
  };
  const view={
    cssFor(name){
      let spec={};
      for(const s of chain(name))spec=mergeSpec(spec,s.custom?s.own:(edits.overrides[s.id]||{}));
      const css=specCss(spec);
      return Object.keys(css).length?css:null;
    },
    /** The same for a style named by its id, as the words of a run name their character style. */
    cssForId(id){
      const s=byId.get(id);
      return s?view.cssFor(s.name):null;
    },
    baseOf(name){
      const c=chain(name);
      return c.length&&c[0].custom===false?c[0].name:null;
    }
  };
  return view;
}
