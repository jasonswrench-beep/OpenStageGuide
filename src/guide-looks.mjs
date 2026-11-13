// Looks for the facilitator guide. A look is the same handful of values the learner workbook uses (the five families, an accent
// colour, an organization's name), turned into the guide's styles. The guide is made of real Word styles, so changing the look
// means changing those styles, and every block follows.
//
// Two things carry a look into the document. Styles the guide has not used yet are made by the next insert, from the active
// look's definitions (the adapter is told the theme). Styles that already exist are changed through Word's style settings
// (applyLook in the Word adapter), because Word keeps the definition it already has when an insert brings its own.
import {resolveWorkbookBrand} from './workbook-presentation.mjs';
import {textContrast} from './workbook-presentation.mjs';
import {lookTheme,WORKBOOK_LOOKS} from './guide-workbook.mjs';
import {styleDefinitions,NS} from './guide-styles.mjs';

// The workbook's five, and one more for designing E-Learning screen by screen: the Practice Lab's layout in a deep teal.
export const STORYBOARD_LOOK=Object.freeze({id:'storyboard',name:'Storyboard',note:'Screen by screen, for E-Learning.',family:'practice-lab',accent:'#0B5D5D'});
export const GUIDE_LOOKS=Object.freeze([...WORKBOOK_LOOKS,STORYBOARD_LOOK]);
export const LOOK_PROPERTY='osg.look';
export const DEFAULT_LOOK=Object.freeze({id:'clearline',accent:'',organization:'',header:false,copyright:'',revised:''});
const HEX=/^#[0-9a-fA-F]{6}$/;
const oneLine=(s,max)=>String(s??'').replace(/[\u0000-\u001f\u007f]+/g,' ').replace(/\s+/g,' ').trim().slice(0,max);

// Whatever was stored (or typed) becomes a complete, safe set of choices. Anything unknown falls back to the default.
export function cleanLook(raw){
  const r=raw&&typeof raw==='object'&&!Array.isArray(raw)?raw:{};
  let accent=String(r.accent??'').trim();
  if(accent&&!accent.startsWith('#'))accent=`#${accent}`;
  return {
    id:GUIDE_LOOKS.some(l=>l.id===r.id)?r.id:DEFAULT_LOOK.id,
    accent:HEX.test(accent)?accent.toUpperCase():'',
    organization:oneLine(r.organization,120),
    header:!!r.header,
    copyright:oneLine(r.copyright,120),
    revised:oneLine(r.revised,40)
  };
}
export const parseLook=text=>{try{return cleanLook(JSON.parse(text));}catch{return cleanLook(null);}};
// The grey a "print in black and white" setting turns a colour into (the usual weights of red, green and blue), as a colour.
const greyOf=color=>{
  const n=parseInt(color.slice(1),16),v=Math.round(0.299*(n>>16&255)+0.587*(n>>8&255)+0.114*(n&255)).toString(16).padStart(2,'0');
  return `#${v}${v}${v}`;
};

// The look's theme (see styleDefinitions), or the reasons it cannot be used. The accent is the colour of every heading, so it
// has to read as text on white.
export function resolveGuideLook(settings){
  const s=cleanLook(settings),brand={},extra=GUIDE_LOOKS.find(l=>l.id===s.id);
  // A look of the guide's own is one of the workbook's families with its own accent unless the author chose one.
  if(s.accent)brand.accent=s.accent;else if(extra?.accent)brand.accent=extra.accent;
  if(s.organization)brand.organizationName=s.organization;
  const r=resolveWorkbookBrand(extra?.family||s.id,brand);
  if(!r.eligible)return {ok:false,settings:s,issues:r.issues};
  const theme=lookTheme(r.effective,'en-US'),accent=`#${theme.navy}`,shade=`#${theme.shade}`,ratio=textContrast(accent,'#FFFFFF');
  if(ratio<4.5)return {ok:false,settings:s,issues:[`The accent colour is too light to read as heading text: its contrast with the page is ${ratio.toFixed(1)} to 1, and it needs at least 4.5 to 1. Choose a darker colour.`]};
  // The accent is also the colour of the labels that sit on the beige boxes (facilitator notes, answer keys), and the guide is printed in
  // black and white as often as not, so it has to read there too, in colour and as the greys a black-and-white printer turns it into.
  const onShade=textContrast(accent,shade);
  if(onShade<4.5)return {ok:false,settings:s,issues:[`The accent colour is too light to read on the beige boxes (facilitator notes and answer keys): its contrast with the box is ${onShade.toFixed(1)} to 1, and it needs at least 4.5 to 1. Choose a darker colour.`]};
  const printed=Math.min(textContrast(greyOf(accent),'#FFFFFF'),textContrast(greyOf(accent),greyOf(shade)));
  if(printed<4.5)return {ok:false,settings:s,issues:[`The accent colour is too faint when the guide is printed in black and white: its contrast would be ${printed.toFixed(1)} to 1, and it needs at least 4.5 to 1. Choose a darker colour.`]};
  return {ok:true,settings:s,theme,effective:r.effective,issues:[]};
}

// ---- styles, as the settings Word's style object takes
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const kids=(n,name)=>Array.from(n.childNodes).filter(c=>c.nodeType===1&&c.localName===name);
// An attribute's value, or null when it is not there (a parser may answer "" for a missing one).
const attr=(n,name)=>{if(!n)return null;const v=n.getAttributeNS?.(W,name)||n.getAttribute?.(`w:${name}`);return v?v:null;};
const on=n=>!!n&&!['false','0','off'].includes(String(attr(n,'val')??'').toLowerCase());
const pt=twips=>Math.round(Number(twips)/20*100)/100;
// Word's border widths come in fixed steps (points).
const BORDER_STEPS=[0.25,0.5,0.75,1,1.5,2.25,3,4.5,6];
const STEP_NAMES={0.25:'Pt025',0.5:'Pt050',0.75:'Pt075',1:'Pt100',1.5:'Pt150',2.25:'Pt225',3:'Pt300',4.5:'Pt450',6:'Pt600'};
const nearestStep=points=>BORDER_STEPS.reduce((best,s)=>Math.abs(s-points)<Math.abs(best-points)?s:best,BORDER_STEPS[0]);
const SIDE={left:'Left',right:'Right',top:'Top',bottom:'Bottom'};
const WORD_NAME={'heading 1':'Heading 1','heading 2':'Heading 2','toc 1':'TOC 1','toc 2':'TOC 2'};

// One spec for each style in the guide's definitions: {name, font, paragraph, shading, borders}. Only what a style says itself
// is in its spec; the rest it inherits from the style it is based on, as it does in the document.
export function styleSpecs(theme,DOMParserClass){
  const doc=new DOMParserClass().parseFromString(`<w:styles ${NS}>${styleDefinitions(theme)}</w:styles>`,'application/xml');
  const entries=Array.from(doc.getElementsByTagNameNS(W,'style')).filter(s=>attr(s,'type')==='paragraph').map(style=>{
    const raw=attr(kids(style,'name')[0],'val'),name=WORD_NAME[raw]||raw;
    const pPr=kids(style,'pPr')[0],rPr=kids(style,'rPr')[0],spec={name,font:{},paragraph:{},shading:null,borders:[]};
    if(rPr){
      const fonts=kids(rPr,'rFonts')[0],size=kids(rPr,'sz')[0],color=kids(rPr,'color')[0];
      if(fonts&&attr(fonts,'ascii'))spec.font.name=attr(fonts,'ascii');
      if(size)spec.font.size=Number(attr(size,'val'))/2;
      if(color&&/^[0-9a-fA-F]{6}$/.test(attr(color,'val')||''))spec.font.color=`#${attr(color,'val').toUpperCase()}`;
      if(kids(rPr,'b')[0])spec.font.bold=on(kids(rPr,'b')[0]);
      if(kids(rPr,'i')[0])spec.font.italic=on(kids(rPr,'i')[0]);
    }
    if(pPr){
      const spacing=kids(pPr,'spacing')[0],ind=kids(pPr,'ind')[0],jc=kids(pPr,'jc')[0],shd=kids(pPr,'shd')[0],bdr=kids(pPr,'pBdr')[0];
      if(spacing){
        if(attr(spacing,'before')!==null)spec.paragraph.spaceBefore=pt(attr(spacing,'before'));
        if(attr(spacing,'after')!==null)spec.paragraph.spaceAfter=pt(attr(spacing,'after'));
        if(attr(spacing,'line')!==null&&(attr(spacing,'lineRule')||'auto')==='auto')spec.paragraph.lineSpacing=Math.round(Number(attr(spacing,'line'))/240*12*100)/100;
      }
      if(ind){
        if(attr(ind,'left')!==null)spec.paragraph.leftIndent=pt(attr(ind,'left'));
        if(attr(ind,'hanging')!==null)spec.paragraph.firstLineIndent=-pt(attr(ind,'hanging'));
        else if(attr(ind,'firstLine')!==null)spec.paragraph.firstLineIndent=pt(attr(ind,'firstLine'));
      }
      if(kids(pPr,'keepNext')[0])spec.paragraph.keepWithNext=on(kids(pPr,'keepNext')[0]);
      if(kids(pPr,'keepLines')[0])spec.paragraph.keepTogether=on(kids(pPr,'keepLines')[0]);
      if(jc&&attr(jc,'val')==='center')spec.paragraph.alignment='Centered';
      if(shd&&/^[0-9a-fA-F]{6}$/.test(attr(shd,'fill')||''))spec.shading=`#${attr(shd,'fill').toUpperCase()}`;
      if(bdr)for(const side of Object.keys(SIDE)){
        const b=kids(bdr,side)[0];
        if(b&&attr(b,'val')!=='nil'&&attr(b,'val')!=='none')spec.borders.push({location:SIDE[side],color:/^[0-9a-fA-F]{6}$/.test(attr(b,'color')||'')?`#${attr(b,'color').toUpperCase()}`:'#000000',width:STEP_NAMES[nearestStep(Number(attr(b,'sz')||4)/8)]});
      }
    }
    // Word writes the font a style had at the time into an imported style, so a guide made before the look keeps its old body font in
    // every body style however Normal changes. Each body style (not a heading's) therefore says the look's body font and size itself.
    const based=attr(kids(style,'basedOn')[0],'val')||'';
    if(/^OSG /.test(name)&&!/^(Heading|Title)/.test(based)){
      if(spec.font.name===undefined&&theme.bodyFont)spec.font.name=theme.bodyFont;
      if(spec.font.size===undefined&&theme.bodyHalfPoints)spec.font.size=theme.bodyHalfPoints/2;
    }
    return {id:attr(style,'styleId'),based,spec};
  });
  // The same flattening hits the guide's own headings (OSG Module, OSG Lesson, OSG Guide Title, OSG Assessment): they say nothing of
  // their own in the definitions, but Word wrote the heading values of the look the guide began with into them. Found in real Word on
  // 2026-11-06: after a look change Heading 1, Heading 2 and Title followed the look and these four did not. Each therefore says the
  // font and paragraph settings of the heading style it is based on, in this look.
  const byId=new Map(entries.map(e=>[e.id,e]));
  for(const e of entries){
    const base=/^OSG /.test(e.spec.name)&&/^(Heading|Title)/.test(e.based)?byId.get(e.based):null;
    if(!base)continue;
    for(const part of ['font','paragraph'])for(const [key,value] of Object.entries(base.spec[part]))if(e.spec[part][key]===undefined)e.spec[part][key]=value;
  }
  return entries.map(e=>e.spec);
}
// The document's own body text: font, size and line spacing come from the look, through the Normal style.
export const normalSpec=theme=>({name:'Normal',font:{name:theme.bodyFont,size:(theme.bodyHalfPoints||24)/2},paragraph:{lineSpacing:Math.round((theme.line||276)/240*12*100)/100,spaceAfter:6},shading:null,borders:[]});
export const lookSpecs=(theme,DOMParserClass)=>[normalSpec(theme),...styleSpecs(theme,DOMParserClass).filter(s=>s.name!=='Normal')];

// ---- header and footer, as packages for Word's insertOoxml
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const run=(text,extra='')=>`<w:r>${extra?`<w:rPr>${extra}</w:rPr>`:''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const field=(code,result)=>`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const pkg=body=>`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document ${NS}><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const small='<w:sz w:val="18"/><w:szCs w:val="18"/>';
// Running header: the guide's title. Running footer: who made it and when it was revised on the left, "Page X of Y" on the right.
export function headerPackage({title}){
  return pkg(`<w:p><w:pPr><w:jc w:val="right"/></w:pPr>${run(oneLine(title,200),small)}</w:p>`);
}
export function footerPackage({organization,copyright,revised}){
  const left=[oneLine(organization,120),oneLine(copyright,120),revised?`Revised ${oneLine(revised,40)}`:''].filter(Boolean).join(' · ');
  return pkg(`<w:p><w:pPr><w:tabs><w:tab w:val="clear" w:pos="4680"/><w:tab w:val="clear" w:pos="9360"/><w:tab w:val="right" w:pos="9360"/></w:tabs></w:pPr>${run(left,small)}<w:r><w:tab/></w:r>${run('Page ',small)}${field('PAGE','1')}${run(' of ',small)}${field('NUMPAGES','1')}</w:p>`);
}
