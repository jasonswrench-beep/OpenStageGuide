// The learner workbook: a separate Word file made from the paragraphs of the guide that are for learners.
// It is built from what is KEPT. Nothing is copied and then hidden or deleted, so a facilitator paragraph that is not
// kept is simply not in the file. A second, independent check (workbook-leak.mjs) looks at the finished file.
// Pure functions: the XML parser and serializer are passed in, so this runs in the pane, in tests and on the command line.
import {STYLE_NAMES,STYLE_IDS,NS,stylesPart} from './guide-styles.mjs';
import {kindOf} from './guide-model.mjs';
import {readStyles,textOf} from './guide-check.mjs';
import {resolveWorkbookBrand,textContrast,WORKBOOK_FAMILIES} from './workbook-presentation.mjs';
import {allCues,cueById} from './guide-cues.mjs';
import {esc,para,run} from './guide-ooxml.mjs';
import {landscapeFlags} from './guide-forms.mjs';

// The looks the author can choose from: the five the product already has, each with a few words on what it is for.
const LOOK_NOTES={clearline:'Spacious, one column.','editorial-study':'Reference style with serif headings.','practice-lab':'Numbered responses.','field-guide':'Step by step.','reflection-journal':'Prompt and response.'};
export const WORKBOOK_LOOKS=WORKBOOK_FAMILIES.map(f=>({id:f.id,name:f.name,note:LOOK_NOTES[f.id]||''}));

export const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14='http://schemas.microsoft.com/office/word/2010/wordml';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG='http://schemas.microsoft.com/office/2006/xmlPackage';
const REL_NS='http://schemas.openxmlformats.org/package/2006/relationships';

const clean=t=>String(t??'').replace(/[\r\n\u0007]+$/,'').trim();
const els=n=>Array.from(n.childNodes).filter(c=>c.nodeType===1);
const isOn=n=>n&&!['false','0','off'].includes((n.getAttributeNS(W,'val')||'').toLowerCase());

// ---- who is a paragraph for?
// The workbook keeps blocks for everyone and the whole of an assessment section except its keys; everything else is left
// out. A cue's extra lines do not say who they are for, so they take it from the cue above them.
const sectionEnds=(p,k)=>k==='module'||k==='lesson'||k==='assessment'||k==='guide-title'||k==='title'||/^heading\b/i.test(String(p.style||''))||/^Heading/.test(String(p.builtIn||''));
const BLOCK_MINUTES=/\s*\(\d+ min\)\s*$/;

// paras: [{text,style,builtIn}] in document order, as Word reports them. Returns one decision for each paragraph.
export function classify(paras){
  const kinds=paras.map(kindOf),n=paras.length;
  const decisions=paras.map(()=>({keep:false,reason:'own'}));
  // The front pages (title, contents, Getting Started) run from the start to the first module, lesson or assessment heading.
  let firstBody=kinds.findIndex(k=>k==='module'||k==='lesson'||k==='assessment');
  if(firstBody<0)firstBody=n;
  // An assessment section has two parts. First its lead: the note to the learner and the information lines the pane writes
  // (plain paragraphs, up to LEAD_LINES of them), which the author may edit. Then its items, where only the paragraphs the
  // pane's own styles mark as an item are kept. Anything else typed there is the author's own text: it is left out and listed.
  // A form (guide-forms.mjs) is a span: its title paragraph says who it is for, and the paragraphs of its elements take that from it.
  let cueAudience=null,formAudience=null,section=null,leadLines=0,lesson='',title='',overviewTable=false;
  const objectives=[];
  for(let i=0;i<n;i++){
    const k=kinds[i],text=clean(paras[i].text),d=decisions[i];
    const keep=reason=>{d.keep=true;d.reason=reason;};
    const drop=reason=>{d.keep=false;d.reason=reason;};
    if(k==='guide-title'){if(!title)title=text;drop('title');continue;}
    if(i<firstBody){drop(text?'front':'blank');continue;}
    // A module's overview is for the facilitator: its labelled lines, and its table of lessons (told by its first header, "Lesson").
    if(k!=='table-head'&&k!=='table-cell')overviewTable=false;
    if(k==='table-head'&&!section&&!formAudience&&/^lesson$/i.test(text))overviewTable=true;
    if(k==='module'){keep('heading');cueAudience=null;formAudience=null;section=null;continue;}
    if(k==='lesson'){keep('heading');cueAudience=null;formAudience=null;section=null;lesson=text;continue;}
    if(k==='assessment'){keep('heading');cueAudience=null;formAudience=null;section='lead';leadLines=0;continue;}
    if(sectionEnds(paras[i],k)){section=null;cueAudience=null;formAudience=null;drop(text?'own':'blank');continue;}
    if(section==='lead'&&k!=='other')section='items';
    switch(k){
      case 'cue':{
        formAudience=null;
        cueAudience=paras[i].style===STYLE_NAMES.cueFacilitator?'facilitator':'everyone';
        if(cueAudience==='facilitator'){drop('facilitator-block');break;}
        keep('block');
        // Word can leave a placeholder character where a picture is, so allow junk before the label.
        const label=cueById('objective')?.label||'Objective',m=new RegExp(`^[^A-Za-z]*${label}\\s+(.*?)\\s*(?:\\(\\d+ min\\))?$`,'s').exec(text);
        if(m&&m[1])objectives.push({lesson,text:m[1].trim()});
        break;
      }
      case 'body':
        if(cueAudience==='everyone')keep('block');
        else drop(cueAudience==='facilitator'?'facilitator-block':'own');
        break;
      case 'form-title':
        formAudience=paras[i].style===STYLE_NAMES.formTitleFacilitator?'facilitator':'everyone';
        cueAudience=null;
        if(formAudience==='everyone')keep('form');else drop('facilitator-form');
        break;
      case 'form-text':case 'form-line':case 'form-field':case 'form-box':case 'form-source':case 'form-card':
        if(formAudience==='everyone')keep('form');
        else drop(formAudience==='facilitator'?'facilitator-form':(text?'own':'blank'));
        break;
      case 'form-print':drop('form-print');break;
      case 'notes':drop('notes');break;
      case 'time':drop('time');break;
      case 'answer-key':case 'key-table-head':case 'key-table-cell':drop('key');break;
      case 'overview':drop('overview');break;
      case 'question':case 'option':case 'answer-line':case 'item-note':case 'table-head':case 'table-cell':case 'rating-box':
        // A table or check box inside a form follows the form's audience.
        if(formAudience&&!section&&(k==='table-head'||k==='table-cell'||k==='rating-box')){if(formAudience==='everyone')keep('form');else drop('facilitator-form');break;}
        if(overviewTable&&!section&&(k==='table-head'||k==='table-cell')){drop('overview');break;}
        if(section)keep('item');else drop(text||k==='answer-line'?'own':'blank');
        break;
      default:
        // Anything else is the author's own paragraph. In a section's lead, a plain paragraph is the note or an information
        // line and is kept; a paragraph in any other style (a pasted key whose style Word renamed, say) ends the lead and is
        // left out. Everywhere else it is left out and listed.
        if(section==='lead'&&text&&isPlain(paras[i])&&leadLines<LEAD_LINES){keep('section');leadLines++;}
        else{if(section==='lead'&&text)section='items';drop(text?'own':'blank');}
    }
  }
  return {decisions,title,objectives,kinds};
}
const LEAD_LINES=8;
const isPlain=p=>p.style==='Normal'||p.builtIn==='Normal';

// What the pane shows before the workbook is made, from the paragraphs alone.
export function summarize(paras){
  const {decisions,title,objectives,kinds}=classify(paras);
  const count=r=>decisions.filter(d=>d.reason===r).length;
  const runs=reason=>decisions.reduce((c,d,i)=>c+(d.reason===reason&&!(i>0&&decisions[i-1].reason===reason)?1:0),0);
  // Headings alone are not something to make a workbook from: it needs a block for learners or an assessment.
  const kept=decisions.filter(d=>d.keep&&d.reason!=='heading').length;
  const own=decisions.map((d,i)=>d.reason==='own'?{index:i,text:clean(paras[i].text)}:null).filter(Boolean);
  const front=decisions.findIndex(d=>d.reason==='front');
  return {
    title,isGuide:kinds.some(k=>k==='guide-title'||k==='module'||k==='lesson'),kept,
    modules:kinds.filter(k=>k==='module').length,lessons:kinds.filter(k=>k==='lesson').length,
    blocks:decisions.filter((d,i)=>d.reason==='block'&&kinds[i]==='cue').length,
    items:decisions.filter((d,i)=>d.keep&&kinds[i]==='question').length,
    forms:decisions.filter((d,i)=>d.reason==='form'&&kinds[i]==='form-title').length,
    left:{
      facilitatorBlocks:decisions.filter((d,i)=>d.reason==='facilitator-block'&&kinds[i]==='cue').length,
      facilitatorForms:decisions.filter((d,i)=>d.reason==='facilitator-form'&&kinds[i]==='form-title').length,
      notes:runs('notes'),keys:runs('key'),timeBlocks:count('time'),front:count('front'),overview:runs('overview'),own:own.length
    },
    frontAt:front,leftOut:own,objectives:objectives.length
  };
}

// ---- reading what Word gives back for the whole document
export function readSource(pkgXml,{DOMParserClass=globalThis.DOMParser}={}){
  const dom=new DOMParserClass().parseFromString(pkgXml,'application/xml');
  const parts=new Map();
  for(const part of Array.from(dom.getElementsByTagNameNS(PKG,'part'))){
    const kids=els(part),xmlData=kids.find(n=>n.localName==='xmlData'),bin=kids.find(n=>n.localName==='binaryData');
    parts.set(part.getAttributeNS(PKG,'name'),{type:part.getAttributeNS(PKG,'contentType'),xml:xmlData?els(xmlData)[0]||null:null,base64:bin?bin.textContent.replace(/\s+/g,''):null});
  }
  const docRoot=parts.get('/word/document.xml')?.xml;
  if(!docRoot)throw new Error('Word did not return the document, so the workbook could not be made. Try again.');
  const body=els(docRoot).find(n=>n.localName==='body');
  const stylesRoot=parts.get('/word/styles.xml')?.xml||null,styles=readStyles(stylesRoot);
  const rels=new Map();
  for(const r of els(parts.get('/word/_rels/document.xml.rels')?.xml||{childNodes:[]}))rels.set(r.getAttribute('Id'),{type:r.getAttribute('Type'),target:r.getAttribute('Target')});
  const lang=stylesRoot?Array.from(stylesRoot.getElementsByTagNameNS(W,'lang'))[0]?.getAttributeNS(W,'val'):'';
  return {parts,body,styles,rels,language:lang||'en-US'};
}

// The body's paragraphs and tables in order, with block-level content controls opened up.
function bodyEntries(body){
  const out=[];
  const visit=node=>{
    for(const c of els(node)){
      if(c.namespaceURI!==W)continue;
      if(c.localName==='p'||c.localName==='tbl')out.push(c);
      else if(c.localName==='sdt'){const content=els(c).find(x=>x.localName==='sdtContent');if(content)visit(content);}
      else if(c.localName==='customXml')visit(c);
    }
  };
  visit(body);
  return out;
}
const paragraphStyleId=p=>{const ppr=els(p).find(n=>n.localName==='pPr'),s=ppr&&els(ppr).find(n=>n.localName==='pStyle');return s?s.getAttributeNS(W,'val'):'Normal';};
const paragraphInfo=(p,styles)=>{
  const id=paragraphStyleId(p);
  return {text:textOf(p),style:styles.get(id)?.name||id,builtIn:/^(Heading\d|Title|Normal)$/.test(id)?id:undefined};
};

// ---- making the file
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC='http://schemas.openxmlformats.org/drawingml/2006/picture';
const MC='http://schemas.openxmlformats.org/markup-compatibility/2006';
const DROP=new Set(['bookmarkStart','bookmarkEnd','proofErr','commentRangeStart','commentRangeEnd','commentReference','del','moveFrom',
  'moveFromRangeStart','moveFromRangeEnd','moveToRangeStart','moveToRangeEnd','footnoteReference','endnoteReference','pict','object','sectPr','permStart','permEnd',
  'delText','delInstrText','smartTagPr','customXmlPr','txbxContent','cellIns','cellDel','cellMerge']);
const UNWRAP=new Set(['ins','moveTo','hyperlink','smartTag','customXml','fldSimple']);

function stripAttributes(el){
  for(const a of Array.from(el.attributes||[]))
    if(/^rsid/i.test(a.localName||'')||a.localName==='paraId'||a.localName==='textId')el.removeAttributeNode(a);
}
// The three ways Word hides text: hidden, hidden when the page is read as a web page, and a paragraph mark that is always hidden.
const HIDING=new Set(['vanish','webHidden','specVanish']);
// Styles (of characters or of paragraphs) whose text is hidden, by themselves or through the style they are based on.
function hiddenStyleIds(styles){
  const hides=s=>{const rpr=s.node&&els(s.node).find(x=>x.localName==='rPr');return !!rpr&&els(rpr).some(x=>HIDING.has(x.localName)&&isOn(x));};
  const out=new Set();
  for(const [id] of styles){
    const seen=new Set();let cur=id;
    while(cur&&styles.has(cur)&&!seen.has(cur)){seen.add(cur);const s=styles.get(cur);if(hides(s)){out.add(id);break;}cur=s.basedOn;}
  }
  return out;
}
// An icon is a plain picture sitting in the line of text. Anything else a drawing can be (a text box, a shape, a chart, a
// picture floating on the page) can hold words of its own, so it does not come through.
function isInlinePicture(drawing){
  const kids=els(drawing);
  if(kids.some(k=>k.namespaceURI===WP&&k.localName==='anchor'))return false;
  const inline=kids.find(k=>k.namespaceURI===WP&&k.localName==='inline');
  if(!inline)return false;
  const data=Array.from(inline.getElementsByTagNameNS(A,'graphicData'))[0];
  return !!data&&data.getAttribute('uri')===PIC&&inline.getElementsByTagNameNS(PIC,'pic').length===1;
}
// Takes out of kept content everything a workbook must not carry: tracked changes (insertions and moves are accepted,
// deletions dropped) and formatting revisions, comments, notes, hidden text (by itself or through a character style), links,
// fields (flattened to their results), legacy pictures, text boxes and other drawings, and the author's own content controls
// (the rating check boxes stay). Works in place, on a copy the caller made.
function sanitize(node,hidden=new Set()){
  const move=(from,to,before)=>{for(const k of Array.from(from.childNodes))to.insertBefore(k,before);};
  for(const c of Array.from(node.childNodes)){
    if(c.nodeType!==1)continue;
    const name=c.localName,isW=c.namespaceURI===W;
    if(c.namespaceURI===MC&&name==='AlternateContent'){node.removeChild(c);continue;}
    if(isW&&(DROP.has(name)||/Change$/.test(name))){node.removeChild(c);continue;}
    if(isW&&UNWRAP.has(name)){sanitize(c,hidden);move(c,node,c);node.removeChild(c);continue;}
    if(isW&&name==='drawing'){if(isInlinePicture(c)){stripAttributes(c);sanitize(c,hidden);}else node.removeChild(c);continue;}
    // Ruby text sits above the words it explains and is easy to overlook; only the words themselves (the base) come through.
    if(isW&&name==='ruby'){
      const base=els(c).find(x=>x.localName==='rubyBase');
      if(base){sanitize(base,hidden);move(base,node,c);}
      node.removeChild(c);continue;
    }
    // A paragraph in a style that hides its text keeps its (empty) place, so a table cell still has a paragraph, but none of its words.
    if(isW&&name==='p'){
      const pr=els(c).find(x=>x.localName==='pPr'),style=pr&&els(pr).find(x=>x.localName==='pStyle');
      if(style&&hidden.has(style.getAttributeNS(W,'val'))){
        for(const k of Array.from(c.childNodes))if(!(k.nodeType===1&&k.localName==='pPr'))c.removeChild(k);
        stripAttributes(c);continue;
      }
    }
    if(isW&&name==='sdt'){
      const pr=els(c).find(x=>x.localName==='sdtPr'),content=els(c).find(x=>x.localName==='sdtContent');
      const isBox=pr&&Array.from(pr.childNodes).some(x=>x.nodeType===1&&x.namespaceURI===W14&&x.localName==='checkbox');
      if(isBox){stripAttributes(c);sanitize(c,hidden);continue;}
      if(content){sanitize(content,hidden);move(content,node,c);}
      node.removeChild(c);continue;
    }
    if(isW&&name==='r'){
      const kids=els(c),rpr=kids.find(x=>x.localName==='rPr');
      const off=rpr&&els(rpr).some(x=>HIDING.has(x.localName)&&isOn(x));
      const style=rpr&&els(rpr).find(x=>x.localName==='rStyle');
      const field=kids.some(x=>x.namespaceURI===W&&(x.localName==='fldChar'||x.localName==='instrText'));
      if(off||field||(style&&hidden.has(style.getAttributeNS(W,'val')))){node.removeChild(c);continue;}
    }
    if(isW&&HIDING.has(name)){node.removeChild(c);continue;}
    stripAttributes(c);
    sanitize(c,hidden);
  }
}
// The source with kept content cleaned before anything is read from it, so words that are hidden or revised away can never
// reach the title page, the contents or the learning objectives.
function prepare(source){
  const body=source.body.cloneNode(true);
  sanitize(body,hiddenStyleIds(source.styles));
  return {...source,body};
}

// The page: Letter or A4 with the look's margins. The footer (page numbers) shows on every page but the first.
const pageSize=effective=>{const a4=effective.pageSize==='A4';return {w:a4?11906:12240,h:a4?16838:15840,m:Math.round(effective.marginIn*1440)};};
// A landscape section (a form printed landscape) has the page turned on its side and a footer of its own, whose page number sits at the wider page's right margin.
// Only the first section's first page goes without a footer (the cover).
const sectPr=(effective,{landscape=false,titlePg=true}={})=>{
  const {w,h,m}=pageSize(effective),[pw,ph]=landscape?[h,w]:[w,h];
  return `<w:sectPr><w:footerReference w:type="default" r:id="${landscape?'rIdFooterWide':'rIdFooter'}"/><w:pgSz w:w="${pw}" w:h="${ph}"${landscape?' w:orient="landscape"':''}/><w:pgMar w:top="${m}" w:right="${m}" w:bottom="${m}" w:left="${m}" w:header="720" w:footer="720" w:gutter="0"/>${titlePg?'<w:titlePg/>':''}</w:sectPr>`;
};
const field=(code,result)=>`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
// The footer: the organization's name and the title on the left, "Page X of Y" on the right.
const footerXml=(effective,title,organization,{landscape=false}={})=>{
  const {w:portraitWidth,h,m}=pageSize(effective),w=landscape?h:portraitWidth,left=[organization,title].filter(Boolean).join(' · ');
  return `<w:ftr ${NS}><w:p><w:pPr><w:pStyle w:val="${WORKBOOK_ONLY_STYLES.footer}"/><w:tabs><w:tab w:val="right" w:pos="${w-2*m}"/></w:tabs></w:pPr>${run(left)}<w:r><w:tab/></w:r>${run('Page ')}${field('PAGE','1')}${run(' of ')}${field('NUMPAGES','1')}</w:p></w:ftr>`;
};
const hex6=c=>String(c).replace('#','').toUpperCase();
const grayOf=hex=>{const [r,g,b]=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)),v=Math.round(0.299*r+0.587*g+0.114*b).toString(16).padStart(2,'0');return `${v}${v}${v}`.toUpperCase();};

// What a look is: the fonts, sizes, colors and page the workbook's styles are written with. In black and white the accent is a
// dark grey and nothing is told by color alone.
export function lookTheme(effective,language='en-US'){
  const gray=effective.printMode==='grayscale',accent=gray?'333333':hex6(effective.accent),text=hex6(effective.text);
  return {navy:accent,ink:text,bar:accent,shade:gray?'EDEDED':'F1EBDD',headingFont:effective.headingFont,bodyFont:effective.bodyFont,
    bodyHalfPoints:Math.round(effective.bodyPt*2),line:Math.round(effective.lineSpacing*240),lang:language,
    // space above each line to write on, so that a line is about responseLineIn inches from the one before
    answerBefore:Math.max(60,Math.round(effective.responseLineIn*1440-effective.bodyPt*20*1.2)),
    sizes:{title:Math.round(effective.headingPt*3.25),h1:Math.round(effective.headingPt*2),h2:Math.round((effective.headingPt-2)*2)}};
}

const WORKBOOK_STYLE_IDS=new Set(['Normal','DefaultParagraphFont','Title','Heading1','Heading2',STYLE_IDS.guideTitle,STYLE_IDS.module,STYLE_IDS.lesson,
  STYLE_IDS.cue,STYLE_IDS.cueLabel,STYLE_IDS.body,STYLE_IDS.assessment,STYLE_IDS.question,STYLE_IDS.option,STYLE_IDS.answerLine,
  STYLE_IDS.tableHead,STYLE_IDS.tableCell,STYLE_IDS.ratingBox,STYLE_IDS.itemNote,
  STYLE_IDS.formTitle,STYLE_IDS.formText,STYLE_IDS.formLine,STYLE_IDS.formField,STYLE_IDS.formBox,STYLE_IDS.formSource,STYLE_IDS.formCard]);
export const WORKBOOK_ONLY_STYLES={subtitle:'OSGWorkbookSubtitle',bullet:'OSGWorkbookBullet',footer:'OSGWorkbookFooter'};
const extraStyles=theme=>`
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${WORKBOOK_ONLY_STYLES.footer}"><w:name w:val="OSG Workbook Footer"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:color w:val="${theme.ink}"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${WORKBOOK_ONLY_STYLES.subtitle}"><w:name w:val="OSG Workbook Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="360"/></w:pPr><w:rPr><w:color w:val="${theme.navy}"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${WORKBOOK_ONLY_STYLES.bullet}"><w:name w:val="OSG Workbook Bullet"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720" w:hanging="360"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="60"/><w:ind w:left="360"/></w:pPr></w:style>`;

const workbookStyles=theme=>{
  // The styles written are the learner ones only, and the lesson style no longer points at the time block that is left out.
  const base=stylesPart(theme,WORKBOOK_STYLE_IDS).replace(/<w:next w:val="OSGTime"\/>/g,'<w:next w:val="Normal"/>');
  return base.replace('</w:styles>',`${extraStyles(theme)}</w:styles>`);
};

const numberingXml=(listCount,bullets)=>{
  const lettered=Array.from({length:listCount},(_,i)=>`<w:num w:numId="${i+1}"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('');
  const bullet=bullets?`<w:num w:numId="${listCount+1}"><w:abstractNumId w:val="1"/></w:num>`:'';
  return `<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr></w:lvl></w:abstractNum>${lettered}${bullet}</w:numbering>`;
};

// A colon in a title becomes " - ", and the other characters Windows will not allow in a file name are taken out.
const safeName=title=>String(title||'Guide').replace(/\s*:\s*/g,' - ').replace(/[\\/*?"<>|]+/g,' ').replace(/\s+/g,' ').trim()||'Guide';
export const workbookFileName=title=>`${safeName(title)} - Participant workbook.docx`;
export const pdfFileName=title=>`${safeName(title)}.pdf`;
export const workbookTitle=title=>`${String(title||'Guide').trim()||'Guide'} - Participant workbook`;

// Every paragraph of a source document, in order, the way Word counts them (a table's cells are paragraphs too),
// and which of them belong to which top-level paragraph or table.
function flatten(source){
  const entries=bodyEntries(source.body),flat=[];
  const owner=entries.map(e=>e.localName==='p'?[flat.push(e)-1]:Array.from(e.getElementsByTagNameNS(W,'p')).map(p=>flat.push(p)-1));
  return {entries,owner,paras:flat.map(p=>paragraphInfo(p,source.styles))};
}
const FACILITATOR_REASONS=['facilitator-block','facilitator-form','form-print','notes','key','time'];
// A facilitator paragraph starts with a label the pane wrote ("Facilitator notes:", "Answer key (...):", a cue's name), and
// someone copying only its words would leave the label behind. So the whole paragraph and what is left after taking off a
// block's minutes and the label are both searched for; the second only when it is long enough to be a safe thing to search for.
// Built each time it is used, so the author's own cues are in it; longest names first, so "Slide notes" is not read as "Slide".
const labelPattern=()=>new RegExp(`^[^A-Za-z]*(?:Facilitator notes:|Answer key \\([^)]*\\):|(?:Why|Points|Objective):|(?:${[...allCues()].sort((a,b)=>b.label.length-a.label.length).map(c=>c.label.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|')})\\s)\\s*`,'i');
const SEARCHABLE=12;
const letters=s=>s.replace(/[^\p{L}\p{N}]+/gu,'').length;
const structural=p=>/^(heading|toc|title)\b/i.test(String(p.style||''))||/^(Heading|Title)/.test(String(p.builtIn||''));
// What the finished file must not contain: every facilitator paragraph, and the author's own text that was left out
// (outside headings and the contents, which the workbook writes again itself).
const facilitatorTextsFrom=(paras,decisions)=>{
  const out=new Set();
  paras.forEach((p,i)=>{
    const reason=decisions[i].reason,full=clean(p.text);
    if(!full)return;
    if(FACILITATOR_REASONS.includes(reason)){
      const bare=full.replace(BLOCK_MINUTES,'').trim(),body=bare.replace(labelPattern(),'').trim();
      for(const v of [full,bare,body])if(v&&(v===full||letters(v)>=SEARCHABLE))out.add(v);
    }else if((reason==='own'||reason==='front'||reason==='overview')&&!structural(p))out.add(full);
  });
  return [...out];
};
// What the file may legitimately say more than once because the guide says it: the kept paragraphs, and the pages the workbook
// writes from them (title, contents entries, learning objectives). A facilitator paragraph's words are only a leak when the
// file holds them more often than these do.
const allowedTextsFrom=(paras,decisions,result)=>[
  ...paras.filter((p,i)=>decisions[i].keep).map(p=>clean(p.text)),
  result.title,'Participant workbook','Name:','Contents','Learning objectives',
  ...paras.filter((p,i)=>decisions[i].keep&&decisions[i].reason==='heading').map(p=>clean(p.text)),
  ...result.objectives.flatMap(o=>[o.text,o.lesson])
].filter(Boolean);
// For checking a saved workbook against its guide: source is {body, styles}, which can be read from Word or from a saved .docx.
export function leakInputsOf(source){
  const {paras}=flatten(prepare(source)),result=classify(paras);
  return {facilitatorTexts:facilitatorTextsFrom(paras,result.decisions),allowedTexts:allowedTextsFrom(paras,result.decisions,result)};
}
export const facilitatorTextsOf=source=>leakInputsOf(source).facilitatorTexts;

// Takes a trailing "(5 min)" off a paragraph however the runs are split, for a block's minutes are the facilitator's.
function trimMinutes(p){
  const ts=Array.from(p.getElementsByTagNameNS(W,'t')),m=BLOCK_MINUTES.exec(ts.map(t=>t.textContent).join(''));
  if(!m)return;
  let n=m[0].length;
  for(let i=ts.length-1;i>=0&&n>0;i--){const s=ts[i].textContent,cut=Math.min(n,s.length);ts[i].textContent=s.slice(0,s.length-cut);n-=cut;}
}

// In black and white, every color in the kept content becomes the grey of the same lightness, so nothing is told by color alone.
function grayify(node){
  for(const el of [node,...Array.from(node.getElementsByTagName('*'))])
    for(const a of Array.from(el.attributes||[]))
      if(a.namespaceURI===W&&(a.localName==='fill'||a.localName==='color'||(a.localName==='val'&&el.localName==='color'))&&/^[0-9a-f]{6}$/i.test(a.value))el.setAttributeNS(W,a.name,grayOf(a.value));
}
// Blocks that ask learners to write or think are followed by lines to write on.
const WRITING_CUES=new Set(['ask','discussion','activity','workbook']);
const escapeRegExp=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
const cueIdOf=text=>[...allCues()].sort((a,b)=>b.label.length-a.label.length).find(c=>new RegExp(`^[^A-Za-z]*${escapeRegExp(c.label)}\\s`,'i').test(text))?.id;
export const MAX_WRITING_LINES=6;

// source: readSource(...). options: familyId (a look), brand (organizationName, accent, pageSize, printMode and the rest of the
// look's own settings, as resolveWorkbookBrand takes them), lines (lines to write on after each question or activity, 0 to 6).
// Returns {files:[{name,text|base64}],summary,facilitatorTexts,allowedTexts,title}.
export function buildWorkbook(source,{familyId='clearline',brand={},lines=3}={},{DOMParserClass=globalThis.DOMParser,SerializerClass=globalThis.XMLSerializer}={}){
  if(!Number.isInteger(lines)||lines<0||lines>MAX_WRITING_LINES)throw new Error(`Lines to write on must be a whole number from 0 to ${MAX_WRITING_LINES}.`);
  const look=resolveWorkbookBrand(familyId,brand);
  if(!look.eligible)throw new Error(`This look cannot be used: ${look.issues.join(' ')}`);
  const effective=look.effective,theme=lookTheme(effective,source.language),gray=effective.printMode==='grayscale';
  // The accent colors the headings, so it must be readable on the page.
  const accentContrast=textContrast(effective.accent,effective.background);
  if(!gray&&accentContrast<4.5)throw new Error(`This look cannot be used: the accent color is too light to read as heading text (its contrast with the page is ${accentContrast.toFixed(1)} to 1, and it needs at least 4.5 to 1).`);
  // Everything is read from a cleaned copy, so hidden or revised-away words cannot reach a page the workbook writes.
  const prepared=prepare(source);
  const {entries,owner,paras}=flatten(prepared);
  const result=classify(paras),{decisions}=result;
  // A table is kept whole or not at all: if any of its words are not for learners, none of the table is.
  entries.forEach((e,ei)=>{
    if(e.localName!=='tbl')return;
    const idx=owner[ei],blocked=idx.some(i=>!decisions[i].keep&&decisions[i].reason!=='blank');
    if(blocked)for(const i of idx)if(decisions[i].keep)decisions[i]={keep:false,reason:'own'};
  });
  const keptEntries=entries.filter((e,ei)=>owner[ei].some(i=>decisions[i].keep));
  if(!decisions.some(d=>d.keep&&d.reason!=='heading'))throw new Error('Nothing in this guide is for learners yet. Add a block for Everyone, or an assessment, first.');
  const title=result.title||'Guide';

  const out=new DOMParserClass().parseFromString(`<w:document ${NS}><w:body/></w:document>`,'application/xml');
  const outBody=els(out.documentElement)[0];
  // A form printed landscape is a section of its own: `flags` says which paragraphs are in one.
  const flags=landscapeFlags(paras);
  let landscapeNow=false,closedSections=0,usedLandscape=false;
  const fragment=xml=>els(new DOMParserClass().parseFromString(`<w:document ${NS}><w:body>${xml}</w:body></w:document>`,'application/xml').documentElement)[0];
  const append=xml=>{for(const n of els(fragment(xml)))outBody.appendChild(out.importNode(n,true));};

  // Lettered lists: each kept option list becomes its own numbering instance in the new file.
  const listIds=new Map();
  const media=[],mediaFor=new Map();

  // The generated pages: title, contents, learning objectives.
  const headings=[];
  for(const e of keptEntries)if(e.localName==='p'){const i=owner[entries.indexOf(e)][0],k=result.kinds[i];if(k==='module'||k==='lesson'||k==='assessment')headings.push({k,text:clean(paras[i].text)});}
  const toc=headings.length?headings.map((h,i)=>{
    const style=h.k==='lesson'?'TOC2':'TOC1',first=i===0,last=i===headings.length-1;
    const begin=first?`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`:'';
    return para(style,begin+run(h.text)+(last?'<w:r><w:fldChar w:fldCharType="end"/></w:r>':''));
  }).join(''):'';
  const needBullets=result.objectives.length>0;
  const pages=[
    para(STYLE_IDS.guideTitle,run(title)),
    para(WORKBOOK_ONLY_STYLES.subtitle,run('Participant workbook')),
    para(STYLE_IDS.answerLine,run('Name:')),
    para('Heading1',run('Contents'),'<w:pageBreakBefore/>'),
    toc||para(null,run('The contents will list the modules and lessons.'))
  ];
  append(pages.join(''));

  const objectivePages=()=>{
    if(!result.objectives.length)return '';
    const parts=[para('Heading1',run('Learning objectives'),'<w:pageBreakBefore/>')];
    let lesson=null;
    for(const o of result.objectives){
      if(o.lesson!==lesson){lesson=o.lesson;if(lesson)parts.push(para(STYLE_IDS.itemNote,run(lesson)));}
      parts.push(para(WORKBOOK_ONLY_STYLES.bullet,run(o.text),`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="@BULLET@"/></w:numPr>`));
    }
    return parts.join('');
  };
  const objectiveXml=objectivePages();
  const objectiveAt=outBody.childNodes.length;

  const remap=numId=>{if(!listIds.has(numId))listIds.set(numId,listIds.size+1);return listIds.get(numId);};
  const own=(n,name)=>els(n).find(x=>x.localName===name);
  const setVal=(el,v)=>el.setAttributeNS(W,'w:val',String(v));
  // Lines to write on after a block that asks learners to write or think, once the block's extra lines are done.
  let pending=0;
  const flush=()=>{
    if(pending<=0)return;
    append(para(STYLE_IDS.answerLine,run('Your notes:'))+Array.from({length:pending-1},()=>para(STYLE_IDS.answerLine,'')).join(''));
    pending=0;
  };
  for(const e of keptEntries){
    const firstKind=result.kinds[owner[entries.indexOf(e)][0]];
    if(firstKind!=='body')flush();
    // Where the orientation changes, the section so far ends in an empty paragraph that carries its page setup.
    const landscape=!!flags[owner[entries.indexOf(e)][0]];
    if(landscape!==landscapeNow){
      flush();
      append(`<w:p><w:pPr>${sectPr(effective,{landscape:landscapeNow,titlePg:closedSections===0&&!landscapeNow})}</w:pPr></w:p>`);
      closedSections++;landscapeNow=landscape;if(landscape)usedLandscape=true;
    }
    const copy=out.importNode(e,true);
    if(gray)grayify(copy);
    const paragraphs=copy.localName==='p'?[copy]:Array.from(copy.getElementsByTagNameNS(W,'p'));
    const idxs=owner[entries.indexOf(e)];
    // Table cells that were blank-and-dropped cannot exist in kept tables, so every paragraph here is kept.
    paragraphs.forEach((p,pi)=>{
      const k=result.kinds[idxs[pi]];
      const ppr=own(p,'pPr');
      if(ppr){
        const numPr=own(ppr,'numPr');
        if(numPr){
          const id=numPr.getElementsByTagNameNS(W,'numId')[0];
          if(k==='option'&&id)setVal(id,remap(Number(id.getAttributeNS(W,'val'))));
          else ppr.removeChild(numPr); // a list the workbook does not carry
        }
        const rpr=own(ppr,'rPr');if(rpr&&!els(rpr).length)ppr.removeChild(rpr);
      }
      // A module or an assessment starts a new page; a block's minutes are for the facilitator, so they go.
      if(k==='module'||k==='assessment'){
        let pp=ppr;
        if(!pp){pp=out.createElementNS(W,'w:pPr');p.insertBefore(pp,p.firstChild);}
        const style=own(pp,'pStyle'),pb=out.createElementNS(W,'w:pageBreakBefore');
        pp.insertBefore(pb,style?style.nextSibling:pp.firstChild);
      }
      if(k==='cue')trimMinutes(p);
      // Pictures: the icons. Each is copied once, under a name of this file's own.
      for(const blip of Array.from(p.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main','blip'))){
        const rid=blip.getAttributeNS(R,'embed'),rel=source.rels.get(rid);
        if(!rel)continue;
        const partName=rel.target.startsWith('/')?rel.target:`/word/${rel.target.replace(/^\.\//,'')}`,part=source.parts.get(partName);
        if(!part?.base64)continue;
        // Each picture keeps its own kind (the icons are PNG; the author's pictures may be JPEG or GIF).
        if(!mediaFor.has(partName)){const n=media.length+1,ext=((/\.([A-Za-z0-9]+)$/.exec(partName)||[])[1]||'png').toLowerCase(),name=`word/media/image${n}.${ext}`;mediaFor.set(partName,{rid:`rIdImg${n}`,name});media.push({name,base64:part.base64});}
        blip.setAttributeNS(R,'r:embed',mediaFor.get(partName).rid);
      }
    });
    outBody.appendChild(copy);
    if(lines>0&&firstKind==='cue'&&WRITING_CUES.has(cueIdOf(clean(textOf(copy)))))pending=lines;
  }
  flush();
  if(objectiveXml){
    const frag=fragment(objectiveXml.replace(/@BULLET@/g,String(listIds.size+1)));
    const before=outBody.childNodes[objectiveAt]||null;
    for(const n of els(frag))outBody.insertBefore(out.importNode(n,true),before);
  }
  append(sectPr(effective,{landscape:landscapeNow,titlePg:closedSections===0&&!landscapeNow}));
  const ser=new SerializerClass(),documentXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${ser.serializeToString(out)}`;

  const lists=listIds.size,needNumbering=lists>0||needBullets;
  const mediaRels=media.map(m=>`<Relationship Id="${[...mediaFor.values()].find(v=>v.name===m.name).rid}" Type="${R}/image" Target="${m.name.replace('word/','')}"/>`).join('');
  const files=[
    {name:'[Content_Types].xml',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>${usedLandscape?'<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>':''}${needNumbering?'<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>':''}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`},
    {name:'_rels/.rels',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`},
    {name:'word/document.xml',text:documentXml},
    {name:'word/_rels/document.xml.rels',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL_NS}"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdSettings" Type="${R}/settings" Target="settings.xml"/><Relationship Id="rIdFooter" Type="${R}/footer" Target="footer1.xml"/>${usedLandscape?`<Relationship Id="rIdFooterWide" Type="${R}/footer" Target="footer2.xml"/>`:''}${needNumbering?`<Relationship Id="rIdNumbering" Type="${R}/numbering" Target="numbering.xml"/>`:''}${mediaRels}</Relationships>`},
    {name:'word/styles.xml',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${workbookStyles(theme)}`},
    {name:'word/settings.xml',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${W}"><w:themeFontLang w:val="${esc(source.language)}"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`},
    {name:'docProps/core.xml',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(workbookTitle(title))}</dc:title><dc:language>${esc(source.language)}</dc:language></cp:coreProperties>`}
  ];
  files.push({name:'word/footer1.xml',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${footerXml(effective,title,effective.organizationName)}`});
  if(usedLandscape)files.push({name:'word/footer2.xml',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${footerXml(effective,title,effective.organizationName,{landscape:true})}`});
  if(needNumbering)files.push({name:'word/numbering.xml',text:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${numberingXml(lists,needBullets)}`});
  for(const m of media)files.push({name:m.name,base64:m.base64});

  return {files,title,workbookTitle:workbookTitle(title),fileName:workbookFileName(title),lookId:familyId,effective,lines,
    facilitatorTexts:facilitatorTextsFrom(paras,decisions),allowedTexts:allowedTextsFrom(paras,decisions,result),
    summary:summarize(paras)};
}

// The files as a .docx, as base64 (what Word's createDocument takes) or as bytes.
export async function zipWorkbook(files,Zip,type='base64'){
  const zip=new Zip();
  // No folder entries: a Word file lists its parts and nothing else.
  for(const f of files)f.base64!==undefined?zip.file(f.name,f.base64,{base64:true,createFolders:false}):zip.file(f.name,f.text,{createFolders:false});
  return zip.generateAsync({type,compression:'DEFLATE'});
}
