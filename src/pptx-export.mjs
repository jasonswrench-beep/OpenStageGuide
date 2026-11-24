// The slide deck writer: one plain data model in, one editable 16:9 PowerPoint file out. Every slide holds real text in real
// placeholders (a title, and a body or subtitle), and the speaker notes are real notes, so the file restyles, reads aloud and
// reopens the way any PowerPoint file does. There are no pictures, so no text sits over a picture.
//
// The model is checked first (checkDeck) and a bad one is refused in plain words. slideFit says which slides hold more than a room can
// read, by a rough rule (see SLIDE_LIMITS). The look is one master: navy titles on white, a thin teal bar across the top, Arial through
// the theme, 36 pt titles (44 pt on the opening slide) and 24 pt bullets. Every colour passes WCAG AA against white.
//
// Layouts: 'Title Slide' for an 'opening' slide (title and subtitle placeholders, plain paragraphs), 'Title and Content' for the rest
// (title and body placeholders, bulleted). A slide with no lines has no body shape, so PowerPoint never shows "Click to add text".
// The deck subtitle is the document subject, and it fills the subtitle of an opening slide that has no lines of its own.
//
// Pure: no file or network access. The only clock read is the time stamp; pass { now } to fix it. JSZip is passed in, as the other
// writers do. A model is never trusted: every text is cleaned of control characters and escaped before it is written.

export const SLIDE_KINDS=Object.freeze(['opening','activity','break','close']);
// fitLines, fitChars and fitLineChars are the slideFit rule: a 24 pt line holds about 60 characters across a 16:9 body and the body
// shows about 7 lines, so over 7 lines or 420 characters is too much, and one line over 140 characters would wrap onto a third row.
export const SLIDE_LIMITS=Object.freeze({titleChars:200,subtitleChars:200,slides:300,lines:12,lineChars:300,notes:40,noteChars:600,keyChars:80,fitLines:7,fitChars:420,fitLineChars:140});
const L=SLIDE_LIMITS,SIZE={cx:12192000,cy:6858000};

// ---- the model -------------------------------------------------------------------------------------------------------------------
const fail=message=>{throw new Error(message);};
const BAD_XML=/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g;
// Line breaks and tabs become one space; every other control character, and every character XML cannot hold, is dropped.
const clean=s=>s.replace(/[\t\r\n\u2028\u2029]+/g,' ').replace(BAD_XML,'').trim();
const size=s=>Array.from(s).length;
function text(value,max,label,empty){
  if(typeof value!=='string')fail(`${label} must be text.`);
  const s=clean(value),n=size(s);
  if(!n)fail(empty||`${label} is empty.`);
  if(n>max)fail(`${label} is ${n} characters long. The most is ${max}.`);
  return s;
}
// A list of texts: at = "Slide 3", noun = "line" or "note", tooMany(n) = the message for a list that is too long.
function items(value,max,at,noun,limit,tooMany){
  if(value===undefined||value===null)return [];
  if(!Array.isArray(value))fail(`${at} ${noun}s must be a list.`);
  if(value.length>max)fail(tooMany(value.length));
  return value.map((v,i)=>text(v,limit,`${at}, ${noun} ${i+1}`));
}
// A cleaned copy of the deck (trimmed text, defaults filled in), or an Error whose message says what to change.
export function checkDeck(model){
  if(!model||typeof model!=='object'||Array.isArray(model))fail('There is no deck to write. Give it a title and a list of slides.');
  const title=text(model.title,L.titleChars,'The deck title','The deck needs a title.');
  let language='en-US';
  if(model.language!==undefined&&model.language!==null){
    const asked=typeof model.language==='string'?clean(model.language):null;   // blank means "use the default"
    if(asked===null||(asked&&(asked.length>35||!/^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8}){0,4}$/.test(asked))))fail('The language must be a code like en-US or fr-CA.');
    if(asked)language=asked;
  }
  const deck={title,language};
  if(model.subtitle!==undefined&&model.subtitle!==null&&(typeof model.subtitle!=='string'||clean(model.subtitle)))deck.subtitle=text(model.subtitle,L.subtitleChars,'The subtitle');
  if(!Array.isArray(model.slides)||!model.slides.length)fail('The deck needs at least one slide.');
  if(model.slides.length>L.slides)fail(`The deck has ${model.slides.length} slides. The most is ${L.slides}. Split it into two decks.`);
  const keys=new Set();
  deck.slides=model.slides.map((slide,i)=>{
    const at=`Slide ${i+1}`;
    if(!slide||typeof slide!=='object'||Array.isArray(slide))fail(`${at} is missing. Each slide needs a kind and a title.`);
    if(!SLIDE_KINDS.includes(slide.kind))fail(`${at} has no valid kind. Use ${SLIDE_KINDS.slice(0,-1).join(', ')} or ${SLIDE_KINDS.at(-1)}.`);
    const out={kind:slide.kind,title:text(slide.title,L.titleChars,`${at} title`,`${at} needs a title.`)};
    out.lines=items(slide.lines,L.lines,at,'line',L.lineChars,n=>`${at} has ${n} lines. The most is ${L.lines}. Move the rest into the speaker notes.`);
    out.notes=items(slide.notes,L.notes,at,'note',L.noteChars,n=>`${at} has ${n} notes. The most is ${L.notes}. Join some of them into one.`);
    if(slide.key!==undefined&&slide.key!==null){
      out.key=text(slide.key,L.keyChars,`${at} key`);
      if(keys.has(out.key))fail(`${at} has the key "${out.key}", which another slide already has. Give each slide its own key.`);
      keys.add(out.key);
    }
    return out;
  });
  return deck;
}

// The slides whose body is more than a room can read at 24 pt, one entry each: [{slide:1-based number, message}].
export function slideFit(model){
  const found=[];
  checkDeck(model).slides.forEach((s,i)=>{
    const n=i+1,chars=s.lines.reduce((sum,line)=>sum+size(line),0),longest=Math.max(0,...s.lines.map(size));
    if(s.lines.length>L.fitLines||chars>L.fitChars)found.push({slide:n,message:`Slide ${n} has too much text to read from the back of a room. Move some of it into the speaker notes.`});
    else if(longest>L.fitLineChars)found.push({slide:n,message:`Slide ${n} has a line that is too long to read from the back of a room. Shorten it, or move the detail into the speaker notes.`});
  });
  return found;
}

// A file name that is safe on Windows, macOS and Linux: letters, digits, spaces, _ and -, at most 65 characters with the .pptx.
export function pptxName(title){
  const room=65-'.pptx'.length;
  let stem=Array.from(String(title??'').normalize('NFC').replace(/[^\p{L}\p{M}\p{N} _-]/gu,'').replace(/\s+/g,' ').trim()).slice(0,room).join('').trim();
  if(!stem||/^(?:con|prn|aux|nul|com\d|lpt\d)$/i.test(stem))stem='Slides';
  return stem+'.pptx';
}

// ---- XML -------------------------------------------------------------------------------------------------------------------------
const NS_A='http://schemas.openxmlformats.org/drawingml/2006/main',NS_P='http://schemas.openxmlformats.org/presentationml/2006/main',NS_R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_REL='http://schemas.openxmlformats.org/package/2006/relationships',NS_TYPES='http://schemas.openxmlformats.org/package/2006/content-types';
const HEAD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',ROOT=`xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
const NAVY='172B4A',TEAL='216861',INK='18283A',WHITE='FFFFFF';
const GROUP='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const xfrm=([x,y,w,h])=>`<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
const rect=g=>xfrm(g)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
// One placeholder shape. ph = the <p:ph> attributes; g = its position (left out to inherit it); inner = the text body's content.
const holder=(id,name,ph,g,inner)=>`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${ph}/></p:nvPr></p:nvSpPr><p:spPr>${g?rect(g):''}</p:spPr><p:txBody>${inner}</p:txBody></p:sp>`;
const para=(t,lang)=>`<a:p><a:r><a:rPr lang="${esc(lang)}"/><a:t>${esc(t)}</a:t></a:r></a:p>`;
const prompts=l=>l.map((t,i)=>`<a:p>${i?`<a:pPr lvl="${i}"/>`:''}<a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p>`).join('');
const LEVELS=['Click to edit Master text styles','Second level','Third level','Fourth level','Fifth level'],TITLE_PROMPT=['Click to edit Master title style'];
const AUTOFIT='<a:bodyPr><a:normAutofit/></a:bodyPr>';
const font=major=>`<a:latin typeface="+m${major?'j':'n'}-lt"/><a:ea typeface="+m${major?'j':'n'}-ea"/><a:cs typeface="+m${major?'j':'n'}-cs"/>`;
const defRPr=(sz,clr,{bold=false,major=false}={})=>`<a:defRPr sz="${sz}"${bold?' b="1"':''} kern="1200"><a:solidFill><a:schemeClr val="${clr}"/></a:solidFill>${font(major)}</a:defRPr>`;
const LEVEL_ATTRS='defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"';
const bulletLevel=n=>`<a:lvl${n}pPr marL="${[342900,742950,1143000,1600200,2057400][n-1]}" indent="-342900" algn="l" ${LEVEL_ATTRS}><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="\u2022"/>${defRPr([2400,2200,2000,1800,1800][n-1],'tx1')}</a:lvl${n}pPr>`;
const TEXT_STYLES=`<p:txStyles><p:titleStyle><a:lvl1pPr algn="l" ${LEVEL_ATTRS}><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>${defRPr(3600,'tx2',{bold:true,major:true})}</a:lvl1pPr></p:titleStyle><p:bodyStyle>${[1,2,3,4,5].map(bulletLevel).join('')}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" ${LEVEL_ATTRS}>${defRPr(1800,'tx1')}</a:lvl1pPr></p:otherStyle></p:txStyles>`;
const CLR_MAP='<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
const INSETS='lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"';
const DECORATIVE='<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>';

function theme(name){
  const clr=(tag,v)=>`<a:${tag}><a:srgbClr val="${v}"/></a:${tag}>`,face=tag=>`<a:${tag}><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`;
  const fill='<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>',line=w=>`<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${fill}<a:prstDash val="solid"/></a:ln>`;
  return HEAD+`<a:theme xmlns:a="${NS_A}" name="${name}"><a:themeElements><a:clrScheme name="OpenStage">`
    +clr('dk1',INK)+clr('lt1',WHITE)+clr('dk2',NAVY)+clr('lt2','EEF2F6')+clr('accent1',TEAL)+clr('accent2',NAVY)+clr('accent3','8A5A00')+clr('accent4','5B3F8C')+clr('accent5','9B2C2C')+clr('accent6','4A5B6B')+clr('hlink','0B5CAD')+clr('folHlink','5B3F8C')
    +`</a:clrScheme><a:fontScheme name="OpenStage">${face('majorFont')}${face('minorFont')}</a:fontScheme><a:fmtScheme name="OpenStage"><a:fillStyleLst>${fill.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst><a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${fill.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

function slideMaster(){
  const bar=`<p:sp><p:nvSpPr><p:cNvPr id="4" name="Accent bar">${DECORATIVE}</p:cNvPr><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr>${rect([0,0,SIZE.cx,109728])}<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
  const title=holder(2,'Title Placeholder 1',' type="title"',[609600,320040,10972800,1188720],`<a:bodyPr vert="horz" ${INSETS} anchor="b"><a:normAutofit/></a:bodyPr><a:lstStyle/>${prompts(TITLE_PROMPT)}`);
  const body=holder(3,'Text Placeholder 2',' type="body" idx="1"',[609600,1645920,10972800,4572000],`<a:bodyPr vert="horz" ${INSETS}><a:normAutofit/></a:bodyPr><a:lstStyle/>${prompts(LEVELS)}`);
  return HEAD+`<p:sldMaster ${ROOT}><p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>${GROUP}${bar}${title}${body}</p:spTree></p:cSld>${CLR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>${TEXT_STYLES}</p:sldMaster>`;
}
const layout=(type,name,shapes)=>HEAD+`<p:sldLayout ${ROOT} type="${type}" preserve="1"><p:cSld name="${name}"><p:spTree>${GROUP}${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
const titleSlideLayout=()=>layout('title','Title Slide',
  holder(2,'Title 1',' type="title"',[914400,1371600,10363200,1463040],`<a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400"/></a:lvl1pPr></a:lstStyle>${prompts(TITLE_PROMPT)}`)
  +holder(3,'Subtitle 2',' type="subTitle" idx="1"',[914400,2926080,10363200,3291840],`${AUTOFIT}<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="l"><a:spcBef><a:spcPts val="600"/></a:spcBef><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>${prompts(['Click to edit Master subtitle style'])}`));
const contentLayout=()=>layout('obj','Title and Content',
  holder(2,'Title 1',' type="title"',null,`<a:bodyPr/><a:lstStyle/>${prompts(TITLE_PROMPT)}`)+holder(3,'Content Placeholder 2',' idx="1"',null,`<a:bodyPr/><a:lstStyle/>${prompts(LEVELS)}`));

function notesMaster(){
  const image=`<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr>${rect([685800,1143000,5486400,3086100])}<a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>`;
  const body=holder(3,'Notes Placeholder 2',' type="body" sz="quarter" idx="3"',[685800,4400550,5486400,3600450],`<a:bodyPr vert="horz" ${INSETS}/><a:lstStyle/>${prompts(['Click to edit Master text styles'])}`);
  const level=`<a:lvl1pPr marL="0" algn="l" ${LEVEL_ATTRS}>${defRPr(1400,'tx1')}</a:lvl1pPr>`;
  return HEAD+`<p:notesMaster ${ROOT}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GROUP}${image}${body}</p:spTree></p:cSld>${CLR_MAP}<p:notesStyle>${level}</p:notesStyle></p:notesMaster>`;
}

// The title is first in reading order, then the body (or the subtitle, on an opening slide). A slide with no text for the body has no body shape.
function slideXml(slide,n,lang,subtitle){
  const lines=slide.kind==='opening'&&!slide.lines.length&&subtitle?[subtitle]:slide.lines,opening=slide.kind==='opening';
  const title=holder(2,'Title 1',' type="title"',null,`${AUTOFIT}<a:lstStyle/>${para(slide.title,lang)}`);
  const body=lines.length?holder(3,opening?'Subtitle 2':'Content Placeholder 2',opening?' type="subTitle" idx="1"':' idx="1"',null,`${AUTOFIT}<a:lstStyle/>${lines.map(t=>para(t,lang)).join('')}`):'';
  return HEAD+`<p:sld ${ROOT}><p:cSld name="${esc(slide.key??`Slide ${n}`)}"><p:spTree>${GROUP}${title}${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}
function notesXml(notes,lang){
  const image='<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>';
  const text=notes.length?notes.map(t=>para(t,lang)).join(''):`<a:p><a:endParaRPr lang="${esc(lang)}"/></a:p>`;
  return HEAD+`<p:notes ${ROOT}><p:cSld><p:spTree>${GROUP}${image}${holder(3,'Notes Placeholder 2',' type="body" idx="1"',null,`<a:bodyPr/><a:lstStyle/>${text}`)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

// ---- the package -----------------------------------------------------------------------------------------------------------------
const RT='http://schemas.openxmlformats.org/officeDocument/2006/relationships/',OFFICE='application/vnd.openxmlformats-officedocument.';
const rels=list=>HEAD+`<Relationships xmlns="${NS_REL}">${list.map(([id,type,target])=>`<Relationship Id="${id}" Type="${type.includes(':')?type:RT+type}" Target="${target}"/>`).join('')}</Relationships>`;
function contentTypes(parts){
  return HEAD+`<Types xmlns="${NS_TYPES}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
    +parts.map(([name,type])=>`<Override PartName="/${name}" ContentType="${type}"/>`).join('')+'</Types>';
}
function coreProperties(deck,now){
  const stamp=now.toISOString().replace(/\.\d{3}Z$/,'Z');
  return HEAD+`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(deck.title)}</dc:title>${deck.subtitle?`<dc:subject>${esc(deck.subtitle)}</dc:subject>`:''}<dc:language>${esc(deck.language)}</dc:language><dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified></cp:coreProperties>`;
}
const appProperties=n=>HEAD+`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Application>OpenStage Guide</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>${n}</Slides><Notes>${n}</Notes><HiddenSlides>0</HiddenSlides></Properties>`;
function presentation(n,lang){
  const ids=Array.from({length:n},(_,i)=>`<p:sldId id="${256+i}" r:id="rId${7+i}"/>`).join('');
  return HEAD+`<p:presentation ${ROOT} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst><p:sldIdLst>${ids}</p:sldIdLst><p:sldSz cx="${SIZE.cx}" cy="${SIZE.cy}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="${esc(lang)}"/></a:defPPr><a:lvl1pPr marL="0" algn="l" ${LEVEL_ATTRS}>${defRPr(1800,'tx1')}</a:lvl1pPr></p:defaultTextStyle></p:presentation>`;
}

// The .pptx as bytes. The model is checked first (checkDeck), so a bad one throws before anything is written.
export async function pptxBytes(model,JSZip,{now=new Date()}={}){
  if(typeof JSZip!=='function')fail('The zip library is missing, so the slides cannot be packed.');
  const deck=checkDeck(model),n=deck.slides.length,zip=new JSZip();
  const put=(name,data)=>zip.file(name,data,{createFolders:false,date:now});
  const parts=[['ppt/presentation.xml',OFFICE+'presentationml.presentation.main+xml'],['ppt/slideMasters/slideMaster1.xml',OFFICE+'presentationml.slideMaster+xml'],
    ['ppt/slideLayouts/slideLayout1.xml',OFFICE+'presentationml.slideLayout+xml'],['ppt/slideLayouts/slideLayout2.xml',OFFICE+'presentationml.slideLayout+xml'],
    ['ppt/notesMasters/notesMaster1.xml',OFFICE+'presentationml.notesMaster+xml'],['ppt/theme/theme1.xml',OFFICE+'theme+xml'],['ppt/theme/theme2.xml',OFFICE+'theme+xml'],
    ['ppt/presProps.xml',OFFICE+'presentationml.presProps+xml'],['ppt/viewProps.xml',OFFICE+'presentationml.viewProps+xml'],['ppt/tableStyles.xml',OFFICE+'presentationml.tableStyles+xml'],
    ...deck.slides.flatMap((_,i)=>[[`ppt/slides/slide${i+1}.xml`,OFFICE+'presentationml.slide+xml'],[`ppt/notesSlides/notesSlide${i+1}.xml`,OFFICE+'presentationml.notesSlide+xml']]),
    ['docProps/core.xml','application/vnd.openxmlformats-package.core-properties+xml'],['docProps/app.xml',OFFICE+'extended-properties+xml']];
  put('[Content_Types].xml',contentTypes(parts));
  put('_rels/.rels',rels([['rId1','officeDocument','ppt/presentation.xml'],['rId2','http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties','docProps/core.xml'],['rId3','extended-properties','docProps/app.xml']]));
  put('docProps/core.xml',coreProperties(deck,now));put('docProps/app.xml',appProperties(n));
  put('ppt/presentation.xml',presentation(n,deck.language));
  put('ppt/_rels/presentation.xml.rels',rels([['rId1','slideMaster','slideMasters/slideMaster1.xml'],['rId2','notesMaster','notesMasters/notesMaster1.xml'],['rId3','presProps','presProps.xml'],['rId4','viewProps','viewProps.xml'],['rId5','theme','theme/theme1.xml'],['rId6','tableStyles','tableStyles.xml'],...deck.slides.map((_,i)=>[`rId${7+i}`,'slide',`slides/slide${i+1}.xml`])]));
  put('ppt/presProps.xml',HEAD+`<p:presentationPr ${ROOT}/>`);
  put('ppt/viewProps.xml',HEAD+`<p:viewPr ${ROOT}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`);
  put('ppt/tableStyles.xml',HEAD+`<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
  put('ppt/theme/theme1.xml',theme('OpenStage Guide'));put('ppt/theme/theme2.xml',theme('OpenStage Notes'));
  put('ppt/slideMasters/slideMaster1.xml',slideMaster());
  put('ppt/slideMasters/_rels/slideMaster1.xml.rels',rels([['rId1','slideLayout','../slideLayouts/slideLayout1.xml'],['rId2','slideLayout','../slideLayouts/slideLayout2.xml'],['rId3','theme','../theme/theme1.xml']]));
  put('ppt/slideLayouts/slideLayout1.xml',titleSlideLayout());put('ppt/slideLayouts/slideLayout2.xml',contentLayout());
  for(const k of [1,2])put(`ppt/slideLayouts/_rels/slideLayout${k}.xml.rels`,rels([['rId1','slideMaster','../slideMasters/slideMaster1.xml']]));
  put('ppt/notesMasters/notesMaster1.xml',notesMaster());
  put('ppt/notesMasters/_rels/notesMaster1.xml.rels',rels([['rId1','theme','../theme/theme2.xml']]));
  deck.slides.forEach((slide,i)=>{
    const k=i+1;
    put(`ppt/slides/slide${k}.xml`,slideXml(slide,k,deck.language,deck.subtitle));
    put(`ppt/slides/_rels/slide${k}.xml.rels`,rels([['rId1','slideLayout',`../slideLayouts/slideLayout${slide.kind==='opening'?1:2}.xml`],['rId2','notesSlide',`../notesSlides/notesSlide${k}.xml`]]));
    put(`ppt/notesSlides/notesSlide${k}.xml`,notesXml(slide.notes,deck.language));
    put(`ppt/notesSlides/_rels/notesSlide${k}.xml.rels`,rels([['rId1','notesMaster','../notesMasters/notesMaster1.xml'],['rId2','slide',`../slides/slide${k}.xml`]]));
  });
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
