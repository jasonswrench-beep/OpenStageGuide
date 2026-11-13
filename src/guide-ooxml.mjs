// Builds the Word content the guide pane inserts. Pure functions: text in, OOXML out, so they can be tested
// without Word. Everything inserted is ordinary visible content in real styles.
import {cueById,iconUrl,TIME_ICON,ACTIVITY_PARTS} from './guide-cues.mjs';
import {NS,STYLE_IDS,stylesPart,numberingPart} from './guide-styles.mjs';

export const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const ICON_EMU=457200; // half an inch square
const TAB='<w:r><w:tab/></w:r>';

const run=(text,rStyle)=>`<w:r>${rStyle?`<w:rPr><w:rStyle w:val="${rStyle}"/></w:rPr>`:''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const para=(style,inner,direct='')=>`<w:p>${style||direct?`<w:pPr>${style?`<w:pStyle w:val="${style}"/>`:''}${direct}</w:pPr>`:''}${inner}</w:p>`;
const lines=text=>String(text??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
export {para,run,lines,TAB};

export const plural=(n,word)=>`${n} ${word}${n===1?'':'s'}`;
export const timeText=(minutes,cumulative)=>`Time: ${plural(minutes,'minute')}. Cumulative time: ${plural(cumulative,'minute')}.`;

// One context per insert: it hands out picture ids and remembers which icons the content uses.
// `lists` counts the lettered lists the content uses; each one is its own numbering instance so it starts again at a.
// `sdt` hands out the numbers of check box content controls.
// `pictures` holds the author's own pictures (not the cue icons) that this insert carries.
// `bullets` is the set of those list numbers that are bulleted lists (an activity's steps) and not lettered ones.
export function makeContext(){return {used:new Set(),docPr:1000,lists:0,bullets:new Set(),sdt:0,pictures:[]};}

const DECORATIVE='<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>';
// One of the author's pictures, in line with the text. pic: {base64,type,mime,size:{cx,cy},alt,decorative,name}. A picture needs
// alt text, or to be marked decorative, so a screen reader is never left with nothing.
export function pictureRun(pic,ctx){
  const alt=String(pic.alt??'').replace(/\s+/g,' ').trim();
  if(!pic.decorative&&!alt)throw new Error('Describe the picture in a few words for people who cannot see it, or tick "This picture is only decoration".');
  ctx.pictures=ctx.pictures||[];
  const n=ctx.pictures.length+1,id=++ctx.docPr,rid=`rIdPic${n}`,{cx,cy}=pic.size;
  ctx.pictures.push({rid,ext:pic.type,mime:pic.mime,base64:pic.base64});
  const name=esc(String(pic.name||`Picture ${n}`).slice(0,80));
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Picture ${id}" descr="${pic.decorative?'':esc(alt)}">${pic.decorative?DECORATIVE:''}</wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

// Decorative: the text label beside the picture carries the meaning, so a screen reader skips it.
// The run is lowered 13.5 pt so the picture's middle sits on the middle of the text line.
export function iconRun(icon,ctx){
  ctx.used.add(icon);
  const id=++ctx.docPr,rid=ridFor(icon);
  return `<w:r><w:rPr><w:position w:val="-27"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${ICON_EMU}" cy="${ICON_EMU}"/><wp:docPr id="${id}" name="Icon ${id}" descr=""><a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst></wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${icon}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${ICON_EMU}" cy="${ICON_EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
const ridFor=icon=>`rIdIcon_${icon.replace(/[^a-z0-9]/gi,'_')}`;

export function moduleHeading(number,title){
  return para(STYLE_IDS.module,run(`Module ${number}: ${title}`));
}
export function lessonHeading(moduleNumber,lessonNumber,title){
  return para(STYLE_IDS.lesson,run(`Lesson ${moduleNumber}.${lessonNumber}: ${title}`));
}
export function timeBlock({minutes,cumulative},ctx){
  return para(STYLE_IDS.time,iconRun(TIME_ICON,ctx)+run(` ${timeText(minutes,cumulative)}`));
}

// A cue block: the cue paragraph (icon, label, first line of text, optional minutes), any further lines,
// then facilitator notes as a shaded paragraph that always starts with the words "Facilitator notes:".
// An activity's labelled parts: each part's first line carries its label, the rest are further paragraphs in the same style.
function activityParts(parts,ctx){
  const out=[];
  if(!lines(parts.instructions).length)throw new Error('Type the instructions for the activity: what learners do, one step to a line.');
  for(const p of ACTIVITY_PARTS){
    const ls=lines(parts[p.id]);
    if(!ls.length)continue;
    const style=p.for==='facilitator'?STYLE_IDS.notes:STYLE_IDS.body;
    // A part of several lines that are steps or items (instructions, materials, preparation) is a bulleted list under its label; the
    // objective and the case study are prose, and a part of one line stays on its label's line (author, 2026-11-07: bullet the lists).
    if(ls.length>1&&p.id!=='objective'&&p.id!=='caseStudy'){
      out.push(para(style,run(`${p.label}:`,STYLE_IDS.cueLabel),'<w:keepNext/>'));
      const list=++ctx.lists;
      (ctx.bullets||(ctx.bullets=new Set())).add(list);
      for(const l of ls)out.push(para(style,run(l),`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${list}"/></w:numPr>`));
      continue;
    }
    ls.forEach((l,i)=>out.push(para(style,(i===0?run(`${p.label}:`,STYLE_IDS.cueLabel)+run(' '):'')+run(l))));
  }
  return out;
}

export function cueBlock({cue,text,notes,slideText,minutes,audience,parts,picture},ctx){
  const c=typeof cue==='string'?cueById(cue):cue;
  if(!c)throw new Error('Choose a cue first.');
  if(c.id==='picture'&&!picture)throw new Error('Choose a picture file for a Picture block.');
  const body=lines(text),first=body.shift()||'';
  const facilitatorOnly=(audience||c.audience)==='facilitator';
  const when=minutes?` (${minutes} min)`:'';
  const out=[para(facilitatorOnly?STYLE_IDS.cueFacilitator:STYLE_IDS.cue,iconRun(c.icon,ctx)+run(` ${c.label}`,STYLE_IDS.cueLabel)+TAB+run(first+when))];
  // The picture sits in its own paragraph straight under the cue line, in the block's body style, so it belongs to the block.
  if(picture)out.push(para(STYLE_IDS.body,pictureRun(picture,ctx)));
  if(parts){
    if(c.id!=='activity')throw new Error('Only an Activity can be written as labelled parts.');
    if(body.length)throw new Error('The title of an activity written as labelled parts is one line. Put the rest in the parts below it.');
    out.push(...activityParts(parts,ctx));
  }
  for(const l of body)out.push(para(STYLE_IDS.body,run(l)));
  // A slide's own words, under its picture, in the shaded facilitator style like its speaker notes (they never reach the workbook).
  lines(slideText).forEach((l,i)=>out.push(para(STYLE_IDS.notes,(i===0?run('Slide text: ',STYLE_IDS.cueLabel):'')+run(l))));
  const n=lines(notes);
  n.forEach((l,i)=>out.push(para(STYLE_IDS.notes,(i===0?run('Facilitator notes: ',STYLE_IDS.cueLabel):'')+run(l))));
  return out.join('');
}

const prompt=text=>para(null,run(text));
export const heading=(level,text,pageBreak=false)=>para(`Heading${level}`,run(text),pageBreak?'<w:pageBreakBefore/>':'');
// The placeholder under "Training at a Glance". The Glance button replaces it, and only it, when it builds the table.
export const GLANCE_PROMPT='Give a short overview of the modules, the time each takes and what happens in each.';

// The standard pages, in the order of the original specification: the title, Acknowledgements, Contents, then Getting Started with
// About This Guide, (Graphic Cues, built by its own button), The Program in Perspective, Program Preparation and Training at a Glance.
// The contents is a real table-of-contents field; the pane refreshes it after inserting. The Appendix is optional and is added at the end.
export const ACK_HEADING='Acknowledgements';
export const ACK_PROMPT='Thank the people and organizations that helped with this guide.';
export const PERSPECTIVE_HEADING='The Program in Perspective';
export const PERSPECTIVE_PROMPT='Say where this program fits: who it is for, what need it meets and how it connects to other training.';
export const APPENDIX_HEADING='Appendix';
export const APPENDIX_PROMPT='Add reference material here: job aids, forms, readings and anything else the facilitator may want to hand.';
export const acknowledgementsBody=()=>heading(1,ACK_HEADING,true)+prompt(ACK_PROMPT);
export const perspectiveBody=()=>heading(2,PERSPECTIVE_HEADING)+prompt(PERSPECTIVE_PROMPT);
export const appendixBody=()=>heading(1,APPENDIX_HEADING,true)+prompt(APPENDIX_PROMPT);
export function standardPages(title){
  return [
    para(STYLE_IDS.guideTitle,run(title)),
    para(null,run('Facilitator guide')),
    acknowledgementsBody(),
    heading(1,'Contents',true),
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run('Update this table to list the modules and lessons.')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    heading(1,'Getting Started',true),
    heading(2,'About This Guide'),
    prompt('Say what this guide is for and who should use it.'),
    perspectiveBody(),
    heading(2,'Program Preparation'),
    prompt('List what the facilitator needs before the session: pre-work, materials, room or technology set-up.'),
    heading(2,'Training at a Glance'),
    prompt(GLANCE_PROMPT)
  ].join('');
}

export function startGuide({title,standard=true},ctx){
  const name=String(title||'').trim();
  if(!name)throw new Error('Type a title for the guide first.');
  return standard?standardPages(name):para(STYLE_IDS.guideTitle,run(name));
}

// A short finished example, so a first-time author can see the result. It is ordinary content.
export function sampleGuide(ctx){
  const m=moduleHeading(1,'Recognize the conflict'),l=lessonHeading(1,1,'Reframe the conversation');
  return [
    standardPages('Sample guide: Handling workplace conflict'),
    m,l,timeBlock({minutes:20,cumulative:0},ctx),
    cueBlock({cue:'objective',text:'Identify two ways to reframe a disagreement.',minutes:2},ctx),
    cueBlock({cue:'say',text:'A disagreement is easier to resolve when both people describe the problem in the same neutral words. Give one example of a loaded phrase and its neutral version.',notes:'Pause after the example and ask the group for a second one.',minutes:5},ctx),
    cueBlock({cue:'ask',text:'What makes a disagreement hard to reframe?',notes:'Listen for blame, assumptions and tone. Write two or three answers on the flip chart. Do not correct them yet.',minutes:5},ctx),
    cueBlock({cue:'activity',text:'In pairs, write two neutral questions that could clarify the disagreement in the case on page 4.',notes:'Walk the room. Allow 8 minutes, then debrief.',minutes:8},ctx),
    cueBlock({cue:'key-point',text:'A neutral question describes the situation without assigning blame.'},ctx),
    cueBlock({cue:'transition',text:'Next we look at how tone changes the way a question is heard.'},ctx)
  ].join('');
}

// Word's insertOoxml drops the formatting of the last paragraph it inserts (its style is not even imported) and can
// join that paragraph to the one that follows. That was seen in real Word: every time block, Transition, Say block and
// set of facilitator notes that came last in an insert lost its style, and a time block ran into the Transition after it.
// So every package ends with one empty paragraph that is allowed to be lost, and the real content keeps its styles.
export const SPARE_PARAGRAPH='<w:p/>';

// A flat Office package for Word's insertOoxml. iconData maps an icon name to its PNG bytes as base64.
// theme (optional) is the active look; styles this insert has to define are written from it (see guide-looks.mjs).
export function buildPackage(body,ctx,iconData,{spare=true,theme}={}){
  const icons=[...ctx.used],pics=ctx.pictures||[];
  for(const i of icons)if(!iconData[i])throw new Error(`The picture for the ${i} cue could not be loaded. Check the connection and try again.`);
  const lists=ctx.lists||0;
  const picName=(p,i)=>`pic${i+1}.${p.ext}`;
  const numbering=lists?`<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"><pkg:xmlData>${numberingPart(lists,ctx.bullets)}</pkg:xmlData></pkg:part>`:'';
  const rels=`<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${lists?'<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>':''}${icons.map(i=>`<Relationship Id="${ridFor(i)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${i}.png"/>`).join('')}${pics.map((p,i)=>`<Relationship Id="${p.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${picName(p,i)}"/>`).join('')}`;
  const media=icons.map(i=>`<pkg:part pkg:name="/word/media/${i}.png" pkg:contentType="image/png" pkg:compression="store"><pkg:binaryData>${iconData[i]}</pkg:binaryData></pkg:part>`).join('')+pics.map((p,i)=>`<pkg:part pkg:name="/word/media/${picName(p,i)}" pkg:contentType="${p.mime}" pkg:compression="store"><pkg:binaryData>${p.base64}</pkg:binaryData></pkg:part>`).join('');
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document ${NS}><w:body>${body}${spare?SPARE_PARAGRAPH:''}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData>${stylesPart(theme||{})}</pkg:xmlData></pkg:part>${numbering}${media}</pkg:package>`;
}

export {iconUrl};
