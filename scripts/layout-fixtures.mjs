// Builds one sample lesson three ways so the layouts can be compared in real Word:
//   cue-labels : cue label with a decorative icon, the content, then shaded facilitator notes (proposed default)
//   narrative  : the same content as plain paragraphs, label on its own line
//   table      : a three-column table (Cue, Content, Notes); for comparison only, not a candidate default
// These are test fixtures for people. They are not part of the add-in build.
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';

export const VARIANTS=['cue-labels','narrative','table'];

export const LESSON={
  module:'Module 1: Recognize the conflict',
  lesson:'Lesson 1.1: Reframe the conversation',
  minutes:20,cumulative:20,
  blocks:[
    {cue:'Objective',icon:'Objective',text:'Identify two ways to reframe a disagreement.'},
    {cue:'Say',icon:'Lecture',text:'A disagreement is easier to resolve when both people describe the problem in the same neutral words. Give one example of a loaded phrase and its neutral version.',notes:'Pause after the example and ask the group for a second one. Allow 3 minutes.'},
    {cue:'Ask',icon:'Questions',text:'What makes a disagreement hard to reframe?',notes:'Listen for blame, assumptions and tone. Write two or three answers on the flip chart. Do not correct them yet.'},
    {cue:'Activity',icon:'Interaction',text:'In pairs, write two neutral questions that could clarify the disagreement in the case on page 4.',prompt:'Write your two neutral questions.',notes:'Walk the room. Allow 8 minutes, then debrief.'},
    {cue:'Key point',icon:'Key Point',text:'A neutral question describes the situation without assigning blame.'},
    {cue:'Transition',icon:'Transitions',text:'Next we look at how tone changes the way a question is heard.'}
  ]
};
const TIME_ICON='Time on Task';
export const ICON_NAMES=[TIME_ICON,...new Set(LESSON.blocks.map(b=>b.icon))];

// A valid 1x1 PNG, for tests that do not need real artwork.
export const TINY_PNG=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==','base64');

const NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const HEAD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

const NAVY='102037',INK='1B2430',SHADE='F1EBDD';
const HANG=2400; // twips; room for the icon and the longest cue label
export const ICON_EMU=457200; // 0.5 inch square; 0.3 inch was too small to tell the icons apart

function stylesXml(){
  const border=side=>`<w:${side} w:val="single" w:sz="4" w:space="0" w:color="8A94A3"/>`;
  return `${HEAD}<w:styles ${NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:color w:val="${NAVY}"/><w:sz w:val="52"/><w:szCs w:val="52"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="${NAVY}"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="${NAVY}"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="OSGTime"><w:name w:val="OSG Time"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="${NAVY}"/></w:pBdr><w:spacing w:before="120" w:after="200"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="OSGCue"><w:name w:val="OSG Cue"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepLines/><w:tabs><w:tab w:val="left" w:pos="${HANG}"/></w:tabs><w:spacing w:before="120"/><w:ind w:left="${HANG}" w:hanging="${HANG}"/></w:pPr></w:style>
<w:style w:type="character" w:customStyle="1" w:styleId="OSGCueLabel"><w:name w:val="OSG Cue Label"/><w:basedOn w:val="DefaultParagraphFont"/><w:qFormat/><w:rPr><w:b/><w:color w:val="${NAVY}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="OSGNotes"><w:name w:val="OSG Facilitator Notes"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="${NAVY}"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="${SHADE}"/><w:ind w:left="${HANG}"/></w:pPr><w:rPr><w:color w:val="${INK}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="OSGPrompt"><w:name w:val="OSG Learner Prompt"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="${HANG}"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="OSGCueLine"><w:name w:val="OSG Cue Line"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="40"/></w:pPr><w:rPr><w:b/><w:color w:val="${NAVY}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="OSGNotesPlain"><w:name w:val="OSG Facilitator Notes Plain"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="360"/></w:pPr></w:style>
<w:style w:type="table" w:customStyle="1" w:styleId="OSGLayoutTable"><w:name w:val="OSG Layout Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="59"/><w:tblPr><w:tblBorders>${border('top')}${border('left')}${border('bottom')}${border('right')}${border('insideH')}${border('insideV')}</w:tblBorders><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${NAVY}"/></w:tcPr></w:tblStylePr></w:style>
</w:styles>`;
}

const run=(text,rStyle)=>`<w:r>${rStyle?`<w:rPr><w:rStyle w:val="${rStyle}"/></w:rPr>`:''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const TAB='<w:r><w:tab/></w:r>';
const para=(style,inner)=>`<w:p>${style?`<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`:''}${inner}</w:p>`;

export async function buildFixture(variant,icons){
  if(!VARIANTS.includes(variant))throw new Error(`Unknown layout "${variant}". Use one of: ${VARIANTS.join(', ')}.`);
  const rid=name=>`rIdIcon${ICON_NAMES.indexOf(name)+1}`;
  let docPrId=0;
  // Decorative: the text label beside the icon carries the meaning, so a screen reader skips the picture.
  // The run is lowered 13.5 pt so the picture's middle sits on the text's middle, not on the baseline.
  const icon=name=>{const id=++docPrId,size=ICON_EMU;return `<w:r><w:rPr><w:position w:val="-27"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${size}" cy="${size}"/><wp:docPr id="${id}" name="Icon ${id}" descr=""><a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst></wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="icon${id}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid(name)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${size}" cy="${size}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;};

  const label=`Layout test: ${variant}`;
  const out=[
    para('Title',run(label)),
    para('Heading1',run(LESSON.module)),
    para('Heading2',run(LESSON.lesson)),
    para('OSGTime',icon(TIME_ICON)+run(` Time: ${LESSON.minutes} minutes. Cumulative time: ${LESSON.cumulative} minutes.`))
  ];

  if(variant==='cue-labels'){
    for(const b of LESSON.blocks){
      out.push(para('OSGCue',icon(b.icon)+run(` ${b.cue}`,'OSGCueLabel')+TAB+run(b.text)));
      if(b.prompt)out.push(para('OSGPrompt',run('Learner prompt: ','OSGCueLabel')+run(b.prompt)));
      if(b.notes)out.push(para('OSGNotes',run('Facilitator notes: ','OSGCueLabel')+run(b.notes)));
    }
  }else if(variant==='narrative'){
    for(const b of LESSON.blocks){
      out.push(para('OSGCueLine',icon(b.icon)+run(` ${b.cue}`)));
      out.push(para(null,run(b.text)));
      if(b.prompt)out.push(para(null,run('Learner prompt: ','OSGCueLabel')+run(b.prompt)));
      if(b.notes)out.push(para('OSGNotesPlain',run('Facilitator notes: ','OSGCueLabel')+run(b.notes)));
    }
  }else{
    const widths=[2500,4360,2500];
    const cell=(w,inner)=>`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
    const head=`<w:tr><w:trPr><w:tblHeader/></w:trPr>${['Cue','Content','Notes'].map((t,i)=>cell(widths[i],para(null,run(t)))).join('')}</w:tr>`;
    const rows=LESSON.blocks.map(b=>`<w:tr>${cell(widths[0],para(null,icon(b.icon)+run(` ${b.cue}`,'OSGCueLabel')))}${cell(widths[1],para(null,run(b.text))+(b.prompt?para(null,run('Learner prompt: ','OSGCueLabel')+run(b.prompt)):''))}${cell(widths[2],para(null,run(b.notes||'No notes.')))}</w:tr>`).join('');
    out.push(`<w:tbl><w:tblPr><w:tblStyle w:val="OSGLayoutTable"/><w:tblW w:w="${widths.reduce((a,b)=>a+b,0)}" w:type="dxa"/><w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/><w:tblCaption w:val="Lesson blocks"/><w:tblDescription w:val="One row per block: the cue, the content and any facilitator notes."/></w:tblPr><w:tblGrid>${widths.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${head}${rows}</w:tbl>`);
  }
  out.push(para(null,run('End of sample lesson.')));

  const document=`${HEAD}<w:document ${NS}><w:body>${out.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const rels=ICON_NAMES.map((n,i)=>`<Relationship Id="rIdIcon${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/icon${i+1}.png"/>`).join('');

  const date=new Date(Date.UTC(2026,9,30)),zip=new JSZip(),put=(name,data)=>zip.file(name,data,{date});
  put('[Content_Types].xml',`${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  put('_rels/.rels',`${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  put('word/document.xml',document);
  put('word/styles.xml',stylesXml());
  put('word/settings.xml',`${HEAD}<w:settings ${NS}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="en-US"/></w:settings>`);
  put('word/_rels/document.xml.rels',`${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>${rels}</Relationships>`);
  put('docProps/core.xml',`${HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>OpenStage Guide ${esc(label)}</dc:title><dc:creator>OpenStage Guide</dc:creator><dc:language>en-US</dc:language><dcterms:created xsi:type="dcterms:W3CDTF">2026-10-30T00:00:00Z</dcterms:created></cp:coreProperties>`);
  put('docProps/app.xml',`${HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>OpenStage Guide Stage 0 test fixture</Application></Properties>`);
  ICON_NAMES.forEach((n,i)=>put(`word/media/icon${i+1}.png`,icons(n)));
  return zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
}

async function main(){
  const args=process.argv.slice(2),opt=n=>{const i=args.indexOf(n);return i>=0?args[i+1]:null;};
  const outDir=opt('--out');
  if(!outDir){console.error('Usage: node scripts/layout-fixtures.mjs --out <folder> [--icons <folder of PNGs named like the icons>]');process.exit(2);}
  const iconDir=opt('--icons')||path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../public/assets/icons');
  const bytes=new Map();
  for(const n of ICON_NAMES)bytes.set(n,await fs.readFile(path.join(iconDir,`${n}.png`)));
  await fs.mkdir(outDir,{recursive:true});
  for(const v of VARIANTS){
    const file=path.join(outDir,`layout-test-${v}.docx`);
    await fs.writeFile(file,await buildFixture(v,n=>bytes.get(n)));
    console.log(file);
  }
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href)main().catch(e=>{console.error(e.message);process.exit(1);});
