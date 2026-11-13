// Word styles for the guide. Everything the guide needs to look right is a real style, so a hand edit,
// a copy and paste or a re-style in Word all behave. Names are what the pane looks for when it reads a guide.
export const STYLE_NAMES=Object.freeze({
  guideTitle:'OSG Guide Title',
  module:'OSG Module',
  lesson:'OSG Lesson',
  time:'OSG Time',
  cue:'OSG Cue',
  cueFacilitator:'OSG Cue Facilitator',
  body:'OSG Cue Body',
  notes:'OSG Facilitator Notes',
  assessment:'OSG Assessment',
  question:'OSG Question',
  option:'OSG Option',
  answerLine:'OSG Answer Line',
  answerKey:'OSG Answer Key',
  tableHead:'OSG Table Head',
  tableCell:'OSG Table Cell',
  ratingBox:'OSG Rating Box',
  itemNote:'OSG Item Note',
  overview:'OSG Overview',
  keyTableHead:'OSG Key Table Head',
  keyTableCell:'OSG Key Table Cell',
  formTitle:'OSG Form Title',
  formTitleFacilitator:'OSG Form Title Facilitator',
  formText:'OSG Form Text',
  formLine:'OSG Form Line',
  formField:'OSG Form Field',
  formBox:'OSG Form Box',
  formSource:'OSG Form Source',
  formPrint:'OSG Form Print',
  formCard:'OSG Form Card'
});
export const STYLE_IDS=Object.freeze({
  guideTitle:'OSGGuideTitle',module:'OSGModule',lesson:'OSGLesson',time:'OSGTime',cue:'OSGCue',cueFacilitator:'OSGCueFacilitator',
  cueLabel:'OSGCueLabel',body:'OSGCueBody',notes:'OSGNotes',
  assessment:'OSGAssessment',question:'OSGQuestion',option:'OSGOption',answerLine:'OSGAnswerLine',answerKey:'OSGAnswerKey',
  tableHead:'OSGTableHead',tableCell:'OSGTableCell',ratingBox:'OSGRatingBox',itemNote:'OSGItemNote',overview:'OSGOverview',keyTableHead:'OSGKeyTableHead',keyTableCell:'OSGKeyTableCell',formTitle:'OSGFormTitle',formTitleFacilitator:'OSGFormTitleFacilitator',
  formText:'OSGFormText',formLine:'OSGFormLine',formField:'OSGFormField',formBox:'OSGFormBox',formSource:'OSGFormSource',formPrint:'OSGFormPrint',formCard:'OSGFormCard'
});

// w14 is Word's 2010 extension namespace, which holds the check box content control.
export const NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const NAVY='102037',INK='1B2430',SHADE='F1EBDD',BAR='5B6B82';
const DEFAULT_LOOK={navy:NAVY,ink:INK,shade:SHADE,bar:BAR};
export const HANG=2400; // twips: room for a half-inch icon and the longest cue label

// A look is a handful of values that restyle everything the guide writes. With none given, this is the facilitator guide's
// own look. The learner workbook passes its look in (colours, heading font, heading sizes in half-points).
export function styleDefinitions(theme={}){
  const {navy:NAVY=DEFAULT_LOOK.navy,ink:INK=DEFAULT_LOOK.ink,shade:SHADE=DEFAULT_LOOK.shade,bar:BAR=DEFAULT_LOOK.bar}=theme;
  const sz={title:52,h1:36,h2:30,...(theme.sizes||{})};
  const AB=theme.answerBefore??240; // twips of space above a line to write on
  const HF=theme.headingFont?`<w:rFonts w:ascii="${theme.headingFont}" w:hAnsi="${theme.headingFont}" w:cs="${theme.headingFont}"/>`:'';
  return `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr>${HF}<w:b/><w:color w:val="${NAVY}"/><w:sz w:val="${sz.title}"/><w:szCs w:val="${sz.title}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.guideTitle}"><w:name w:val="${STYLE_NAMES.guideTitle}"/><w:basedOn w:val="Title"/><w:next w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr>${HF}<w:b/><w:color w:val="${NAVY}"/><w:sz w:val="${sz.h1}"/><w:szCs w:val="${sz.h1}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr>${HF}<w:b/><w:color w:val="${NAVY}"/><w:sz w:val="${sz.h2}"/><w:szCs w:val="${sz.h2}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.module}"><w:name w:val="${STYLE_NAMES.module}"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.lesson}"><w:name w:val="${STYLE_NAMES.lesson}"/><w:basedOn w:val="Heading2"/><w:next w:val="${STYLE_IDS.time}"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.time}"><w:name w:val="${STYLE_NAMES.time}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="${NAVY}"/></w:pBdr><w:spacing w:before="120" w:after="200"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.cue}"><w:name w:val="${STYLE_NAMES.cue}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepLines/><w:tabs><w:tab w:val="left" w:pos="${HANG}"/></w:tabs><w:spacing w:before="120"/><w:ind w:left="${HANG}" w:hanging="${HANG}"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.cueFacilitator}"><w:name w:val="${STYLE_NAMES.cueFacilitator}"/><w:basedOn w:val="${STYLE_IDS.cue}"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="12" w:space="6" w:color="${BAR}"/></w:pBdr></w:pPr></w:style>
<w:style w:type="character" w:customStyle="1" w:styleId="${STYLE_IDS.cueLabel}"><w:name w:val="OSG Cue Label"/><w:basedOn w:val="DefaultParagraphFont"/><w:qFormat/><w:rPr><w:b/><w:color w:val="${NAVY}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.body}"><w:name w:val="${STYLE_NAMES.body}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="${HANG}"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.notes}"><w:name w:val="${STYLE_NAMES.notes}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="${NAVY}"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="${SHADE}"/><w:ind w:left="${HANG}"/></w:pPr><w:rPr><w:color w:val="${INK}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formTitle}"><w:name w:val="${STYLE_NAMES.formTitle}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr>${HF}<w:b/><w:color w:val="${NAVY}"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formTitleFacilitator}"><w:name w:val="${STYLE_NAMES.formTitleFacilitator}"/><w:basedOn w:val="${STYLE_IDS.formTitle}"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="${SHADE}"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formText}"><w:name w:val="${STYLE_NAMES.formText}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="40" w:after="80"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formLine}"><w:name w:val="${STYLE_NAMES.formLine}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="${BAR}"/></w:pBdr><w:spacing w:before="${AB}" w:after="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formField}"><w:name w:val="${STYLE_NAMES.formField}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="${AB}" w:after="40"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formBox}"><w:name w:val="${STYLE_NAMES.formBox}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:top w:val="single" w:sz="6" w:space="4" w:color="${BAR}"/><w:left w:val="single" w:sz="6" w:space="4" w:color="${BAR}"/><w:bottom w:val="single" w:sz="6" w:space="4" w:color="${BAR}"/><w:right w:val="single" w:sz="6" w:space="4" w:color="${BAR}"/></w:pBdr><w:spacing w:before="120" w:after="120"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formSource}"><w:name w:val="${STYLE_NAMES.formSource}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="60" w:after="120"/></w:pPr><w:rPr><w:i/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formPrint}"><w:name w:val="${STYLE_NAMES.formPrint}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="${SHADE}"/><w:spacing w:before="0" w:after="120"/></w:pPr><w:rPr><w:i/><w:color w:val="${INK}"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.formCard}"><w:name w:val="${STYLE_NAMES.formCard}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepLines/><w:pBdr><w:top w:val="dashed" w:sz="6" w:space="8" w:color="${BAR}"/><w:left w:val="dashed" w:sz="6" w:space="8" w:color="${BAR}"/><w:bottom w:val="dashed" w:sz="6" w:space="8" w:color="${BAR}"/><w:right w:val="dashed" w:sz="6" w:space="8" w:color="${BAR}"/><w:between w:val="dashed" w:sz="6" w:space="8" w:color="${BAR}"/></w:pBdr><w:spacing w:before="160" w:after="160"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.assessment}"><w:name w:val="${STYLE_NAMES.assessment}"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.question}"><w:name w:val="${STYLE_NAMES.question}"/><w:basedOn w:val="Normal"/><w:next w:val="${STYLE_IDS.option}"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.option}"><w:name w:val="${STYLE_NAMES.option}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:after="60"/><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.answerLine}"><w:name w:val="${STYLE_NAMES.answerLine}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="${BAR}"/></w:pBdr><w:spacing w:before="${AB}" w:after="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.answerKey}"><w:name w:val="${STYLE_NAMES.answerKey}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepLines/><w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="${BAR}"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="${SHADE}"/><w:spacing w:before="0" w:after="0"/><w:ind w:left="360"/></w:pPr><w:rPr><w:color w:val="${INK}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.tableHead}"><w:name w:val="${STYLE_NAMES.tableHead}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:b/><w:color w:val="${INK}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.tableCell}"><w:name w:val="${STYLE_NAMES.tableCell}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.ratingBox}"><w:name w:val="${STYLE_NAMES.ratingBox}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.itemNote}"><w:name w:val="${STYLE_NAMES.itemNote}"/><w:basedOn w:val="Normal"/><w:next w:val="${STYLE_IDS.option}"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="60"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.overview}"><w:name w:val="${STYLE_NAMES.overview}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepLines/><w:pBdr><w:left w:val="single" w:sz="12" w:space="6" w:color="${BAR}"/></w:pBdr><w:spacing w:before="60" w:after="80"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.keyTableHead}"><w:name w:val="${STYLE_NAMES.keyTableHead}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:b/><w:color w:val="${INK}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_IDS.keyTableCell}"><w:name w:val="${STYLE_NAMES.keyTableCell}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:color w:val="${INK}"/></w:rPr></w:style>`;
}

// The lettered list that gives each multiple-choice option its a, b, c. Each item gets its own copy that starts again at a.
// `bullets` names the lists that are bulleted (an activity's steps) instead of lettered; they use the second definition below.
export function numberingPart(listCount,bullets=new Set()){
  const nums=Array.from({length:listCount},(_,i)=>bullets.has(i+1)
    ?`<w:num w:numId="${i+1}"><w:abstractNumId w:val="1"/></w:num>`
    :`<w:num w:numId="${i+1}"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('');
  const lettered=`<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`;
  const bulleted=`<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${HANG+360}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>`;
  return `<w:numbering ${NS}>${lettered}${bullets.size?bulleted:''}${nums}</w:numbering>`;
}

// The style definitions as a list, so a document can carry only the ones it uses.
export const styleList=theme=>(styleDefinitions(theme).match(/<w:style [\s\S]*?<\/w:style>/g)||[]).map(xml=>({id:/w:styleId="([^"]+)"/.exec(xml)[1],xml}));

// theme.bodyFont, theme.bodyHalfPoints, theme.line (240ths of a line) and theme.lang set the document defaults;
// `only` is a set of style ids to write, in place of all of them.
export function stylesPart(theme={},only=null){
  const font=theme.bodyFont||'Calibri',size=theme.bodyHalfPoints||24,line=theme.line||276,lang=theme.lang||'en-US';
  const defs=only?styleList(theme).filter(s=>only.has(s.id)).map(s=>s.xml).join('\n'):styleDefinitions(theme);
  return `<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/><w:lang w:val="${lang}" w:eastAsia="${lang}" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="${line}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>${defs}</w:styles>`;
}
