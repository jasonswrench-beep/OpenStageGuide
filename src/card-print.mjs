// Card print sheet: a deck of two-sided cards laid out as an editable Word file to print and cut apart, plus a plain-text view and a check
// for text that will not fit on a card. Nothing is stored; the deck is a plain model handed in each time (checkDeck says what it must hold).
//
// The sheet. Every page is one fixed-layout table of cards and nothing else: the title and the printing instructions live in the header of the
// first page only (w:titlePg), so no page is spent on them. Each card is two rows, the card itself and a thin strip under it with its label
// ("Card 4 of 12", and "Card 4 of 12 · Back" on a back); both rows have an exact height, so every card is the same size and a row never splits
// across pages. The cell borders are dashed light grey: they are the lines to cut along. Cards are real text (14 pt for 1 to 4 to a page, 12 pt
// for 6 and 8), the first line of a front in bold. A page ends where the next one says w:pageBreakBefore, which, unlike a break in a paragraph of
// its own, never leaves a stray line at the top of the next page and so never moves the grid.
//
// Page order: every front page first, then every back page (only when some card has a back). That order suits printing in two passes; a printer
// that prints both sides itself wants front and back pages alternating, which is a change to pageSequence() alone. The header says which pages
// are fronts and which are backs.
//
// Duplex order. The back of each card must sit behind its front. Flip on the LONG edge (turn the sheet like the page of a book): a card at column c
// on the front prints at column cols-1-c on the back, in the same row, so each row is mirrored left to right (with one column nothing mirrors).
// Flip on the SHORT edge (turn it like a calendar): a card at row r on the front prints at row rows-1-r on the back, in the same column, so
// the rows are mirrored top to bottom (with one row nothing mirrors). Pages are portrait; the printer's own flip setting does the rest.
// The grid sits in the middle of the page both ways (equal side margins, and equal space above and below it), so the two sides line up
// whichever edge is flipped.
//
// Geometry. Side margins are 0.5 inch. The space above and below the grid is at least 1 inch: the header needs it on the first page and
// the crop marks need 0.25 inch of it. The bottom margin is a little smaller than the top one, only to leave a slack for rounding so the
// last row is never pushed onto a page of its own; the grid itself is centred. One page holds ANCHOR + the grid + ANCHOR twips, ANCHOR being a
// one-point paragraph before the grid (it carries the crop marks and the page break) and one after the last grid (a document cannot end in a table).
//
// Crop marks (cropMarks: true). Short dark lines in the margin, 0.25 inch long, drawn as DrawingML straight-line shapes (wps:wsp, preset
// geometry "line", behind the text) anchored to the page in the one-point paragraph before each grid. They sit at the extension of every card
// edge: above and below each column line, left and right of each row line. They are wrapped in mc:AlternateContent (Requires="wps"), the way
// Word writes them, so a program that does not know the shapes leaves them out and the page is still right.
//
// Overflow. A card cannot grow (its row has an exact height), so text that does not fit is cut off when printed. cardFit counts the text of a side,
// plus 40 for each line break, plus the letters of its own label, and compares it with CARD_CAPACITY. The label has a strip of its own, so
// counting it only makes the warning a little early. The capacities are what the layout below holds at the sizes above, in the narrowest
// (A4) width and the shortest (Letter) height, with a third of the room kept back for words that wrap early and wide letters:
//
//   cards to a page   1     2     4     6     8
//   characters        1700  800   390   320   240
import {para,run,esc} from './guide-ooxml.mjs';
import {NS,stylesPart} from './guide-styles.mjs';

export const CARD_LIMITS=Object.freeze({cards:120,lines:8,line:300});
export const PAPERS=Object.freeze(['Letter','A4']);
export const PER_PAGE=Object.freeze([1,2,4,6,8]);
export const CARD_CAPACITY=Object.freeze({1:1700,2:800,4:390,6:320,8:240});

const TYPES=['activity','toolkit','facilitator'];
const FLIPS=['long','short'];
// 'alternate': each front is followed by its back, for a printer that prints both sides itself. 'separate': every front, then every back, for printing in two passes.
export const SEQUENCES=Object.freeze(['alternate','separate']);
const TITLE_MAX=200,HEAD_TITLE_MAX=56,BREAK_COST=40;
const LANG=/^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){0,3}$/;
const PAGE={Letter:{w:12240,h:15840},A4:{w:11906,h:16838}};
const GRID={1:[1,1],2:[1,2],4:[2,2],6:[2,3],8:[2,4]}; // [columns, rows] on a portrait page
// Twips. SIDE: left and right margin. EDGE: least space above and below the grid. MARK: crop mark length (0.25 inch). LABEL_H: the strip under a card.
// ANCHOR: the one-point paragraph before a grid. SLACK: how much smaller the bottom margin is than the top one.
const SIDE=720,EDGE=1440,MARK=360,LABEL_H=300,ANCHOR=20,SLACK=200,EMU=635;
const CELL_MAR={top:140,left:200,bottom:100,right:200};
const GREY='A6A6A6';

// ---- the deck
// Control characters, line and paragraph separators, and half a surrogate pair have no place in a Word file; each becomes a space.
const BAD=/[\u0000-\u001F\u007F-\u009F\u2028\u2029\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
const tidy=t=>String(t).replace(BAD,' ').replace(/\s+/g,' ').trim();

function checkSide(raw,n,name,required){
  if(raw==null||(Array.isArray(raw)&&!raw.length&&!required))return null;
  if(!Array.isArray(raw))throw new Error(`The ${name} of card ${n} must be a list of lines.`);
  if(!raw.length)throw new Error(`Card ${n} needs at least one line on its ${name}.`);
  if(raw.length>CARD_LIMITS.lines)throw new Error(`The ${name} of card ${n} has ${raw.length} lines. A side holds up to ${CARD_LIMITS.lines} lines.`);
  return raw.map((l,i)=>{
    const t=typeof l==='string'?tidy(l):'';
    if(!t)throw new Error(`Line ${i+1} on the ${name} of card ${n} is empty.`);
    if(t.length>CARD_LIMITS.line)throw new Error(`Line ${i+1} on the ${name} of card ${n} has ${t.length} letters. A line holds up to ${CARD_LIMITS.line} letters.`);
    return t;
  });
}

// The deck as the sheet will use it: a cleaned copy (text tidied, a blank back made null), or a plain-English error.
export function checkDeck(model){
  if(!model||typeof model!=='object')throw new Error('There is no deck to print.');
  const title=typeof model.title==='string'?tidy(model.title):'';
  if(!title||title.length>TITLE_MAX)throw new Error(`Give the deck a title of 1 to ${TITLE_MAX} letters.`);
  const language=model.language==null?'en-US':model.language;
  if(typeof language!=='string'||!LANG.test(language))throw new Error('The language must be a code such as en-US.');
  if(!TYPES.includes(model.type))throw new Error('A deck is for an activity, a toolkit or a facilitator.');
  if(!PAPERS.includes(model.paper))throw new Error('Choose Letter or A4 paper.');
  if(!PER_PAGE.includes(model.perPage))throw new Error('Cards go 1, 2, 4, 6 or 8 to a page.');
  if(model.cropMarks!=null&&typeof model.cropMarks!=='boolean')throw new Error('Crop marks are either on or off.');
  if(!FLIPS.includes(model.flip))throw new Error('Choose to flip on the long edge or the short edge.');
  const sequence=model.sequence==null?'alternate':model.sequence;
  if(!SEQUENCES.includes(sequence))throw new Error('Choose to print each front with its back, or all the fronts and then all the backs.');
  if(!Array.isArray(model.cards)||!model.cards.length)throw new Error('Add at least one card.');
  if(model.cards.length>CARD_LIMITS.cards)throw new Error(`A deck holds up to ${CARD_LIMITS.cards} cards.`);
  const cards=model.cards.map((c,i)=>{
    if(!c||typeof c!=='object')throw new Error(`Card ${i+1} needs a front.`);
    return {front:checkSide(c.front,i+1,'front',true),back:checkSide(c.back,i+1,'back',false)};
  });
  return {title,language,type:model.type,paper:model.paper,perPage:model.perPage,cropMarks:model.cropMarks===true,flip:model.flip,sequence,cards};
}

const labelOf=(n,total,side)=>`Card ${n} of ${total}${side==='back'?' · Back':''}`;
const countOf=(lines,label)=>lines.reduce((a,l)=>a+l.length,0)+BREAK_COST*(lines.length-1)+label.length;

// The sides whose text will not fit the card at this print size, each with a sentence the author can act on.
export function cardFit(model){
  const m=checkDeck(model),limit=CARD_CAPACITY[m.perPage],out=[];
  m.cards.forEach((c,i)=>{
    for(const [side,lines] of [['front',c.front],['back',c.back]]){
      if(!lines)continue;
      const chars=countOf(lines,labelOf(i+1,m.cards.length,side));
      if(chars<=limit)continue;
      const lead=`Card ${i+1} (${side}) has ${chars} characters, which is more than fits`;
      out.push({card:i+1,side,chars,limit,message:m.perPage===1?`${lead} on a page by itself. Shorten it.`:`${lead} when ${m.perPage} cards go on a page. Shorten it, or print fewer cards to a page.`});
    }
  });
  return out;
}

// ---- the plain text view
export function deckText(model){
  const m=checkDeck(model);
  const side=(name,lines)=>lines?`${name}: ${lines.join('\n  ')}`:`${name}: (blank)`;
  return `${m.cards.map((c,i)=>`Card ${i+1}\n${side('Front',c.front)}\n${side('Back',c.back)}`).join('\n\n')}\n`;
}

export function deckFileName(title,paper){
  if(!PAPERS.includes(paper))throw new Error('Choose Letter or A4 paper.');
  const tail=` - print sheet (${paper}).docx`;
  const base=String(title??'').replace(/[^\p{L}\p{N} _()-]/gu,'').replace(/\s+/g,' ').trim().slice(0,80-tail.length).trim();
  return `${base||'Cards'}${tail}`;
}

// ---- the page
function layout(m){
  const {w,h}=PAGE[m.paper],[cols,rows]=GRID[m.perPage];
  const cellW=Math.floor((w-2*SIDE)/cols),tableW=cellW*cols;
  const pitch=Math.floor((h-2*EDGE)/rows),gridH=pitch*rows;       // pitch: one row of cards, with its strip
  const gridTop=Math.floor((h-gridH)/2),top=gridTop-ANCHOR;
  return {w,h,cols,rows,cellW,tableW,pitch,gridH,gridTop,left:Math.floor((w-tableW)/2),top,bottom:top-SLACK};
}

// What each slot of a page shows (null for an empty slot). A back slot shows the card whose front sits where this slot mirrors to; mirroring is
// its own inverse, so the same step finds the slot from the card and the card from the slot.
function slotOf(m,L,side,page,s){
  const r=Math.floor(s/L.cols),c=s%L.cols;
  let fr=r,fc=c;
  if(side==='back'){if(m.flip==='long')fc=L.cols-1-c;else fr=L.rows-1-r;}
  const i=page*m.perPage+fr*L.cols+fc,card=m.cards[i];
  const lines=card&&(side==='front'?card.front:card.back);
  return lines?{lines,head:side==='front',label:labelOf(i+1,m.cards.length,side)}:null;
}

const nil=side=>`<w:tcBorders><w:${side} w:val="nil"/></w:tcBorders>`;
const tc=(w,pr,inner)=>`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${pr}</w:tcPr>${inner}</w:tc>`;
const tr=(h,cells)=>`<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${h}" w:hRule="exact"/></w:trPr>${cells.join('')}</w:tr>`;
const cardCell=(w,s)=>tc(w,nil('bottom'),s?s.lines.map((l,i)=>para(s.head&&i===0?'OSGCardHead':'OSGCardText',run(l))).join(''):para('OSGCardText',''));
const labelCell=(w,s)=>tc(w,`${nil('top')}<w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/></w:tcMar>`,para('OSGCardLabel',s?run(s.label):''));
const BORDERS=['top','left','bottom','right','insideH','insideV'].map(s=>`<w:${s} w:val="dashed" w:sz="4" w:space="0" w:color="${GREY}"/>`).join('');
const NOUN={activity:'activity cards',toolkit:'toolkit cards',facilitator:'facilitator cards'};

function gridXml(m,L,side,page,pages){
  const rows=[];
  for(let r=0;r<L.rows;r++){
    const slots=Array.from({length:L.cols},(_,c)=>slotOf(m,L,side,page,r*L.cols+c));
    rows.push(tr(L.pitch-LABEL_H,slots.map(s=>cardCell(L.cellW,s))),tr(LABEL_H,slots.map(s=>labelCell(L.cellW,s))));
  }
  const caption=`${side==='front'?'Fronts':'Backs'} of the cards, page ${page+1} of ${pages}`;
  const about=`This table only lays out the ${side==='front'?'fronts':'backs'} of the ${NOUN[m.type]} to print and cut apart. It is not a table of data.`;
  const mar=`<w:tblCellMar><w:top w:w="${CELL_MAR.top}" w:type="dxa"/><w:left w:w="${CELL_MAR.left}" w:type="dxa"/><w:bottom w:w="${CELL_MAR.bottom}" w:type="dxa"/><w:right w:w="${CELL_MAR.right}" w:type="dxa"/></w:tblCellMar>`;
  return `<w:tbl><w:tblPr><w:tblW w:w="${L.tableW}" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>${BORDERS}</w:tblBorders><w:tblLayout w:type="fixed"/>${mar}<w:tblCaption w:val="${esc(caption)}"/><w:tblDescription w:val="${esc(about)}"/></w:tblPr><w:tblGrid>${Array.from({length:L.cols},()=>`<w:gridCol w:w="${L.cellW}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
}

// ---- crop marks
function markList(L){
  const out=[],end=L.gridTop+L.gridH;
  for(let c=0;c<=L.cols;c++)out.push({x:L.left+c*L.cellW,y:L.gridTop-MARK,vertical:true},{x:L.left+c*L.cellW,y:end,vertical:true});
  for(let r=0;r<=L.rows;r++)out.push({x:L.left-MARK,y:L.gridTop+r*L.pitch,vertical:false},{x:L.left+L.tableW,y:L.gridTop+r*L.pitch,vertical:false});
  return out;
}
function markXml(id,{x,y,vertical}){
  const cx=vertical?0:MARK*EMU,cy=vertical?MARK*EMU:0;
  return `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${x*EMU}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${y*EMU}</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="Crop mark ${id}" descr="A mark in the margin that shows where to cut. It is not part of a card."/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;
}

// ---- the parts
// Each front followed by its back (the default: a printer that prints both sides itself takes the pages in this order), or every front and then every back (to print in two
// passes), when any card has a back.
function pageSequence(pages,hasBacks,sequence){
  const fronts=Array.from({length:pages},(_,p)=>['front',p]);
  if(!hasBacks)return fronts;
  return sequence==='separate'?fronts.concat(fronts.map(([,p])=>['back',p])):fronts.flatMap(([,p])=>[['front',p],['back',p]]);
}

const style=(id,name,based,pPr,rPr,custom=true)=>`<w:style w:type="paragraph"${custom?' w:customStyle="1"':''} w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="${based}"/>${custom?'<w:qFormat/>':'<w:uiPriority w:val="99"/><w:unhideWhenUsed/>'}${pPr?`<w:pPr>${pPr}</w:pPr>`:''}${rPr?`<w:rPr>${rPr}</w:rPr>`:''}</w:style>`;
function stylesXml(m,language){
  const hp=m.perPage<=4?28:24,size=n=>`<w:sz w:val="${n}"/><w:szCs w:val="${n}"/>`;
  const flat='<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
  const extra=[
    style('OSGCardText','OSG Card Text','Normal','<w:spacing w:before="0" w:after="80" w:line="240" w:lineRule="auto"/>',size(hp)),
    style('OSGCardHead','OSG Card Head','OSGCardText','','<w:b/><w:bCs/>'),
    style('OSGCardLabel','OSG Card Label','Normal',flat,`<w:color w:val="6B6B6B"/>${size(16)}`),
    style('OSGCardAnchor','OSG Card Page Anchor','Normal','<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>',size(2)),
    style('Header','header','Normal',flat,`<w:color w:val="404040"/>${size(16)}`,false)
  ].join('');
  return stylesPart({lang:language,bodyHalfPoints:hp},new Set(['Normal','DefaultParagraphFont','Title'])).replace('</w:styles>',`${extra}</w:styles>`);
}

function headerXml(m,pages,hasBacks){
  const shown=m.title.length>HEAD_TITLE_MAX?`${m.title.slice(0,HEAD_TITLE_MAX-1).trimEnd()}…`:m.title;
  const small='<w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>';
  const order=m.sequence==='separate'
    ?(pages===1?'Page 1 is the front. Page 2 is the back.':`Pages 1 to ${pages} are the fronts. Pages ${pages+1} to ${2*pages} are the backs.`)
    :'Each front page is followed by its back.';
  const lines=hasBacks
    ?[`Print on ${m.paper} paper, two-sided, flip on the ${m.flip} edge. Cut along the dashed lines.`,order]
    :[`Print on ${m.paper} paper. Cut along the dashed lines.`];
  return `<w:hdr ${NS}>${para('Title',`<w:r>${small}<w:t xml:space="preserve">${esc(`Print sheet: ${shown}`)}</w:t></w:r>`,'<w:spacing w:before="0" w:after="0"/>')}${lines.map(l=>para('Header',run(l))).join('')}</w:hdr>`;
}

// The .docx parts for a deck, as {name,text}, ready to be zipped (the same shape as agendaFiles).
export function cardSheetFiles(model,{language}={}){
  const m=checkDeck(model),lang=language==null?m.language:language;
  if(typeof lang!=='string'||!LANG.test(lang))throw new Error('The language must be a code such as en-US.');
  const L=layout(m),pages=Math.ceil(m.cards.length/m.perPage),hasBacks=m.cards.some(c=>c.back);
  let id=0;
  const body=pageSequence(pages,hasBacks,m.sequence).map(([side,p],k)=>{
    const marks=m.cropMarks?markList(L).map(mk=>markXml(++id,mk)).join(''):'';
    return para('OSGCardAnchor',marks,k?'<w:pageBreakBefore/>':'')+gridXml(m,L,side,p,pages);
  }).join('');
  const sect=`<w:sectPr><w:headerReference w:type="first" r:id="rIdHeader1"/><w:pgSz w:w="${L.w}" w:h="${L.h}"/><w:pgMar w:top="${L.top}" w:right="${SIDE}" w:bottom="${L.bottom}" w:left="${SIDE}" w:header="320" w:footer="320" w:gutter="0"/><w:titlePg/></w:sectPr>`;
  const x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const docNs=`${NS} xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
  const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return [
    {name:'[Content_Types].xml',text:`${x}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`},
    {name:'_rels/.rels',text:`${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`},
    {name:'word/document.xml',text:`${x}<w:document ${docNs}><w:body>${body}${para('OSGCardAnchor','')}${sect}</w:body></w:document>`},
    {name:'word/_rels/document.xml.rels',text:`${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/></Relationships>`},
    {name:'word/styles.xml',text:`${x}${stylesXml(m,lang)}`},
    {name:'word/header1.xml',text:`${x}${headerXml(m,pages,hasBacks)}`},
    {name:'docProps/core.xml',text:`${x}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(`Print sheet: ${m.title}`)}</dc:title><dc:language>${esc(lang)}</dc:language></cp:coreProperties>`}
  ];
}
