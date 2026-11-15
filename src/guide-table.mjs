// Changing the cells of a table in the guide, without disturbing anything else in it. Pure functions: Word-format XML in, Word-format XML out.
// A cell the editor can change holds one paragraph of plain words (or nothing); a cell that holds a check box, a picture, several paragraphs or
// anything else is shown and left exactly as it is. The words of a cell are changed by the same code that changes the words of a paragraph
// (guide-edit.mjs), so a cell is held to the same rules as any other paragraph.
import {readParagraph,setWords} from './guide-edit.mjs';
import {readPicture,setPicture} from './guide-picture.mjs';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const TABLE_LIMITS=Object.freeze({rows:200});
const kids=el=>Array.from(el.childNodes).filter(n=>n.nodeType===1);
const named=(el,name)=>kids(el).find(n=>n.localName===name&&n.namespaceURI===W);
const attr=(el,name)=>el?.getAttributeNS(W,name)||'';
const CELL_STYLE='OSG Table Cell';

function open(xml,io){
  const root=io.parse(xml).documentElement;
  if(!root||root.namespaceURI!==W||root.localName!=='tbl')throw new Error('That is not a table.');
  return root;
}
const rowsOf=root=>kids(root).filter(n=>n.localName==='tr'&&n.namespaceURI===W);
const cellsOf=tr=>kids(tr).filter(n=>n.localName==='tc'&&n.namespaceURI===W);
const isHeader=tr=>{const pr=named(tr,'trPr');return !!(pr&&named(pr,'tblHeader'));};

// The one paragraph a cell can be changed through, or why there is none.
function cellParagraph(tc,io){
  const ps=kids(tc).filter(n=>n.localName==='p'&&n.namespaceURI===W);
  const others=kids(tc).filter(n=>n.localName!=='tcPr'&&n.localName!=='p');
  if(others.length||ps.length!==1)return {editable:false,reason:'This cell holds more than words, so it is changed with the tools beside the guide.'};
  const info=readParagraph(io.serialize(ps[0]),CELL_STYLE,io);
  return info.editable?{editable:true,paragraph:ps[0],text:info.text}:{editable:false,reason:info.reason,paragraph:ps[0]};
}

/**
 * The table as rows of cells.
 * @returns {{caption:string, description:string, rows:{header:boolean, cells:{editable:boolean, text:string, span:number, reason?:string}[]}[]}}
 */
export function readTable(xml,io){
  const root=open(xml,io);
  const pr=named(root,'tblPr');
  const caption=attr(pr&&named(pr,'tblCaption'),'val'),description=attr(pr&&named(pr,'tblDescription'),'val');
  const rows=rowsOf(root).map(tr=>({header:isHeader(tr),cells:cellsOf(tr).map(tc=>{
    const c=cellParagraph(tc,io),span=spanOf(tc);
    if(c.editable)return {editable:true,text:c.text,span};
    const text=c.paragraph?Array.from(c.paragraph.getElementsByTagNameNS(W,'t')).map(t=>t.textContent??'').join(''):'';
    return {editable:false,text,reason:c.reason,span};
  })}));
  return {caption,description,rows};
}

function cellAt(root,r,c){
  const tr=rowsOf(root)[r];
  const tc=tr&&cellsOf(tr)[c];
  if(!tc)throw new Error('That cell is no longer in the table.');
  return tc;
}

/** The table with other words in one cell. When the words are the same, the very same XML comes back. */
export function setCellWords(xml,r,c,words,io){
  const root=open(xml,io);
  const cell=cellParagraph(cellAt(root,r,c),io);
  if(!cell.editable)throw new Error(cell.reason);
  const old=io.serialize(cell.paragraph);
  const fresh=setWords(old,CELL_STYLE,words,io);
  if(fresh===old)return xml;
  const imported=root.ownerDocument.importNode(io.parse(fresh).documentElement,true);
  cell.paragraph.parentNode.replaceChild(imported,cell.paragraph);
  return io.serialize(root);
}

// A row of empty cells made from a row that is there: the same widths and kinds of cell, no words, check boxes unchecked.
function emptyRowFrom(tr,doc,io,nextId){
  const copy=tr.cloneNode(true);
  const pr=named(copy,'trPr');
  const header=!!(pr&&named(pr,'tblHeader'));
  if(header&&pr)pr.removeChild(named(pr,'tblHeader'));
  for(const tc of cellsOf(copy)){
    const tcPr=named(tc,'tcPr');
    const shade=tcPr&&named(tcPr,'shd');
    if(header&&shade)tcPr.removeChild(shade);
    const c=cellParagraph(tc,io);
    if(c.editable){
      const p=c.paragraph;
      for(const run of kids(p).filter(n=>n.localName==='r'))p.removeChild(run);
      if(header){
        const pStyle=named(named(p,'pPr')||{childNodes:[]},'pStyle');
        if(pStyle)pStyle.setAttributeNS(W,'w:val','OSGTableCell');
      }
    }
  }
  for(const id of Array.from(copy.getElementsByTagNameNS(W,'id')))id.setAttributeNS(W,'w:val',String(nextId()));
  for(const checked of Array.from(copy.getElementsByTagNameNS('http://schemas.microsoft.com/office/word/2010/wordml','checked')))checked.setAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml','w14:val','0');
  return copy;
}

/** The table with an empty row after row `r` (after the header, a new first row of the body). */
export function addRow(xml,r,io){
  const root=open(xml,io);
  const rows=rowsOf(root);
  if(!rows[r])throw new Error('That row is no longer in the table.');
  if(rows.length>=TABLE_LIMITS.rows)throw new Error(`A table holds up to ${TABLE_LIMITS.rows} rows.`);
  const ids=Array.from(root.getElementsByTagNameNS(W,'id')).map(n=>Number(attr(n,'val'))).filter(Number.isFinite);
  let next=Math.max(700000,...ids);
  // The row to copy: the one the new row goes after, unless that is the header, then the first row of the body (or the header made plain).
  const model=isHeader(rows[r])?(rows.find(x=>!isHeader(x))||rows[r]):rows[r];
  const fresh=emptyRowFrom(model,root.ownerDocument,io,()=>++next);
  rows[r].parentNode.insertBefore(fresh,rows[r].nextSibling);
  return io.serialize(root);
}

/** The table without row `r`. The header row, which names the columns, stays. */
export function deleteRow(xml,r,io){
  const root=open(xml,io);
  const rows=rowsOf(root);
  if(!rows[r])throw new Error('That row is no longer in the table.');
  if(isHeader(rows[r]))throw new Error('The header row names the columns, so it stays.');
  rows[r].parentNode.removeChild(rows[r]);
  return io.serialize(root);
}

// ---- columns, merged cells, and the title and description of a table
// A cell in a row, with the first column it covers and how many columns it spans (a merged cell spans more than one).
const spanOf=tc=>{const pr=named(tc,'tcPr'),s=pr&&named(pr,'gridSpan');return s?Math.max(1,Math.min(64,Number(attr(s,'val'))||1)):1;};
function cellStarts(tr){
  let start=0;
  return cellsOf(tr).map(tc=>{const span=spanOf(tc),x={tc,start,span};start+=span;return x;});
}
const gridCols=root=>{const g=named(root,'tblGrid');return g?kids(g).filter(n=>n.localName==='gridCol'&&n.namespaceURI===W):[];};
const widthsOf=root=>gridCols(root).map(g=>Number(attr(g,'w'))||0);
const sum=a=>a.reduce((x,y)=>x+y,0);
const setW=(el,w)=>el.setAttributeNS(W,'w:w',String(Math.max(1,Math.round(w))));
// The widths scaled to add up to `target`, each at least 1, the last one taking up what rounding left over.
function scaleTo(widths,target){
  const total=sum(widths)||widths.length;
  const out=widths.map(w=>Math.max(1,Math.round((w||1)*target/total)));
  out[out.length-1]=Math.max(1,out[out.length-1]+(target-sum(out)));
  return out;
}
function refuseMergedDown(root){
  if(root.getElementsByTagNameNS(W,'vMerge').length)throw new Error('This table has cells merged downwards, which this program does not change. Change it in Word.');
}
function setSpan(tc,span){
  let pr=named(tc,'tcPr');
  if(!pr){pr=tc.ownerDocument.createElementNS(W,'w:tcPr');tc.insertBefore(pr,tc.firstChild);}
  let s=named(pr,'gridSpan');
  if(span<=1){if(s)pr.removeChild(s);return;}
  if(!s){
    s=tc.ownerDocument.createElementNS(W,'w:gridSpan');
    // gridSpan follows the cell's width in a cell's properties.
    const w=named(pr,'tcW');
    pr.insertBefore(s,w?w.nextSibling:pr.firstChild);
  }
  s.setAttributeNS(W,'w:val',String(span));
}
// Every cell's width is the width of the columns it covers.
function fixWidths(tr,widths){
  for(const {tc,start,span} of cellStarts(tr)){
    let pr=named(tc,'tcPr');
    if(!pr){pr=tc.ownerDocument.createElementNS(W,'w:tcPr');tc.insertBefore(pr,tc.firstChild);}
    let w=named(pr,'tcW');
    if(!w){w=tc.ownerDocument.createElementNS(W,'w:tcW');pr.insertBefore(w,pr.firstChild);}
    w.setAttributeNS(W,'w:w',String(sum(widths.slice(start,start+span))));
    w.setAttributeNS(W,'w:type','dxa');
  }
}
// An empty cell made from one that is there: the same look (shading, alignment, the paragraph's style), no words, no check box, no picture.
function blankCell(tc){
  const copy=tc.cloneNode(true);
  const pr=named(copy,'tcPr');
  if(pr){const s=named(pr,'gridSpan');if(s)pr.removeChild(s);}
  const paragraphs=kids(copy).filter(n=>n.localName==='p'&&n.namespaceURI===W);
  for(const n of kids(copy))if(n!==pr)copy.removeChild(n);
  const p=copy.ownerDocument.createElementNS(W,'w:p');
  const ppr=paragraphs[0]&&named(paragraphs[0],'pPr');
  if(ppr)p.appendChild(ppr.cloneNode(true));
  copy.appendChild(p);
  return copy;
}
const TABLE_COLUMNS=12;
export const COLUMN_LIMIT=TABLE_COLUMNS;

/** The table with an empty column after grid column `c`: inside a merged cell it makes the cell wider. The other columns give up the room it needs. */
export function addColumn(xml,c,io){
  const root=open(xml,io);
  refuseMergedDown(root);
  const g=widthsOf(root),n=g.length;
  if(!n)throw new Error('This table has no column widths, so a column cannot be added to it.');
  if(!Number.isInteger(c)||c<0||c>=n)throw new Error('That column is no longer in the table.');
  if(n>=TABLE_COLUMNS)throw new Error(`A table holds up to ${TABLE_COLUMNS} columns.`);
  const total=sum(g),added=Math.round(total/(n+1));
  const scaled=scaleTo(g,total-added),widths=[...scaled.slice(0,c+1),added,...scaled.slice(c+1)];
  const rows=rowsOf(root);
  for(const tr of rows)if(sum(cellStarts(tr).map(x=>x.span))!==n)throw new Error('A row of this table does not match its columns, so a column cannot be added.');
  const cols=gridCols(root);
  cols[c].parentNode.insertBefore(cols[c].cloneNode(true),cols[c].nextSibling);
  gridCols(root).forEach((el,i)=>setW(el,widths[i]));
  for(const tr of rows){
    const cover=cellStarts(tr).find(x=>x.start<=c&&c<x.start+x.span);
    if(cover.span>1&&c<cover.start+cover.span-1)setSpan(cover.tc,cover.span+1);
    else tr.insertBefore(blankCell(cover.tc),cover.tc.nextSibling);
    fixWidths(tr,widths);
  }
  return io.serialize(root);
}

/** The table without grid column `c`: a merged cell that covers it becomes narrower, any other cell in it is taken out with its words. The rest take up its room. */
export function deleteColumn(xml,c,io){
  const root=open(xml,io);
  refuseMergedDown(root);
  const g=widthsOf(root),n=g.length;
  if(!n)throw new Error('This table has no column widths, so a column cannot be taken out of it.');
  if(!Number.isInteger(c)||c<0||c>=n)throw new Error('That column is no longer in the table.');
  if(n<=1)throw new Error('A table needs at least one column, so the last one stays.');
  const widths=scaleTo(g.filter((_,i)=>i!==c),sum(g));
  const rows=rowsOf(root);
  for(const tr of rows)if(sum(cellStarts(tr).map(x=>x.span))!==n)throw new Error('A row of this table does not match its columns, so a column cannot be taken out.');
  for(const tr of rows){
    const cover=cellStarts(tr).find(x=>x.start<=c&&c<x.start+x.span);
    if(cover.span>1)setSpan(cover.tc,cover.span-1);else tr.removeChild(cover.tc);
  }
  const col=gridCols(root)[c];col.parentNode.removeChild(col);
  gridCols(root).forEach((el,i)=>setW(el,widths[i]));
  for(const tr of rows)fixWidths(tr,widths);
  return io.serialize(root);
}

/** The cell `c` of row `r` joined with the cell after it: their words are put together, and the two make one cell as wide as both. Header cells are not joined. */
export function mergeCell(xml,r,c,io){
  const root=open(xml,io);
  refuseMergedDown(root);
  const tr=rowsOf(root)[r];
  if(!tr)throw new Error('That row is no longer in the table.');
  if(isHeader(tr))throw new Error('The header row names each column, so its cells are not joined.');
  const cells=cellsOf(tr),a=cells[c],b=cells[c+1];
  if(!a)throw new Error('That cell is no longer in the table.');
  if(!b)throw new Error('There is no cell after this one to join it with.');
  const pa=cellParagraph(a,io),pb=cellParagraph(b,io);
  if(!pa.editable||!pb.editable)throw new Error('A cell that holds more than words cannot be joined. Take its check box or picture out first.');
  const words=[pa.text,pb.text].filter(t=>t.trim()).join(' ');
  if(words!==pa.text){
    const old=io.serialize(pa.paragraph),fresh=setWords(old,CELL_STYLE,words,io);
    if(fresh!==old)a.replaceChild(root.ownerDocument.importNode(io.parse(fresh).documentElement,true),pa.paragraph);
  }
  setSpan(a,spanOf(a)+spanOf(b));
  tr.removeChild(b);
  fixWidths(tr,widthsOf(root));
  return io.serialize(root);
}

/** A joined cell made into cells again: its words stay in the first, and the others are empty. */
export function splitCell(xml,r,c,io){
  const root=open(xml,io);
  refuseMergedDown(root);
  const tr=rowsOf(root)[r];
  if(!tr)throw new Error('That row is no longer in the table.');
  const a=cellsOf(tr)[c];
  if(!a)throw new Error('That cell is no longer in the table.');
  const span=spanOf(a);
  if(span<2)throw new Error('This cell is not joined to another, so there is nothing to split.');
  setSpan(a,1);
  let after=a;
  for(let i=1;i<span;i++){const cell=blankCell(a);tr.insertBefore(cell,after.nextSibling);after=cell;}
  fixWidths(tr,widthsOf(root));
  return io.serialize(root);
}

const clean=t=>String(t??'').replace(/\s+/g,' ').trim();
export const ALT_LIMITS=Object.freeze({caption:200,description:500});
/** The table with another title and description (what a screen reader says about it). The title is needed; the description may be empty. */
export function setTableAlt(xml,{caption,description},io){
  const root=open(xml,io);
  const title=clean(caption),about=clean(description);
  if(!title)throw new Error('Give the table a title: a few words naming it, which a screen reader says first.');
  if(title.length>ALT_LIMITS.caption)throw new Error(`The title is too long: it holds up to ${ALT_LIMITS.caption} letters.`);
  if(about.length>ALT_LIMITS.description)throw new Error(`The description is too long: it holds up to ${ALT_LIMITS.description} letters.`);
  const pr=named(root,'tblPr');
  if(!pr)throw new Error('This table has no properties to change.');
  const doc=root.ownerDocument;
  let cap=named(pr,'tblCaption'),desc=named(pr,'tblDescription');
  if(!cap){cap=doc.createElementNS(W,'w:tblCaption');pr.insertBefore(cap,desc||null);}
  cap.setAttributeNS(W,'w:val',title);
  if(about){
    if(!desc){desc=doc.createElementNS(W,'w:tblDescription');pr.appendChild(desc);}
    desc.setAttributeNS(W,'w:val',about);
  }else if(desc)pr.removeChild(desc);
  return io.serialize(root);
}

// ---- pictures in cells
// A cell holds at most one picture, in the cell's one paragraph, after its words. The picture is changed with the same code as a picture in a paragraph.
const WP_NS='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const NS_DECLARATIONS=`xmlns:w="${W}" xmlns:wp="${WP_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const EMU_PER_TWIP=635;
/** The room a picture has in a cell, in EMU: the cell's width less its margins, and at least half an inch. */
export const cellRoom=twips=>Math.max(457200,(Number(twips)-240)*EMU_PER_TWIP);

function onePictureCell(root,r,c){
  const tc=cellAt(root,r,c);
  const ps=kids(tc).filter(n=>n.localName==='p'&&n.namespaceURI===W);
  const others=kids(tc).filter(n=>n.localName!=='tcPr'&&n.localName!=='p');
  return {tc,paragraph:ps.length===1&&!others.length?ps[0]:null};
}
/**
 * What a cell has to say about pictures: how much room it gives one (EMU), whether a picture can be put in it, and the picture it has.
 * @returns {{room:number, canAdd:boolean, why:string, picture:ReturnType<typeof readPicture>}}
 */
export function cellPictureInfo(xml,r,c,io){
  const root=open(xml,io);
  const {tc,paragraph}=onePictureCell(root,r,c);
  const w=named(named(tc,'tcPr')||{childNodes:[]},'tcW'),twips=Number(attr(w,'w'))||2000;
  const room=cellRoom(twips);
  if(!paragraph)return {room,canAdd:false,why:'This cell holds more than words, so it is changed in Word.',picture:null};
  const picture=readPicture(io.serialize(paragraph),io);
  if(picture)return {room,canAdd:false,why:'This cell already has a picture. Change it with Picture details.',picture};
  const cell=cellParagraph(tc,io);
  if(!cell.editable)return {room,canAdd:false,why:'This cell holds more than words, so a picture cannot be put in it here.',picture:null};
  return {room,canAdd:true,why:'',picture:null};
}

/** The table with a picture in a cell: `runXml` is the picture's run (as pictureRun makes it), put after the cell's words. */
export function addCellPicture(xml,r,c,runXml,io){
  const info=cellPictureInfo(xml,r,c,io);
  if(!info.canAdd)throw new Error(info.why);
  const root=open(xml,io);
  const {paragraph}=onePictureCell(root,r,c);
  let run;
  try{run=io.parse(`<w:holder ${NS_DECLARATIONS}>${runXml}</w:holder>`).documentElement.firstChild;}catch{run=null;}
  if(!run||run.localName!=='r'||!run.getElementsByTagNameNS(W,'drawing').length)throw new Error('That picture could not be put in the cell.');
  paragraph.appendChild(root.ownerDocument.importNode(run,true));
  return io.serialize(root);
}

/** The table with the cell's picture described, marked as decoration or not, or made another width (kept inside the cell). */
export function changeCellPicture(xml,r,c,change,io){
  const root=open(xml,io);
  const {paragraph}=onePictureCell(root,r,c);
  if(!paragraph||!readPicture(io.serialize(paragraph),io))throw new Error('This cell has no picture of yours to change.');
  const room=cellPictureInfo(xml,r,c,io).room;
  const fresh=setPicture(io.serialize(paragraph),{...change,maxCx:room},io);
  paragraph.parentNode.replaceChild(root.ownerDocument.importNode(io.parse(fresh).documentElement,true),paragraph);
  return io.serialize(root);
}

/** The table with the cell's picture taken out; the words of the cell stay. */
export function removeCellPicture(xml,r,c,io){
  const root=open(xml,io);
  const {paragraph}=onePictureCell(root,r,c);
  if(!paragraph||!readPicture(io.serialize(paragraph),io))throw new Error('This cell has no picture of yours to take out.');
  for(const run of kids(paragraph).filter(n=>n.localName==='r'&&n.getElementsByTagNameNS(W,'drawing').length))paragraph.removeChild(run);
  return io.serialize(root);
}
