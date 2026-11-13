// Changing the cells of a table in the guide, without disturbing anything else in it. Pure functions: Word-format XML in, Word-format XML out.
// A cell the editor can change holds one paragraph of plain words (or nothing); a cell that holds a check box, a picture, several paragraphs or
// anything else is shown and left exactly as it is. The words of a cell are changed by the same code that changes the words of a paragraph
// (guide-edit.mjs), so a cell is held to the same rules as any other paragraph.
import {readParagraph,setWords} from './guide-edit.mjs';
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
 * @returns {{caption:string, rows:{header:boolean, cells:{editable:boolean, text:string, reason?:string}[]}[]}}
 */
export function readTable(xml,io){
  const root=open(xml,io);
  const pr=named(root,'tblPr');
  const caption=attr(pr&&named(pr,'tblCaption'),'val');
  const rows=rowsOf(root).map(tr=>({header:isHeader(tr),cells:cellsOf(tr).map(tc=>{
    const c=cellParagraph(tc,io);
    if(c.editable)return {editable:true,text:c.text};
    const text=c.paragraph?Array.from(c.paragraph.getElementsByTagNameNS(W,'t')).map(t=>t.textContent??'').join(''):'';
    return {editable:false,text,reason:c.reason};
  })}));
  return {caption,rows};
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
