// Two more layouts for the Word file (the Layout tool): the icon of each block in a column of its own at the left, and the picture of a slide beside the script that goes
// with it. Both are made at the time the Word file is made, from the guide's paragraphs: a block (its cue line and the paragraphs that belong to it) becomes one table of one
// row, with the icon or the slide's picture in a cell on the left and the words in a cell on the right. The guide itself does not change. The table has no borders and says
// in its title that it is for layout; its cells keep the paragraphs as they were, so the shading of the facilitator's notes and the cue's look are as in the guide.
// Pure: the guide's entries ({xml,text,style}) in, the entries for the Word file out.
import {kindOf,blockSpan} from './guide-model.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PAGE_WIDTH=9360,ICON_COLUMN=700,SLIDE_COLUMN=2900;
const kids=el=>Array.from(el.childNodes).filter(n=>n.nodeType===1);
const nameOf=drawing=>Array.from(drawing.getElementsByTagNameNS(WP,'docPr'))[0]?.getAttribute('name')||'';
const runsWith=(p,prefix)=>kids(p).filter(r=>r.localName==='r'&&r.namespaceURI===W).filter(r=>{
  const d=Array.from(r.getElementsByTagNameNS(W,'drawing'))[0];
  return d&&nameOf(d).startsWith(prefix);
});
const hasText=p=>Array.from(p.getElementsByTagNameNS(W,'t')).some(t=>(t.textContent||'').trim());

// The run of a picture as it goes in a cell of its own: without the lowering that sits an icon on the line of text it was beside.
function bareRun(run,doc,io){
  const copy=run.cloneNode(true);
  for(const pr of kids(copy).filter(n=>n.localName==='rPr'))copy.removeChild(pr);
  return io.serialize(copy);
}
const cell=(width,inner)=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${inner||'<w:p/>'}</w:tc>`;
const cellParagraph=(runXml,center)=>`<w:p xmlns:w="${W}">${center?'<w:pPr><w:jc w:val="center"/></w:pPr>':''}${runXml}</w:p>`;
const SEPARATOR=`<w:p xmlns:w="${W}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="40" w:lineRule="exact"/></w:pPr></w:p>`;

/**
 * The table for one block, or null when there is nothing to put at its left (a block with no icon, a slide with no picture, when that is what was asked for).
 * @param {{xml?:string,text?:string,style?:string}[]} block @param {{marginIcons:boolean,slideBeside:boolean}} layout
 */
function blockTable(block,layout,io){
  const cue=io.parse(block[0].xml).documentElement;
  if(!cue||cue.localName!=='p'||cue.namespaceURI!==W)return null;
  const doc=cue.ownerDocument;
  const left=[];let widths=[];
  const rest=[];
  // The icon at the front of the cue line goes to a column of its own.
  const icon=layout.marginIcons?runsWith(cue,'Icon')[0]:null;
  if(icon){left.push(cellParagraph(bareRun(icon,doc,io),true));widths.push(ICON_COLUMN);cue.removeChild(icon);}
  // A slide's picture (the paragraph straight under its cue line that holds nothing but a picture of the author's) goes beside the script.
  let slide=null,slideAt=-1;
  if(layout.slideBeside&&block.length>1&&block[1].xml){
    const p=io.parse(block[1].xml).documentElement;
    const pics=runsWith(p,'Picture');
    if(p.localName==='p'&&pics.length===1&&!hasText(p)){slide=pics[0];slideAt=1;}
  }
  if(slide){
    const holder=io.parse(block[1].xml).documentElement;
    left.push(cellParagraph(bareRun(runsWith(holder,'Picture')[0],doc,io),false));widths.push(SLIDE_COLUMN);
  }
  if(!left.length)return null;
  rest.push(io.serialize(cue));
  block.slice(1).forEach((e,i)=>{if(i+1!==slideAt)rest.push(e.xml);});
  const used=widths.reduce((a,b)=>a+b,0),columns=icon&&slide?[ICON_COLUMN,SLIDE_COLUMN]:widths;
  const cells=[];
  if(icon&&slide){cells.push(cell(ICON_COLUMN,left[0]),cell(SLIDE_COLUMN,left[1]));}
  else cells.push(cell(widths[0],left[0]));
  cells.push(cell(PAGE_WIDTH-used,rest.join('')));
  const grid=[...columns,PAGE_WIDTH-used].map(w=>`<w:gridCol w:w="${w}"/>`).join('');
  return `<w:tbl xmlns:w="${W}"><w:tblPr><w:tblW w:w="${PAGE_WIDTH}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/><w:tblCaption w:val="Layout of a block"/><w:tblDescription w:val="This table only lays out one block of the guide: ${slide?'the picture of the slide':'its icon'} is in the left cell and the words are in the right cell. It has no header row because it holds no data."/></w:tblPr><w:tblGrid>${grid}</w:tblGrid><w:tr>${cells.join('')}</w:tr></w:tbl>`;
}

/**
 * The entries for the Word file with each block laid out as asked: `marginIcons` puts its icon in a column at the left, `slideBeside` puts a slide's picture beside the script.
 * A block that has nothing for the left column is left as it is. The entries come back in the same order; a block's entries are replaced by two: its table, and a small paragraph after it.
 * @param {{xml?:string,text?:string,style?:string,builtIn?:string}[]} entries
 * @param {{marginIcons?:boolean,slideBeside?:boolean}} layout
 * @param {{parse:(x:string)=>Document, serialize:(e:Element)=>string}} io
 */
export function layoutBlocks(entries,layout,io){
  const want={marginIcons:!!layout?.marginIcons,slideBeside:!!layout?.slideBeside};
  if(!want.marginIcons&&!want.slideBeside)return entries;
  const out=[];
  for(let i=0;i<entries.length;i++){
    const e=entries[i];
    if(e?.xml&&kindOf(e)==='cue'){
      const span=blockSpan(entries,i);
      if(span&&span.start===i){
        const block=entries.slice(i,span.end+1);
        // A block is its own paragraphs only: a table or a paragraph with no XML inside it is not laid out.
        if(block.every(b=>b.xml&&/^<w:p[\s>]/.test(b.xml))){
          let table=null;
          try{table=blockTable(block,want,io);}catch{table=null;}
          // A table is followed by a small paragraph of its own, so two of them never run together into one table in Word.
          if(table){out.push({xml:table,text:block[0].text,style:'OSG Block Layout'},{xml:SEPARATOR,text:'',style:'OSG Block Layout'});i=span.end;continue;}
        }
      }
    }
    out.push(e);
  }
  return out;
}
