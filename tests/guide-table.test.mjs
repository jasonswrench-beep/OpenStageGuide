// Changing the cells of a table (src/guide-table.mjs) and replacing a table in the document (replaceTable).
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {readTable,setCellWords,addRow,deleteRow,TABLE_LIMITS} from '../src/guide-table.mjs';
import {formTable,checkList,ratingGrid} from '../src/guide-forms-ooxml.mjs';
import {para,run} from '../src/guide-ooxml.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
async function withTable(make){
  const a=demo();
  await a.insert(ctx=>make(ctx)+para('OSGCueBody',run('After the table')),{where:'start',spare:false});
  const ps=a.exportState().paragraphs;
  return {a,ps,table:ps.find(p=>/^<w:tbl[\s>]/.test(p.xml)),at:ps.findIndex(p=>/^<w:tbl[\s>]/.test(p.xml))};
}
const sample=ctx=>formTable({columns:[{name:'Topic',weight:2},{name:'Notes',weight:3},{name:'Done',weight:1,kind:'check'}],rows:[['Plan','First note']],blankRows:2,ctx,caption:'Session notes'});

test('a table reads as rows of cells: the header row, the words, and which cells can be changed',async()=>{
  const {table}=await withTable(sample);
  const t=readTable(table.xml,io);
  assert.equal(t.caption,'Session notes');
  assert.equal(t.rows.length,4);
  assert.deepEqual(t.rows.map(r=>r.header),[true,false,false,false]);
  assert.deepEqual(t.rows[0].cells.map(c=>c.text),['Topic','Notes','Done']);
  assert.deepEqual(t.rows[1].cells.map(c=>[c.text,c.editable]),[['Plan',true],['First note',true],['☐',false]]);
  assert.deepEqual(t.rows[2].cells.map(c=>[c.text,c.editable]),[['',true],['',true],['☐',false]],'empty cells can be written in; check boxes cannot');
  assert.match(t.rows[1].cells[2].reason,/more than words/);
  assert.throws(()=>readTable('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',io),/not a table/);
});

test('a cell\'s words change and nothing else does: widths, header mark, check boxes, other cells and the caption stay',async()=>{
  const {table}=await withTable(sample);
  const out=setCellWords(table.xml,1,1,'A <better> note & more',io);
  const t=readTable(out,io);
  assert.equal(t.rows[1].cells[1].text,'A <better> note & more');
  assert.deepEqual(t.rows[1].cells.map(c=>c.text),['Plan','A <better> note & more','☐']);
  assert.equal(t.caption,'Session notes');
  assert.equal(out.includes('<w:tblHeader/>'),true);
  assert.equal((out.match(/<w14:checkbox/g)||[]).length,(table.xml.match(/<w14:checkbox/g)||[]).length,'every check box is still there');
  assert.equal(out.match(/<w:gridCol[^>]*>/g).join(''),table.xml.match(/<w:gridCol[^>]*>/g).join(''),'the column widths are as they were');
  const empty=setCellWords(table.xml,2,0,'Filled in',io);
  assert.equal(readTable(empty,io).rows[2].cells[0].text,'Filled in','an empty cell can be written in');
  assert.equal(setCellWords(table.xml,1,0,'Plan',io),table.xml,'the same words give back the very same XML');
  const header=readTable(setCellWords(table.xml,0,0,'Subject',io),io);
  assert.equal(header.rows[0].cells[0].text,'Subject');assert.equal(header.rows[0].header,true,'a header cell is still a header cell');
});

test('a cell that holds a check box or is not there is refused',async()=>{
  const {table}=await withTable(sample);
  assert.throws(()=>setCellWords(table.xml,1,2,'x',io),/more than words/);
  assert.throws(()=>setCellWords(table.xml,9,0,'x',io),/no longer in the table/);
  assert.throws(()=>setCellWords(table.xml,0,9,'x',io),/no longer in the table/);
  assert.throws(()=>setCellWords(table.xml,-1,0,'x',io),/no longer in the table/);
});

test('every table the builders make can be read and each of its editable cells changed and changed back',async()=>{
  for(const [what,make] of [['a table',sample],['a check list',ctx=>checkList({items:['One','Two'],columns:['Who','When'],ctx})],['a rating grid',ctx=>ratingGrid({statements:['I plan.'],evidence:true,ctx})]]){
    const {table}=await withTable(make);
    const t=readTable(table.xml,io);
    let changed=0;
    t.rows.forEach((row,r)=>row.cells.forEach((cell,c)=>{
      if(!cell.editable)return;
      const there=setCellWords(table.xml,r,c,'Typed here',io);
      assert.equal(readTable(there,io).rows[r].cells[c].text,'Typed here',`${what} ${r},${c}`);
      assert.deepEqual(readTable(setCellWords(there,r,c,cell.text,io),io),t,`${what} ${r},${c} is as it was`);
      changed++;
    }));
    assert.ok(changed>=2,`${what}: some cells can be changed`);
  }
});

test('a row can be added after any row: empty, check boxes unchecked and numbered on their own, the header row is not copied',async()=>{
  const {table}=await withTable(sample);
  const out=addRow(table.xml,1,io);
  const t=readTable(out,io);
  assert.equal(t.rows.length,5);
  assert.deepEqual(t.rows[2].cells.map(c=>[c.text,c.editable]),[['',true],['',true],['☐',false]],'the new row is after the row it was added to, and empty');
  assert.deepEqual(t.rows[1].cells.map(c=>c.text),['Plan','First note','☐'],'the row before is as it was');
  const ids=[...out.matchAll(/<w:id w:val="(\d+)"\/>/g)].map(m=>m[1]);
  assert.equal(new Set(ids).size,ids.length,'every check box has an id of its own: '+ids);
  assert.equal(t.rows.filter(r=>r.header).length,1);
  const afterHeader=readTable(addRow(table.xml,0,io),io);
  assert.deepEqual(afterHeader.rows.map(r=>r.header),[true,false,false,false,false],'a row added under the header is a body row');
  assert.deepEqual(afterHeader.rows[1].cells.map(c=>c.text),['','','☐']);
  assert.throws(()=>addRow(table.xml,9,io),/no longer in the table/);
});

test('a row added to a table that has only a header makes a plain row from it; the row limit holds',async()=>{
  const {table}=await withTable(ctx=>formTable({columns:[{name:'A',weight:1},{name:'B',weight:1}],blankRows:0,ctx}));
  assert.equal(readTable(table.xml,io).rows.length,1);
  const out=addRow(table.xml,0,io);
  const t=readTable(out,io);
  assert.deepEqual(t.rows.map(r=>r.header),[true,false]);
  assert.ok(!/E4E8EE/.test(out.split('</w:tr>')[1]),'the new row is not shaded like a header');
  assert.match(out.split('</w:tr>')[1],/OSGTableCell/,'and its cells are table cells');
  let x=out;
  while(readTable(x,io).rows.length<TABLE_LIMITS.rows)x=addRow(x,1,io);
  assert.throws(()=>addRow(x,1,io),/up to 200 rows/);
});

test('a row can be deleted; the header row cannot',async()=>{
  const {table}=await withTable(sample);
  const out=deleteRow(table.xml,1,io);
  const t=readTable(out,io);
  assert.equal(t.rows.length,3);
  assert.deepEqual(t.rows.map(r=>r.cells[0].text),['Topic','','']);
  assert.throws(()=>deleteRow(table.xml,0,io),/header row/);
  assert.throws(()=>deleteRow(table.xml,9,io),/no longer in the table/);
  assert.equal(readTable(deleteRow(deleteRow(deleteRow(table.xml,3,io),2,io),1,io),io).rows.length,1,'down to the header row');
});

test('replaceTable puts a changed table in the document: its cells follow it, Undo puts it back, a stale or wrong request is refused',async()=>{
  const {a,ps,at,table}=await withTable(sample);
  const first=ps[at].text;
  const cells=ps.slice(at).findIndex((p,i)=>i>0&&p.xml)>0?ps.slice(at).findIndex((p,i)=>i>0&&p.xml):0;
  assert.ok(cells>=5,'the table covers its cells: '+cells);
  const before=a.exportState().paragraphs.length;
  await a.replaceTable({start:at,guard:first,xml:addRow(table.xml,1,io)});
  const after=a.exportState().paragraphs;
  assert.equal(after.length,before+3,'three more cells');
  assert.equal(after[at+cells+3].text,'After the table','what came after the table is still after it');
  assert.equal(readTable(after[at].xml,io).rows.length,5);
  a.undo();
  assert.equal(a.exportState().paragraphs.length,before);
  await assert.rejects(a.replaceTable({start:at,guard:'something else',xml:table.xml}),/changed since/);
  await assert.rejects(a.replaceTable({start:at+1,guard:'',xml:table.xml}),/changed since/,'a cell is not a table');
  await assert.rejects(a.replaceTable({start:at,guard:first,xml:'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'}),/could not be made/);
  await assert.rejects(a.replaceTable({start:at,guard:first,xml:'nope'}),/could not be made/);
  await assert.rejects(a.replaceTable({start:at,guard:first,xml:5}),/could not be made/);
  assert.equal(a.exportState().paragraphs.length,before,'nothing changed');
});

test('typing in one cell is one step of Undo',async()=>{
  const {a,ps,at,table}=await withTable(sample);
  const first=ps[at].text;
  let xml=table.xml;
  const base=a.historyState().undo;
  for(const t of ['A','AB','ABC']){xml=setCellWords(xml,1,0,t,io);await a.replaceTable({start:at,guard:first,xml,group:'cell:1:0'});}
  assert.equal(a.historyState().undo,base+1);
  a.undo();
  assert.equal(readTable(a.exportState().paragraphs[at].xml,io).rows[1].cells[0].text,'Plan');
});
