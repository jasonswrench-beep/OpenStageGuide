// Columns, joined cells and the title and description of a table (src/guide-table.mjs), and putting the changed table back into the guide (replaceTable).
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {readTable,addColumn,deleteColumn,mergeCell,splitCell,setTableAlt,setCellWords,COLUMN_LIMIT} from '../src/guide-table.mjs';
import {formTable} from '../src/guide-forms-ooxml.mjs';
import {para,run} from '../src/guide-ooxml.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function withTable(make){
  const a=demo();
  await a.insert(ctx=>make(ctx)+para('OSGCueBody',run('After the table')),{where:'start',spare:false});
  const ps=a.exportState().paragraphs,at=ps.findIndex(p=>/^<w:tbl[\s>]/.test(p.xml));
  return {a,ps,table:ps[at],at};
}
const sample=ctx=>formTable({columns:[{name:'Topic',weight:2},{name:'Notes',weight:3},{name:'Done',weight:1,kind:'check'}],rows:[['Plan','First note']],blankRows:2,ctx,caption:'Session notes'});
const widths=xml=>Array.from(io.parse(xml).documentElement.getElementsByTagNameNS(W,'gridCol')).map(g=>Number(g.getAttributeNS(W,'w')));
const rowWidths=xml=>Array.from(io.parse(xml).documentElement.getElementsByTagNameNS(W,'tr')).map(tr=>Array.from(tr.childNodes).filter(n=>n.localName==='tc').map(tc=>Number(tc.getElementsByTagNameNS(W,'tcW')[0].getAttributeNS(W,'w'))));
const sum=a=>a.reduce((x,y)=>x+y,0);
const texts=xml=>readTable(xml,io).rows.map(r=>r.cells.map(c=>c.text));

test('a column is added after another: every row gets an empty cell there, the header cell is a header cell, and the widths still fill the table',async()=>{
  const {table}=await withTable(sample);
  const out=addColumn(table.xml,1,io);
  const t=readTable(out,io);
  assert.deepEqual(t.rows.map(r=>r.cells.length),[4,4,4,4]);
  assert.deepEqual(texts(out)[0],['Topic','Notes','','Done'],'the new header cell is empty, to be named');
  assert.deepEqual(texts(out)[1],['Plan','First note','','☐'],'the other cells did not move');
  assert.equal(t.rows[0].header,true);
  assert.ok(out.match(/<w:shd [^>]*E4E8EE/g).length>=4,'the new header cell has the header shading');
  const g=widths(out);
  assert.equal(g.length,4);assert.equal(sum(g),9360,'the grid still adds up to the table\'s width');
  for(const row of rowWidths(out))assert.deepEqual(row,g,'every cell is as wide as its column');
  assert.ok(g[0]<3120&&g[1]<4680,'the others made room');
  assert.equal((out.match(/<w14:checkbox/g)||[]).length,(table.xml.match(/<w14:checkbox/g)||[]).length,'no check box was lost or doubled');
  assert.equal(readTable(out,io).caption,'Session notes');
  assert.equal(addColumn(table.xml,2,io).includes('OSGRatingBox'),true);
  const last=texts(addColumn(table.xml,2,io));assert.equal(last[0].length,4);assert.equal(last[1][3],'','the new cell at the end has no check box');
});

test('a column is taken out: its cells go with their words, the others take up its room, and the last column stays',async()=>{
  const {table}=await withTable(sample);
  const out=deleteColumn(table.xml,1,io);
  assert.deepEqual(texts(out)[0],['Topic','Done']);
  assert.deepEqual(texts(out)[1],['Plan','☐']);
  const g=widths(out);assert.equal(g.length,2);assert.equal(sum(g),9360);
  for(const row of rowWidths(out))assert.deepEqual(row,g);
  assert.throws(()=>deleteColumn(deleteColumn(out,1,io),0,io),/needs at least one column/);
  assert.throws(()=>deleteColumn(table.xml,9,io),/no longer in the table/);
  assert.throws(()=>addColumn(table.xml,-1,io),/no longer in the table/);
});

test('a table holds up to twelve columns',async()=>{
  const {table}=await withTable(sample);
  let xml=table.xml;
  for(let i=3;i<COLUMN_LIMIT;i++)xml=addColumn(xml,0,io);
  assert.equal(widths(xml).length,COLUMN_LIMIT);assert.equal(sum(widths(xml)),9360);
  assert.throws(()=>addColumn(xml,0,io),/up to 12 columns/);
});

test('two cells are joined into one: their words are put together, the width is both, and the header row is not joined',async()=>{
  const {table}=await withTable(sample);
  const out=mergeCell(table.xml,1,0,io);
  assert.deepEqual(texts(out)[1],['Plan First note','☐']);
  assert.equal(readTable(out,io).rows[1].cells[0].span,2);
  assert.deepEqual(rowWidths(out)[1],[7800,1560]);
  assert.deepEqual(rowWidths(out)[2],[3120,4680,1560],'the other rows are as they were');
  assert.throws(()=>mergeCell(table.xml,0,0,io),/header row names each column/);
  assert.throws(()=>mergeCell(table.xml,1,1,io),/holds more than words/,'a check box is not joined');
  assert.throws(()=>mergeCell(table.xml,1,2,io),/no cell after this one/);
  const empty=mergeCell(table.xml,2,0,io);assert.deepEqual(texts(empty)[2],['','☐'],'two empty cells make an empty one');
  // a joined cell is edited like any other
  assert.equal(readTable(setCellWords(out,1,0,'Joined words',io),io).rows[1].cells[0].text,'Joined words');
});

test('a joined cell is split again, and a column added or taken out through a joined cell keeps every row in step',async()=>{
  const {table}=await withTable(sample);
  const joined=mergeCell(table.xml,1,0,io);
  const split=splitCell(joined,1,0,io);
  assert.deepEqual(texts(split)[1],['Plan First note','','☐'],'the words stay in the first cell');
  assert.deepEqual(rowWidths(split)[1],[3120,4680,1560]);
  assert.throws(()=>splitCell(table.xml,1,0,io),/not joined to another/);
  const more=addColumn(joined,0,io);                       // inside the joined cell: it only gets wider
  assert.equal(readTable(more,io).rows[1].cells[0].span,3);
  assert.equal(texts(more)[2].length,4);
  const fewer=deleteColumn(joined,1,io);                   // the second column is covered by the joined cell: it only gets narrower
  assert.equal(readTable(fewer,io).rows[1].cells[0].span,1);
  assert.equal(readTable(fewer,io).rows[1].cells.length,2);
  assert.equal(sum(widths(fewer)),9360);
  for(const row of rowWidths(fewer))assert.equal(sum(row),9360);
});

test('a table merged downwards is left as it is, with the reason',async()=>{
  const {table}=await withTable(sample);
  const down=table.xml.replace('<w:tcW w:w="3120" w:type="dxa"/><w:vAlign','<w:tcW w:w="3120" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign');
  assert.notEqual(down,table.xml);
  for(const f of [()=>addColumn(down,0,io),()=>deleteColumn(down,0,io),()=>mergeCell(down,1,0,io),()=>splitCell(down,1,0,io)])assert.throws(f,/merged downwards/);
});

test('the title and the description of a table are changed, the description may be taken away, and the title is needed',async()=>{
  const {table}=await withTable(sample);
  assert.equal(readTable(table.xml,io).description,'A table to fill in, with the columns: Topic, Notes, Done.');
  const out=setTableAlt(table.xml,{caption:'  Notes for the session ',description:'Three columns:\n topic, notes and a box to tick.'},io);
  const t=readTable(out,io);
  assert.equal(t.caption,'Notes for the session');assert.equal(t.description,'Three columns: topic, notes and a box to tick.');
  assert.equal(widths(out).join(),widths(table.xml).join(),'nothing else changed');
  assert.equal(readTable(setTableAlt(out,{caption:'Notes',description:''},io),io).description,'');
  assert.ok(!setTableAlt(out,{caption:'Notes',description:''},io).includes('tblDescription'));
  assert.throws(()=>setTableAlt(table.xml,{caption:'   ',description:'x'},io),/Give the table a title/);
  assert.throws(()=>setTableAlt(table.xml,{caption:'x'.repeat(201),description:''},io),/title is too long/);
  assert.throws(()=>setTableAlt(table.xml,{caption:'x',description:'y'.repeat(501)},io),/description is too long/);
  const bare=table.xml.replace(/<w:tblCaption[^>]*\/>/,'').replace(/<w:tblDescription[^>]*\/>/,'');
  const made=setTableAlt(bare,{caption:'Fresh',description:'Described'},io);
  assert.ok(made.indexOf('<w:tblCaption')<made.indexOf('<w:tblDescription'),'in the order Word keeps them');
});

test('a changed table is put back into the guide as the table, with its cells read again',async()=>{
  const {a,table,at}=await withTable(sample);
  const out=addColumn(table.xml,0,io);
  await a.replaceTable({start:at,guard:table.text,xml:out,select:at});
  const after=a.exportState().paragraphs;
  const head=after.slice(at,at+4).map(p=>p.style);
  assert.deepEqual(head,['OSG Table Head','OSG Table Head','OSG Table Head','OSG Table Head'],'four header cells now');
  const end=after.findIndex(p=>p.text==='After the table');
  assert.equal(end-at,16,'four columns and four rows, each a cell of its own');
  assert.ok(after.some(p=>p.text==='After the table'),'what is after the table is still there');
  assert.ok(a.exportState().paragraphs[at].xml.startsWith('<w:tbl'),'the table\'s XML is on its first cell');
});
