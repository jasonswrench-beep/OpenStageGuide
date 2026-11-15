// A picture in a table cell (src/guide-table.mjs): put after the cell's words, kept inside the cell, described and changed like any picture, taken out again, and carried
// into the guide with the table (replaceTable keeps its pictures under ids of their own) and into the Word file.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {readTable,cellRoom,cellPictureInfo,addCellPicture,changeCellPicture,removeCellPicture} from '../src/guide-table.mjs';
import {readPicture} from '../src/guide-picture.mjs';
import {docxFromPackage} from '../src/guide-docx.mjs';
import {formTable} from '../src/guide-forms-ooxml.mjs';
import {para,run,makeContext,pictureRun} from '../src/guide-ooxml.mjs';
import {pictureSize} from '../src/image-prepare.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const sample=ctx=>formTable({columns:[{name:'Topic',weight:2},{name:'Notes',weight:3},{name:'Done',weight:1,kind:'check'}],rows:[['Plan','First note']],blankRows:1,ctx,caption:'Session notes'});
async function withTable(){
  const a=demo();
  await a.insert(ctx=>sample(ctx)+para('OSGCueBody',run('After the table')),{where:'start',spare:false});
  const ps=a.exportState().paragraphs,at=ps.findIndex(p=>/^<w:tbl[\s>]/.test(p.xml));
  return {a,table:ps[at],at};
}
const made=(alt='A flip chart',decorative=false,width='large',cx)=>{
  const ctx=makeContext();
  const size=pictureSize({width:300,height:200},width);
  const runXml=pictureRun({base64:PNG,type:'png',mime:'image/png',size:cx?{cx,cy:Math.round(cx*2/3)}:size,alt,decorative},ctx);
  return {runXml,pictures:ctx.pictures};
};

test('a cell says how much room it gives a picture and whether one can go in: words and empty cells yes, a check box and a cell that has one no',async()=>{
  const {table}=await withTable();
  const info=cellPictureInfo(table.xml,1,0,io);
  assert.equal(info.canAdd,true);assert.equal(info.room,cellRoom(3120),'the cell is 3120 twips wide');
  assert.ok(info.room<3120*635,'less than the whole cell: its margins are left');
  assert.equal(cellPictureInfo(table.xml,2,1,io).canAdd,true,'an empty cell');
  const box=cellPictureInfo(table.xml,1,2,io);assert.equal(box.canAdd,false);assert.match(box.why,/more than words/);
  assert.throws(()=>cellPictureInfo(table.xml,9,0,io),/no longer in the table/);
  assert.equal(cellRoom(100),457200,'at least half an inch, however narrow');
});

test('a picture goes in after the cell\'s words, nothing else in the table changes, and a second picture is refused',async()=>{
  const {table}=await withTable();
  const {runXml}=made();
  const out=addCellPicture(table.xml,1,0,runXml,io);
  const t=readTable(out,io);
  assert.equal(t.rows[1].cells[0].text,'Plan','the words are still there');assert.equal(t.rows[1].cells[0].editable,false,'a cell with a picture is not typed in');
  assert.equal(t.rows[1].cells[1].text,'First note');assert.equal(t.caption,'Session notes');
  const pic=cellPictureInfo(out,1,0,io);
  assert.equal(pic.picture.description,'A flip chart');assert.equal(pic.canAdd,false);assert.match(pic.why,/already has a picture/);
  assert.throws(()=>addCellPicture(out,1,0,runXml,io),/already has a picture/);
  assert.throws(()=>addCellPicture(table.xml,1,2,runXml,io),/more than words/);
  assert.throws(()=>addCellPicture(table.xml,1,0,'<w:r><w:t>not a picture</w:t></w:r>',io),/could not be put in the cell/);
  const text=out.indexOf('Plan'),drawing=out.indexOf('<w:drawing');assert.ok(text>0&&drawing>text,'the picture is after the words');
});

test('a picture in a cell is changed like any picture, but never made wider than the cell, and can be taken out with the words left',async()=>{
  const {table}=await withTable();
  const withPic=addCellPicture(table.xml,1,2-2,made('x',false,'small').runXml,io);
  const room=cellPictureInfo(withPic,1,0,io).room;
  const wide=changeCellPicture(withPic,1,0,{description:'A wide picture',decorative:false,width:'large'},io);
  const p=cellPictureInfo(wide,1,0,io).picture;
  assert.equal(p.description,'A wide picture');assert.ok(p.cx<=room,`kept inside the cell: ${p.cx} is not more than ${room}`);
  assert.equal(p.cy,Math.round(p.cx*2/3),'and still the shape it was');
  const deco=changeCellPicture(wide,1,0,{description:'',decorative:true},io);
  assert.equal(cellPictureInfo(deco,1,0,io).picture.decorative,true);
  assert.throws(()=>changeCellPicture(wide,1,0,{description:'',decorative:false},io),/Describe the picture/);
  assert.throws(()=>changeCellPicture(table.xml,1,0,{description:'x',decorative:false},io),/no picture of yours/);
  const gone=removeCellPicture(wide,1,0,io);
  assert.equal(readTable(gone,io).rows[1].cells[0].text,'Plan');assert.equal(readTable(gone,io).rows[1].cells[0].editable,true,'the cell can be typed in again');
  assert.ok(!gone.includes('<w:drawing'));
  assert.throws(()=>removeCellPicture(gone,1,0,io),/no picture of yours/);
});

test('the table comes into the guide with its picture under an id of its own, and the Word file carries the picture',async()=>{
  const {a,table,at}=await withTable();
  const {runXml,pictures}=made();
  const xml=addCellPicture(table.xml,2,0,runXml,io);
  await a.replaceTable({start:at,guard:table.text,xml,select:at,pictures});
  const kept=a.exportState();
  const id=/r:embed="(rIdPic[A-Za-z0-9]+)"/.exec(kept.paragraphs[at].xml)?.[1];
  assert.ok(id,'the picture has an id in the table');assert.ok(kept.media[id],'and the guide holds the picture under it');
  assert.equal(kept.media[id].mime,'image/png');assert.equal(kept.media[id].base64,PNG);
  assert.notEqual(id,pictures[0].rid,'not the id it was made with, which another picture may have');
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:kept.paragraphs}));
  assert.ok(Object.keys(zip.files).some(n=>/^word\/media\/image\d+\.png$/.test(n)),'the Word file has the picture');
  assert.ok(!(await a.sourcePackage()).includes('rIdPic'+'undefined'));
  // two tables each given "their first picture" do not share one
  const second=made('Another picture');
  const again=addCellPicture(kept.paragraphs[at].xml,1,1,second.runXml,io);
  await a.replaceTable({start:at,guard:kept.paragraphs[at].text,xml:again,select:at,pictures:second.pictures});
  const ids=[...new Set([...a.exportState().paragraphs[at].xml.matchAll(/r:embed="(rIdPic[A-Za-z0-9]+)"/g)].map(m=>m[1]))];
  assert.equal(ids.length,2,'each picture has its own id');
  assert.equal(readPicture('<x/>',io),null);
});

test('a picture that is not one of the kinds a Word file holds is refused before the table changes',async()=>{
  const {a,table,at}=await withTable();
  const {runXml}=made();
  const xml=addCellPicture(table.xml,2,0,runXml,io);
  await assert.rejects(a.replaceTable({start:at,guard:table.text,xml,pictures:[{rid:'rIdPic1',mime:'image/svg+xml',base64:PNG}]}),/./);
  assert.equal(a.exportState().paragraphs[at].xml,table.xml,'the table is as it was');
});
