// The icon of each block in a column at the left, and a slide's picture beside its script (src/guide-margin-layout.mjs, chosen in the Layout tool): made when the Word file is made,
// from the guide's own paragraphs, which do not change. A block becomes a table of one row; nothing in the block is lost, and every picture is still named by a relationship.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {layoutBlocks} from '../src/guide-margin-layout.mjs';
import {docxFromPackage} from '../src/guide-docx.mjs';
import {cleanLayout,describeLayout,LAYOUT_DEFAULTS} from '../src/guide-output-layout.mjs';
import {readSettings,writeSettings} from '../src/guide-forms.mjs';
import {cueBlock,moduleHeading,lessonHeading,timeBlock} from '../src/guide-ooxml.mjs';
import {pictureSize} from '../src/image-prepare.mjs';
import {STYLE_NAMES} from '../src/guide-styles.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const slidePicture={base64:PNG,type:'png',mime:'image/png',size:pictureSize({width:320,height:180},'medium'),alt:'Slide 1: Welcome',decorative:false};
async function guide(){
  const a=demo();
  await a.insert(ctx=>moduleHeading(1,'One')+lessonHeading(1,1,'First')+timeBlock({minutes:10,cumulative:10},ctx)
    +cueBlock({cue:'say',text:'Welcome everyone.',notes:'Smile.'},ctx)
    +cueBlock({cue:'slide',text:'Slide 1: Welcome',picture:slidePicture,slideText:'Words on the slide',notes:'Click twice.'},ctx)
    +cueBlock({cue:'ask',text:'What makes it hard?'},ctx),{where:'start',select:'last',spare:false});
  return a;
}
const count=(s,re)=>(s.match(re)||[]).length;

test('with no layout asked for, the entries come back as they were',async()=>{
  const a=await guide(),entries=a.exportState().paragraphs;
  assert.equal(layoutBlocks(entries,{},io),entries);assert.equal(layoutBlocks(entries,{marginIcons:false,slideBeside:false},io),entries);
});

test('a column of icons: each block\'s icon is in the left cell, the rest of the block is in the right, and nothing is lost',async()=>{
  const a=await guide(),entries=a.exportState().paragraphs;
  const out=layoutBlocks(entries,{marginIcons:true},io);
  const tables=out.filter(e=>/^<w:tbl[\s>]/.test(e.xml));
  assert.equal(tables.length,3,'three blocks, three tables');
  assert.ok(out.length<entries.length,'a block is one entry now');
  for(const t of tables){
    const doc=io.parse(t.xml).documentElement,cells=Array.from(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','tc'));
    assert.equal(cells.length,2);
    assert.equal(count(io.serialize(cells[0]),/<wp:docPr[^>]*name="Icon /g),1,'one icon on the left');
    assert.equal(count(io.serialize(cells[1]),/<wp:docPr[^>]*name="Icon /g),0,'none left in the cue line');
    assert.match(t.xml,/w:tblCaption w:val="Layout of a block"/);assert.match(t.xml,/no header row because it holds no data/);
  }
  assert.ok(out.some(e=>/One/.test(e.text)&&!/^<w:tbl/.test(e.xml)),'the module heading is as it was');assert.ok(out.some(e=>/First/.test(e.text)&&!/^<w:tbl/.test(e.xml)),'and the lesson heading');
  assert.equal(out.filter(e=>e.style==='OSG Block Layout'&&!/^<w:tbl/.test(e.xml)).length,3,'a small paragraph follows each table, as an entry of its own');
  for(const t of tables)assert.equal(io.parse(t.xml).documentElement.localName,'tbl'),assert.ok(t.xml.trimEnd().endsWith('</w:tbl>'),'a table entry holds one table and nothing else');
  const allWords=out.map(e=>e.xml).join('');
  for(const w of ['Welcome everyone.','Smile.','Click twice.','Words on the slide','What makes it hard?','Slide 1: Welcome'])assert.ok(allWords.includes(w),`${w} is still there`);
  assert.equal(count(allWords,/<wp:docPr[^>]*name="Picture /g),1,'and the slide\'s picture is still in the guide, once');
});

test('a slide\'s picture beside its script: the picture in the left cell, the cue line (with its icon), the slide\'s words and the notes in the right; a block with no picture is left alone',async()=>{
  const a=await guide(),entries=a.exportState().paragraphs;
  const out=layoutBlocks(entries,{slideBeside:true},io);
  const tables=out.filter(e=>/^<w:tbl[\s>]/.test(e.xml));
  assert.equal(tables.length,1,'only the slide block has a picture');
  const cells=Array.from(io.parse(tables[0].xml).documentElement.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','tc')).map(c=>io.serialize(c));
  assert.equal(cells.length,2);
  assert.equal(count(cells[0],/<wp:docPr[^>]*name="Picture /g),1);assert.equal(count(cells[1],/<wp:docPr[^>]*name="Picture /g),0,'the picture paragraph is not repeated on the right');
  assert.equal(count(cells[1],/<wp:docPr[^>]*name="Icon /g),1,'the icon stays beside the cue\'s name');
  for(const w of ['Slide 1: Welcome','Words on the slide','Click twice.'])assert.ok(cells[1].includes(w),w);
  assert.match(tables[0].xml,/the picture of the slide is in the left cell/);
  assert.equal(out.filter(e=>/Welcome everyone/.test(e.text)).length,1,'the block with no picture is an ordinary paragraph still');
});

test('both together: the icon, the slide\'s picture and then the words, in three cells',async()=>{
  const a=await guide(),out=layoutBlocks(a.exportState().paragraphs,{marginIcons:true,slideBeside:true},io);
  const slide=out.find(e=>/^<w:tbl/.test(e.xml)&&/Click twice/.test(e.xml));
  const cells=Array.from(io.parse(slide.xml).documentElement.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','tc')).map(c=>io.serialize(c));
  assert.equal(cells.length,3);
  assert.equal(count(cells[0],/<wp:docPr[^>]*name="Icon /g),1);assert.equal(count(cells[1],/<wp:docPr[^>]*name="Picture /g),1);assert.ok(cells[2].includes('Click twice.'));
  const grid=[...slide.xml.matchAll(/<w:gridCol w:w="(\d+)"/g)].map(m=>Number(m[1]));
  assert.equal(grid.length,3);assert.equal(grid.reduce((x,y)=>x+y,0),9360,'the columns fill the page');
});

test('the Word file holds the laid-out blocks as tables, is well formed, and every picture is still named by a relationship',async()=>{
  const a=await guide();
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:a.exportState().paragraphs,layout:{marginIcons:true,slideBeside:true}}));
  const doc=await zip.file('word/document.xml').async('string'),rels=await zip.file('word/_rels/document.xml.rels').async('string');
  assert.equal(io.parse(doc).getElementsByTagName('parsererror').length,0);
  assert.ok(count(doc,/<w:tbl>/g)+count(doc,/<w:tbl /g)>=3,'the blocks are tables in the file');
  const named=[...doc.matchAll(/r:embed="([^"]+)"/g)].map(m=>m[1]);
  assert.ok(named.length>=4&&named.every(id=>rels.includes(`Id="${id}"`)),'every picture the file uses has its relationship');
  const plain=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:a.exportState().paragraphs}));
  assert.ok(!(await plain.file('word/document.xml').async('string')).includes('Layout of a block'),'without the choice the file is as it was');
  const template=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{template:true,layout:{marginIcons:true}}));
  assert.ok(!(await template.file('word/document.xml').async('string')).includes('Layout of a block'),'a template has none of it');
});

test('the choices are kept with the guide, named in what the Word file says it did, and the Layout tool offers them',async()=>{
  assert.deepEqual(cleanLayout({marginIcons:true,slideBeside:true,dividers:'x'}),{...LAYOUT_DEFAULTS,marginIcons:true,slideBeside:true});
  assert.match(describeLayout({marginIcons:true,slideBeside:true}),/icon of each block is in a column at the left; the picture of each slide is beside its script/);
  assert.deepEqual(readSettings(writeSettings({layout:{marginIcons:true}})).layout,{...LAYOUT_DEFAULTS,marginIcons:true});
  const g=await started('Margin guide');
  await addModule(g,'One');await addLesson(g,'First','5');await addBlock(g,'say','Hello.');
  g.p.useTool('layout');await g.p.refresh();
  assert.equal(g.n('guide-layout-icons').checked,false);assert.equal(g.n('guide-layout-slides').checked,false);
  g.n('guide-layout-icons').checked=true;await g.n('guide-layout-icons').onchange();
  g.n('guide-layout-slides').checked=true;await g.n('guide-layout-slides').onchange();
  await new Promise(r=>setTimeout(r,400));
  const kept=readSettings(await g.adapter.formSettings()).layout;
  assert.equal(kept.marginIcons,true);assert.equal(kept.slideBeside,true);
  assert.match(g.n('guide-layout-says').textContent,/icon of each block is in a column at the left/);
  assert.ok(STYLE_NAMES.cue);
});
