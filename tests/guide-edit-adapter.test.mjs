// The document's operations for the editor: replacing paragraphs under a guard, Undo and Redo with typing grouped into one step.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {createProjectAdapter} from '../src/guide-project-adapter.mjs';
import {setWords,splitWords,joinWords} from '../src/guide-edit.mjs';
import {formTable} from '../src/guide-forms-ooxml.mjs';
import {moduleHeading,lessonHeading,para,run} from '../src/guide-ooxml.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
async function seeded(){
  const a=demo();
  await a.insert(()=>moduleHeading(1,'Naming it')+lessonHeading(1,1,'First')+para('OSGCueBody',run('One two three'))+para('OSGCueBody',run('Four five')),{where:'start',spare:false});
  return a;
}
const rows=a=>a.exportState().paragraphs;
const at=(a,text)=>rows(a).findIndex(p=>p.text===text);

test('one paragraph\'s words are replaced; its text is read again from the new XML; Undo puts them back',async()=>{
  const a=await seeded();
  const i=at(a,'One two three'),p=rows(a)[i];
  await a.replaceParagraphs({start:i,count:1,guard:p.text,xmls:[setWords(p.xml,p.style,'One two',io)],select:i});
  assert.equal(rows(a)[i].text,'One two');
  assert.equal(rows(a)[i].style,'OSG Cue Body');
  assert.equal(a.cursor(),i);
  assert.equal(a.undo(),true);
  assert.equal(rows(a)[i].text,'One two three');
});

test('a paragraph that is not what the editor last saw is refused, and nothing changes',async()=>{
  const a=await seeded();
  const i=at(a,'One two three'),p=rows(a)[i];
  const before=JSON.stringify(rows(a));
  await assert.rejects(a.replaceParagraphs({start:i,count:1,guard:'Something else',xmls:[p.xml]}),/changed since you chose it/);
  await assert.rejects(a.replaceParagraphs({start:999,count:1,guard:'x',xmls:[p.xml]}),/changed since you chose it/);
  await assert.rejects(a.replaceParagraphs({start:-1,count:1,guard:'x',xmls:[p.xml]}),/changed since/);
  await assert.rejects(a.replaceParagraphs({start:i,count:1.5,guard:p.text,xmls:[p.xml]}),/changed since/);
  await assert.rejects(a.replaceParagraphs({start:i,count:1,guard:p.text,xmls:'nope'}),/changed since/);
  assert.equal(JSON.stringify(rows(a)),before);
  assert.equal(a.historyState().undo,1,'only the seeding is in the history: a refused change is not a step');
});

test('only a paragraph can be made: markup that is not one, two at once, or a damaged one is refused',async()=>{
  const a=await seeded();
  const i=at(a,'One two three'),p=rows(a)[i];
  const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  // (A mismatched closing tag is refused by a browser's parser, which flags it; this test's parser is lenient, and the paragraph that comes out is
  // always the parse tree written out again, so it is well formed either way.)
  for(const bad of ['not xml',`<w:tbl ${W}/>`,`<w:p ${W}/><w:p ${W}/>`,'']){
    await assert.rejects(a.replaceParagraphs({start:i,count:1,guard:p.text,xmls:[bad]}),/could not be made/,bad);
  }
  assert.equal(rows(a)[i].text,'One two three');
});

test('a table is never touched: its cells carry no XML of their own, and the table is changed with the tools',async()=>{
  const a=demo();
  await a.insert(ctx=>formTable({columns:[{name:'Goal',weight:3}],rows:[['Plan']],ctx,caption:'Goals'})+para('OSGCueBody',run('After')),{where:'start',spare:false});
  const ps=rows(a);
  const t=ps.findIndex(p=>/^<w:tbl[\s>]/.test(p.xml));
  assert.ok(t>=0);
  await assert.rejects(a.replaceParagraphs({start:t,count:1,guard:ps[t].text,xmls:[ps.at(-1).xml]}),/table/);
  await assert.rejects(a.replaceParagraphs({start:t+1,count:1,guard:ps[t+1].text,xmls:[ps.at(-1).xml]}),/table/,'a cell');
});

test('Enter, Backspace and a new line: one paragraph becomes two, two become one, none becomes one',async()=>{
  const a=await seeded();
  const i=at(a,'One two three'),p=rows(a)[i];
  const cut=splitWords(p.xml,p.style,7,io);
  await a.replaceParagraphs({start:i,count:1,guard:p.text,xmls:[cut.first,cut.second],select:i+1});
  assert.deepEqual(rows(a).slice(i,i+3).map(r=>r.text),['One two',' three','Four five']);
  assert.equal(a.cursor(),i+1);
  const [x,y]=[rows(a)[i],rows(a)[i+1]];
  const joined=joinWords(x,y,io);
  await a.replaceParagraphs({start:i,count:2,guard:x.text,xmls:[joined.xml],select:i});
  assert.deepEqual(rows(a).slice(i,i+2).map(r=>r.text),['One two three','Four five']);
  await a.replaceParagraphs({start:i+1,count:0,guard:'',xmls:['<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'],select:i+1});
  assert.equal(rows(a)[i+1].text,'');
  assert.equal(rows(a).length>=4,true);
});

test('typing in one paragraph is one step of Undo; a different paragraph, or another kind of change, starts a new one',async()=>{
  const a=await seeded();
  const i=at(a,'One two three'),j=at(a,'Four five');
  const type=async(index,text,group)=>{const p=rows(a)[index];await a.replaceParagraphs({start:index,count:1,guard:p.text,xmls:[setWords(p.xml,p.style,text,io)],select:index,group});};
  const base=a.historyState().undo;
  for(const t of ['O','On','One','One t','One tw'])await type(i,t,`type:${i}`);
  assert.equal(a.historyState().undo,base+1,'five keystrokes, one step');
  await type(j,'Four','type:'+j);
  assert.equal(a.historyState().undo,base+2,'another paragraph is another step');
  await type(j,'Four f','type:'+j);
  assert.equal(a.historyState().undo,base+2);
  a.undo();
  assert.equal(rows(a)[j].text,'Four five','undoing the second paragraph\'s typing');
  assert.equal(rows(a)[i].text,'One tw');
  a.undo();
  assert.equal(rows(a)[i].text,'One two three','undoing all of the first paragraph\'s typing at once');
});

test('typing in a group stops being one step after a pause, and a change by the tools in between ends it',async()=>{
  const a=await seeded();
  const i=at(a,'One two three');
  const type=async(text,group)=>{const p=rows(a)[i];await a.replaceParagraphs({start:i,count:1,guard:p.text,xmls:[setWords(p.xml,p.style,text,io)],select:i,group});};
  const real=Date.now;
  try{
    let now=1_000_000;Date.now=()=>now;
    const base=a.historyState().undo;
    await type('A','g');now+=500;await type('AB','g');
    assert.equal(a.historyState().undo,base+1);
    now+=5000;await type('ABC','g');
    assert.equal(a.historyState().undo,base+2,'a pause starts a new step');
    now+=100;
    await a.insert(()=>para('OSGCueBody',run('From the tools')),{where:'after',after:i,spare:false});
    now+=100;
    await type('ABCD','g');
    assert.equal(a.historyState().undo,base+4,'a change by the tools ends the group');
  }finally{Date.now=real;}
});

test('Redo brings back what Undo took away; a new change ends what could be redone; nothing to undo or redo says so',async()=>{
  const a=await seeded();
  const i=at(a,'One two three'),p=rows(a)[i];
  assert.equal(a.redo(),false);
  await a.replaceParagraphs({start:i,count:1,guard:p.text,xmls:[setWords(p.xml,p.style,'Changed',io)],select:i});
  a.undo();
  assert.deepEqual(a.historyState(),{undo:1,redo:1});
  assert.equal(a.redo(),true);
  assert.equal(rows(a)[i].text,'Changed');
  assert.equal(a.cursor(),i);
  a.undo();
  await a.replaceParagraphs({start:i,count:1,guard:p.text,xmls:[setWords(p.xml,p.style,'Different',io)],select:i});
  assert.equal(a.historyState().redo,0,'a new change ends Redo');
  assert.equal(a.redo(),false);
  a.undo();a.undo();
  assert.equal(a.undo(),false,'nothing more to undo');
});

test('the history is capped, and a project that is opened starts with none',async()=>{
  const a=await seeded();
  const i=at(a,'One two three');
  for(let n=0;n<350;n++){const p=rows(a)[i];await a.replaceParagraphs({start:i,count:1,guard:p.text,xmls:[setWords(p.xml,p.style,'v'+n,io)],select:i});}
  assert.equal(a.historyState().undo,300);
  const state=a.exportState();
  const b=demo();
  b.importState(JSON.parse(JSON.stringify(state)));
  assert.deepEqual(b.historyState(),{undo:0,redo:0});
});

test('with the program\'s saving: editing, Undo and Redo are changes that save and tell the page; moving the cursor is not',async()=>{
  const saved=[],changes=[];
  const guide=createProjectAdapter(demo(),{read:async()=>null,write:async t=>{saved.push(t);},delayMs:0,soon:fn=>fn(),cancel:()=>{},onChange:()=>changes.push(1)});
  await guide.ready;
  await guide.insert(()=>para('OSGCueBody',run('Hello')),{where:'start',spare:false});
  await guide.flush();
  const base=changes.length,savedBase=saved.length;
  guide.place(0);guide.fireSelection();guide.historyState();
  assert.equal(changes.length,base,'the cursor and the history\'s state are not changes');
  const p=guide.exportState().paragraphs[0];
  await guide.replaceParagraphs({start:0,count:1,guard:p.text,xmls:[setWords(p.xml,p.style,'Hello there',io)],select:0});
  assert.equal(changes.length,base+1);
  guide.undo();guide.redo();
  assert.equal(changes.length,base+3);
  await guide.flush();
  assert.ok(saved.length>savedBase&&saved.at(-1).includes('Hello there'));
});
