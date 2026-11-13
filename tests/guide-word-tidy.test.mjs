// moveRange and deleteRange on the Word adapter, against a hand-made stand-in for the few Office.js objects they touch.
// The stand-in can imitate what the author's Word did in the first native test (the last paragraph of an insert loses its style and
// joins the next one) and leave stray empty paragraphs. It is not Word: it proves the adapter asks Word for the right things at the
// right moments and checks the result. Whether real Word keeps every style and picture through getOoxml and insertOoxml is NOT proven here.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {createWordAdapter,withSpareParagraph} from '../src/guide-word.mjs';
import {STYLE_NAMES,STYLE_IDS,stylesPart,NS} from '../src/guide-styles.mjs';
import {readStyles,textOf} from '../src/guide-check.mjs';
import {esc} from '../src/guide-ooxml.mjs';
import {parse,partXml} from './helpers/guide-package.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const idOf=name=>Object.entries(STYLE_NAMES).map(([k,n])=>n===name?STYLE_IDS[k]:null).find(Boolean)||(name==='Normal'?null:name.replace(/\s+/g,''));

function fakeWord({rows,quirk='none'}){
  const doc={paras:rows.map(r=>({text:r.text,style:r.style||'Normal',table:r.table||null})),selected:null,errors:[],ooxml:[],tablesDeleted:0,
    tableCount:Math.max(0,...rows.map(r=>r.table||0)),
    styles:new Set(['Normal','heading 1','heading 2',...Object.values(STYLE_NAMES)])};
  const fromPackage=pkg=>{
    const styles=readStyles(parse(partXml(pkg,'/word/styles.xml'))),d=parse(partXml(pkg,'/word/document.xml'));
    const entry=(p,table)=>{
      const id=Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val');
      return {text:textOf(p),style:id?(styles.get(id)?.name||id.replace(/^Heading(\d)$/,'heading $1')):'Normal',table};
    };
    return Array.from(d.getElementsByTagNameNS(W,'body')[0].childNodes).filter(n=>n.nodeType===1&&(n.localName==='p'||n.localName==='tbl')).flatMap(n=>{
      if(n.localName==='p')return [entry(n,null)];
      const t=(doc.tableCount=(doc.tableCount||0)+1);
      return Array.from(n.getElementsByTagNameNS(W,'p')).map(p=>entry(p,t));
    });
  };
  // What getOoxml gives for paragraphs a..b: a flat package of their words and styles, ending with the section properties.
  const ooxmlOf=(a,b)=>{
    const slice=doc.paras.slice(a,b+1);
    let body='',i=0;
    const p=q=>`<w:p>${q.style==='Normal'?'':`<w:pPr><w:pStyle w:val="${idOf(q.style)}"/></w:pPr>`}<w:r><w:t xml:space="preserve">${esc(q.text)}</w:t></w:r></w:p>`;
    while(i<slice.length){
      if(!slice[i].table){body+=p(slice[i]);i++;continue;}
      const t=slice[i].table;let cells='';
      while(i<slice.length&&slice[i].table===t){cells+=`<w:tc>${p(slice[i])}</w:tc>`;i++;}
      body+=`<w:tbl><w:tr>${cells}</w:tr></w:tbl>`;
    }
    const part=(name,type,inner)=>`<pkg:part pkg:name="${name}" pkg:contentType="${type}"><pkg:xmlData>${inner}</pkg:xmlData></pkg:part>`;
    return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">${part('/word/document.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',`<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`)}${part('/word/styles.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',stylesPart())}</pkg:package>`;
  };
  const insertAt=(pkg,at)=>{
    doc.ooxml.push(pkg);
    let added=fromPackage(pkg);
    if(quirk==='join-last'){
      const last=added.pop();
      if(last.text.trim()){if(doc.paras[at])doc.paras[at].text=last.text+doc.paras[at].text;else added.push({...last,style:'Normal'});}
    }
    if(quirk==='drop-middle')added.splice(1,1);
    if(quirk==='lose-styles')for(const q of added)if(q.style===STYLE_NAMES.cue)q.style='Normal';
    if(quirk==='stray-before')added.unshift({text:'',style:'Normal',table:null});
    if(quirk==='stray-after')added.push({text:'',style:'Normal',table:null});
    doc.paras.splice(at,0,...added);
  };
  const tableOf=q=>q.table?{
    isNullObject:false,load(){},
    getRange(){return {insertOoxml(pkg){insertAt(pkg,Math.max(...doc.paras.map((x,i)=>x.table===q.table?i:-1))+1);}};},
    delete(){doc.tablesDeleted++;doc.paras=doc.paras.filter(x=>x.table!==q.table);}
  }:{isNullObject:true,load(){}};
  const rangeOf=q=>({
    insertOoxml(pkg,where){assert.equal(where,'After');insertAt(pkg,doc.paras.indexOf(q)+1);},
    expandTo(other){const a=doc.paras.indexOf(q);return {getOoxml(){return {value:ooxmlOf(a,doc.paras.indexOf(other.q))};}};},
    q
  });
  const facade=q=>({
    get text(){return q.text;},get style(){return q.style;},
    set style(v){if(!doc.styles.has(v))doc.errors.push(new Error('ItemNotFound'));else q.style=v;},
    get parentTableOrNullObject(){return tableOf(q);},
    getRange(){return rangeOf(q);},
    select(){doc.selected=doc.paras.indexOf(q);},
    insertText(t){q.text=t;},
    delete(){const i=doc.paras.indexOf(q);if(i===doc.paras.length-1)doc.errors.push(new Error('Cannot delete the last paragraph'));else doc.paras.splice(i,1);}
  });
  const ctx={
    document:{body:{get paragraphs(){return {items:doc.paras.map(facade),load(){}};},insertOoxml(pkg,where){assert.equal(where,'Start');insertAt(pkg,0);}}},
    sync:async()=>{if(doc.errors.length)throw doc.errors.shift();}
  };
  return {api:{run:fn=>fn(ctx)},doc};
}
const adapterFor=word=>createWordAdapter({Word:word.api,fetchFn:async()=>({ok:true,arrayBuffer:async()=>new Uint8Array([1]).buffer}),DOMParserClass:DOMParser});
const shape=word=>word.doc.paras.map(p=>`${p.table?'[t] ':''}${p.style}: ${p.text}`);

const R=(style,text,table)=>({style,text,table});
const BASE=()=>[
  R('OSG Guide Title','Guide'),
  R('OSG Module','Module 1: One'),
  R('OSG Lesson','Lesson 1.1: Alpha'),R('OSG Time','Time: 4 minutes. Cumulative time: 0 minutes.'),R('OSG Cue',' Ask Alpha? (4 min)'),
  R('OSG Lesson','Lesson 1.2: Beta'),R('OSG Time','Time: 3 minutes. Cumulative time: 4 minutes.'),R('OSG Facilitator Notes','Facilitator notes: slow down'),
  R('OSG Cue Facilitator',' Transition On to module 2.'),
  R('OSG Module','Module 2: Two'),R('OSG Lesson','Lesson 2.1: Delta'),R('OSG Time','Time: 8 minutes. Cumulative time: 7 minutes.'),
  R('Normal','The end.')
];
const QUIRKS=['none','join-last','stray-after','stray-before'];

test('a lesson moves down past the next one, whatever Word does to the end of a pasted block',async()=>{
  for(const quirk of QUIRKS){
    const word=fakeWord({rows:BASE(),quirk}),adapter=adapterFor(word);
    const out=await adapter.moveRange({start:2,end:4,guard:'Lesson 1.1: Alpha',after:7});
    assert.deepEqual(shape(word),[
      'OSG Guide Title: Guide','OSG Module: Module 1: One',
      'OSG Lesson: Lesson 1.2: Beta','OSG Time: Time: 3 minutes. Cumulative time: 4 minutes.','OSG Facilitator Notes: Facilitator notes: slow down',
      'OSG Lesson: Lesson 1.1: Alpha','OSG Time: Time: 4 minutes. Cumulative time: 0 minutes.','OSG Cue:  Ask Alpha? (4 min)',
      'OSG Cue Facilitator:  Transition On to module 2.','OSG Module: Module 2: Two','OSG Lesson: Lesson 2.1: Delta','OSG Time: Time: 8 minutes. Cumulative time: 7 minutes.','Normal: The end.'
    ],`quirk ${quirk}`);
    assert.equal(out.at,5,'the moved block now starts here');
    assert.equal(word.doc.selected,5,`the cursor is on the moved heading (${quirk})`);
  }
});

test('a lesson moves up past the one above it',async()=>{
  for(const quirk of QUIRKS){
    const word=fakeWord({rows:BASE(),quirk}),adapter=adapterFor(word);
    const out=await adapter.moveRange({start:5,end:7,guard:'Lesson 1.2: Beta',after:1});
    assert.deepEqual(shape(word).slice(1,9),[
      'OSG Module: Module 1: One','OSG Lesson: Lesson 1.2: Beta','OSG Time: Time: 3 minutes. Cumulative time: 4 minutes.','OSG Facilitator Notes: Facilitator notes: slow down',
      'OSG Lesson: Lesson 1.1: Alpha','OSG Time: Time: 4 minutes. Cumulative time: 0 minutes.','OSG Cue:  Ask Alpha? (4 min)','OSG Cue Facilitator:  Transition On to module 2.'
    ],`quirk ${quirk}`);
    assert.equal(word.doc.paras.length,BASE().length,`nothing gained or lost (${quirk})`);
    assert.equal(out.at,2);assert.equal(word.doc.selected,2);
  }
});

test('a module moves with everything in it, including a table, and a block can go to the very top',async()=>{
  const rows=[
    R('OSG Guide Title','Guide'),
    R('OSG Module','Module 1: One'),R('OSG Table Head','Lesson',1),R('OSG Table Head','Minutes',1),R('OSG Table Cell','Alpha',1),R('OSG Table Cell','4',1),
    R('OSG Lesson','Lesson 1.1: Alpha'),R('OSG Time','Time: 4 minutes. Cumulative time: 0 minutes.'),
    R('OSG Module','Module 2: Two'),R('OSG Lesson','Lesson 2.1: Delta'),R('OSG Time','Time: 8 minutes. Cumulative time: 4 minutes.'),R('Normal','The end.')
  ];
  const word=fakeWord({rows,quirk:'join-last'}),adapter=adapterFor(word);
  await adapter.moveRange({start:1,end:7,guard:'Module 1: One',after:10});
  assert.deepEqual(shape(word),[
    'OSG Guide Title: Guide','OSG Module: Module 2: Two','OSG Lesson: Lesson 2.1: Delta','OSG Time: Time: 8 minutes. Cumulative time: 4 minutes.',
    'OSG Module: Module 1: One','[t] OSG Table Head: Lesson','[t] OSG Table Head: Minutes','[t] OSG Table Cell: Alpha','[t] OSG Table Cell: 4',
    'OSG Lesson: Lesson 1.1: Alpha','OSG Time: Time: 4 minutes. Cumulative time: 0 minutes.','Normal: The end.'
  ]);
  assert.equal(word.doc.tablesDeleted,1,'the old table was deleted as a table');
  // To the very top.
  const top=fakeWord({rows:BASE()});
  await adapterFor(top).moveRange({start:9,end:11,guard:'Module 2: Two',after:-1});
  assert.equal(shape(top)[0],'OSG Module: Module 2: Two');
  assert.equal(top.doc.paras.length,BASE().length);
  assert.equal(top.doc.selected,0);
});

test('a style Word dropped in the copy is put back',async()=>{
  const word=fakeWord({rows:BASE(),quirk:'lose-styles'});
  await adapterFor(word).moveRange({start:2,end:4,guard:'Lesson 1.1: Alpha',after:7});
  assert.ok(shape(word).includes('OSG Cue:  Ask Alpha? (4 min)'),'the cue style was restored');
});

test('if Word does not bring every paragraph across, nothing is deleted and the message says so',async()=>{
  const word=fakeWord({rows:BASE(),quirk:'drop-middle'});
  await assert.rejects(adapterFor(word).moveRange({start:2,end:4,guard:'Lesson 1.1: Alpha',after:7}),/did not bring every paragraph across, so nothing was moved/);
  for(const row of ['OSG Lesson: Lesson 1.1: Alpha','OSG Time: Time: 4 minutes. Cumulative time: 0 minutes.','OSG Cue:  Ask Alpha? (4 min)'])
    assert.ok(shape(word).includes(row),`${row} is still there`);
});

test('a block that has changed since it was chosen is refused before anything is touched',async()=>{
  const word=fakeWord({rows:BASE()}),before=shape(word);
  await assert.rejects(adapterFor(word).moveRange({start:2,end:4,guard:'Lesson 1.9: Someone else',after:7}),/has changed since you chose it/);
  await assert.rejects(adapterFor(word).deleteRange({start:2,end:4,guard:'Lesson 1.9: Someone else'}),/has changed since you chose it/);
  await assert.rejects(adapterFor(word).moveRange({start:2,end:4,guard:'Lesson 1.1: Alpha',after:3}),/cannot be moved into itself/);
  await assert.rejects(adapterFor(word).moveRange({start:2,end:4,guard:'Lesson 1.1: Alpha',after:99}),/no longer there/);
  assert.deepEqual(shape(word),before);
  assert.equal(word.doc.ooxml.length,0,'Word was never asked to insert anything');
});

test('a block that starts or ends in a table is not moved from here',async()=>{
  const rows=[R('OSG Guide Title','G'),R('OSG Module','Module 1: A'),R('OSG Table Head','h',1),R('OSG Table Cell','c',1),R('OSG Module','Module 2: B'),R('Normal','x')];
  const word=fakeWord({rows});
  await assert.rejects(adapterFor(word).moveRange({start:1,end:2,guard:'Module 1: A',after:4}),/starts or ends inside a table/);
  assert.equal(word.doc.ooxml.length,0);
});

test('deleteRange removes the block and selects what took its place',async()=>{
  const word=fakeWord({rows:BASE()});
  await adapterFor(word).deleteRange({start:2,end:4,guard:'Lesson 1.1: Alpha'});
  assert.deepEqual(shape(word).slice(1,4),['OSG Module: Module 1: One','OSG Lesson: Lesson 1.2: Beta','OSG Time: Time: 3 minutes. Cumulative time: 4 minutes.']);
  assert.equal(word.doc.paras.length,BASE().length-3);
  assert.equal(word.doc.selected,2);
});

test('deleteRange deletes a table in the block as a table, and empties the document\'s last paragraph instead of deleting it',async()=>{
  const rows=[R('OSG Guide Title','G'),R('OSG Module','Module 1: A'),R('OSG Table Head','h',1),R('OSG Table Cell','c',1),R('OSG Lesson','Lesson 1.1: L'),R('OSG Time','Time: 1 minute.'),R('Normal','last words')];
  const word=fakeWord({rows});
  await adapterFor(word).deleteRange({start:1,end:6,guard:'Module 1: A'});
  assert.equal(word.doc.tablesDeleted,1);
  assert.deepEqual(shape(word),['OSG Guide Title: G','Normal: '],'only the title and an empty last paragraph are left');
  assert.equal(word.doc.selected,1);
});

test('the spare paragraph goes just before the section properties of a package, or at the end of the body',()=>{
  const body='<w:p><w:r><w:t>a</w:t></w:r></w:p>';
  assert.equal(withSpareParagraph(`<w:body>${body}<w:sectPr><w:pgSz/></w:sectPr></w:body>`),`<w:body>${body}<w:p/><w:sectPr><w:pgSz/></w:sectPr></w:body>`);
  assert.equal(withSpareParagraph(`<w:body>${body}</w:body>`),`<w:body>${body}<w:p/></w:body>`);
  const inPara='<w:p><w:pPr><w:sectPr><w:pgSz/></w:sectPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>';
  assert.equal(withSpareParagraph(`<w:body>${inPara}</w:body>`),`<w:body>${inPara}<w:p/></w:body>`,'a section break inside a paragraph is left alone');
  assert.equal(withSpareParagraph('no body here'),'no body here');
});

// ---- applyEdits: the time block is rewritten without asking Word for an inline picture's range
// In the author's Word (2026-11-03) InlinePicture.getRange('After') is refused with InvalidArgument, which made every time update in
// Update guide fail. This stand-in refuses it the same way, so the adapter must not use it.
function fakeEdits(rows){
  const state={paras:rows.map(text=>({text})),pending:[],picturesAsked:0};
  const facade=p=>({
    get text(){return p.text;},
    get inlinePictures(){state.picturesAsked++;return {items:[{getRange(){state.pending.push(Object.assign(new Error('InvalidArgument'),{code:'InvalidArgument',debugInfo:{errorLocation:'InlinePicture.getRange'}}));return {};}}],load(){}};},
    search(q){
      const at=p.text.indexOf(q);
      return {items:at<0?[]:[{expandTo(){return {insertText(t,how){assert.equal(how,'Replace');if(p.fail)state.pending.push(Object.assign(new Error('Refused'),{code:'GeneralException'}));else p.text=p.text.slice(0,at)+t;}};}}],load(){}};
    },
    getRange(){return {getRange(){return {};}};},
    insertText(t,how){assert.equal(how,'Replace');p.text=t;}
  });
  const ctx={document:{body:{paragraphs:{get items(){return state.paras.map(facade);},load(){}}}},sync:async()=>{if(state.pending.length)throw state.pending.shift();}};
  return {api:{run:fn=>fn(ctx)},state};
}
const edits=word=>createWordAdapter({Word:word.api,DOMParserClass:DOMParser});

test('Update guide rewrites a time block after its icon, and never asks Word for the picture\'s range',async()=>{
  const word=fakeEdits(['Lesson 1.2: X','\u0001 Time: 20 minutes. Cumulative time: 0 minutes.','\u0001 Time: 5 minutes. Cumulative time: 20 minutes.']);
  await edits(word).applyEdits([
    {kind:'heading',index:0,text:'Lesson 1.1: X'},
    {kind:'time',index:1,minutes:7,cumulative:10},
    {kind:'time',index:2,minutes:1,cumulative:17}
  ]);
  assert.deepEqual(word.state.paras.map(p=>p.text),['Lesson 1.1: X','\u0001 Time: 7 minutes. Cumulative time: 10 minutes.','\u0001 Time: 1 minute. Cumulative time: 17 minutes.']);
  assert.equal(word.state.picturesAsked,0,'the inline picture is not touched');
});

test('a time block with no "Time:" in it is rewritten whole',async()=>{
  const word=fakeEdits(['Lesson 1.1: X','hand-typed words']);
  await edits(word).applyEdits([{kind:'time',index:1,minutes:3,cumulative:0}]);
  assert.equal(word.state.paras[1].text,'Time: 3 minutes. Cumulative time: 0 minutes.');
});

test('when Word refuses one change, the others are still made and the message says which one and where',async()=>{
  const word=fakeEdits(['Lesson 1.2: X','\u0001 Time: 20 minutes. Cumulative time: 0 minutes.','Lesson 1.3: Y']);
  word.state.paras[1].fail=true;
  await assert.rejects(edits(word).applyEdits([
    {kind:'heading',index:0,text:'Lesson 1.1: X'},{kind:'time',index:1,minutes:7,cumulative:10},{kind:'heading',index:2,text:'Lesson 1.2: Y'}
  ]),/Word would not make 1 of 3 changes to numbers and times, and made the rest\. To finish by hand: change the time line "Time: 20 minutes\. Cumulative time: 0 minutes\." to read "Time: 7 minutes\. Cumulative time: 10 minutes\." \(select the words after the small picture and type over them\)\. \(Word's code: GeneralException\.\)$/);
  assert.deepEqual([word.state.paras[0].text,word.state.paras[2].text],['Lesson 1.1: X','Lesson 1.2: Y']);
});
