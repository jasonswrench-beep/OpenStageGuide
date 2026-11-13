import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {createWordAdapter} from '../src/guide-word.mjs';
import {startGuide,moduleHeading,lessonHeading,timeBlock,cueBlock} from '../src/guide-ooxml.mjs';
import {itemBlock} from '../src/guide-assess-ooxml.mjs';
import {readStyles,textOf} from '../src/guide-check.mjs';
import {parse,partXml} from './helpers/guide-package.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// A hand-made stand-in for the few Office.js objects the adapter touches. It can imitate what the author's Word did
// in the first native test (quirk 'join-last': the last paragraph an insert brings in loses its style, and the
// style is not imported, and its words join the next paragraph) and leave strays. It is not Word: it proves the
// adapter reaches Word at the right moment, asks for the right things and checks the result, nothing more.
function fakeWord({rows=[],quirk='none'}={}){
  const doc={paras:rows.map(r=>({text:r.text,style:r.style||'Normal'})),styles:new Set(['Normal','Heading 1','Heading 2','heading 1','heading 2']),selected:null,errors:[],cursor:rows.length-1,packages:[],custom:new Map()};
  const names=pkgStyles=>readStyles(parse(pkgStyles));
  // A table's cells come back as paragraphs, one after another, each marked with its table, as Word's paragraphs do.
  const fromPackage=pkg=>{
    const styles=names(partXml(pkg,'/word/styles.xml')),d=parse(partXml(pkg,'/word/document.xml'));
    const entry=(p,table)=>{
      const id=Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val');
      return {text:textOf(p),style:id?styles.get(id).name:'Normal',table};
    };
    return Array.from(d.getElementsByTagNameNS(W,'body')[0].childNodes).filter(n=>n.nodeType===1&&(n.localName==='p'||n.localName==='tbl')).flatMap(n=>{
      if(n.localName==='p')return [entry(n,null)];
      const id=(doc.tableCount=(doc.tableCount||0)+1);
      return Array.from(n.getElementsByTagNameNS(W,'p')).map(p=>entry(p,id));
    });
  };
  const tableOf=p=>p.table?{
    isNullObject:false,load(){},
    getRange(){return {insertOoxml(pkg,where){assert.equal(where,'After');insertAt(pkg,Math.max(...doc.paras.map((q,i)=>q.table===p.table?i:-1))+1);}};},
    delete(){doc.tablesDeleted=(doc.tablesDeleted||0)+1;doc.paras=doc.paras.filter(q=>q.table!==p.table);}
  }:{isNullObject:true,load(){}};
  const insertAt=(pkg,at)=>{
    doc.packages.push(pkg);
    let added=fromPackage(pkg);
    const last=added.at(-1);
    if(quirk==='join-last'||quirk==='join-last-and-stray'){
      added=added.slice(0,-1);
      // Styles used only by the lost paragraph are not imported.
      if(last.text.trim()){
        if(doc.paras[at])doc.paras[at].text=last.text+doc.paras[at].text;
        else added.push({text:last.text,style:'Normal'});
      }
    }
    // 'notes-lose-style': the words arrive in their own paragraph but Word does not keep the notes style.
    if(quirk==='notes-lose-style')for(const p of added)if(p.style==='OSG Facilitator Notes')p.style='Normal';
    for(const p of added)doc.styles.add(p.style);
    if(quirk==='stray-before')added.unshift({text:'',style:'Normal'});
    if(quirk==='join-last-and-stray')added.push({text:'',style:'Normal'});
    doc.paras.splice(at,0,...added);
  };
  const facade=p=>({
    get text(){return p.text;},get style(){return p.style;},
    set style(v){if(!doc.styles.has(v))doc.errors.push(new Error('ItemNotFound'));else p.style=v;},
    get parentTableOrNullObject(){return tableOf(p);},
    getRange(kind){return {p,kind,insertOoxml(pkg,where){assert.equal(where,'After');insertAt(pkg,doc.paras.indexOf(p)+1);},compareLocationWith(o){return {value:o.p===p?'Equal':'Other'};}};},
    select(){doc.selected=doc.paras.indexOf(p);},
    delete(){const i=doc.paras.indexOf(p);if(i===doc.paras.length-1)doc.errors.push(new Error('Cannot delete the last paragraph'));else doc.paras.splice(i,1);}
  });
  const collection=()=>({items:doc.paras.map(facade),load(){}});
  const ctx={
    document:{
      body:{get paragraphs(){return collection();},insertOoxml(pkg,where){assert.equal(where,'Start');insertAt(pkg,0);},insertParagraph(text,where){assert.equal(where,'End');const p={text,style:doc.paras.at(-1)?.style||'Normal'};doc.paras.push(p);return facade(p);}},
      getSelection:()=>({paragraphs:{getFirst:()=>facade(doc.paras[doc.cursor])}}),
      properties:{title:'Old',load(){},customProperties:{
        getItemOrNullObject:key=>({isNullObject:!doc.custom.has(key),value:doc.custom.get(key),load(){}}),
        add:(key,value)=>{doc.custom.set(key,value);} // Word's add also replaces an existing property
      }}
    },
    sync:async()=>{if(doc.errors.length)throw doc.errors.shift();}
  };
  return {api:{run:fn=>fn(ctx)},doc};
}
const response={ok:true,arrayBuffer:async()=>new Uint8Array([137,80,78,71]).buffer};
const adapterFor=word=>createWordAdapter({Word:word.api,fetchFn:async()=>response,DOMParserClass:DOMParser});
const reachWord=()=>{const had=Object.getOwnPropertyDescriptor(globalThis,'Word');return()=>{if(had)Object.defineProperty(globalThis,'Word',had);else delete globalThis.Word;};};
const rows=word=>word.doc.paras.map(p=>`${p.style}: ${p.text.trim()}`);

test('the adapter finds Word when it is used, not when it is created',async()=>{
  const restore=reachWord();
  try{
    delete globalThis.Word;
    const adapter=createWordAdapter({fetchFn:async()=>response,DOMParserClass:DOMParser});
    await assert.rejects(adapter.paragraphs(),/Word is not ready yet\. Close the OpenStage Guide pane/);
    globalThis.Word={run:fn=>fn({document:{body:{paragraphs:{items:[{text:'Hello',style:'Normal',styleBuiltIn:'Normal'}],load(){}}}},sync:async()=>{}})};
    assert.deepEqual(await adapter.paragraphs(),[{text:'Hello',style:'Normal',builtIn:'Normal'}]);
  }finally{restore();}
});

test('a missing Word gives one plain message, never a property error',async()=>{
  const restore=reachWord();
  try{
    delete globalThis.Word;
    const adapter=createWordAdapter({fetchFn:async()=>response,DOMParserClass:DOMParser});
    for(const call of [()=>adapter.paragraphs(),()=>adapter.title(),()=>adapter.cursorIndex(),()=>adapter.go(0),()=>adapter.applyEdits([]),()=>adapter.reviewParts()])
      await assert.rejects(call(),e=>/Word is not ready yet/.test(e.message)&&!/undefined/.test(e.message));
    assert.equal(await adapter.refreshFields(),false,'refreshing the contents table is optional and never throws');
    assert.equal(await adapter.keepPaneOpen(),false);
    assert.equal(adapter.onSelectionChanged(()=>{}),false,'hooking events never throws, it just reports that it could not without Office');
  }finally{restore();}
});

test('your own cue names are kept as one custom document property and read back',async()=>{
  const word=fakeWord(),adapter=adapterFor(word);
  assert.deepEqual(await adapter.customCues(),[],'a new document has none');
  await adapter.setCustomCues(['Debrief','','Wrap up']);
  assert.equal(word.doc.custom.get('osg.customCues'),'["Debrief","","Wrap up"]');
  assert.deepEqual(await adapter.customCues(),['Debrief','','Wrap up']);
  await adapter.setCustomCues(['Recap','','']);
  assert.deepEqual(await adapter.customCues(),['Recap','',''],'saving again replaces the names');
  word.doc.custom.set('osg.customCues','not json');
  assert.deepEqual(await adapter.customCues(),[],'a damaged property reads as no cues, never an error');
  word.doc.custom.set('osg.customCues','{"a":1}');
  assert.deepEqual(await adapter.customCues(),[],'a property that is not a list reads as no cues');
});

test('the cursor is the paragraph whose start matches the start of the selection',async()=>{
  const word=fakeWord({rows:[{text:'a'},{text:'b'},{text:'c'}]});
  word.doc.cursor=1;
  assert.equal(await adapterFor(word).cursorIndex(),1);
});

test('go moves to the paragraph, finds the same words if the document has moved on, and refuses if they are gone',async()=>{
  const word=fakeWord({rows:[{text:'one'},{text:'two'},{text:'three'}]}),adapter=adapterFor(word);
  await adapter.go(1,'two');
  assert.equal(word.doc.selected,1);
  word.doc.paras.unshift({text:'new first line',style:'Normal'});
  await adapter.go(1,'two');
  assert.equal(word.doc.selected,2,'it followed the words, not the old number');
  await assert.rejects(adapter.go(0,'words that are not there'),/has changed since this list was made/);
  await assert.rejects(adapter.go(9),/no longer there/);
});

const lesson=(c)=>lessonHeading(1,1,'Lesson')+timeBlock({minutes:10,cumulative:0},c);

test('a block inserted after a paragraph keeps every style, even the last paragraph, because the package ends with a spare one',async()=>{
  for(const quirk of ['none','join-last','join-last-and-stray','stray-before']){
    const word=fakeWord({rows:[{text:'Module 1: One',style:'OSG Module'},{text:'Transition Say how this leads on.',style:'OSG Cue Facilitator'}],quirk});
    word.doc.styles.add('OSG Module');word.doc.styles.add('OSG Cue Facilitator');
    const adapter=adapterFor(word);
    await adapter.insert(c=>lesson(c),{where:'after',after:0,select:'last'});
    assert.deepEqual(rows(word),[
      'OSG Module: Module 1: One','OSG Lesson: Lesson 1.1: Lesson','OSG Time: Time: 10 minutes. Cumulative time: 0 minutes.','OSG Cue Facilitator: Transition Say how this leads on.',
      'Normal: '
    ],`quirk ${quirk}`);
    assert.equal(word.doc.selected,2,`cursor on the time block (${quirk})`);
    assert.ok(word.doc.styles.has('OSG Time'),'the time block\'s style was imported');
  }
});

test('notes, which come last in a block, keep their shaded style',async()=>{
  const word=fakeWord({rows:[{text:'Lesson',style:'OSG Lesson'}],quirk:'join-last'});
  word.doc.styles.add('OSG Lesson');
  await adapterFor(word).insert(c=>cueBlock({cue:'ask',text:'Why?',notes:'Listen.',minutes:4},c),{where:'after',after:0});
  assert.deepEqual(rows(word),['OSG Lesson: Lesson','OSG Cue: Ask Why? (4 min)','OSG Facilitator Notes: Facilitator notes: Listen.','Normal: ']);
});

test('when an insert ends the document in a block\'s own paragraph (shaded notes), a plain empty paragraph is left below it so the cursor can get out',async()=>{
  // The author's recording of 2026-11-08: slides put in at the end of the guide left the last notes paragraph as the last paragraph of the
  // document, and Enter there only made more notes lines.
  const word=fakeWord({rows:[{text:'Lesson',style:'OSG Lesson'}]});
  word.doc.styles.add('OSG Lesson');
  await adapterFor(word).insert(c=>cueBlock({cue:'slide',text:'Slide 1: Intro',notes:'Say hello.',audience:'facilitator'},c),{where:'after',after:0});
  const end=word.doc.paras.at(-1);
  assert.equal(end.style,'Normal');assert.equal(end.text,'');
  assert.match(word.doc.paras.at(-2).style,/^OSG Facilitator Notes$/,'the notes keep their style');
  // Something plain of the author's below the block is left as it is: nothing is added.
  const kept=fakeWord({rows:[{text:'Lesson',style:'OSG Lesson'},{text:'The author wrote this.',style:'Normal'}]});
  kept.doc.styles.add('OSG Lesson');
  const before=kept.doc.paras.length;
  await adapterFor(kept).insert(c=>cueBlock({cue:'ask',text:'Why?',minutes:4},c),{where:'after',after:0});
  assert.equal(kept.doc.paras.at(-1).text,'The author wrote this.');
  assert.ok(kept.doc.paras.length>before);
});

test('without the spare paragraph the same Word joins the time block to the Transition and drops the styles (what the author saw)',async()=>{
  const word=fakeWord({rows:[{text:'Module 1: One',style:'OSG Module'},{text:'Transition Say how this leads on.',style:'OSG Cue Facilitator'}],quirk:'join-last'});
  word.doc.styles.add('OSG Module');word.doc.styles.add('OSG Cue Facilitator');
  await assert.rejects(adapterFor(word).insert(c=>lesson(c),{where:'after',after:0,spare:false}),/joined part of this block to another paragraph/);
  assert.match(word.doc.paras[2].text,/Time: 10 minutes\. Cumulative time: 0 minutes\..*Transition/s);
  assert.ok(!word.doc.styles.has('OSG Time'),'and the style never arrived');
});

test('when Word keeps the words but drops a style that exists in the document, the style is put back',async()=>{
  const word=fakeWord({rows:[{text:'Lesson',style:'OSG Lesson'},{text:'Later',style:'Normal'}],quirk:'notes-lose-style'});
  word.doc.styles.add('OSG Lesson');word.doc.styles.add('OSG Facilitator Notes');
  await adapterFor(word).insert(c=>cueBlock({cue:'ask',text:'Why?',notes:'Listen.'},c),{where:'after',after:0});
  assert.deepEqual(rows(word),['OSG Lesson: Lesson','OSG Cue: Ask Why?','OSG Facilitator Notes: Facilitator notes: Listen.','Normal: Later']);
});

test('when the style is missing from the document too, it says so plainly and tells the author how to take the block back',async()=>{
  const word=fakeWord({rows:[{text:'Lesson',style:'OSG Lesson'}],quirk:'notes-lose-style'});
  word.doc.styles.add('OSG Lesson');
  await assert.rejects(adapterFor(word).insert(c=>cueBlock({cue:'ask',text:'Why?',notes:'Listen.'},c),{where:'after',after:0}),
    e=>/could not finish tidying this block or apply the guide's styles/.test(e.message)&&/press Ctrl\+Z/.test(e.message)&&/\[Diagnostic: inserted 2 paragraphs/.test(e.message));
});

test('an empty paragraph Word leaves before the block is removed, and the cursor still lands on the block',async()=>{
  const word=fakeWord({rows:[{text:'Lesson',style:'OSG Lesson'},{text:'Later text',style:'Normal'}],quirk:'stray-before'});
  word.doc.styles.add('OSG Lesson');
  await adapterFor(word).insert(c=>cueBlock({cue:'ask',text:'Why?'},c),{where:'after',after:0,select:'last'});
  assert.deepEqual(rows(word),['OSG Lesson: Lesson','OSG Cue: Ask Why?','Normal: Later text']);
  assert.equal(word.doc.selected,1);
});

test('the author\'s own empty paragraph is never removed',async()=>{
  const word=fakeWord({rows:[{text:'Lesson',style:'OSG Lesson'},{text:'',style:'Normal'},{text:'Later text',style:'Normal'}]});
  word.doc.styles.add('OSG Lesson');
  await adapterFor(word).insert(c=>cueBlock({cue:'ask',text:'Why?'},c),{where:'after',after:0});
  assert.deepEqual(rows(word),['OSG Lesson: Lesson','OSG Cue: Ask Why?','Normal: ','Normal: Later text']);
});

test('replacing a block inserts the new one after it and removes the old one, and refuses if the old one has changed',async()=>{
  const word=fakeWord({rows:[{text:'Lesson',style:'OSG Lesson'},{text:' Ask Old?',style:'OSG Cue'},{text:'Facilitator notes: old',style:'OSG Facilitator Notes'},{text:'Later',style:'Normal'}]});
  ['OSG Lesson','OSG Cue','OSG Facilitator Notes'].forEach(s=>word.doc.styles.add(s));
  const adapter=adapterFor(word);
  await assert.rejects(adapter.insert(c=>cueBlock({cue:'ask',text:'New?'},c),{where:'after',after:2,remove:{start:1,end:2,guard:'Ask Something else'}}),/has changed since you chose it/);
  assert.equal(word.doc.paras.length,4,'nothing was inserted');
  await adapter.insert(c=>cueBlock({cue:'ask',text:'New?'},c),{where:'after',after:2,select:'last',remove:{start:1,end:2,guard:' Ask Old?'}});
  assert.deepEqual(rows(word),['OSG Lesson: Lesson','OSG Cue: Ask New?','Normal: Later']);
  assert.equal(word.doc.selected,1);
});

// Word only opens a pane with its document when the manifest has a ShowTaskpane action named
// Office.AutoShowTaskpaneWithDocument. Asking without one made Word say "We can't find the task pane to open" on every reopen
// (the author's test, 2026-10-31). Giving the Open button that action (manifest 1.1.0.0) made it do nothing in Word (2026-11-01), so
// manifest 1.1.0.1 has no such action and the pane does not ask by default; asking stays possible for a manifest that has one.
const AUTO='Office.AutoShowTaskpaneWithDocument';
const settingsBag=(saved='succeeded')=>({store:new Map(),saves:0,get(k){return this.store.get(k);},set(k,v){this.store.set(k,v);},remove(k){this.store.delete(k);},saveAsync(cb){this.saves++;cb({status:saved});}});
const adapterWith=settings=>createWordAdapter({Word:{run:()=>{}},Office:{context:{document:{settings}}}});

test('by default the pane does not ask Word to reopen it, and takes out a request an earlier build left',async()=>{
  const clean=settingsBag();
  assert.equal(await adapterWith(clean).keepPaneOpen(),false);
  assert.equal(clean.get(AUTO),undefined);
  assert.equal(clean.saves,0,'and it does not even dirty the document');
  const left=settingsBag();left.set(AUTO,true);
  assert.equal(await adapterWith(left).keepPaneOpen(),false);
  assert.equal(left.get(AUTO),undefined,'the request 1.1.0.0 left in a document is taken out');
  assert.equal(left.saves,1);
});

test('given a manifest that can open the pane with its document, the pane asks once',async()=>{
  const settings=settingsBag(),adapter=adapterWith(settings);
  assert.equal(await adapter.keepPaneOpen({manifestSupportsIt:true}),true);
  assert.equal(settings.get(AUTO),true);
  assert.equal(settings.saves,1);
  assert.equal(await adapter.keepPaneOpen({manifestSupportsIt:true}),false,'already asked: nothing more to save');
  assert.equal(settings.saves,1);
  assert.equal(await adapterWith(settingsBag('failed')).keepPaneOpen({manifestSupportsIt:true}),false,'a failed save is not reported as done');
});

test('told the manifest cannot open it that way, the pane does not ask, and takes out a request an earlier build left, once',async()=>{
  const clean=settingsBag(),a=adapterWith(clean);
  assert.equal(await a.keepPaneOpen({manifestSupportsIt:false}),false);
  assert.equal(clean.get(AUTO),undefined);
  assert.equal(clean.saves,0,'and it does not even dirty the document');
  const left=settingsBag(),b=adapterWith(left);
  left.set(AUTO,true);
  assert.equal(await b.keepPaneOpen({manifestSupportsIt:false}),false);
  assert.equal(left.get(AUTO),undefined);
  assert.equal(left.saves,1);
  await b.keepPaneOpen({manifestSupportsIt:false});
  assert.equal(left.saves,1,'nothing more to save the second time');
});

test('Word\'s cursor-moved event reaches the pane through the adapter',()=>{
  const handlers=[],office={EventType:{DocumentSelectionChanged:'sel'},context:{document:{addHandlerAsync:(type,fn)=>handlers.push([type,fn])}}};
  let calls=0;
  assert.equal(createWordAdapter({Office:office}).onSelectionChanged(()=>calls++),true);
  assert.equal(handlers[0][0],'sel');
  handlers[0][1]();
  assert.equal(calls,1);
});

test('if Word refuses to hook the cursor event (a document still opening), it is tried again',async()=>{
  const attempts=[],office={EventType:{DocumentSelectionChanged:'sel'},context:{document:{addHandlerAsync:(type,fn,done)=>{attempts.push(fn);done({status:attempts.length<3?'failed':'succeeded'});}}}};
  let calls=0;
  createWordAdapter({Office:office}).onSelectionChanged(()=>calls++,{delayMs:1});
  await new Promise(resolve=>setTimeout(resolve,40));
  assert.equal(attempts.length,3,'two refusals, then it took');
  attempts.at(-1)();
  assert.equal(calls,1);
});

test('it gives up after a few tries, and never throws when Word has no events at all',async()=>{
  let tries=0;
  const refusing={EventType:{DocumentSelectionChanged:'sel'},context:{document:{addHandlerAsync:(t,f,done)=>{tries++;done({status:'failed'});}}}};
  createWordAdapter({Office:refusing}).onSelectionChanged(()=>{},{retries:2,delayMs:1});
  await new Promise(resolve=>setTimeout(resolve,40));
  assert.equal(tries,3);
  assert.equal(createWordAdapter({Office:{context:{document:{}}}}).onSelectionChanged(()=>{}),false);
});

test('an assessment item keeps every style, its key included, in a Word that drops the last paragraph\'s style',async()=>{
  const word=fakeWord({rows:[{text:'Level Two Assessment',style:'OSG Assessment'},{text:'Later',style:'Normal'}],quirk:'join-last'});
  word.doc.styles.add('OSG Assessment');
  await adapterFor(word).insert(c=>itemBlock({type:'multiple-choice',number:1,stem:'Q?',options:['a','b'],correct:[1],why:'W'},c),{where:'after',after:0});
  assert.deepEqual(rows(word),[
    'OSG Assessment: Level Two Assessment','OSG Question: Question 1. Q?','OSG Option: a','OSG Option: b','OSG Answer Line: Your answer:',
    'OSG Answer Key: Answer key (multiple choice): b.','OSG Answer Key: Why: W','Normal: Later'
  ]);
  assert.ok(word.doc.styles.has('OSG Answer Key'),'the key\'s style reached the document');
});

test('start of document insert puts the pages first and the standard pages keep their text',async()=>{
  const word=fakeWord({rows:[{text:'',style:'Normal'}],quirk:'join-last'});
  word.doc.styles.add('OSG Guide Title');
  await adapterFor(word).insert(c=>startGuide({title:'A guide',standard:true},c),{where:'start',select:'last'});
  assert.equal(word.doc.paras[0].style,'OSG Guide Title');
  assert.equal(word.doc.paras[0].text.trim(),'A guide');
  assert.ok(word.doc.paras.some(p=>p.text.includes('Training at a Glance')));
});

// ---- items with a table (matching, rating statements)
const RATING={type:'rating-statements',stem:'Trainer.',statements:['Prepared.','Clear.'],scale:'success'};
const tableStyles=word=>{for(const s of ['OSG Assessment','OSG Question','OSG Table Head','OSG Table Cell','OSG Rating Box','OSG Item Note','OSG Option','OSG Answer Key'])word.doc.styles.add(s);};

test('a rating item arrives as a table in order, with no stray empty paragraph, whichever way this Word misbehaves',async()=>{
  for(const quirk of ['none','join-last','join-last-and-stray','stray-before']){
    const word=fakeWord({rows:[{text:'Level One Assessment',style:'OSG Assessment'},{text:'Later',style:'Normal'}],quirk});
    tableStyles(word);
    await adapterFor(word).insert(c=>itemBlock(RATING,c),{where:'after',after:0,select:'last'});
    const got=rows(word);
    assert.equal(got[0],'OSG Assessment: Level One Assessment',quirk);
    assert.equal(got[1],'OSG Question: Trainer.',quirk);
    assert.deepEqual(got.slice(2,8),['OSG Table Head: Statement','OSG Table Head: Not at all','OSG Table Head: Slightly','OSG Table Head: Moderately','OSG Table Head: Mostly','OSG Table Head: Completely'],quirk);
    assert.equal(got.filter(r=>r.startsWith('OSG Rating Box')).length,10,quirk);
    assert.equal(got.at(-1),'Normal: Later',`the next paragraph is untouched (${quirk})`);
    assert.equal(got.filter(r=>r==='Normal:').length,0,`no stray empty paragraph (${quirk})`);
    assert.equal(word.doc.selected,got.length-2,`the cursor is on the last paragraph of the item (${quirk})`);
  }
});

test('replacing a rating item deletes its table as a table, once, and the new item takes its place',async()=>{
  const word=fakeWord({rows:[{text:'Level One Assessment',style:'OSG Assessment'},{text:'Later',style:'Normal'}],quirk:'join-last'});
  tableStyles(word);
  const adapter=adapterFor(word);
  await adapter.insert(c=>itemBlock(RATING,c),{where:'after',after:0});
  const end=word.doc.paras.findIndex(p=>p.text==='Later')-1;
  await adapter.insert(c=>itemBlock({...RATING,stem:'Facilitator.',statements:['Warm.']},c),{where:'after',after:end,select:'last',remove:{start:1,end,guard:'Trainer.'}});
  const got=rows(word);
  assert.equal(word.doc.tablesDeleted,1,'one table deleted, not a paragraph at a time');
  assert.deepEqual(got.filter(r=>r.startsWith('OSG Question')),['OSG Question: Facilitator.']);
  assert.equal(got.filter(r=>r.startsWith('OSG Table Cell')).length,1);
  assert.equal(got.at(-1),'Normal: Later');
  assert.equal(word.doc.selected,got.length-2,'the cursor is on the last paragraph of the new item');
});

test('an item added with the cursor inside a table goes after the whole table',async()=>{
  const word=fakeWord({rows:[{text:'Level One Assessment',style:'OSG Assessment'},{text:'Later',style:'Normal'}],quirk:'join-last'});
  tableStyles(word);
  const adapter=adapterFor(word);
  await adapter.insert(c=>itemBlock(RATING,c),{where:'after',after:0});
  const inside=word.doc.paras.findIndex(p=>p.text==='Prepared.');
  await adapter.insert(c=>itemBlock({type:'open-comment',stem:'Anything else?'},c),{where:'after',after:inside});
  const got=rows(word),q=got.indexOf('OSG Question: Anything else?'),lastBox=got.map((r,i)=>r.startsWith('OSG Rating Box')?i:-1).filter(i=>i>=0).at(-1);
  assert.equal(q,lastBox+1,'right after the last cell of the table');
});

// What the guide remembers of the presentations its slides came from (slides-sync.mjs) is one setting kept in the document.
test('what the guide remembers about its slides is kept in the document\'s settings, saved when it is written, and its absence is not an error',async()=>{
  const settings=settingsBag(),adapter=adapterWith(settings);
  assert.equal(await adapter.slideMemory(),null,'nothing remembered yet');
  assert.equal(await adapter.setSlideMemory('{"v":1,"decks":[]}'),true);
  assert.equal(settings.get('osg.slides.v1'),'{"v":1,"decks":[]}');
  assert.equal(settings.saves,1);
  assert.equal(await adapter.slideMemory(),'{"v":1,"decks":[]}');
  await assert.rejects(adapterWith(settingsBag('failed')).setSlideMemory('x'),/did not save what the guide remembers/);
  // no settings at all (a browser with no Word): reading gives nothing, writing says it could not
  const bare=createWordAdapter({Word:{run:()=>{}},Office:{context:{document:{}}}});
  assert.equal(await bare.slideMemory(),null);assert.equal(await bare.setSlideMemory('x'),false);
});
