// Update from Slides: the comparison of a changed presentation with the slides already in the guide. The plan is pure, so it is tested
// on plain data here, and once on slide blocks read back from a guide made the way the pane makes them.
import test from 'node:test';
import assert from 'node:assert/strict';
import {fingerprint,normNotes,notesPrint,deckKey,readMemory,writeMemory,recordFor,remember,slideBlocksIn,labelOf,planSlideUpdate,planSummary,operations} from '../src/slides-sync.mjs';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {cueBlock} from '../src/guide-ooxml.mjs';
import {slideBlock} from '../src/guide-slides.mjs';

const slide=(order,title,{id=255+order,clicks=0,notes=''}={})=>({order,sourceId:id,title,clicks,notes});
const block=(order,title,{index=order*10,clicks=0,notes='',hasPicture=true,audience='facilitator'}={})=>({index,span:{start:index,end:index+2},firstText:`Slide\t${labelOf({order,title})}`,label:labelOf({order,title}),order,title,clicks,notes,audience,hasPicture});
const entry=(order,title,{id=255+order,notes='',pic='p'+id}={})=>({id,label:labelOf({order,title}),ns:notesPrint(notes),pic});
const record=(entries,seen)=>({name:'deck.pptx',key:'deck',seen:seen??entries.map(e=>e.id),slides:entries});
// the fingerprint of each slide's new picture: it follows the slide's permanent number, so a slide that moved keeps its picture
const pics=deck=>new Map(deck.slides.map(s=>[s.order,'p'+s.sourceId]));
const find=(plan,key)=>plan.items.find(i=>i.key===key);

test('fingerprints are the same for the same text or bytes, and different for a change',()=>{
  assert.equal(fingerprint('abc'),fingerprint('abc'));
  assert.notEqual(fingerprint('abc'),fingerprint('abd'));
  assert.equal(fingerprint(new Uint8Array([1,2,3])),fingerprint(new Uint8Array([1,2,3])));
  assert.notEqual(fingerprint(new Uint8Array([1,2,3])),fingerprint(new Uint8Array([1,2,4])));
  assert.equal(notesPrint('  one\n\n two  '),notesPrint('one\ntwo'),'blank lines and spaces do not count');
  assert.equal(normNotes('a\r\n\r\nb  c'),'a\nb c');
});

test('a presentation is known by its name without .pptx or a copy number',()=>{
  assert.equal(deckKey('AI Deck.pptx'),'ai deck');
  assert.equal(deckKey('AI Deck (2).PPTX'),'ai deck');
  assert.equal(deckKey('ai deck - Copy.pptx'),'ai deck - copy');
});

test('the record is read back, junk is ignored, and remembering adds to what is there',()=>{
  assert.deepEqual(readMemory('not json'),{v:1,decks:[]});
  assert.deepEqual(readMemory('{"v":2}'),{v:1,decks:[]});
  let m=remember({v:1,decks:[]},{name:'A.pptx',seen:[256,257,258],entries:[{id:256,label:'Slide 1: A',ns:'x',pic:'y'}]});
  m=remember(m,{name:'a (2).pptx',seen:[259],entries:[{id:257,label:'Slide 2: B',ns:'',pic:''}]});
  const r=recordFor(m,'A.pptx');
  assert.deepEqual(r.seen,[256,257,258,259]);
  assert.deepEqual(r.slides.map(s=>s.id),[256,257],'importing more slides of the same presentation adds to its record');
  assert.equal(m.decks.length,1);
  assert.deepEqual(readMemory(writeMemory(m)),m);
  const many=Array.from({length:12},(_,i)=>i).reduce((acc,i)=>remember(acc,{name:`d${i}.pptx`,seen:[],entries:[]}),{v:1,decks:[]});
  assert.equal(many.decks.length,8,'the record keeps the eight most recent presentations');
  assert.equal(many.decks[0].key,'d11');
});

test('nothing differs: every slide matched, nothing offered',()=>{
  const blocks=[block(1,'Intro'),block(2,'Topic',{clicks:2,notes:'say hi'})];
  const deck={slides:[slide(1,'Intro'),slide(2,'Topic',{clicks:2,notes:'say hi'})]};
  const plan=planSlideUpdate({blocks,deck,pictures:pics(deck),record:record([entry(1,'Intro'),entry(2,'Topic',{notes:'say hi'})])});
  assert.deepEqual(plan.items,[]);assert.equal(plan.unchanged,2);
  assert.match(planSummary(plan),/^Nothing differs\. 2 slides unchanged, left alone\.$/);
});

test('a slide that moved and was renamed is still found by its permanent number, and what changed is listed',()=>{
  const blocks=[block(1,'Intro'),block(2,'Old title',{clicks:1}),block(3,'Wrap up')];
  // in the presentation now: Intro, then "Wrap up" (permanent number 258) moved up to 2, then the renamed slide (257) moved down to 3
  const deck={slides:[slide(1,'Intro'),slide(2,'Wrap up',{id:258}),slide(3,'New title',{id:257,clicks:3})]};
  const plan=planSlideUpdate({blocks,deck,pictures:pics(deck),record:record([entry(1,'Intro'),entry(2,'Old title'),entry(3,'Wrap up',{id:258})])});
  const renamed=plan.items.find(i=>i.slide.sourceId===257);
  assert.equal(renamed.how,'memory');
  assert.deepEqual(renamed.changes.map(c=>c.id),['title','number','clicks']);
  assert.ok(renamed.changes.find(c=>c.id==='title').text.includes('"Old title" to "New title"'));
  assert.ok(renamed.changes.find(c=>c.id==='number').text.includes('slide 2 is now slide 3'));
  assert.ok(renamed.changes.find(c=>c.id==='clicks').text.includes('1 click to 3 clicks'));
  assert.ok(renamed.changes.every(c=>c.on));
  const wrap=plan.items.find(i=>i.slide.sourceId===258);
  assert.deepEqual(wrap.changes.map(c=>c.id),['number'],'only its number changed: it moved');
  assert.equal(plan.unchanged,1);
});

test('speaker notes the author rewrote in Word are kept unless ticked; notes still as imported follow the presentation',()=>{
  const deck={slides:[slide(1,'A',{notes:'new deck notes'}),slide(2,'B',{notes:'new B notes'})]};
  const blocks=[block(1,'A',{notes:'my own words'}),block(2,'B',{notes:'old B notes'})];
  const plan=planSlideUpdate({blocks,deck,record:record([entry(1,'A',{notes:'old A notes'}),entry(2,'B',{notes:'old B notes'})])});
  const a=find(plan,'s1').changes.find(c=>c.id==='notes'),b=find(plan,'s2').changes.find(c=>c.id==='notes');
  assert.equal(a.on,false);assert.match(a.text,/You rewrote them in Word, so yours are kept unless you tick this/);
  assert.equal(b.on,true);assert.match(b.text,/changed in the presentation/);
  // with nothing remembered, no one's notes are overwritten unasked
  const none=planSlideUpdate({blocks,deck});
  assert.equal(find(none,'s1').changes.find(c=>c.id==='notes').on,false);assert.equal(find(none,'s2').changes.find(c=>c.id==='notes').on,false);
});

test('pictures: a changed fingerprint is a change, an unchanged one is not, and one that cannot be compared is offered unticked',()=>{
  const deck={slides:[slide(1,'A'),slide(2,'B'),slide(3,'C')]};
  const blocks=[block(1,'A'),block(2,'B'),block(3,'C',{hasPicture:false})];
  const plan=planSlideUpdate({blocks,deck,pictures:new Map([[1,'p256'],[2,'NEW'],[3,'p258']]),record:record([entry(1,'A'),entry(2,'B'),{...entry(3,'C'),pic:''}])});
  assert.equal(find(plan,'s1'),undefined,'same fingerprint: left alone');
  assert.deepEqual(find(plan,'s2').changes.map(c=>[c.id,c.on]),[['picture',true]]);
  assert.deepEqual(find(plan,'s3').changes.map(c=>[c.id,c.on]),[['picture',true]],'a text-only block gets its picture');
  const d1={slides:[slide(1,'A')]},unsure=planSlideUpdate({blocks:[block(1,'A')],deck:d1,pictures:pics(d1),record:null});
  assert.deepEqual(find(unsure,'s1').changes.map(c=>[c.id,c.on,!!c.unsure]),[['picture',false,true]]);
});

test('without the PDF, a slide with a picture cannot be rebuilt, and says so; one without a picture can',()=>{
  const deck={slides:[slide(1,'A',{clicks:2}),slide(2,'B',{clicks:2})]};
  const blocks=[block(1,'A'),block(2,'B',{hasPicture:false})];
  const plan=planSlideUpdate({blocks,deck,pictures:null,record:record([entry(1,'A'),entry(2,'B')])});
  assert.match(find(plan,'s1').blocked,/Choose the PDF of the presentation too/);
  assert.equal(find(plan,'s2').blocked,undefined);
  const ops=operations(plan);
  assert.deepEqual(ops.map(o=>o.key),['s2'],'only the slide that can be rebuilt is touched');
});

test('new slides are offered ticked when the guide remembers the presentation; slides left out earlier are not offered again',()=>{
  const deck={slides:[slide(1,'A'),slide(2,'Brand new',{id:300}),slide(3,'B',{id:258}),slide(4,'Left out',{id:301})]};
  const blocks=[block(1,'A'),block(2,'B')];
  const plan=planSlideUpdate({blocks,deck,record:record([entry(1,'A'),entry(2,'B',{id:258})],[256,258,301])});
  const added=plan.items.filter(i=>i.kind==='added');
  assert.deepEqual(added.map(i=>i.slide.title),['Brand new']);
  assert.equal(added[0].changes[0].on,true);
  assert.equal(plan.leftOut,1);
  assert.match(planSummary(plan),/1 slide you left out of the guide earlier was not offered again/);
  // nothing remembered: they are offered, but not ticked
  const none=planSlideUpdate({blocks,deck});
  assert.deepEqual(none.items.filter(i=>i.kind==='added').map(i=>[i.slide.title,i.changes[0].on]),[['Brand new',false],['Left out',false]]);
});

test('a slide gone from the presentation is offered for removal, unticked, only when it is known to have come from it',()=>{
  const blocks=[block(1,'A'),block(2,'Gone'),block(3,'From another deck')];
  const deck={slides:[slide(1,'A')]};
  const withRecord=planSlideUpdate({blocks,deck,record:record([entry(1,'A'),entry(2,'Gone')])});
  assert.deepEqual(withRecord.items.filter(i=>i.kind==='removed').map(i=>[i.label,i.changes[0].on]),[['Slide 2: Gone',false]],'the block from another deck is not listed');
  const without=planSlideUpdate({blocks,deck});
  assert.deepEqual(without.items.filter(i=>i.kind==='removed').map(i=>i.label),['Slide 2: Gone','Slide 3: From another deck'],'with no record, every block matched to nothing is listed, unticked');
});

test('with no record, slides are matched by title, then by place when the title changed',()=>{
  const blocks=[block(1,'Intro'),block(2,'Old name'),block(3,'Moved topic')];
  const deck={slides:[slide(1,'Intro',{id:0}),slide(2,'New name'),slide(3,'Other'),slide(4,'Moved topic')].map(s=>({...s,sourceId:undefined}))};
  const plan=planSlideUpdate({blocks,deck});
  const by=(order)=>plan.pairs.find(p=>p.slide.order===order);
  assert.equal(by(1).how,'title');assert.equal(by(4).how,'title');assert.equal(by(2).how,'number');
  assert.equal(by(3),undefined,'no block for "Other"');
  assert.deepEqual(find(plan,'s4').changes.map(c=>c.id),['number']);
  assert.deepEqual(find(plan,'s2').changes.map(c=>c.id),['title']);
});

test('the operations are done from the bottom of the guide up, an insert before a replace at the same place, and a new slide goes after the one before it',()=>{
  const blocks=[block(1,'A',{index:10}),block(2,'B',{index:20})];
  const deck={slides:[slide(1,'A',{clicks:1}),slide(2,'New',{id:300}),slide(3,'B',{clicks:2,id:257})]};
  const plan=planSlideUpdate({blocks,deck,pictures:pics(deck),record:record([entry(1,'A'),entry(2,'B',{id:257})],[256,257])});
  const ops=operations(plan),keys=ops.map(o=>`${o.type}:${o.key}`);
  // B (guide slide 2, now slide 3) is below A, so it is done first; the new slide 2 goes after A's block, before A is rebuilt
  assert.equal(keys[0],'replace:s3');
  const insert=ops.find(o=>o.type==='insert');
  assert.equal(insert.after,blocks[0].span.end);assert.deepEqual(insert.slides.map(s=>s.title),['New']);
  assert.ok(keys.indexOf(`insert:${insert.key}`)<keys.indexOf('replace:s1'),'the insert after A is done before A is rebuilt');
  assert.ok(keys.indexOf('replace:s3')<keys.indexOf(`insert:${insert.key}`),'what is lower in the guide is done first');
  // unticking removes the operation
  for(const c of find(plan,'s1').changes)c.on=false;
  assert.ok(!operations(plan).some(o=>o.key==='s1'));
});

test('a rebuilt slide takes the new values only for what is ticked and keeps the rest, including who it is shown to',()=>{
  const blocks=[block(2,'Old',{index:20,clicks:1,notes:'mine',audience:'everyone'})];
  const deck={slides:[slide(2,'New',{clicks:4,notes:'deck'})]};
  const plan=planSlideUpdate({blocks,deck,pictures:pics(deck),record:record([entry(2,'Old',{notes:'older'})])});
  const it=find(plan,'s2');
  assert.equal(it.changes.find(c=>c.id==='notes').on,false);
  it.changes.find(c=>c.id==='clicks').on=false;
  const [op]=operations(plan);
  assert.deepEqual(op.slide,{order:2,title:'New',clicks:1,notes:'mine',text:''},'a block with no slide text of its own is rebuilt without any');
  assert.equal(op.share,true);assert.equal(op.wantsPicture,true);
  it.changes.find(c=>c.id==='notes').on=true;
  assert.equal(operations(plan)[0].slide.notes,'deck');
});

test('removal and a first-slide insert have usable positions',()=>{
  const blocks=[block(2,'Gone',{index:20})];
  const deck={slides:[slide(1,'First',{id:300}),slide(2,'Other',{id:301})]};
  const plan=planSlideUpdate({blocks,deck,record:record([entry(2,'Gone',{id:257})],[257])});
  for(const it of plan.items)for(const c of it.changes)c.on=true;
  const ops=operations(plan);
  assert.ok(ops.some(o=>o.type==='delete'&&o.span.start===20));
  const insert=ops.find(o=>o.type==='insert');
  assert.equal(insert.after,null,'with no slide of the presentation in the guide, the pane decides where (the cursor)');
  assert.deepEqual(insert.slides.map(s=>s.title),['First','Other']);
});

// ---- blocks read back from a document made the way the pane makes it
const quirks={};
async function guideWith(blocks){
  const adapter=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer,quirks});
  adapter.seed([{text:'Guide',style:'OSG Guide Title'},{text:'Module 1: M',style:'OSG Module'}]);
  await adapter.insert(ctx=>blocks.map(b=>cueBlock(b,ctx)).join(''),{where:'after',after:1,select:'last'});
  return adapter;
}
test('slide blocks are read back from the document: number, title, clicks, notes, audience',async()=>{
  const adapter=await guideWith([slideBlock({order:1,title:'Intro to AI',clicks:2,notes:'Say hello.\nAsk who has used it.'},{share:true}),{cue:'say',text:'Not a slide',audience:'facilitator',notes:'',minutes:0},slideBlock({order:2,title:'',clicks:0,notes:''},{})]);
  const blocks=slideBlocksIn(await adapter.paragraphs());
  assert.equal(blocks.length,2);
  assert.deepEqual([blocks[0].order,blocks[0].title,blocks[0].clicks,blocks[0].notes,blocks[0].audience,blocks[0].label],[1,'Intro to AI',2,'Say hello.\nAsk who has used it.','everyone','Slide 1: Intro to AI']);
  assert.deepEqual([blocks[1].order,blocks[1].title,blocks[1].clicks,blocks[1].audience,blocks[1].label],[2,'',0,'facilitator','Slide 2']);
  assert.ok(blocks[0].span.start<blocks[0].span.end||blocks[0].span.start===blocks[0].span.end);
  assert.ok(blocks[1].index>blocks[0].span.end,'the block between them is not a slide');
});

// ---- a slide's own text, kept under its picture (author, 2026-11-07)
import {slideTextLines} from '../src/guide-slides.mjs';
import {entriesAfter} from '../src/slides-sync.mjs';

test('the words on a slide, apart from its title, are what goes under its picture',()=>{
  assert.equal(slideTextLines({title:'Why we care',text:'Why we care\nThree reasons\n  It saves   time \n\nIt helps'}),'Three reasons\nIt saves time\nIt helps');
  assert.equal(slideTextLines({title:'Only a title',text:'Only a title'}),'');
  assert.equal(slideTextLines({title:'',text:'Words with no title'}),'Words with no title');
  assert.equal(slideTextLines({title:'A',text:'a\nB'}),'B','the title is dropped whatever its capitals');
  assert.equal(slideTextLines({title:'T'}),'');
});

test('a slide block carries its text and a block read back keeps it apart from the speaker notes',async()=>{
  const sb=slideBlock({order:3,title:'Why',text:'Why\nFirst point\nSecond point',notes:'Say hello.',clicks:0});
  assert.equal(sb.slideText,'First point\nSecond point');assert.equal(sb.notes,'Say hello.');
  const a=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  await a.insert(ctx=>cueBlock(sb,ctx),{where:'start'});
  const blocks=slideBlocksIn(await a.paragraphs());
  assert.equal(blocks.length,1);
  assert.equal(blocks[0].slideText,'First point\nSecond point');
  assert.equal(blocks[0].notes,'Say hello.','the notes are not mixed up with the slide text');
});

test('slide text that changed in the presentation is offered, and kept when the author rewrote it in Word; guides made before it are left alone',()=>{
  const text=t=>({text:`S\n${t}`});
  const old=(order)=>({...block(order,'S'),slideText:'Body old'});
  // untouched since the import: follows the presentation, ticked
  let plan=planSlideUpdate({blocks:[old(2)],deck:{slides:[{...slide(2,'S'),...text('Body new')}]},pictures:pics({slides:[slide(2,'S')]}),record:record([{...entry(2,'S'),tx:notesPrint('Body old')}])});
  let c=find(plan,'s2').changes.find(x=>x.id==='text');
  assert.equal(c.on,true);assert.equal(c.text,'Slide text changed in the presentation');
  // rewritten in Word: kept unless ticked
  plan=planSlideUpdate({blocks:[{...block(2,'S'),slideText:'Mine'}],deck:{slides:[{...slide(2,'S'),...text('Body new')}]},pictures:pics({slides:[slide(2,'S')]}),record:record([{...entry(2,'S'),tx:notesPrint('Body old')}])});
  c=find(plan,'s2').changes.find(x=>x.id==='text');
  assert.equal(c.on,false);assert.match(c.text,/You changed it in Word, so yours is kept unless you tick this/);
  // a guide made before slide text was kept: no text in the block and none remembered, so the slide is not reported as changed
  plan=planSlideUpdate({blocks:[block(2,'S')],deck:{slides:[{...slide(2,'S'),...text('Some body')}]},pictures:pics({slides:[slide(2,'S')]}),record:record([entry(2,'S')])});
  assert.equal(plan.items.length,0);assert.equal(plan.unchanged,1);
  // the same words: nothing to report
  plan=planSlideUpdate({blocks:[old(2)],deck:{slides:[{...slide(2,'S'),...text('Body old')}]},pictures:pics({slides:[slide(2,'S')]}),record:record([{...entry(2,'S'),tx:notesPrint('Body old')}])});
  assert.equal(plan.items.length,0);
});

test('a rebuilt slide takes the new slide text only when ticked, and the guide remembers what it holds',()=>{
  const deck={slides:[{...slide(2,'S',{notes:'n'}),text:'S\nBody new'}]};
  const plan=planSlideUpdate({blocks:[{...block(2,'S',{notes:'n'}),slideText:'Body old'}],deck,pictures:pics(deck),record:record([{...entry(2,'S',{notes:'n'}),tx:notesPrint('Body old')}])});
  const it=find(plan,'s2');
  it.changes.find(x=>x.id==='text').on=false;
  assert.equal(operations(plan)[0]?.slide.text??'none','none','nothing else changed, so nothing is rebuilt');
  it.changes.find(x=>x.id==='text').on=true;
  const ops=operations(plan);
  assert.equal(ops[0].slide.text,'Body new');
  const {entries}=entriesAfter(plan,ops,{deck,pictures:pics(deck)});
  assert.equal(entries[0].tx,notesPrint('Body new'));
  // inserted slides remember theirs too
  const plan2=planSlideUpdate({blocks:[],deck:{slides:[{...slide(5,'Fresh'),text:'Fresh\nNew words'}]},record:record([])});
  for(const i of plan2.items)for(const x of i.changes)x.on=true;
  const ops2=operations(plan2);
  assert.equal(entriesAfter(plan2,ops2,{deck:{slides:[{...slide(5,'Fresh'),text:'Fresh\nNew words'}]}}).entries[0].tx,notesPrint('New words'));
});

test('the record keeps what it knows of each slide\'s text, and an older record reads with none',()=>{
  const m=remember({v:1,decks:[]},{name:'A.pptx',seen:[256],entries:[{id:256,label:'Slide 1: A',ns:'x',pic:'y',tx:'zz'}]});
  assert.equal(recordFor(readMemory(writeMemory(m)),'A.pptx').slides[0].tx,'zz');
  const olderText=JSON.stringify({v:1,decks:[{name:'B',key:'b',seen:[1],slides:[{id:1,label:'Slide 1: B',ns:'',pic:''}]}]});
  assert.equal(readMemory(olderText).decks[0].slides[0].tx,'');
});
