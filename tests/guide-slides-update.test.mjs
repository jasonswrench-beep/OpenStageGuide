// Update from Slides, in the pane: a presentation imported, then changed in a presentation program, then compared with the guide and brought in.
// The real pane module and the real page ids, an in-memory document, and a stand-in for the file window. Not native Word.
import test from 'node:test';
import assert from 'node:assert/strict';
import {TINY} from './helpers/guide-package.mjs';
import {started,addModule,addLesson,guideRows,click} from './helpers/guide-harness.mjs';
import {readMemory,recordFor,slideBlocksIn,notesPrint,fingerprint} from '../src/slides-sync.mjs';

const PNG=Buffer.from(TINY,'base64');
// Each picture is the tiny PNG with its own byte after it, so every slide has its own fingerprint.
const picture=(n,variant=n)=>({name:`Slide${n}.PNG`,bytes:new Uint8Array([...PNG,variant]),type:'png',mime:'image/png',width:1,height:1,shrunk:false});
const slide=(order,sourceId,title,{notes='',clicks=0}={})=>({order,sourceId,title,text:title,notes,clicks,hidden:false,titleGuessed:false});
const choose=(a,deck,files=[])=>a.adapter.setPicker(async opts=>{assert.equal(opts.kind,'slides');return {deck,files};});

// the first version of the presentation: three slides, permanent numbers 256 to 258
const V1=()=>({name:'AI deck.pptx',slides:[slide(1,256,'Intro',{notes:'Welcome.'}),slide(2,257,'Topic',{notes:'Explain it.',clicks:1}),slide(3,258,'Wrap up',{notes:'Thank them.'})]});
const P1=()=>[picture(1),picture(2),picture(3)];

async function imported(){
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  choose(a,V1(),P1());
  await click(a,'guide-slides-choose');
  for(const n of [1,2,3]){const box=a.n('guide-slides-list').children[n-1].children[0];box.checked=true;box.onchange();}
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  return a;
}
const rows=a=>guideRows(a);
const slideRows=a=>rows(a).filter(r=>/Slide\s+Slide \d/.test(r)).map(r=>r.replace(/^OSG Cue( Facilitator)?: .*?Slide\s+/,''));
const memory=async a=>readMemory(await a.adapter.slideMemory());
// The checkboxes are made by the pane, so they are found through the list rather than by a fixed id.
const boxes=a=>a.n('guide-upd-list').children.flatMap(g=>g.children.filter(l=>l.children?.[0]?.type==='checkbox').map(l=>l.children[0]));
const box=(a,key,id)=>boxes(a).find(b=>b.id===`guide-upd-${key}-${id}`);
async function update(a,deck,files){choose(a,deck,files);await click(a,'guide-upd-choose');}

test('importing slides leaves the guide a record of where they came from: permanent numbers, lines, notes and pictures',async()=>{
  const a=await imported();
  const m=await memory(a),r=recordFor(m,'AI deck.pptx');
  assert.deepEqual(r.seen,[256,257,258]);
  assert.deepEqual(r.slides.map(s=>[s.id,s.label]),[[256,'Slide 1: Intro'],[257,'Slide 2: Topic'],[258,'Slide 3: Wrap up']]);
  assert.equal(r.slides[0].ns,notesPrint('Welcome.'));
  assert.equal(r.slides[1].pic,fingerprint(picture(2).bytes));
  assert.deepEqual(slideBlocksIn(await a.adapter.paragraphs()).map(b=>b.label),['Slide 1: Intro','Slide 2: Topic','Slide 3: Wrap up']);
});

test('an unchanged presentation lists nothing and changes nothing',async()=>{
  const a=await imported(),before=rows(a);
  await update(a,V1(),P1());
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/Nothing differs\. 3 slides unchanged, left alone\./);
  assert.equal(a.n('guide-upd-fieldset').hidden,true);
  assert.deepEqual(rows(a),before);
});

test('a renamed slide, a changed picture and a new slide: the list, then the update in place',async()=>{
  const a=await imported();
  // in a presentation program: slide 2 renamed and given a second click, a new slide put between 2 and 3, and the Wrap up slide's picture redrawn
  const v2={name:'AI deck.pptx',slides:[slide(1,256,'Intro',{notes:'Welcome.'}),slide(2,257,'Topic, revised',{notes:'Explain it.',clicks:2}),slide(3,300,'A new idea',{notes:'Brand new.'}),slide(4,258,'Wrap up',{notes:'Thank them.'})]};
  const files=[picture(1),picture(2),picture(3,33),picture(4,99)].map((p,i)=>({...p,name:`Slide${i+1}.PNG`}));
  await update(a,v2,files);
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-upd-fieldset').hidden,false);
  const items=a.n('guide-upd-list').children.map(g=>g.children[0].textContent);
  assert.deepEqual(items,['Slide 2: Topic: changed','Slide 3: A new idea: not in the guide','Slide 3: Wrap up: changed']);
  assert.equal(box(a,'s2','title').checked,true);assert.equal(box(a,'s2','clicks').checked,true);
  assert.equal(box(a,'n3','add').checked,true,'the guide remembers the presentation, so a new slide is ticked');
  assert.equal(box(a,'s4','number').checked,true);assert.equal(box(a,'s4','picture').checked,true);
  assert.match(a.n('guide-upd-go').textContent,/Update the guide \(3 changes\)/);
  await click(a,'guide-upd-go');
  assert.equal(a.err(),'');
  assert.deepEqual(slideRows(a).map(r=>r.split('\n')[0]),['Slide 1: Intro','Slide 2: Topic, revised','Slide 3: A new idea','Slide 4: Wrap up'],'in place, with the new slide after the slide before it');
  const blocks=slideBlocksIn(await a.adapter.paragraphs());
  assert.deepEqual(blocks.map(b=>[b.label,b.clicks,b.hasPicture]),[['Slide 1: Intro',0,true],['Slide 2: Topic, revised',2,true],['Slide 3: A new idea',0,true],['Slide 4: Wrap up',0,true]]);
  assert.equal(blocks[0].notes,'Welcome.');assert.equal(blocks[3].notes,'Thank them.');
  assert.match(a.n('guide-status').textContent,/3 changes made\./);
  assert.equal(a.n('guide-upd-fieldset').hidden,true);
  // what the guide remembers now matches, so the same presentation lists nothing
  const r=recordFor(await memory(a),'AI deck.pptx');
  assert.deepEqual(r.slides.map(s=>[s.id,s.label]),[[256,'Slide 1: Intro'],[257,'Slide 2: Topic, revised'],[258,'Slide 4: Wrap up'],[300,'Slide 3: A new idea']]);
  await update(a,v2,files);
  assert.match(a.n('guide-status').textContent,/Nothing differs\. 4 slides unchanged/);
});

test('notes rewritten in Word are kept unless ticked; ticking them takes the presentation\'s',async()=>{
  const a=await imported();
  // the author rewrites slide 1's notes in Word
  const snap=a.adapter.snapshot().map(r=>({text:r.text.replace('Facilitator notes: Welcome.','Facilitator notes: My own welcome.'),style:r.style}));
  a.adapter.seed(snap);await a.p.refresh();
  assert.equal(slideBlocksIn(await a.adapter.paragraphs())[0].notes,'My own welcome.');
  const v2=V1();v2.slides[0].notes='Welcome, new words.';v2.slides[2].notes='Thank them warmly.';
  await update(a,v2,P1());
  assert.equal(a.err(),'');
  assert.equal(box(a,'s1','notes').checked,false,'rewritten notes are kept unless ticked');
  assert.equal(box(a,'s3','notes').checked,true,'notes still as imported follow the presentation');
  await click(a,'guide-upd-go');
  const blocks=slideBlocksIn(await a.adapter.paragraphs());
  assert.equal(blocks[0].notes,'My own welcome.');assert.equal(blocks[2].notes,'Thank them warmly.');
  // and again: the author's notes are still theirs, still unticked
  await update(a,v2,P1());
  assert.equal(box(a,'s1','notes').checked,false);
  box(a,'s1','notes').checked=true;box(a,'s1','notes').onchange();
  await click(a,'guide-upd-go');
  assert.equal(slideBlocksIn(await a.adapter.paragraphs())[0].notes,'Welcome, new words.');
});

test('unticking a change leaves that part of the slide as it was; Tick none and Tick all do what they say',async()=>{
  const a=await imported();
  const v2=V1();v2.slides[1].title='Topic renamed';v2.slides[1].clicks=3;
  await update(a,v2,P1());
  box(a,'s2','clicks').checked=false;box(a,'s2','clicks').onchange();
  await click(a,'guide-upd-go');
  const b=slideBlocksIn(await a.adapter.paragraphs())[1];
  assert.deepEqual([b.title,b.clicks],['Topic renamed',1],'the title followed the presentation, the clicks stayed');
  // tick none, then nothing can be done
  await update(a,v2,P1());
  await click(a,'guide-upd-none');
  assert.equal(box(a,'s2','clicks').checked,false);
  await click(a,'guide-upd-go');
  assert.match(a.err(),/Nothing is ticked, so nothing was changed\./);
  await click(a,'guide-upd-all');
  assert.equal(box(a,'s2','clicks').checked,true);
});

test('without the PDF a slide that has a picture cannot be rebuilt, and says so, and nothing is done to it',async()=>{
  const a=await imported();
  const v2=V1();v2.slides[1].title='Topic renamed';
  await update(a,v2,[]);
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/No PDF was chosen, so a slide that has a picture cannot be rebuilt/);
  assert.match(a.n('guide-upd-list').children[0].children[1].textContent,/Choose the PDF of the presentation too/);
  assert.equal(box(a,'s2','title').disabled,true);
  const before=rows(a);
  await click(a,'guide-upd-go');
  assert.match(a.err(),/Nothing is ticked/);
  assert.deepEqual(rows(a),before);
});

test('a slide gone from the presentation is offered for removal, unticked; ticking it takes the block out',async()=>{
  const a=await imported();
  const v2={name:'AI deck.pptx',slides:[V1().slides[0],V1().slides[2]].map((s,i)=>({...s,order:i+1}))};
  await update(a,v2,[picture(1),picture(3)].map((p,i)=>({...p,name:`Slide${i+1}.PNG`})));
  assert.equal(a.err(),'');
  const items=a.n('guide-upd-list').children.map(g=>g.children[0].textContent);
  assert.ok(items.includes('Slide 2: Topic: no longer in the presentation'),items.join(' | '));
  const remove=boxes(a).find(x=>x.id.endsWith('-remove'));
  assert.equal(remove.checked,false);
  // Wrap up moved up to slide 2 (a renumbering, ticked); the removal is the author's to tick
  remove.checked=true;remove.onchange();
  await click(a,'guide-upd-go');
  assert.equal(a.err(),'');
  assert.deepEqual(slideBlocksIn(await a.adapter.paragraphs()).map(b=>b.label),['Slide 1: Intro','Slide 2: Wrap up']);
});

test('a guide with no slides says to import first; a document changed since the list was made is not touched',async()=>{
  const a=await started();
  await update(a,V1(),P1());
  assert.match(a.err(),/There are no slides from a presentation in this guide yet\. Use Import slides first\./);
  const b=await imported();
  const v2=V1();v2.slides[1].title='Topic renamed';
  await update(b,v2,P1());
  // the author deletes the very block that was to be rebuilt, in Word, before pressing the button
  const list=await b.adapter.paragraphs(),second=slideBlocksIn(list)[1];
  await b.adapter.deleteRange({start:second.span.start,end:second.span.end,guard:second.firstText});
  const before=rows(b);
  await click(b,'guide-upd-go');
  assert.match(b.err(),/Your guide changed since the list was made/);
  assert.deepEqual(rows(b),before);
});

test('a copy with another name is a different presentation until the author says it is a new version of the first',async()=>{
  const a=await imported();
  const v2=V1();v2.name='AI deck v2.pptx';v2.slides[1].title='Topic renamed';
  await update(a,v2,P1());
  assert.equal(a.n('guide-upd-which-box').hidden,false);
  assert.equal(a.n('guide-upd-which').value,'','not assumed to be the same presentation');
  // matched by title instead: slide 2's new title matches nothing, so it is matched by its place
  assert.ok(a.n('guide-upd-list').children.some(g=>/Slide 2: Topic: changed/.test(g.children[0].textContent)));
  a.n('guide-upd-which').value='ai deck';await a.n('guide-upd-which').onchange();
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-upd-which').value,'ai deck');
  assert.equal(box(a,'s2','title').checked,true);
  await click(a,'guide-upd-go');
  assert.equal(slideBlocksIn(await a.adapter.paragraphs())[1].title,'Topic renamed');
  const m=await memory(a);
  assert.ok(recordFor(m,'AI deck v2.pptx'),'the new name is remembered from now on');
});

test('slides imported before the guide kept a record are still found, by title, and nothing is offered as new without asking',async()=>{
  const a=await imported();
  await a.adapter.setSlideMemory(null); // as if the guide were made by an earlier version
  const v2=V1();v2.slides[1].title='Topic';v2.slides[1].clicks=2;v2.slides.push(slide(4,259,'One more'));
  await update(a,v2,P1());
  assert.equal(a.err(),'');
  assert.equal(box(a,'s2','clicks').checked,true,'matched by its title');
  assert.equal(box(a,'n4','add').checked,false,'cannot tell whether it is new or left out, so it is not ticked');
});

test('an update that stops partway says how far it got',async()=>{
  const a=await imported();
  const v2=V1();v2.slides[0].title='Intro two';v2.slides[1].title='Topic two';
  await update(a,v2,P1());
  const real=a.adapter.insert.bind(a.adapter);let calls=0;
  a.adapter.insert=async(...args)=>{if(++calls===2)throw new Error('Word did not answer.');return real(...args);};
  await click(a,'guide-upd-go');
  assert.match(a.err(),/1 of 2 changes were made before this stopped: Word did not answer\. To take back what was done, choose Edit, then Undo, several times\./);
  assert.equal(a.n('guide-upd-fieldset').hidden,true,'the list is dropped: it is no longer true');
});
