import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {SLIDE_KINDS,SLIDE_LIMITS,checkDeck,slideFit,pptxBytes,pptxName} from '../src/pptx-export.mjs';
import {readPptx} from '../src/pptx.mjs';
globalThis.DOMParser=DOMParser;   // the repo's own reader (src/pptx.mjs) parses with the global one

const A='http://schemas.openxmlformats.org/drawingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG='http://schemas.openxmlformats.org/package/2006/relationships',CT='http://schemas.openxmlformats.org/package/2006/content-types',DC='http://purl.org/dc/elements/1.1/';
const NOW=new Date('2026-11-23T09:00:00Z');
const SAMPLE=Object.freeze({title:'Skill Practice Workshop',subtitle:'Practice and feedback',slides:[
  {kind:'opening',title:'Skill Practice Workshop',lines:['Leave with: a rehearsed response and one revision','Agenda: briefing, demonstration, rehearsal, feedback, exit check'],notes:['Welcome people and say how to take part.']},
  {kind:'activity',title:'Outcome and agenda briefing (09:00, 5 min)'},
  {kind:'activity',title:'Scenario rehearsal (09:17, 18 min)',lines:['In threes: performer, counterpart, observer','Produce: a rehearsed response and one revision'],notes:['3 min: assign roles','6 min: run the scenario and record evidence','Debrief: What changed when you repeated the moment?'],key:'A13:1'},
  {kind:'break',title:'Break (09:35, 10 min)',lines:['Back at 09:45']},
  {kind:'activity',title:'Feedback and revision (09:45, 7 min)',lines:['Observer: one thing that worked, one thing to try','Performer: say the revised response aloud'],notes:['Keep to time.']},
  {kind:'close',title:'Close (09:52, 8 min)',lines:['Write one insight and one next step']}]});
const copy=()=>structuredClone(SAMPLE);

// ---- reading a package back, by parsing it (never by searching its text) --------------------------------------------------------
function parse(xml,name='part'){
  const errors=[],doc=new DOMParser({errorHandler:{warning:m=>errors.push(m),error:m=>errors.push(m),fatalError:m=>errors.push(m)}}).parseFromString(xml,'text/xml');
  assert.deepEqual(errors,[],name+' is not well-formed');assert.equal(doc.getElementsByTagName('parsererror').length,0,name+' has a parse error');
  return doc;
}
const ns=(node,space,name)=>Array.from(node.getElementsByTagNameNS(space,name));
const attr=(el,name)=>el?.getAttribute(name)||'';
async function build(model=SAMPLE){const bytes=await pptxBytes(model,JSZip,{now:NOW}),zip=await JSZip.loadAsync(bytes);return {bytes,zip};}
const part=async(zip,name)=>parse(await zip.file(name).async('string'),name);
const parasOf=sp=>ns(sp,A,'p').map(p=>ns(p,A,'t').map(t=>t.textContent).join(''));
// The shapes of a slide or notes page in document order: its placeholder (type, idx), its paragraphs, and the element.
const shapesOf=doc=>ns(doc,P,'sp').map(sp=>{const ph=ns(sp,P,'ph')[0];return {type:attr(ph,'type'),idx:attr(ph,'idx'),isPlaceholder:!!ph,paras:parasOf(sp),sp};});
const slideShapes=async(zip,n)=>shapesOf(await part(zip,`ppt/slides/slide${n}.xml`));
const noteShapes=async(zip,n)=>shapesOf(await part(zip,`ppt/notesSlides/notesSlide${n}.xml`));
const relsOf=async(zip,name)=>{const dir=path.posix.dirname(name),base=path.posix.basename(name);return ns(await part(zip,`${dir}/_rels/${base}.rels`),PKG,'Relationship').map(r=>({id:attr(r,'Id'),type:attr(r,'Type').split('/').pop(),target:attr(r,'Target'),mode:attr(r,'TargetMode')}));};
const deckOf=(titles,extra={})=>({title:'Deck',slides:titles.map(title=>({kind:'activity',title,...extra}))});

test('the sample deck writes a package that is well-formed, complete and 16:9',async()=>{
  const {zip}=await build(),names=Object.keys(zip.files);
  assert.ok(names.every(n=>!zip.files[n].dir),'no folder entries in the zip');
  assert.equal(names[0],'[Content_Types].xml');
  const docs=new Map();
  for(const n of names)if(/\.(xml|rels)$/.test(n))docs.set(n,await part(zip,n));
  for(const n of ['_rels/.rels','docProps/core.xml','docProps/app.xml','ppt/presentation.xml','ppt/_rels/presentation.xml.rels','ppt/presProps.xml','ppt/viewProps.xml','ppt/tableStyles.xml','ppt/theme/theme1.xml','ppt/theme/theme2.xml','ppt/slideMasters/slideMaster1.xml','ppt/slideMasters/_rels/slideMaster1.xml.rels','ppt/slideLayouts/slideLayout1.xml','ppt/slideLayouts/slideLayout2.xml','ppt/notesMasters/notesMaster1.xml','ppt/notesMasters/_rels/notesMaster1.xml.rels'])assert.ok(docs.has(n),'missing '+n);
  // every relationship target exists, and is not external
  for(const [n,doc] of docs){
    if(!n.endsWith('.rels'))continue;
    const source=n==='_rels/.rels'?'':n.replace(/(^|\/)_rels\//,'$1').replace(/\.rels$/,''),dir=source?path.posix.dirname(source):'';
    for(const r of ns(doc,PKG,'Relationship')){
      assert.equal(attr(r,'TargetMode'),'',n+': no external targets');
      const target=path.posix.normalize(path.posix.join(dir,attr(r,'Target')));
      assert.ok(zip.file(target),`${n}: ${attr(r,'Target')} does not exist`);
    }
  }
  // every part has a content type, and every override names a part that exists
  const types=docs.get('[Content_Types].xml'),defaults=new Set(ns(types,CT,'Default').map(d=>attr(d,'Extension'))),overrides=new Set(ns(types,CT,'Override').map(o=>attr(o,'PartName').slice(1)));
  for(const n of names)if(n!=='[Content_Types].xml')assert.ok(overrides.has(n)||defaults.has(n.split('.').pop()),n+' has no content type');
  for(const o of overrides)assert.ok(zip.file(o),'content type for a missing part: '+o);
  // the deck itself
  const pres=docs.get('ppt/presentation.xml'),ids=ns(pres,P,'sldId'),size=ns(pres,P,'sldSz')[0];
  assert.equal(ids.length,SAMPLE.slides.length);assert.deepEqual(ids.map(i=>Number(attr(i,'id'))),[256,257,258,259,260,261]);
  assert.equal(attr(size,'cx'),'12192000');assert.equal(attr(size,'cy'),'6858000');
  assert.equal(ns(pres,P,'sldMasterId').length,1);assert.equal(ns(pres,P,'notesMasterId').length,1);assert.ok(ns(pres,P,'notesSz').length);
  const deckRels=await relsOf(zip,'ppt/presentation.xml');
  for(const [i,id] of ids.entries()){
    const rel=deckRels.find(r=>r.id===id.getAttributeNS(R,'id'));assert.equal(rel.type,'slide');assert.equal(rel.target,`slides/slide${i+1}.xml`);
    // every slide has a notes slide that points back at it
    const slideRels=await relsOf(zip,`ppt/slides/slide${i+1}.xml`),notes=slideRels.find(r=>r.type==='notesSlide');
    assert.equal(notes.target,`../notesSlides/notesSlide${i+1}.xml`);assert.equal(slideRels.filter(r=>r.type==='slideLayout').length,1);
    const back=await relsOf(zip,`ppt/notesSlides/notesSlide${i+1}.xml`);
    assert.deepEqual(back.map(r=>[r.type,r.target]).sort(),[['notesMaster','../notesMasters/notesMaster1.xml'],['slide',`../slides/slide${i+1}.xml`]]);
  }
  assert.equal(names.filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).length,6);assert.equal(names.filter(n=>/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n)).length,6);
  // two themes, so the notes master does not share the slide master's
  assert.equal((await relsOf(zip,'ppt/slideMasters/slideMaster1.xml')).find(r=>r.type==='theme').target,'../theme/theme1.xml');
  assert.equal((await relsOf(zip,'ppt/notesMasters/notesMaster1.xml')).find(r=>r.type==='theme').target,'../theme/theme2.xml');
  // the document properties
  const core=docs.get('docProps/core.xml');assert.equal(ns(core,DC,'title')[0].textContent,'Skill Practice Workshop');assert.equal(ns(core,DC,'language')[0].textContent,'en-US');assert.equal(ns(core,DC,'subject')[0].textContent,'Practice and feedback');
  assert.equal(ns(docs.get('docProps/app.xml'),'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties','Slides')[0].textContent,'6');
  assert.ok(!names.some(n=>/media|\.(png|jpe?g|gif|emf|wmf)$/i.test(n)),'no pictures in the package');
});

test('titles, body lines and notes are native text in the right slides, in reading order',async()=>{
  const {zip}=await build();
  for(const [i,s] of SAMPLE.slides.entries()){
    const doc=await part(zip,`ppt/slides/slide${i+1}.xml`),shapes=shapesOf(doc),[title,body]=shapes;
    assert.equal(shapes.length,s.lines?.length?2:1,`slide ${i+1}: a body shape only when there are lines`);
    assert.deepEqual([title.type,title.idx,title.paras],['title','',[s.title]],`slide ${i+1} title placeholder, first in reading order`);
    if(s.lines?.length){
      assert.deepEqual(s.kind==='opening'?[body.type,body.idx]:[body.type,body.idx],s.kind==='opening'?['subTitle','1']:['','1']);
      assert.deepEqual(body.paras,s.lines,`slide ${i+1} body paragraphs, one per line`);
    }
    for(const sp of shapes.map(x=>x.sp)){
      assert.equal(ns(ns(sp,A,'bodyPr')[0],A,'normAutofit').length,1,'the text shape autofits');
      for(const r of ns(sp,A,'r'))assert.equal(attr(ns(r,A,'rPr')[0],'lang'),'en-US','every run carries the language');
    }
    // the notes: a slide image and one paragraph per notes line
    const notes=await noteShapes(zip,i+1);
    assert.equal(notes.length,2);assert.deepEqual([notes[0].type,notes[0].paras.join('')],['sldImg','']);
    assert.deepEqual([notes[1].type,notes[1].idx],['body','1']);assert.deepEqual(notes[1].paras,s.notes?.length?s.notes:['']);
    assert.equal(ns(notes[1].sp,A,'p').length,Math.max(1,s.notes?.length||0));
    // layouts: Title Slide for the opening slide, Title and Content for the rest
    const layout=(await relsOf(zip,`ppt/slides/slide${i+1}.xml`)).find(r=>r.type==='slideLayout').target,layoutDoc=await part(zip,layout.replace('..','ppt'));
    assert.equal(attr(ns(layoutDoc,P,'cSld')[0],'name'),s.kind==='opening'?'Title Slide':'Title and Content');
  }
  // no text carries a size of its own: the sizes below come from the master and layouts
  for(let i=1;i<=6;i++)assert.equal(ns(await part(zip,`ppt/slides/slide${i}.xml`),A,'rPr').filter(r=>r.getAttribute('sz')).length,0);
});

test('the key survives in the slide name, and a slide with no key is named by its number',async()=>{
  const {zip}=await build(),names=await Promise.all([1,2,3,4,5,6].map(async n=>attr(ns(await part(zip,`ppt/slides/slide${n}.xml`),P,'cSld')[0],'name')));
  assert.deepEqual(names,['Slide 1','Slide 2','A13:1','Slide 4','Slide 5','Slide 6']);
});

test('XML special characters, injected markup, control characters and non-Latin text all round-trip as plain text',async()=>{
  const nasty='</a:t></a:r></a:p><p:sp>x',key='k&<>"\'';
  const model={title:`Fish & Chips <"quoted"> it's`,subtitle:'R&D > "Q&A"',language:'ja-JP',slides:[
    {kind:'opening',title:'日本語のタイトル',lines:['東京 — 大阪 & 京都'],notes:['メモ <1> & "2"']},
    {kind:'activity',title:'العنوان بالعربية',lines:[nasty,`a < b && c > d "quote" 'single'`,'Emoji 😀 and é (e + combining: é)'],notes:[nasty,`it's & "that's" <all>`],key},
    {kind:'close',title:'Ωmega ΑΒΓ — Привет'}]};
  const {zip}=await build(model),slide2=await slideShapes(zip,2);
  assert.equal(slide2.length,2,'injected markup did not add a shape');assert.deepEqual(slide2[0].paras,['العنوان بالعربية']);
  assert.deepEqual(slide2[1].paras,[nasty,`a < b && c > d "quote" 'single'`,'Emoji 😀 and é (e + combining: é)']);
  assert.deepEqual((await noteShapes(zip,2))[1].paras,[nasty,`it's & "that's" <all>`]);
  assert.equal(attr(ns(await part(zip,'ppt/slides/slide2.xml'),P,'cSld')[0],'name'),key);
  assert.deepEqual((await slideShapes(zip,1)).map(s=>s.paras),[['日本語のタイトル'],['東京 — 大阪 & 京都']]);assert.deepEqual((await noteShapes(zip,1))[1].paras,['メモ <1> & "2"']);
  assert.deepEqual((await slideShapes(zip,3))[0].paras,['Ωmega ΑΒΓ — Привет']);
  const core=await part(zip,'docProps/core.xml');assert.equal(ns(core,DC,'title')[0].textContent,`Fish & Chips <"quoted"> it's`);assert.equal(ns(core,DC,'subject')[0].textContent,'R&D > "Q&A"');assert.equal(ns(core,DC,'language')[0].textContent,'ja-JP');
  for(const r of ns(await part(zip,'ppt/slides/slide1.xml'),A,'rPr'))assert.equal(attr(r,'lang'),'ja-JP');
  const cleaned=checkDeck({title:'A'+String.fromCharCode(0,7,27,127)+'B\tC\r\nD'+String.fromCharCode(0xd800)+'E',slides:[{kind:'break',title:'x'+String.fromCharCode(0xfffe)+'y'}]});
  assert.equal(cleaned.title,'AB C DE');assert.equal(cleaned.slides[0].title,'xy');
  const {zip:zip2}=await build({title:'A'+String.fromCharCode(0,7,27)+'B',slides:[{kind:'break',title:'T'+String.fromCharCode(8)+'1',notes:['n'+String.fromCharCode(1)+'1']}]});
  assert.deepEqual((await slideShapes(zip2,1))[0].paras,['T1']);assert.deepEqual((await noteShapes(zip2,1))[1].paras,['n1']);
});

test('checkDeck returns a cleaned copy and leaves the input alone',()=>{
  const input={title:'  Spaced  ',slides:[{kind:'activity',title:' One ',lines:['  a  ',' b'],key:' k '}]},before=structuredClone(input),out=checkDeck(input);
  assert.deepEqual(input,before);assert.notEqual(out,input);
  assert.deepEqual(out,{title:'Spaced',language:'en-US',slides:[{kind:'activity',title:'One',lines:['a','b'],notes:[],key:'k'}]});
  assert.equal(checkDeck({title:'T',language:'fr-CA',subtitle:'  S ',slides:[{kind:'close',title:'C'}]}).subtitle,'S');
  for(const blank of ['','  ',null,undefined])assert.equal(checkDeck({title:'T',language:blank,slides:[{kind:'close',title:'C'}]}).language,'en-US','a blank language means the default');
  assert.equal(checkDeck({title:'T',language:' zh-Hans-CN ',slides:[{kind:'close',title:'C'}]}).language,'zh-Hans-CN');
  assert.ok(!('subtitle' in checkDeck({title:'T',subtitle:'   ',slides:[{kind:'close',title:'C'}]})));
  assert.deepEqual(SLIDE_KINDS,['opening','activity','break','close']);assert.ok(Object.isFrozen(SLIDE_LIMITS));
  assert.deepEqual([SLIDE_LIMITS.titleChars,SLIDE_LIMITS.slides,SLIDE_LIMITS.lines,SLIDE_LIMITS.lineChars,SLIDE_LIMITS.notes,SLIDE_LIMITS.noteChars,SLIDE_LIMITS.keyChars],[200,300,12,300,40,600,80]);
  assert.equal(checkDeck(SAMPLE).slides.length,6);
});

// ---- bad input is refused, in plain words ------------------------------------------------------------------------------------------
const refused=(name,change,pattern)=>test('refuses '+name,async()=>{
  const m=copy();change(m);
  const plain=e=>e instanceof Error&&pattern.test(e.message)&&!/undefined|null|NaN|\[object|TypeError|Cannot read/.test(e.message)&&e.message.length<200;
  assert.throws(()=>checkDeck(m),plain);await assert.rejects(pptxBytes(m,JSZip),plain);
});
refused('an empty deck title',m=>{m.title='   ';},/^The deck needs a title\.$/);
refused('a missing deck title',m=>{delete m.title;},/deck title must be text/);
refused('a deck title over 200 characters',m=>{m.title='x'.repeat(201);},/deck title is 201 characters long\. The most is 200\./);
refused('an empty slide title',m=>{m.slides[1].title='';},/^Slide 2 needs a title\.$/);
refused('an unknown kind',m=>{m.slides[1].kind='lecture';},/^Slide 2 has no valid kind\. Use opening, activity, break or close\.$/);
refused('a missing kind',m=>{delete m.slides[3].kind;},/Slide 4 has no valid kind/);
refused('13 lines',m=>{m.slides[2].lines=Array.from({length:13},(_,i)=>'line '+i);},/^Slide 3 has 13 lines\. The most is 12\. Move the rest into the speaker notes\.$/);
refused('301 slides',m=>{m.slides=Array.from({length:301},(_,i)=>({kind:'activity',title:'Slide '+i}));},/^The deck has 301 slides\. The most is 300\./);
refused('no slides',m=>{m.slides=[];},/at least one slide/);
refused('slides that are not a list',m=>{m.slides='none';},/at least one slide/);
refused('a slide that is not an object',m=>{m.slides[2]=null;},/^Slide 3 is missing/);
refused('a 301-character line',m=>{m.slides[2].lines[0]='x'.repeat(301);},/^Slide 3, line 1 is 301 characters long\. The most is 300\.$/);
refused('an empty line',m=>{m.slides[2].lines[1]='   ';},/^Slide 3, line 2 is empty\.$/);
refused('a line that is not text',m=>{m.slides[2].lines[0]=42;},/^Slide 3, line 1 must be text\.$/);
refused('lines that are not a list',m=>{m.slides[2].lines='one line';},/lines must be a list/);
refused('41 notes',m=>{m.slides[2].notes=Array.from({length:41},(_,i)=>'n'+i);},/^Slide 3 has 41 notes\. The most is 40\./);
refused('a 601-character note',m=>{m.slides[2].notes[0]='y'.repeat(601);},/^Slide 3, note 1 is 601 characters long\. The most is 600\.$/);
refused('an 81-character key',m=>{m.slides[2].key='k'.repeat(81);},/^Slide 3 key is 81 characters long\. The most is 80\.$/);
refused('an empty key',m=>{m.slides[2].key='  ';},/Slide 3 key is empty/);
refused('two slides with the same key',m=>{m.slides[4].key='A13:1';},/^Slide 5 has the key "A13:1", which another slide already has\./);
refused('a bad language',m=>{m.language='english please';},/language must be a code like en-US/);
refused('a language that is not text',m=>{m.language=5;},/language must be a code like en-US/);
refused('a subtitle that is too long',m=>{m.subtitle='s'.repeat(201);},/subtitle is 201 characters long/);
refused('a subtitle that is not text',m=>{m.subtitle=5;},/^The subtitle must be text\.$/);
for(const [name,value] of [['null',null],['a string','deck'],['an array',[]],['nothing',undefined]])test('refuses a model that is '+name,async()=>{
  const plain=e=>e instanceof Error&&e.message==='There is no deck to write. Give it a title and a list of slides.';
  assert.throws(()=>checkDeck(value),plain);await assert.rejects(pptxBytes(value,JSZip),plain);assert.throws(()=>slideFit(value),plain);
});
test('limits are exact: 300 slides, 12 lines, 300 and 600 characters, 40 notes and an 80-character key are all accepted',async()=>{
  const model={title:'x'.repeat(200),slides:Array.from({length:300},(_,i)=>({kind:'activity',title:'T'.repeat(200),lines:i?[]:Array.from({length:12},(_,k)=>String(k%10).repeat(300)),notes:i?['n']:Array.from({length:40},(_,k)=>String(k%10).repeat(600)),key:i?'k'+i:'k'.repeat(80)}))};
  const {bytes,zip}=await build(model);
  assert.equal(ns(await part(zip,'ppt/presentation.xml'),P,'sldId').length,300);assert.equal((await slideShapes(zip,1))[1].paras.length,12);assert.equal((await noteShapes(zip,1))[1].paras.length,40);
  const read=await readPptx(bytes,'big.pptx',JSZip);assert.equal(read.slides.length,300);assert.equal(read.slides[0].notes.split('\n').length,40);
});
test('refuses a missing zip library',async()=>{await assert.rejects(pptxBytes(SAMPLE,undefined),/zip library is missing/);});

// ---- slideFit ------------------------------------------------------------------------------------------------------------------------
test('slideFit flags a crowded slide and passes a normal one',()=>{
  assert.deepEqual(slideFit(SAMPLE),[]);
  const lines=n=>Array.from({length:n},(_,i)=>'Point '+(i+1));
  const m=deckOf(['ok','eight lines','long text','one long line','edge lines','edge chars','edge line']);
  m.slides[0].lines=lines(7);                                              // 7 short lines: fits
  m.slides[1].lines=lines(8);                                              // 8 lines: too many
  m.slides[2].lines=['x'.repeat(110),'y'.repeat(110),'z'.repeat(110),'w'.repeat(110)];   // 440 characters in 4 lines: too much text
  m.slides[3].lines=['Short','q'.repeat(141)];                              // one line over 140
  m.slides[4].lines=Array.from({length:7},()=>'a'.repeat(60));              // 7 lines of 60 = 420: fits
  m.slides[5].lines=['b'.repeat(140),'c'.repeat(140),'d'.repeat(141)];      // 421 characters
  m.slides[6].lines=['e'.repeat(140)];                                      // exactly 140: fits
  assert.deepEqual(slideFit(m),[
    {slide:2,message:'Slide 2 has too much text to read from the back of a room. Move some of it into the speaker notes.'},
    {slide:3,message:'Slide 3 has too much text to read from the back of a room. Move some of it into the speaker notes.'},
    {slide:4,message:'Slide 4 has a line that is too long to read from the back of a room. Shorten it, or move the detail into the speaker notes.'},
    {slide:6,message:'Slide 6 has too much text to read from the back of a room. Move some of it into the speaker notes.'}]);
  assert.throws(()=>slideFit({title:'',slides:[]}),/needs a title/);
});

// ---- pptxName ------------------------------------------------------------------------------------------------------------------------
test('pptxName makes a safe file name',()=>{
  assert.equal(pptxName('Skill Practice Workshop'),'Skill Practice Workshop.pptx');
  assert.equal(pptxName('A/B: "Plan" <v2>?'),'AB Plan v2.pptx');
  assert.equal(pptxName('  spaced   out  '),'spaced out.pptx');
  assert.equal(pptxName('under_score-and-dash 7'),'under_score-and-dash 7.pptx');
  for(const empty of ['','   ','***',null,undefined,'..','???'])assert.equal(pptxName(empty),'Slides.pptx');
  for(const reserved of ['CON','nul','Com1','lpt9'])assert.equal(pptxName(reserved),'Slides.pptx');
  assert.equal(pptxName('..\\..\\evil/../name'),'evilname.pptx');
  assert.equal(pptxName('日本語のタイトル'),'日本語のタイトル.pptx');
  assert.equal(pptxName('Café'),'Café.pptx','accents are kept, composed');
  const long=pptxName('x'.repeat(100)),astral=pptxName('\u{20BB7}'.repeat(100));
  assert.equal(long.length,65);assert.ok(long.endsWith('.pptx'));assert.equal(Array.from(astral).length,65,'cut by characters, not by halves of one');assert.ok(!/[\ud800-\udfff]/.test(astral.replace(/[\ud800-\udbff][\udc00-\udfff]/g,'')));
  for(const n of [long,astral,pptxName('A/B: "Plan" <v2>?'),pptxName('x'.repeat(59)+' '+'y'.repeat(20))])assert.match(n,/^[\p{L}\p{M}\p{N} _-]+\.pptx$/u);
  assert.ok(!/ \.pptx$/.test(pptxName('x'.repeat(59)+' '+'y'.repeat(20))),'no space left before the extension');
});

// ---- the look ---------------------------------------------------------------------------------------------------------------------------
const luminance=hex=>{const [r,g,b]=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)/255).map(v=>v<=0.03928?v/12.92:((v+0.055)/1.055)**2.4);return 0.2126*r+0.7152*g+0.0722*b;};
const contrast=(a,b)=>{const [hi,lo]=[luminance(a),luminance(b)].sort((x,y)=>y-x);return (hi+0.05)/(lo+0.05);};
test('the look: navy titles, teal accent bar, Arial through the theme, big text, WCAG AA colours, no pictures',async()=>{
  const {zip}=await build(),theme=await part(zip,'ppt/theme/theme1.xml'),colour=name=>attr(ns(ns(theme,A,name)[0],A,'srgbClr')[0],'val');
  assert.deepEqual(['dk1','lt1','dk2','accent1'].map(colour),['18283A','FFFFFF','172B4A','216861']);
  for(const name of ['dk1','dk2','accent1','hlink','folHlink'])assert.ok(contrast(colour(name),colour('lt1'))>=4.5,`${name} is AA against white`);
  assert.ok(contrast('172B4A','FFFFFF')>=7&&contrast('18283A','FFFFFF')>=7,'titles and body text are AAA');
  assert.deepEqual(['majorFont','minorFont'].map(f=>attr(ns(ns(theme,A,f)[0],A,'latin')[0],'typeface')),['Arial','Arial']);
  const master=await part(zip,'ppt/slideMasters/slideMaster1.xml'),title=ns(ns(master,P,'titleStyle')[0],A,'defRPr')[0],body=ns(master,P,'bodyStyle')[0],level1=ns(body,A,'lvl1pPr')[0];
  assert.equal(attr(title,'sz'),'3600');assert.equal(attr(ns(title,A,'schemeClr')[0],'val'),'tx2');assert.equal(attr(ns(level1,A,'defRPr')[0],'sz'),'2400');
  assert.equal(attr(ns(level1,A,'buChar')[0],'char'),'•','bullets come from a:buChar');
  for(const sz of ns(master,A,'defRPr').map(d=>d.getAttribute('sz')).filter(Boolean))assert.ok(Number(sz)>=1800,'no text under 18 pt on a slide: '+sz);
  const bar=ns(master,P,'sp').find(sp=>attr(ns(sp,P,'cNvPr')[0],'name')==='Accent bar');
  assert.equal(attr(ns(bar,A,'schemeClr')[0],'val'),'accent1');assert.ok(Number(attr(ns(bar,A,'ext')[0],'cy'))<=137160,'a thin bar');assert.ok(bar.toString().includes('decorative'),'the bar is marked decorative');
  const opening=await part(zip,'ppt/slideLayouts/slideLayout1.xml'),shapes=shapesOf(opening);
  assert.equal(attr(ns(shapes[0].sp,A,'defRPr')[0],'sz'),'4400','the opening title is 44 pt');
  assert.deepEqual([shapes[1].type,shapes[1].idx],['subTitle','1']);assert.equal(ns(shapes[1].sp,A,'buNone').length,1,'the subtitle has no bullets');assert.equal(ns(shapes[1].sp,A,'buChar').length,0);
  for(const n of [1,2,3,4,5,6]){const d=await part(zip,`ppt/slides/slide${n}.xml`);for(const tag of ['pic','blip','blipFill','graphicFrame'])assert.equal(d.getElementsByTagName('p:'+tag).length+d.getElementsByTagName('a:'+tag).length,0,'no pictures or tables');}
});

test('an opening slide with no lines shows the deck subtitle, and with no subtitle has no body shape',async()=>{
  const lone=(subtitle)=>({title:'T',...(subtitle?{subtitle}:{}),slides:[{kind:'opening',title:'Welcome'},{kind:'activity',title:'Next'}]});
  const {zip}=await build(lone('The subtitle'));assert.deepEqual((await slideShapes(zip,1)).map(s=>[s.type,s.paras]),[['title',['Welcome']],['subTitle',['The subtitle']]]);assert.equal((await slideShapes(zip,2)).length,1);
  const {zip:plain}=await build(lone());assert.equal((await slideShapes(plain,1)).length,1);
});

test('the same deck and time give the same bytes, and the writer never changes the model',async()=>{
  const model=copy(),before=structuredClone(model),a=await pptxBytes(model,JSZip,{now:NOW}),b=await pptxBytes(model,JSZip,{now:NOW});
  assert.deepEqual(model,before);assert.ok(a instanceof Uint8Array);assert.deepEqual(a,b);assert.equal(a[0],0x50);assert.equal(a[1],0x4b);
});

// ---- the repo's own reader -----------------------------------------------------------------------------------------------------------
test('round trip: readPptx finds the same slides, titles, body text and notes',async()=>{
  const {bytes}=await build(),before=bytes.slice(),read=await readPptx(bytes,'sample.pptx',JSZip);
  assert.deepEqual(bytes,before);assert.equal(read.slides.length,SAMPLE.slides.length);
  assert.deepEqual(read.slides.map(s=>s.title),SAMPLE.slides.map(s=>s.title));
  assert.deepEqual(read.slides.map(s=>s.notes),SAMPLE.slides.map(s=>(s.notes||[]).join('\n')));
  assert.deepEqual(read.slides.map(s=>s.text),SAMPLE.slides.map(s=>[s.title,...(s.lines||[])].join('\n')));
  assert.deepEqual(read.slides.map(s=>s.sourceId),[256,257,258,259,260,261]);assert.deepEqual(read.slides.map(s=>s.order),[1,2,3,4,5,6]);
  assert.deepEqual(read.slides[2].warnings,[],'a slide with lines and notes has nothing to warn about');
  assert.deepEqual(read.slides[1].warnings,['No speaker notes text found.']);
  const details=await readPptx(bytes,'sample.pptx',JSZip,{details:true});assert.deepEqual(details.slides.map(s=>s.title),read.slides.map(s=>s.title));assert.ok(details.slides.every(s=>!s.titleGuessed&&!s.hidden&&s.clicks===0));
});
test('round trip: special characters and several notes come back through the reader',async()=>{
  const model={title:'T',slides:[{kind:'activity',title:`R&D <"Q&A"> it's`,lines:['x < y & z'],notes:['first','second & third','日本語']}]},read=await readPptx((await build(model)).bytes,'x.pptx',JSZip);
  assert.equal(read.slides[0].title,`R&D <"Q&A"> it's`);assert.equal(read.slides[0].notes,'first\nsecond & third\n日本語');assert.equal(read.slides[0].text,`R&D <"Q&A"> it's\nx < y & z`);
});

// ---- the schema validator (opt-in) ---------------------------------------------------------------------------------------------------
// OPENSTAGE_PPTX_VALIDATOR names the pptx skill's validate.py. `python -I validate.py` cannot import the validator's own sibling
// folders (isolated mode leaves the script's folder off the path), so a short bootstrap, still run with -I, puts that folder on the
// path and then runs validate.py as __main__. It needs the Python packages lxml and defusedxml; if they are not installed for Python,
// OPENSTAGE_PPTX_VALIDATOR_LIBS may name a folder that holds them (it is added to the path, nothing else is). Without them the test skips.
const BOOT='import os,runpy,sys;v,libs=sys.argv[1:3];sys.path[:0]=[p for p in (libs,os.path.dirname(v)) if p];sys.argv=[v]+sys.argv[3:];runpy.run_path(v,run_name="__main__")';
const NEEDS='import sys;sys.path[:0]=[p for p in sys.argv[1:] if p];import lxml,defusedxml';
function python(args){for(const exe of ['python','python3']){const r=spawnSync(exe,args,{encoding:'utf8',timeout:180000});if(!r.error||r.error.code!=='ENOENT')return r;}return {error:new Error('Python is not installed or not on the path')};}
test('the pptx schema validator accepts the sample deck',{skip:process.env.OPENSTAGE_PPTX_VALIDATOR?false:'set OPENSTAGE_PPTX_VALIDATOR to the path of the pptx skill\'s validate.py to run the schema check'},async t=>{
  const validator=path.resolve(process.env.OPENSTAGE_PPTX_VALIDATOR),libs=process.env.OPENSTAGE_PPTX_VALIDATOR_LIBS?path.resolve(process.env.OPENSTAGE_PPTX_VALIDATOR_LIBS):'';
  assert.ok(fs.existsSync(validator),'OPENSTAGE_PPTX_VALIDATOR names a file that does not exist: '+validator);
  const needs=python(['-I','-c',NEEDS,libs]);
  if(needs.error||needs.status!==0){const why=needs.error?needs.error.message:String(needs.stderr).trim().split('\n').pop();t.skip('the validator needs the Python packages lxml and defusedxml ('+why+')');return;}
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'openstage-pptx-')),file=path.join(dir,'deck.pptx');
  try{
    fs.writeFileSync(file,(await build()).bytes);
    const run=python(['-I','-c',BOOT,validator,libs,file]);
    assert.ok(!run.error,run.error?.message);assert.equal(run.status,0,`${run.stdout}\n${run.stderr}`);assert.match(run.stdout,/All validations PASSED/);
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});
