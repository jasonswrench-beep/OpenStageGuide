// Changing the words of one paragraph (src/guide-edit.mjs): what can be changed, what stays, and that nothing else is ever touched.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readParagraph,setWords,wordsOf,splitWords,joinWords,nextParagraphSpec,paragraphWith,cleanWords,isHeading,isLocked,MAX_WORDS,KINDS,kindChange,withStyle} from '../src/guide-edit.mjs';
import {pictureRun,para,run,makeContext} from '../src/guide-ooxml.mjs';
import {formTitleBlock,citationControl,checkList,formTable} from '../src/guide-forms-ooxml.mjs';
import {STYLE_IDS} from '../src/guide-styles.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function guide(){
  const a=await started('Editing');
  await addModule(a,'Naming it');await addLesson(a,'First','10');
  await addBlock(a,'ask','What makes it hard?\nSecond line.',{notes:'Listen for blame.\nMore.',minutes:'5'});
  const ps=a.adapter.exportState().paragraphs;
  const at=ps.findIndex(p=>p.style==='OSG Module');
  return {all:ps,block:ps.slice(at),by:style=>ps.find(p=>p.style===style)};
}
const read=p=>readParagraph(p.xml,p.style,io);

test('a cue line reads as a lead that is shown and the words that are typed; "(5 min)" stays in the words',async()=>{
  const g=await guide();
  const cue=read(g.by('OSG Cue'));
  assert.equal(cue.editable,true);
  assert.deepEqual(cue.lead.map(l=>l.kind),['icon','label','tab']);
  assert.equal(cue.lead[1].text.trim(),'Ask');
  assert.equal(cue.text,'What makes it hard? (5 min)');
});

test('notes read with their label in front, continuation lines and body lines as plain words, and Normal text as words',async()=>{
  const g=await guide();
  const notes=g.block.filter(p=>p.style==='OSG Facilitator Notes').map(read);
  assert.deepEqual(notes.map(n=>n.lead.map(l=>l.kind)),[['label'],[]]);
  assert.equal(notes[0].lead[0].text,'Facilitator notes: ');
  assert.deepEqual(notes.map(n=>n.text),['Listen for blame.','More.']);
  assert.equal(read(g.by('OSG Cue Body')).text,'Second line.');
  assert.equal(read(g.all.find(p=>p.style==='Normal'&&p.text==='Facilitator guide')).text,'Facilitator guide');
  assert.equal(read(g.all.find(p=>p.style==='heading 1')).text,'Acknowledgements');
});

test('a module or lesson heading reads as a number the program writes and a title that is typed',async()=>{
  const g=await guide();
  const m=read(g.by('OSG Module')),l=read(g.by('OSG Lesson'));
  assert.deepEqual([m.prefix,m.text],['Module 1: ','Naming it']);
  assert.deepEqual([l.prefix,l.text],['Lesson 1.1: ','First']);
  assert.equal(isHeading('OSG Module'),true);
  assert.equal(isHeading('OSG Cue Body'),false);
});

test('what the program writes, and what holds more than words, is not changed by the editor, and says why',async()=>{
  const g=await guide();
  const time=read(g.by('OSG Time'));
  assert.equal(time.editable,false);assert.match(time.reason,/program writes/);
  assert.equal(isLocked('OSG Time'),true);assert.equal(isLocked('TOC 2'),true);
  assert.equal(read({xml:'',style:'TOC 1'}).editable,false);
  assert.equal(read({xml:'',style:'Normal'}).editable,false,'no XML: a table cell or a generated line');
  // One of each kind of thing that is not words, made by the builders the program uses and read back as the document keeps them.
  const a=await started('Atoms');
  await a.adapter.insert(ctx=>
    para(STYLE_IDS.body,pictureRun({base64:'AAAA',type:'png',mime:'image/png',size:{cx:100,cy:100},alt:'A cat',name:'cat'},ctx))
    +para(STYLE_IDS.body,run('Give out ')+citationControl({id:'f1',text:'Handout 1.1: Goals'},ctx)+run(' now.'))
    +formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})
    +formTable({columns:[{name:'Goal',weight:3}],rows:[['Plan']],ctx,caption:'Goals'})
    +checkList({items:['Read it'],columns:[],ctx})
    +`<w:p><w:r><w:t>One</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Two</w:t></w:r></w:p>`,
    {where:'start',spare:false});
  const everything=a.adapter.exportState().paragraphs;
  const made=everything.slice(0,everything.findIndex(p=>p.style==='OSG Guide Title')).filter(p=>p.xml); // what was inserted comes before the guide's own title
  assert.ok(made.length>=6,'several paragraphs were made: '+made.length);
  for(const p of made){
    const info=read(p);
    assert.equal(info.editable,false,`${p.style}: ${p.text}`);
    assert.throws(()=>setWords(p.xml,p.style,'x',io),/./,p.text);
    assert.equal(splitWords(p.xml,p.style,0,io),null);
  }
});

test('the same words give back the very same XML, for every paragraph the guide has',async()=>{
  const g=await guide();
  let changed=0;
  for(const p of g.all){
    const info=read(p);
    if(!info.editable)continue;
    assert.equal(setWords(p.xml,p.style,info.text,io),p.xml,p.style);
    changed++;
  }
  assert.ok(changed>=10,'a good many paragraphs were checked: '+changed);
});

test('new words keep the lead, the paragraph\'s properties and the heading\'s number; only the words change',async()=>{
  const g=await guide();
  const cue=g.by('OSG Cue');
  const edited=setWords(cue.xml,cue.style,'Why is it hard?',io);
  const info=readParagraph(edited,cue.style,io);
  assert.deepEqual(info.lead.map(l=>l.kind),['icon','label','tab']);
  assert.equal(info.text,'Why is it hard?');
  assert.ok(edited.includes('OSGCue'),'the paragraph\'s style');
  assert.equal((edited.match(/<w:drawing>/g)||[]).length,1,'the picture is still there once');
  const notes=g.block.find(p=>p.style==='OSG Facilitator Notes');
  const n2=readParagraph(setWords(notes.xml,notes.style,'Ask what they notice.',io),notes.style,io);
  assert.deepEqual([n2.lead[0].text,n2.text],['Facilitator notes: ','Ask what they notice.']);
  const mod=g.by('OSG Module');
  const edited2=setWords(mod.xml,mod.style,'Naming feelings',io);
  assert.equal(readParagraph(edited2,mod.style,io).prefix+readParagraph(edited2,mod.style,io).text,'Module 1: Naming feelings');
  assert.match(edited2,/Module 1: Naming feelings/);
});

test('paragraph properties other than the style (a page break, a list) are kept as they were',()=>{
  const xml=`<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="Heading1"/><w:pageBreakBefore/></w:pPr><w:r><w:t xml:space="preserve">Contents</w:t></w:r></w:p>`;
  const out=setWords(xml,'heading 1','Table of contents',io);
  assert.match(out,/<w:pageBreakBefore\/>/);
  assert.match(out,/Table of contents/);
  const list=`<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="OSGCueBody"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Step one</w:t></w:r></w:p>`;
  assert.match(setWords(list,'OSG Cue Body','Step 1',io),/<w:numId w:val="3"\/>/);
});

test('words are text: markup characters, quotes, accents, other scripts and emoji go in and come back exactly; line ends become spaces; characters XML cannot hold are dropped',()=>{
  const xml=`<w:p xmlns:w="${W}"><w:r><w:t>x</w:t></w:r></w:p>`;
  for(const words of ['<script>alert(1)</script> & "quoted" \'single\'','Guía de facilitación — ¿qué sentimos?','日本語のテキスト','مرحبا بالعالم','Emoji 😀👍🏽 and a tab\there','  leading and trailing  ','a'.repeat(5000)]){
    const out=setWords(xml,'Normal',words,io);
    assert.equal(readParagraph(out,'Normal',io).text,words,words.slice(0,30));
    assert.ok(!/<script>/.test(out),'markup in the words is text, not markup');
    assert.ok(io.parse(out).documentElement,'the XML is still well formed');
  }
  assert.equal(wordsOf(setWords(xml,'Normal','one\r\ntwo\nthree\u2028four',io),'Normal',io),'one two three four');
  assert.equal(cleanWords('a\u0000b\u0008c\u000Bd\uFFFEe'),'abcde');
  assert.equal(wordsOf(setWords(xml,'Normal','',io),'Normal',io),'','empty words leave an empty paragraph');
  assert.throws(()=>setWords(xml,'Normal','x'.repeat(MAX_WORDS+1),io),/up to/);
  assert.throws(()=>setWords(xml,'Normal',42,io),/as text/);
});

test('an empty paragraph can be given words, and a cue line with no words can be given some',()=>{
  const empty=`<w:p xmlns:w="${W}"/>`;
  assert.equal(readParagraph(empty,'Normal',io).editable,true);
  assert.equal(wordsOf(setWords(empty,'Normal','Hello',io),'Normal',io),'Hello');
});

test('Enter in the middle of a cue line: the cue keeps the words before, a body paragraph takes the rest',async()=>{
  const g=await guide();
  const cue=g.by('OSG Cue');
  const out=splitWords(cue.xml,cue.style,'What makes '.length,io);
  assert.equal(wordsOf(out.first,'OSG Cue',io),'What makes ');
  assert.deepEqual(readParagraph(out.first,'OSG Cue',io).lead.map(l=>l.kind),['icon','label','tab']);
  assert.equal(wordsOf(out.second,'OSG Cue Body',io),'it hard? (5 min)');
  assert.match(out.second,/w:val="OSGCueBody"/);
  const end=splitWords(cue.xml,cue.style,10000,io);
  assert.equal(wordsOf(end.second,'OSG Cue Body',io),'','Enter at the end makes an empty paragraph');
  assert.equal(wordsOf(end.first,'OSG Cue',io),'What makes it hard? (5 min)');
});

test('Enter goes on in the same kind of paragraph: notes in notes, text in text, a list item in its list; a heading is not split',async()=>{
  const g=await guide();
  const notes=g.block.filter(p=>p.style==='OSG Facilitator Notes')[0];
  const n=splitWords(notes.xml,notes.style,6,io);
  assert.match(n.second,/w:val="OSGNotes"/);
  assert.equal(readParagraph(n.second,'OSG Facilitator Notes',io).lead.length,0,'the label stays with the first paragraph only');
  assert.equal(wordsOf(n.first,notes.style,io),'Listen');
  assert.equal(wordsOf(n.second,'OSG Facilitator Notes',io),' for blame.');
  const normal=`<w:p xmlns:w="${W}"><w:r><w:t>One two</w:t></w:r></w:p>`;
  const s=splitWords(normal,'Normal',3,io);
  assert.ok(!/pStyle/.test(s.second),'Normal text goes on as Normal text');
  assert.deepEqual([wordsOf(s.first,'Normal',io),wordsOf(s.second,'Normal',io)],['One',' two']);
  const list=`<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="OSGCueBody"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Step one</w:t></w:r></w:p>`;
  assert.match(nextParagraphSpec(list,'OSG Cue Body',io).numPr,/numId/);
  assert.match(splitWords(list,'OSG Cue Body',4,io).second,/<w:numId w:val="3"\/>/);
  assert.equal(splitWords(g.by('OSG Module').xml,'OSG Module',3,io),null);
  assert.equal(splitWords(g.by('OSG Time').xml,'OSG Time',3,io),null);
  assert.equal(nextParagraphSpec(g.by('OSG Lesson').xml,'OSG Lesson',io),null);
});

test('a paragraph made of words in a style is well formed whatever the words are',()=>{
  const xml=paragraphWith({styleId:'OSGNotes',numPr:null},'a < b & c > d',io);
  assert.equal(readParagraph(xml,'OSG Facilitator Notes',io).text,'a < b & c > d');
  assert.equal(readParagraph(paragraphWith({styleId:null,numPr:null},'',io),'Normal',io).editable,true);
});

test('Backspace at the start joins a paragraph to the one before; not into a heading, and not a paragraph that has a lead of its own',async()=>{
  const g=await guide();
  const body=g.by('OSG Cue Body'),cue=g.by('OSG Cue'),notes=g.block.filter(p=>p.style==='OSG Facilitator Notes');
  const j=joinWords(cue,body,io);
  assert.equal(wordsOf(j.xml,'OSG Cue',io),'What makes it hard? (5 min)Second line.');
  assert.equal(j.offset,'What makes it hard? (5 min)'.length);
  assert.deepEqual(readParagraph(j.xml,'OSG Cue',io).lead.map(l=>l.kind),['icon','label','tab'],'the cue line keeps its lead');
  assert.equal(wordsOf(joinWords(notes[0],notes[1],io).xml,'OSG Facilitator Notes',io),'Listen for blame.More.');
  assert.equal(joinWords(g.by('OSG Lesson'),cue,io),null,'a cue line has a lead of its own');
  assert.equal(joinWords(g.by('OSG Lesson'),body,io),null,'not into a heading');
  assert.equal(joinWords(body,g.by('OSG Time'),io),null,'not a locked line');
});

test('words typed in any order give back what was typed (a thousand random edits)',()=>{
  let seed=7;const rand=n=>{seed=(seed*1103515245+12345)&0x7fffffff;return seed%n;};
  const alphabet=['a','b',' ','<','>','&','"','é','日','😀','1','.',':','-'];
  const make=n=>Array.from({length:n},()=>alphabet[rand(alphabet.length)]).join('');
  let xml=`<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="OSGCueBody"/></w:pPr><w:r><w:t>start</w:t></w:r></w:p>`;
  for(let i=0;i<1000;i++){
    const words=make(rand(40));
    xml=setWords(xml,'OSG Cue Body',words,io);
    assert.equal(wordsOf(xml,'OSG Cue Body',io),words);
  }
  assert.match(xml,/w:val="OSGCueBody"/);
  assert.equal((xml.match(/<w:r>/g)||[]).length<=1,true,'one run, however many edits');
});

test('the kinds a paragraph can be changed to are styles the guide already has; the structure the program writes is not among them',()=>{
  assert.deepEqual(KINDS.map(k=>k.label),['Normal text','Block text','Facilitator note','Heading 1','Heading 2']);
  assert.ok(KINDS.every(k=>k.note&&k.note.length>5));
  for(const k of KINDS)if(k.styleId!==null)assert.match(k.styleId,/^[A-Za-z][A-Za-z0-9]*$/);
  assert.deepEqual(KINDS.map(k=>k.style),['Normal','OSG Cue Body','OSG Facilitator Notes','heading 1','heading 2'],'the names the guide reads them by');
});

test('which paragraphs can change kind, and why the others cannot',async()=>{
  const g=await guide();
  assert.equal(kindChange(g.by('OSG Cue Body').xml,'OSG Cue Body',io).ok,true);
  assert.equal(kindChange(g.all.find(p=>p.style==='Normal'&&p.text==='Facilitator guide').xml,'Normal',io).ok,true);
  assert.equal(kindChange(g.all.find(p=>p.style==='heading 1').xml,'heading 1',io).ok,true);
  for(const style of ['OSG Module','OSG Lesson','OSG Guide Title']){
    const c=kindChange(g.by(style).xml,style,io);
    assert.equal(c.ok,false,style);assert.match(c.reason,/made by the tools/);
  }
  const cue=kindChange(g.by('OSG Cue').xml,'OSG Cue',io);
  assert.equal(cue.ok,false);assert.match(cue.reason,/cue line/);
  const labelled=g.block.find(p=>p.style==='OSG Facilitator Notes');
  assert.equal(kindChange(labelled.xml,labelled.style,io).ok,false,'a note with its label in front keeps its kind');
  assert.equal(kindChange(g.by('OSG Time').xml,'OSG Time',io).ok,false,'locked');
  assert.equal(kindChange('',  'Normal',io).ok,false,'nothing to change');
});

test('changing the kind changes the style and nothing else: the words, other properties and a list stay (a heading is not a list)',()=>{
  const W2=`xmlns:w="${W}"`;
  const body=`<w:p ${W2}><w:pPr><w:pStyle w:val="OSGCueBody"/><w:keepNext/></w:pPr><w:r><w:t>Hi there</w:t></w:r></w:p>`;
  const note=withStyle(body,'OSG Cue Body','OSGNotes',io);
  assert.match(note,/w:val="OSGNotes"/);assert.match(note,/<w:keepNext\/>/);
  assert.equal(readParagraph(note,'OSG Facilitator Notes',io).text,'Hi there');
  const plain=withStyle(body,'OSG Cue Body',null,io);
  assert.ok(!/pStyle/.test(plain));assert.match(plain,/<w:keepNext\/>/,'the other property is kept');
  const bare=`<w:p ${W2}><w:r><w:t>Hi</w:t></w:r></w:p>`;
  assert.match(withStyle(bare,'Normal','Heading2',io),/<w:pPr><w:pStyle w:val="Heading2"\/><\/w:pPr>/,'a paragraph with no properties gets some');
  assert.ok(!/pPr/.test(withStyle(withStyle(bare,'Normal','OSGNotes',io),'OSG Facilitator Notes',null,io)),'and loses them again when nothing is left');
  const list=`<w:p ${W2}><w:pPr><w:pStyle w:val="OSGCueBody"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Step</w:t></w:r></w:p>`;
  assert.match(withStyle(list,'OSG Cue Body','OSGNotes',io),/<w:numId w:val="3"\/>/,'a list item stays a list item');
  assert.ok(!/numPr/.test(withStyle(list,'OSG Cue Body','Heading1',io)),'a heading is not a list item');
});

test('a kind that is not a plain id, or a paragraph that cannot change kind, is refused',()=>{
  const W2=`xmlns:w="${W}"`;
  const body=`<w:p ${W2}><w:pPr><w:pStyle w:val="OSGCueBody"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>`;
  for(const bad of ['"><w:evil/>','a b','1x','','x"y'])assert.throws(()=>withStyle(body,'OSG Cue Body',bad,io),/not known/,bad);
  const mod=`<w:p ${W2}><w:pPr><w:pStyle w:val="OSGModule"/></w:pPr><w:r><w:t>Module 1: A</w:t></w:r></w:p>`;
  assert.throws(()=>withStyle(mod,'OSG Module',null,io),/made by the tools/);
});
