import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {backupFileName,inspectBackup,copyProblem,backupFacts,sizeText,MAX_BACKUP_BYTES} from '../src/guide-backup.mjs';
import {STYLE_IDS,stylesPart,NS} from '../src/guide-styles.mjs';
import {pane,started,click,addModule,addLesson} from './helpers/guide-harness.mjs';
globalThis.DOMParser=DOMParser;

const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;');
// A real .docx made the way Word makes one: the guide's own styles, paragraphs that name them, some pictures in word/media.
async function docx(paras,{images=2,title='A guide',modified='2026-11-08T09:15:00Z'}={}){
  const zip=new JSZip();
  zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  zip.file('word/styles.xml',stylesPart({}));
  const body=paras.map(([id,text])=>`<w:p>${id?`<w:pPr><w:pStyle w:val="${id}"/></w:pPr>`:''}<w:r><w:t>${esc(text)}</w:t></w:r></w:p>`).join('');
  zip.file('word/document.xml',`<w:document ${NS}><w:body>${body}</w:body></w:document>`);
  zip.file('docProps/core.xml',`<cp:coreProperties xmlns:cp="x" xmlns:dc="y" xmlns:dcterms="z"><dc:title>${esc(title)}</dc:title><dcterms:modified>${modified}</dcterms:modified></cp:coreProperties>`);
  for(let i=0;i<images;i++)zip.file(`word/media/image${i+1}.png`,new Uint8Array([137,80,78,71,i]));
  return zip.generateAsync({type:'uint8array'});
}
const GUIDE=[[STYLE_IDS.guideTitle,'Reframing conflict'],[STYLE_IDS.module,'Module 1: Basics'],[STYLE_IDS.lesson,'Lesson 1.1: Start'],[STYLE_IDS.time,' Time: 10 minutes. Cumulative time: 0 minutes.'],
  [STYLE_IDS.cueFacilitator,' Slide\tSlide 1: Intro'],[STYLE_IDS.notes,'Facilitator notes: Say hello.'],[STYLE_IDS.lesson,'Lesson 1.2: Next'],[STYLE_IDS.module,'Module 2: More'],[null,'Plain text.']];

test('a backup is named for the guide and the minute it was made, with nothing in the name a file cannot hold',()=>{
  const when=new Date(2026,10,8,9,5);
  assert.equal(backupFileName('Reframing conflict',when),'Reframing conflict - Backup 2026-11-08 0905.docx');
  assert.equal(backupFileName('A/B: "C"? <D>',when),'AB C D - Backup 2026-11-08 0905.docx');
  assert.equal(backupFileName('',when),'OpenStage Guide - Backup 2026-11-08 0905.docx');
  assert.ok(backupFileName('x'.repeat(200),when).length<=90);
  assert.equal(sizeText(3*1024*1024+200000),'3.2 MB');assert.equal(sizeText(2048),'2 KB');
});

test('a Word file is read for what it holds: its guide, its pictures, its dates',async()=>{
  const info=await inspectBackup(await docx(GUIDE),JSZip,{DOMParserClass:DOMParser});
  assert.equal(info.isGuide,true);
  assert.equal(info.title,'Reframing conflict');
  assert.deepEqual([info.modules,info.lessons,info.slides,info.paragraphs,info.images],[2,2,1,9,2]);
  assert.equal(info.modified,'2026-11-08T09:15:00Z');
  assert.match(info.fingerprint,/^[0-9a-f]{64}$/);
  const facts=backupFacts(info);
  assert.equal(facts[0],'Title: Reframing conflict');
  assert.equal(facts[1],'2 modules, 2 lessons, 1 slide.');
  assert.match(facts[2],/^9 paragraphs, 2 images, \d+ KB\.$/);
  assert.ok(facts.some(f=>/^Last saved in Word: .*2026/.test(f)));
  assert.ok(facts.some(f=>/^Fingerprint: [0-9a-f]{12} /.test(f)));
});

test('a Word file that is not a guide is said to be one that was not made with the guide tools',async()=>{
  const info=await inspectBackup(await docx([[null,'Just a letter.']],{title:'Letter'}),JSZip,{DOMParserClass:DOMParser});
  assert.equal(info.isGuide,false);assert.equal(info.title,'Letter');
  assert.match(backupFacts(info)[1],/not made with the guide tools/);
});

test('what is not a readable Word file is refused in plain words, never with a code',async()=>{
  const bad=[
    [new Uint8Array(0),/That file is empty/],
    [new Uint8Array([1,2,3,4,5]),/not a Word document/],
    [await new JSZip().file('hello.txt','hi').generateAsync({type:'uint8array'}),/not a Word document/],
    [new Uint8Array(MAX_BACKUP_BYTES+1),/Backups up to 64\.0 MB/],
    [await docx([]),/no paragraphs/]
  ];
  for(const [bytes,pattern] of bad)await assert.rejects(inspectBackup(bytes,JSZip,{DOMParserClass:DOMParser}),pattern);
  await assert.rejects(inspectBackup(await docx(GUIDE),null),/could not load what it needs/);
});

test('a copy that says something different from the document is caught before it is offered',async()=>{
  const info=await inspectBackup(await docx(GUIDE),JSZip,{DOMParserClass:DOMParser});
  assert.equal(copyProblem(info,{isGuide:true,modules:2,lessons:2}),'');
  assert.match(copyProblem(info,{isGuide:true,modules:3,lessons:2}),/has 2 modules and 2 lessons, but the document has 3 modules and 2 lessons.*not saved/);
  const plain=await inspectBackup(await docx([[null,'x']]),JSZip,{DOMParserClass:DOMParser});
  assert.match(copyProblem(plain,{isGuide:true,modules:1,lessons:0}),/does not read as a guide/);
  assert.equal(copyProblem(plain,{isGuide:false,modules:0,lessons:0}),'');
});

// ---- the pane

const file=(bytes,name)=>({name,size:bytes.length,arrayBuffer:async()=>bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.length)});
const base64Of=bytes=>Buffer.from(bytes).toString('base64');
const withZip=async()=>{
  const a=await started('Reframing conflict');
  await addModule(a,'Basics');await addLesson(a,'Start');
  return a;
};
test('Save Backup and Restore Backup are tools of the guide now, each with its own panel and nothing else showing',async()=>{
  const a=await withZip();
  for(const [id,section,other] of [['saveBackup','guide-sec-backup','guide-sec-restore'],['restoreBackup','guide-sec-restore','guide-sec-backup']]){
    a.p.useTool(id);
    assert.equal(a.n(section).hidden,false,`${id} shows its panel`);
    assert.equal(a.n(other).hidden,true,`${id} hides the other`);
    assert.equal(a.n('guide-sec-module').hidden,true);
  }
});

test('Save a backup takes a copy of the whole document, reads it back, checks it against the guide, and downloads it',async()=>{
  const a=await withZip();
  const bytes=await docx([[STYLE_IDS.guideTitle,'Reframing conflict'],[STYLE_IDS.module,'Module 1: Basics'],[STYLE_IDS.lesson,'Lesson 1.1: Start'],[STYLE_IDS.time,' Time: 10 minutes. Cumulative time: 0 minutes.']],{title:'Reframing conflict'});
  a.adapter.setCapturedFile(bytes);
  a.p.useTool('saveBackup');
  await click(a,'guide-bk-save');
  assert.equal(a.err(),'');
  const saved=a.adapter.downloads();
  assert.equal(saved.length,1);
  assert.match(saved[0].fileName,/^Reframing conflict - Backup \d{4}-\d{2}-\d{2} \d{4}\.docx$/);
  assert.equal(saved[0].base64,base64Of(bytes),'the file offered is the copy Word made, byte for byte');
  assert.match(a.n('guide-status').textContent,/Backup made: "Reframing conflict - Backup .*\.docx" \(\d+ KB\)\. It was checked: it opens as a Word file and holds 1 module, 1 lesson, 0 slides and 4 paragraphs, the same guide as the document\. Look in your Downloads folder.*Restore Backup\./);
});

test('a copy that does not match the document is not saved, and says so',async()=>{
  const a=await withZip();
  a.adapter.setCapturedFile(await docx([[STYLE_IDS.guideTitle,'Reframing conflict'],[STYLE_IDS.module,'Module 1: Basics']]));
  a.p.useTool('saveBackup');
  await click(a,'guide-bk-save');
  assert.match(a.err(),/has 1 module and 0 lessons, but the document has 1 module and 1 lesson, so the copy was not saved/);
  assert.equal(a.adapter.downloads().length,0);
});

test('when Word cannot give the file, the author is told and nothing is downloaded',async()=>{
  const a=await withZip();
  a.p.useTool('saveBackup');
  await click(a,'guide-bk-save');
  assert.match(a.err(),/no Word document to back up/);
  assert.equal(a.adapter.downloads().length,0);
});

test('a saved backup is read and described, then opens as a separate document',async()=>{
  const a=await withZip();
  const bytes=await docx(GUIDE);
  a.p.useTool('restoreBackup');
  a.n('guide-bk-file').files=[file(bytes,'Reframing conflict - Backup 2026-11-08 0915.docx')];
  await a.n('guide-bk-file').onchange();
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-bk-found').hidden,false);
  const facts=a.n('guide-bk-facts').children.map(li=>li.textContent);
  assert.equal(facts[0],'Title: Reframing conflict');assert.equal(facts[1],'2 modules, 2 lessons, 1 slide.');
  assert.match(a.n('guide-status').textContent,/is a readable backup\. Check that this is the one you want, then choose Open this backup in a new window\./);
  assert.equal(a.adapter.openedDocuments().length,0,'choosing the file opens nothing');
  await click(a,'guide-bk-open');
  assert.equal(a.err(),'');
  const opened=a.adapter.openedDocuments();
  assert.equal(opened.length,1);
  assert.equal(opened[0].base64,base64Of(bytes));
  assert.match(a.n('guide-status').textContent,/open in a new Word window\. The document you were working in has not changed/);
  assert.equal(a.adapter.snapshot().some(p=>/Module 2: More/.test(p.text)),false,'the working document was not touched');
});

test('a file that is not a readable backup is refused in plain words and offers nothing to open',async()=>{
  const a=await withZip();
  a.p.useTool('restoreBackup');
  a.n('guide-bk-file').files=[file(new Uint8Array([1,2,3,4,5,6]),'notes.docx')];
  await a.n('guide-bk-file').onchange();
  assert.match(a.err(),/not a Word document that can be opened here/);
  assert.equal(a.n('guide-bk-found').hidden,true);
  // and a good file chosen after it works again
  a.n('guide-bk-file').files=[file(await docx(GUIDE),'good.docx')];
  await a.n('guide-bk-file').onchange();
  assert.equal(a.n('guide-bk-found').hidden,false);
  // an older backup package that is not one is told so by the older reader's own words
  a.n('guide-bk-file').files=[file(await new JSZip().file('x.txt','y').generateAsync({type:'uint8array'}),'old.zip')];
  await a.n('guide-bk-file').onchange();
  assert.match(a.err(),/backup manifest is missing/);
  assert.equal(a.n('guide-bk-found').hidden,true);
  await click(a,'guide-bk-open');
  assert.match(a.err(),/Choose a backup file first/);
});
