// Timing & Agenda on the guide (rebuilt 2026-11-07 so it needs no older course record): clock times for each lesson and the timed
// blocks in it, from a start time, read from the document. Drives the real pane module in memory (not native Word) and checks the
// Word file it makes.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {guideAgenda,agendaSummary,agendaFiles,agendaFileName,START_TIME,blocksOf} from '../src/guide-agenda.mjs';
import {readGuide} from '../src/guide-model.mjs';
import {started,click,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const memoryStorage=()=>{const data={};return {data,getItem:k=>k in data?data[k]:null,setItem:(k,v)=>{data[k]=String(v);}};};
async function planned(opts){
  const a=await started('Agenda course',opts);
  await addModule(a,'Reframing');
  await addLesson(a,'Opening','20');
  await addLesson(a,'Practice','30');
  await addBlock(a,'say','Warm up',{minutes:'5'});
  await addBlock(a,'activity','Role play',{minutes:'25'});
  await addModule(a,'Closing');
  await addLesson(a,'Wrap-up','10');
  return a;
}
const rowsOf=a=>a.n('guide-agenda-body').children.map(tr=>tr.children.map(c=>c.textContent));
const open=async a=>{a.p.useTool('openDesigner');await a.p.openTool('timingAndAgenda');};

test('the start time must be a clock time, and the agenda says so otherwise',async()=>{
  const a=await planned();
  const paras=await a.adapter.paragraphs();
  for(const bad of ['','9','25:00','09:60','nine','09:00:00'])assert.throws(()=>guideAgenda(paras,bad),/Choose a valid start time\./,bad);
  for(const good of ['00:00','9:05','23:59','09:00'])assert.ok(START_TIME.test(good),good);
});

test('each lesson starts when the one before ends, and timed blocks are laid out inside their lesson',async()=>{
  const a=await planned();
  const g=guideAgenda(await a.adapter.paragraphs(),'09:00');
  assert.deepEqual(g.rows.filter(r=>r.kind==='module').map(r=>r.text),['Module 1: Reframing','Module 2: Closing']);
  const lessons=g.rows.filter(r=>r.kind==='lesson').map(r=>[r.text,r.start,r.end,r.minutes]);
  assert.deepEqual(lessons,[['Lesson 1.1: Opening','09:00','09:20',20],['Lesson 1.2: Practice','09:20','09:50',30],['Lesson 2.1: Wrap-up','09:50','10:00',10]]);
  const blocks=g.rows.filter(r=>r.kind==='block').map(r=>[r.text,r.start,r.end,r.minutes]);
  assert.deepEqual(blocks,[['Say: Warm up','09:20','09:25',5],['Activity: Role play','09:25','09:50',25]]);
  assert.equal(g.lessons,3);assert.equal(g.minutes,60);assert.equal(g.start,'09:00');assert.equal(g.end,'10:00');
  assert.deepEqual(g.notes,[]);
  assert.equal(agendaSummary(g),'3 lessons, 60 minutes, 09:00 to 10:00.');
});

test('a day that runs past midnight says which day, and a lesson with no minutes or a time line that disagrees is noted',async()=>{
  const a=await started('Late');
  await addModule(a,'Night');
  await addLesson(a,'Long one','90');
  await addLesson(a,'No time','0');
  await addLesson(a,'Disagree','15');
  await addBlock(a,'say','Ten only',{minutes:'10'});
  const g=guideAgenda(await a.adapter.paragraphs(),'23:30');
  const [first,second,third]=g.rows.filter(r=>r.kind==='lesson');
  assert.deepEqual([first.start,first.end],['23:30','Day 2 · 01:00']);
  assert.equal(second.minutes,0);
  assert.equal(third.minutes,10,'a lesson with block minutes takes their sum');
  assert.deepEqual(g.notes.map(n=>n.text),[
    'Lesson 1.2: No time has no minutes, so it takes no time here. Give it a time, or minutes on its blocks.',
    'Lesson 1.3: Disagree: its time line says 15 minutes but its blocks add up to 10. The agenda uses 10. Choose Update guide to bring them together.'
  ]);
  assert.equal(g.end,'Day 2 · 01:10');
});

test('a guide with no lessons has nothing to schedule, and says what to do',async()=>{
  const a=await started('Empty');
  const g=guideAgenda(await a.adapter.paragraphs(),'09:00');
  assert.equal(g.lessons,0);assert.equal(g.minutes,0);
  assert.match(agendaSummary(g),/There are no lessons yet\. Add a module and a lesson, then give them minutes\./);
});

test('the tool opens on its own panel with the agenda already worked out, as a table with a header row',async()=>{
  const a=await planned();await open(a);
  assert.equal(a.n('guide-sec-agenda').hidden,false);assert.equal(a.n('guide-sec-module').hidden,true);
  assert.equal(a.n('guide-agenda-summary').textContent,'3 lessons, 60 minutes, 09:00 to 10:00.');
  assert.equal(a.n('guide-agenda-table').hidden,false);
  assert.deepEqual(rowsOf(a),[
    ['Module 1: Reframing'],
    ['09:00 to 09:20','Lesson 1.1: Opening','20'],
    ['09:20 to 09:50','Lesson 1.2: Practice','30'],
    ['09:20 to 09:25','Say: Warm up','5'],
    ['09:25 to 09:50','Activity: Role play','25'],
    ['Module 2: Closing'],
    ['09:50 to 10:00','Lesson 2.1: Wrap-up','10']
  ]);
  assert.equal(a.n('guide-agenda-body').children[0].children[0].attributes.scope,'rowgroup','module rows are row group headers');
  assert.equal(a.err(),'');
});

test('a new start time moves every time, and is remembered for next time',async()=>{
  const storage=memoryStorage();
  const a=await planned({storage});await open(a);
  a.n('guide-agenda-start').value='13:15';a.n('guide-agenda-start').onchange();
  assert.equal(a.n('guide-agenda-summary').textContent,'3 lessons, 60 minutes, 13:15 to 14:15.');
  assert.deepEqual(rowsOf(a)[1],['13:15 to 13:35','Lesson 1.1: Opening','20']);
  assert.equal(storage.data['osg.agendaStart.v1'],'13:15');
  const b=await planned({storage});
  assert.equal(b.n('guide-agenda-start').value,'13:15','the next pane starts where the author left it');
  a.n('guide-agenda-start').value='';a.n('guide-agenda-start').onchange();
  assert.equal(a.n('guide-agenda-summary').textContent,'3 lessons, 60 minutes, 09:00 to 10:00.','an empty start time is the usual 09:00');
});

test('the agenda is made as a Word file in a new window: a table for each module, header rows, lessons first and blocks indented',async()=>{
  const a=await planned();await open(a);
  await click(a,'guide-agenda-download');
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/The agenda is open in a new Word window\. Save it as "Agenda course - agenda\.docx"/);
  const doc=a.adapter.openedDocuments().at(-1);
  assert.equal(doc.fileName,'Agenda course - agenda.docx');
  const zip=await JSZip.loadAsync(doc.base64,{base64:true});
  assert.deepEqual(Object.keys(zip.files).sort(),['[Content_Types].xml','_rels/.rels','docProps/core.xml','word/_rels/document.xml.rels','word/document.xml','word/styles.xml']);
  const xml=await zip.file('word/document.xml').async('string');
  const dom=new DOMParser().parseFromString(xml,'text/xml');
  assert.equal(dom.getElementsByTagName('w:tbl').length,2,'one table for each module');
  assert.equal(dom.getElementsByTagName('w:tblHeader').length,2,'each with a header row that repeats');
  const text=Array.from(dom.getElementsByTagName('w:t')).map(t=>t.textContent).join('|');
  assert.ok(text.includes('Agenda course - agenda'));assert.ok(text.includes('60 minutes · 09:00 to 10:00'));
  for(const want of ['Time','What','Minutes','Module 1: Reframing','09:20 to 09:50','Lesson 1.2: Practice','Say: Warm up','Activity: Role play','09:50 to 10:00'])assert.ok(text.includes(want),want);
  assert.doesNotMatch(await zip.file('docProps/core.xml').async('string'),/Project ID|revision/i);
});

test('the Word file refuses an empty guide, and an agenda that changed since it was shown is shown again instead',async()=>{
  const e=await started('Empty');await open(e);
  await click(e,'guide-agenda-download');
  assert.match(e.err(),/There are no lessons to put in an agenda yet\./);
  const a=await planned();await open(a);
  await addLesson(a,'Surprise','15');                              // the guide changed after the agenda was shown
  a.n('guide-agenda-start').value='10:00';                          // and so did the start time, without the pane being told
  await click(a,'guide-agenda-download');
  assert.match(a.err(),/changed since the agenda was shown\. It is shown again above/);
  assert.equal(a.adapter.openedDocuments().length,0,'nothing was made from an agenda the author had not seen');
  assert.match(a.n('guide-agenda-summary').textContent,/10:00 to /);
});

test('the file name is safe, and the file reads blocks only inside lessons',async()=>{
  assert.equal(agendaFileName('Fall: "Wrench" / 2026?'),'Fall Wrench  2026 - agenda.docx');
  assert.equal(agendaFileName(''),'Guide - agenda.docx');
  const a=await planned();
  const paras=await a.adapter.paragraphs(),guide=readGuide(paras);
  assert.deepEqual(blocksOf(paras,guide.modules[0].lessons[1]).map(b=>[b.label,b.title,b.minutes]),[['Say','Warm up',5],['Activity','Role play',25]]);
  assert.equal(agendaFiles(guideAgenda(paras,'09:00')).length,6);
});
