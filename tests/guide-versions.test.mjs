// Delivery Versions on the guide (rebuilt 2026-11-07 so it needs no older course record): a shorter or longer session planned from the
// same guide, kept in the document. Drives the real pane module in memory (not native Word) and checks the Word file it makes.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {readVersions,writeVersions,newVersion,toggleLesson,moveLesson,versionAgenda,versionSummary,parseTarget,lessonList,lessonKey,MAX_VERSIONS,VERSIONS_KEY} from '../src/guide-versions.mjs';
import {started,click,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

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
const open=async a=>{a.p.useTool('openDesigner');await a.p.openTool('deliveryVersions');};
const lessonRows=a=>a.n('guide-ver-lessons').children;
const box=li=>li.children[0].children[0];
const buttonOf=(li,text)=>li.children[1].children.find(b=>b.textContent===text);
const tableRows=a=>a.n('guide-ver-body').children.map(tr=>tr.children.map(c=>c.textContent));
const stored=a=>a.adapter.versionMemory();

test('a new version holds every lesson in the guide\'s order, with a name that is not already taken',async()=>{
  const a=await planned(),all=lessonList(await a.adapter.paragraphs());
  const v=newVersion(all,[],{now:()=>1000});
  assert.deepEqual(v.lessons,['Opening','Practice','Wrap-up']);
  assert.equal(v.name,'New version');assert.equal(v.target,0);assert.equal(v.start,'09:00');
  assert.equal(newVersion(all,[v]).name,'New version 2');
  assert.throws(()=>newVersion(all,Array.from({length:MAX_VERSIONS},(_,i)=>({...v,id:'x'+i,name:'n'+i}))),/most\. Delete one first\./);
});

test('lessons are taken in and out, and a lesson put back goes where the guide has it; they can be moved earlier or later',async()=>{
  const a=await planned(),all=lessonList(await a.adapter.paragraphs());
  let v=newVersion(all,[]);
  v=toggleLesson(v,all,'Practice');
  assert.deepEqual(v.lessons,['Opening','Wrap-up']);
  v=toggleLesson(v,all,'Practice');
  assert.deepEqual(v.lessons,['Opening','Practice','Wrap-up'],'back in its place, not at the end');
  v=moveLesson(v,'Wrap-up',-1);assert.deepEqual(v.lessons,['Opening','Wrap-up','Practice']);
  assert.equal(moveLesson(v,'Opening',-1),v,'the first cannot go earlier');
  assert.equal(moveLesson(v,'Practice',1),v,'nor the last later');
  assert.equal(moveLesson(v,'No such lesson',1),v);
  assert.equal(lessonKey('  Wrap-UP '),'wrap-up');
});

test('the version\'s agenda follows its own order and start time, repeats a module heading when the order goes back to it, and compares with the target',async()=>{
  const a=await planned(),paras=await a.adapter.paragraphs(),all=lessonList(paras);
  const base={...newVersion(all,[]),start:'13:00',name:'Half day'};
  const v={...base,lessons:['Wrap-up','Opening'],target:45};
  const g=versionAgenda(paras,v,{title:'Agenda course'});
  assert.equal(g.title,'Agenda course - Half day');
  assert.deepEqual(g.rows.filter(r=>r.kind!=='block').map(r=>r.kind==='module'?r.text:[r.text,r.start,r.end,r.minutes]),[
    'Module 2: Closing',['Lesson 2.1: Wrap-up','13:00','13:10',10],
    'Module 1: Reframing',['Lesson 1.1: Opening','13:10','13:30',20]
  ]);
  assert.equal(g.minutes,30);assert.equal(g.variance,-15);
  assert.equal(versionSummary(g),'2 lessons, 30 minutes, 13:00 to 13:30. 15 minutes under the target of 45.');
  assert.equal(versionSummary(versionAgenda(paras,{...v,target:30},{})),'2 lessons, 30 minutes, 13:00 to 13:30. Exactly on the target of 30 minutes.');
  assert.match(versionSummary(versionAgenda(paras,{...v,target:20},{})),/10 minutes over the target of 20\./);
  assert.match(versionSummary(versionAgenda(paras,{...v,target:0},{})),/No target set\.$/);
  assert.match(versionSummary(versionAgenda(paras,{...v,lessons:[]},{})),/No lessons are in this version yet/);
  const withBlocks=versionAgenda(paras,{...base,lessons:['Practice']},{});
  assert.deepEqual(withBlocks.rows.filter(r=>r.kind==='block').map(r=>[r.text,r.start,r.end]),[['Say: Warm up','13:00','13:05'],['Activity: Role play','13:05','13:30']]);
  assert.throws(()=>versionAgenda(paras,{...base,start:'nope'},{}),/Choose a valid start time\./);
});

test('a lesson that is no longer in the guide is named and left out, not guessed at',async()=>{
  const a=await planned(),paras=await a.adapter.paragraphs(),all=lessonList(paras);
  const g=versionAgenda(paras,{...newVersion(all,[]),lessons:['Opening','Deleted lesson']},{});
  assert.deepEqual(g.missing,['Deleted lesson']);
  assert.deepEqual(g.notes.map(n=>n.text),['"Deleted lesson" is no longer in the guide, so it is left out of this version.']);
  assert.equal(g.lessons,1);
});

test('versions are read back from the document whole or not at all, and a target is whole minutes from 1 to 10080',()=>{
  assert.deepEqual(readVersions(null),[]);assert.deepEqual(readVersions('not json'),[]);assert.deepEqual(readVersions('{"v":2,"versions":[]}'),[]);
  const good={id:'a',name:'  One  hour ',target:60,start:'09:30',lessons:['Opening',' Practice ','',7]};
  const out=readVersions(JSON.stringify({v:1,versions:[good,{...good},{id:'b',name:''},{name:'No id'},null,{id:'c',name:'Odd',target:-4,start:'25:99',lessons:'x'}]}));
  assert.deepEqual(out,[{id:'a',name:'One hour',target:60,start:'09:30',lessons:['Opening','Practice']},{id:'c',name:'Odd',target:0,start:'09:00',lessons:[]}]);
  assert.deepEqual(readVersions(writeVersions(out)),out,'what is written is what is read');
  assert.equal(parseTarget(''),0);assert.equal(parseTarget(' 180 '),180);assert.equal(parseTarget('10080'),10080);
  for(const bad of ['0','10081','1.5','-3','an hour','1e3'])assert.throws(()=>parseTarget(bad),/Enter a target from 1 to 10080 minutes, or leave it empty\./,bad);
});

test('the tool opens on its own panel, with no version until one is made, and a new version has every lesson in it',async()=>{
  const a=await planned();await open(a);
  assert.equal(a.n('guide-sec-versions').hidden,false);assert.equal(a.n('guide-sec-module').hidden,true);
  assert.equal(a.n('guide-ver-edit').hidden,true);assert.equal(a.n('guide-ver-delete').disabled,true);
  await click(a,'guide-ver-new');
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-ver-edit').hidden,false);
  assert.equal(a.n('guide-ver-name').value,'New version');
  assert.equal(lessonRows(a).length,3);
  assert.ok(lessonRows(a).every(li=>box(li).checked));
  assert.match(lessonRows(a)[0].children[0].children[1].textContent,/Lesson 1\.1: Opening \(20 min\)/);
  assert.equal(a.n('guide-ver-summary').textContent,'3 lessons, 60 minutes, 09:00 to 10:00. No target set.');
  assert.deepEqual(tableRows(a)[1],['09:00 to 09:20','Lesson 1.1: Opening','20']);
  assert.ok(JSON.parse(await stored(a)).versions.length===1,'kept in the document');
});

test('the target, the start time and the lessons all change the plan, and each change is kept in the document',async()=>{
  const a=await planned();await open(a);await click(a,'guide-ver-new');
  a.n('guide-ver-target').value='45';await a.n('guide-ver-target').onchange();
  assert.equal(a.n('guide-ver-summary').textContent,'3 lessons, 60 minutes, 09:00 to 10:00. 15 minutes over the target of 45.');
  a.n('guide-ver-start').value='14:00';await a.n('guide-ver-start').onchange();
  assert.match(a.n('guide-ver-summary').textContent,/14:00 to 15:00/);
  await box(lessonRows(a)[2]).onchange();                               // take Wrap-up out
  assert.equal(a.n('guide-ver-summary').textContent,'2 lessons, 50 minutes, 14:00 to 14:50. 5 minutes over the target of 45.');
  assert.equal(box(lessonRows(a).at(-1)).checked,false,'a lesson that is out is listed after the ones that are in');
  await buttonOf(lessonRows(a)[0],'Later').onclick();
  assert.deepEqual(tableRows(a).filter(r=>r.length===3).map(r=>r[1]),['Lesson 1.2: Practice','Say: Warm up','Activity: Role play','Lesson 1.1: Opening']);
  const kept=readVersions(await stored(a))[0];
  assert.deepEqual([kept.target,kept.start,kept.lessons],[45,'14:00',['Practice','Opening']]);
  assert.equal(buttonOf(lessonRows(a)[0],'Earlier').disabled,true);assert.equal(buttonOf(lessonRows(a)[1],'Later').disabled,true);
  assert.equal(a.err(),'');
});

test('a name must be given and be new; a target must be a whole number; a bad entry puts the old one back and says why',async()=>{
  const a=await planned();await open(a);await click(a,'guide-ver-new');
  a.n('guide-ver-name').value='   ';await a.n('guide-ver-name').onchange();
  assert.match(a.err(),/A version needs a name\./);assert.equal(a.n('guide-ver-name').value,'New version');
  a.n('guide-ver-name').value='One hour';await a.n('guide-ver-name').onchange();
  assert.equal(a.err(),'');assert.deepEqual(a.n('guide-ver-select').children.map(o=>o.textContent),['One hour']);
  await click(a,'guide-ver-new');
  assert.equal(a.n('guide-ver-name').value,'New version');
  a.n('guide-ver-name').value='one HOUR';await a.n('guide-ver-name').onchange();
  assert.match(a.err(),/Another version already has that name\./);assert.equal(a.n('guide-ver-name').value,'New version');
  a.n('guide-ver-target').value='about 90';await a.n('guide-ver-target').onchange();
  assert.match(a.err(),/Enter a target from 1 to 10080 minutes/);assert.equal(a.n('guide-ver-target').value,'');
  a.n('guide-ver-start').value='';await a.n('guide-ver-start').onchange();
  assert.equal(a.n('guide-ver-start').value,'09:00','an empty start time is the usual 09:00');
});

test('versions are chosen from a list, and deleting one takes two presses so it is never an accident',async()=>{
  const a=await planned();await open(a);
  await click(a,'guide-ver-new');a.n('guide-ver-name').value='Three hours';await a.n('guide-ver-name').onchange();
  await click(a,'guide-ver-new');a.n('guide-ver-name').value='One hour';await a.n('guide-ver-name').onchange();
  assert.deepEqual(a.n('guide-ver-select').children.map(o=>o.textContent),['Three hours','One hour']);
  a.n('guide-ver-select').value=a.n('guide-ver-select').children[0].value;a.n('guide-ver-select').onchange();
  assert.equal(a.n('guide-ver-name').value,'Three hours');
  await click(a,'guide-ver-delete');
  assert.equal(a.n('guide-ver-delete').textContent,'Yes, delete it');
  assert.match(a.n('guide-status').textContent,/Choose "Yes, delete it" to delete Three hours\. Anything else keeps it\./);
  assert.equal(readVersions(await stored(a)).length,2,'nothing is deleted by the first press');
  a.n('guide-ver-target').value='60';await a.n('guide-ver-target').onchange();   // doing anything else cancels it
  assert.equal(a.n('guide-ver-delete').textContent,'Delete this version');
  await click(a,'guide-ver-delete');await click(a,'guide-ver-delete');
  assert.match(a.n('guide-status').textContent,/Three hours was deleted\./);
  assert.deepEqual(readVersions(await stored(a)).map(v=>v.name),['One hour']);
  assert.equal(a.n('guide-ver-name').value,'One hour','the other version is chosen');
});

test('the versions are still there in the next pane on the same document, and a lesson renamed since is reported',async()=>{
  const a=await planned();await open(a);await click(a,'guide-ver-new');
  a.n('guide-ver-name').value='Keep me';await a.n('guide-ver-name').onchange();
  await a.p.refresh();await open(a);
  assert.deepEqual(a.n('guide-ver-select').children.map(o=>o.textContent),['Keep me']);
  await addLesson(a,'Added later','5');
  await a.p.refresh();
  assert.equal(lessonRows(a).length,4,'a lesson added since is listed, not in the version');
  assert.equal(box(lessonRows(a).at(-1)).checked,false);
});

test('Word is told to keep a change, and if it will not, the pane goes back to what it had',async()=>{
  const a=await planned();await open(a);await click(a,'guide-ver-new');
  const real=a.adapter.setVersionMemory;
  a.adapter.setVersionMemory=async()=>{throw new Error('The document did not keep your delivery versions.');};
  a.n('guide-ver-target').value='60';await a.n('guide-ver-target').onchange();
  assert.match(a.err(),/The document did not keep your delivery versions\./);
  assert.equal(a.n('guide-ver-target').value,'','the screen shows what is kept, not what was refused');
  a.adapter.setVersionMemory=real;
  a.n('guide-ver-target').value='60';await a.n('guide-ver-target').onchange();
  assert.equal(a.n('guide-ver-target').value,'60');
});

test('the version\'s agenda is made as a Word file named for the guide and the version',async()=>{
  const a=await planned();await open(a);
  await click(a,'guide-ver-make');
  assert.match(a.err(),/Choose or make a version first\./);
  await click(a,'guide-ver-new');a.n('guide-ver-name').value='Half day';await a.n('guide-ver-name').onchange();
  await box(lessonRows(a)[0]).onchange();                                // without Opening
  await click(a,'guide-ver-make');
  assert.equal(a.err(),'');
  const doc=a.adapter.openedDocuments().at(-1);
  assert.equal(doc.fileName,'Agenda course - Half day - agenda.docx');
  const zip=await JSZip.loadAsync(doc.base64,{base64:true});
  const text=Array.from(new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'),'text/xml').getElementsByTagName('w:t')).map(t=>t.textContent).join('|');
  assert.ok(text.includes('Agenda course - Half day - agenda'));
  assert.ok(text.includes('Lesson 1.2: Practice')&&text.includes('Lesson 2.1: Wrap-up'));
  assert.ok(!text.includes('Lesson 1.1: Opening'),'a lesson that is out is not in the file');
  assert.ok(text.includes('40 minutes · 09:00 to 09:40'));
  // a version with nothing in it is refused
  await box(lessonRows(a)[0]).onchange();await box(lessonRows(a)[0]).onchange();
  const count=a.adapter.openedDocuments().length;
  await click(a,'guide-ver-make');
  assert.match(a.err(),/No lessons are in this version yet\. Tick the lessons to teach\./);
  assert.equal(a.adapter.openedDocuments().length,count);
  assert.equal(VERSIONS_KEY,'osg.versions.v1');
});
