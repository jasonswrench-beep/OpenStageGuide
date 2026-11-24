// The run sheet and the preparation checklist (src/session-materials.mjs): made from a guide that holds a session template, read back as Word files.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {runSheetModel,runSheetFiles,runSheetFileName,prepModel,prepFiles,prepFileName,supplyKind} from '../src/session-materials.mjs';
import {textOf} from '../src/guide-check.mjs';
import {makeContext,buildPackage,moduleHeading,lessonHeading,timeBlock,cueBlock} from '../src/guide-ooxml.mjs';
import {paragraphsOf} from './helpers/guide-package.mjs';
import {started,addModule,addLesson,submit} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const parse=t=>{const errors=[];const d=new DOMParser({errorHandler:{warning:()=>{},error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(t,'application/xml');assert.deepEqual(errors,[]);return d;};
const part=(files,name)=>files.find(f=>f.name===name)?.text;
const paragraphs=files=>Array.from(parse(part(files,'word/document.xml')).getElementsByTagNameNS(W,'p')).map(p=>textOf(p));

async function guide(tpl=0){
  const a=await started('Skill day');
  await addModule(a,'Practice');await addLesson(a,'Warm up','5');
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  a.n('guide-tpl-list').children[tpl].children[2].children[0].onclick();
  a.type('guide-tpl-title','Skill practice');
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  return a;
}

test('a run sheet has the timeline with clock times and, for each activity, its steps with the clock time of each',async()=>{
  const a=await guide();
  const m=runSheetModel(await a.adapter.paragraphs(),{start:'09:00'});
  const lesson=m.lessons.find(l=>l.title==='Lesson 1.2: Skill practice');
  assert.equal(lesson.blocks.length,6,'the six activities have minutes; the objective block has none and is not on the timeline');
  const act=lesson.blocks.filter(b=>b.detail);
  assert.equal(act.length,6);
  const rehearsal=act.find(b=>b.title==='Scenario rehearsal');
  assert.equal(rehearsal.phase,'Practice');
  assert.equal(rehearsal.minutes,18);
  assert.deepEqual(rehearsal.steps.map(s=>s.min),[3,6,6,3]);
  assert.deepEqual(rehearsal.supplies,['Scenario','Cue card','Observation checklist']);
  assert.ok(rehearsal.access.includes('Offer a written simulation.'));
  assert.equal(rehearsal.output,'A rehearsed response and one revision');
  assert.equal(rehearsal.notes.length,0,'private notes are only in it when asked for');
});

test('the run sheet Word file is well formed, has real tables with header rows and bullets, and the steps start at the right clock times',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const files=runSheetFiles(runSheetModel(paras,{start:'09:00'}));
  for(const f of files)if(/\.(xml|rels)$/.test(f.name))parse(f.text);
  const doc=part(files,'word/document.xml');
  assert.match(doc,/<w:tblHeader\/>/);
  assert.match(doc,/<w:numPr>/);
  assert.match(doc,/w:tblCaption w:val="Timeline of Lesson 1.2: Skill practice"/);
  const text=paragraphs(files).join('\n');
  assert.match(text,/Run sheet: Skill day/);
  assert.match(text,/Skill practice/);
  assert.match(text,/Supplies:/);assert.match(text,/Debrief prompts:/);assert.match(text,/If something goes wrong:/);
  assert.match(text,/Other ways to take part, such as without a screen:/);
  assert.doesNotMatch(text,/Private notes/);
  // Skill practice starts after the 5-minute lesson one: 09:05. Scenario rehearsal starts 17 minutes in, its steps at 09:22, 09:25, 09:31, 09:37.
  assert.match(text,/09:22 Scenario rehearsal \(18 min\)/);
  const cells=Array.from(parse(doc).getElementsByTagNameNS(W,'tc')).map(c=>textOf(c));
  const at=cells.findIndex(c=>c==='Assign performer, counterpart, and observer');
  assert.deepEqual(cells.slice(at-2,at+1),['09:22','3','Assign performer, counterpart, and observer']);
  assert.deepEqual(cells.slice(at+1,at+4),['09:25','6','Run the scenario and record evidence']);
  assert.equal(runSheetFileName('Skill: day'),'Skill day - run sheet.docx');
});

test('private notes are in the run sheet only when they are asked for, and are said to be private',()=>{
  const ctx=makeContext();
  const xml=moduleHeading(1,'M')+lessonHeading(1,1,'L')+timeBlock({minutes:5,cumulative:0},ctx)
    +cueBlock({cue:'activity',text:'Round',minutes:5,parts:{instructions:'Go.',steps:'5 min: Go around'},notes:'Keep a spare scenario in your bag.'},ctx)+cueBlock({cue:'transition',text:'On.'},ctx);
  const paras=paragraphsOf(buildPackage(xml,ctx,Object.fromEntries([...ctx.used].map(i=>[i,'AAAA']))));
  const plain=paragraphs(runSheetFiles(runSheetModel(paras,{start:'09:00'}))).join('\n');
  const noted=paragraphs(runSheetFiles(runSheetModel(paras,{start:'09:00',includeNotes:true}))).join('\n');
  assert.doesNotMatch(plain,/Keep a spare scenario/);
  assert.doesNotMatch(plain,/Private notes/);
  assert.match(noted,/Private notes \(facilitator only\):/);
  assert.match(noted,/Keep a spare scenario in your bag\./);
  assert.match(noted,/This copy includes the facilitator's private notes\./);
});
test('one lesson can be made on its own',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const lessonAt=paras.findIndex(p=>/^Lesson 1\.2/.test(p.text));
  const m=runSheetModel(paras,{start:'09:00',lessonIndex:lessonAt});
  assert.equal(m.lessons.length,1);
  assert.equal(m.lessons[0].title,'Lesson 1.2: Skill practice');
  assert.equal(m.minutes,60);
});

test('supplies are told apart as printed, brought or opened on a screen, and a word of print wins',()=>{
  assert.equal(supplyKind('Scenario cards'),'Print');
  assert.equal(supplyKind('One-page decision log'),'Print');
  assert.equal(supplyKind('Shared board or worksheet'),'Print');
  assert.equal(supplyKind('Agenda and outcome slide'),'Open digitally');
  assert.equal(supplyKind('Shared text document'),'Open digitally');
  assert.equal(supplyKind('Dot stickers or accessible ballot'),'Print');
  assert.equal(supplyKind('Markers'),'Bring');
  assert.equal(supplyKind('A timer'),'Bring');
});

test('the preparation checklist lists each supply once, with every activity that uses it, how many to make from the people and teams, and the paper for what is printed',async()=>{
  const a=await guide(1);        // Team or Project Kickoff
  const paras=await a.adapter.paragraphs();
  const m=prepModel(paras,{group:{people:12,teams:3,tables:null},paper:'A4'});
  assert.ok(m.rows.length>=5);
  assert.equal(new Set(m.rows.map(r=>r.item.toLowerCase())).size,m.rows.length,'each once');
  const log=m.rows.find(r=>/action log/i.test(r.item));
  assert.ok(log);
  assert.equal(log.kind,'Print');
  assert.equal(log.paper,'A4');
  assert.equal(log.status,'To do');
  assert.ok(log.activities.length>=1);
  const slide=m.rows.find(r=>/slide/i.test(r.item));
  assert.equal(slide.kind,'Open digitally');assert.equal(slide.paper,'');
  assert.equal(slide.per,'room');assert.equal(slide.quantity,1);
  const pair=m.rows.find(r=>/Prompt card/.test(r.item));
  assert.equal(pair.per,'pair');assert.equal(pair.quantity,6,'twelve people in pairs');
  const groups=m.rows.find(r=>/Shared board or worksheet/.test(r.item));
  assert.equal(groups.per,'group');assert.equal(groups.quantity,3,'three teams');
  // kinds are in the order print, bring, open digitally
  const kinds=m.rows.map(r=>r.kind);
  assert.deepEqual(kinds,[...kinds].sort((x,y)=>['Print','Bring','Open digitally'].indexOf(x)-['Print','Bring','Open digitally'].indexOf(y)));
  // without counts the quantity is left to fill in
  const none=prepModel(paras,{});
  assert.equal(none.counted,false);
  assert.equal(none.rows.find(r=>/Prompt card/.test(r.item)).quantity,null);
  assert.equal(none.rows.find(r=>/slide/i.test(r.item)).quantity,1);
});

test('the preparation checklist Word file has a table for each kind of supply, with header rows, and says when the counts are missing',async()=>{
  const a=await guide(1);
  const paras=await a.adapter.paragraphs();
  const files=prepFiles(prepModel(paras,{}),{title:'Skill day'});
  for(const f of files)if(/\.(xml|rels)$/.test(f.name))parse(f.text);
  const text=paragraphs(files).join('\n');
  assert.match(text,/Preparation checklist: Skill day/);
  assert.match(text,/To print/);assert.match(text,/To open on a screen/);
  assert.match(text,/does not say how many people and teams there are/);
  const doc=part(files,'word/document.xml');
  assert.ok((doc.match(/<w:tblHeader\/>/g)||[]).length>=2);
  assert.match(doc,/w:tblCaption w:val="Supplies to print"/);
  const counted=paragraphs(prepFiles(prepModel(paras,{group:{people:8}}),{title:'x'})).join('\n');
  assert.match(counted,/quantities come from the number of people/);
  assert.equal(prepFileName('Skill: day'),'Skill day - preparation checklist.docx');
  const empty=paragraphs(prepFiles(prepModel([{text:'x',style:'Normal'}],{}),{title:'x'})).join('\n');
  assert.match(empty,/No activity names any supplies yet/);
  // it zips into a Word file
  const zip=new JSZip();for(const f of files)zip.file(f.name,f.text,{createFolders:false});
  assert.ok((await zip.generateAsync({type:'nodebuffer'})).length>1000);
});
