import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {cuesInGroup,cueById,CUE_GROUPS,ICON_FILES} from '../src/guide-cues.mjs';
import {PROMPTS,screensOf,narrationStats,narrationMinutes,elearningIssues,screenBlocks,scriptText,scriptFiles,scriptFileName,WORDS_PER_MINUTE} from '../src/guide-elearning.mjs';
import {makeContext,buildPackage} from '../src/guide-ooxml.mjs';
import {zipWorkbook} from '../src/guide-workbook.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {GUIDE_LOOKS,resolveGuideLook} from '../src/guide-looks.mjs';
import {TINY,paragraphsOf} from './helpers/guide-package.mjs';
import {pane,started,click,submit,guideRows,addModule,addLesson} from './helpers/guide-harness.mjs';

const P=(style,text)=>({style,text});
const MOD=P('OSG Module','Module 1: Basics'),LES=(n,t)=>P('OSG Lesson',`Lesson 1.${n}: ${t}`),CUE=(label,text,fac=false)=>P(fac?'OSG Cue Facilitator':'OSG Cue',`${label}\t${text}`);
const FAC=new Set(['Narration','Visual direction','Branch','Animation note','Developer note']);
const cue=(label,text)=>CUE(label,text,FAC.has(label));
const iconData=Object.fromEntries(ICON_FILES.map(i=>[i,TINY]));

test('the E-Learning group has the twelve cues of the plan, and the designer-only ones are facilitator only',()=>{
  const ids=cuesInGroup('elearning').map(c=>c.id);
  assert.deepEqual(ids,['screen-title','visual-direction','on-screen-text','narration','interaction','knowledge-check','feedback','branch','animation-note','media-asset','accessibility-note','developer-note']);
  for(const id of ['narration','visual-direction','animation-note','developer-note','branch'])assert.equal(cueById(id).audience,'facilitator',id);
  for(const id of ['screen-title','on-screen-text','interaction','knowledge-check','feedback','media-asset','accessibility-note'])assert.equal(cueById(id).audience,'everyone',id);
  assert.deepEqual(CUE_GROUPS.map(g=>g.id).slice(-2),['elearning','you']);
});

test('screens are read from the blocks: each starts at a Screen title and runs to the next one, a lesson or the Transition',()=>{
  const paras=[
    P('OSG Guide Title','Course'),MOD,LES(1,'Welcome'),P('OSG Time','Time: 5 minutes. Cumulative time: 0 minutes.'),
    cue('Screen title','Welcome (2 min)'),cue('Visual direction','A sunrise over a desk.'),cue('On-screen text','Welcome to the course.'),
    cue('Narration','Hello, and welcome. Today you’ll learn three things.'),P('OSG Cue Body','We start with the first.'),
    cue('Screen title','Your goals'),cue('Narration','What do you want from this course?'),
    LES(2,'Practice'),cue('Screen title','Try it'),cue('Narration','Now you try.'),CUE('Transition','On to module two.',true),
    cue('Narration','This one belongs to no screen.')
  ];
  const screens=screensOf(paras);
  assert.deepEqual(screens.map(s=>[s.lesson,s.number,s.title]),[['Lesson 1.1: Welcome',1,'Welcome'],['Lesson 1.1: Welcome',2,'Your goals'],['Lesson 1.2: Practice',1,'Try it']]);
  assert.equal(screens[0].narration,'Hello, and welcome. Today you’ll learn three things.\nWe start with the first.');
  assert.equal(screens[0].onScreen,'Welcome to the course.');
  assert.equal(screens[0].narrationWords,13,'a contraction is one word');
  assert.equal(screens[2].blocks.length,2,'the Transition ends the screen, and what follows it belongs to none');
  assert.deepEqual(screens[0].blocks.map(b=>b.cue),['screen-title','visual-direction','on-screen-text','narration']);
});

test('narration is counted and timed at a steady pace',()=>{
  assert.equal(WORDS_PER_MINUTE,150);
  assert.equal(narrationMinutes(150),1);assert.equal(narrationMinutes(75),0.5);assert.equal(narrationMinutes(0),0);assert.equal(narrationMinutes(200),1.3);
  const screens=[{lesson:'A',narrationWords:300},{lesson:'A',narrationWords:150},{lesson:'B',narrationWords:75}];
  assert.deepEqual(narrationStats(screens),{screens:3,words:525,minutes:3.5,lessons:[{lesson:'A',words:450,minutes:3},{lesson:'B',words:75,minutes:0.5}]});
});

test('the review names what a screen still lacks, and says nothing about a finished one',()=>{
  const done=[MOD,LES(1,'L'),cue('Screen title','Done'),cue('On-screen text','Words.'),cue('Narration','Say it.'),cue('Interaction','Click a card.'),cue('Feedback','Right!'),cue('Knowledge check','Question 2')];
  assert.deepEqual(elearningIssues(done,new Set([1,2])),[]);
  assert.deepEqual(elearningIssues([MOD,LES(1,'L'),cue('Say','Not a screen.')]),[],'a guide with no screens has nothing to say');
  const rules=paras=>elearningIssues(paras,new Set([1])).map(i=>i.rule);
  assert.deepEqual(rules([MOD,LES(1,'L'),cue('Screen title','Bare')]),['EL-NARRATION','EL-ONSCREEN']);
  assert.deepEqual(rules([MOD,LES(1,'L'),cue('Screen title','S'),cue('On-screen text',PROMPTS.onScreen),cue('Narration','Say it.')]),['EL-ONSCREEN','EL-PROMPT'],'a sample sentence left in place counts as missing');
  assert.deepEqual(rules([...done.slice(0,5),cue('Interaction','Click.')]),['EL-FEEDBACK']);
  assert.deepEqual(rules([...done.slice(0,5),cue('Knowledge check','Check it.')]),['EL-KC-LINK']);
  assert.deepEqual(rules([...done.slice(0,5),cue('Knowledge check','Question 9')]),['EL-KC-MISSING']);
  const lack=elearningIssues([MOD,LES(1,'L'),cue('Screen title','Bare')]);
  assert.match(lack[0].message,/The screen "Bare" has no narration yet\./);assert.equal(lack[0].paragraph,2);
});

test('Add a screen writes the usual run of blocks with a sample sentence under each part',()=>{
  const ctx=makeContext(),pkg=buildPackage(screenBlocks({title:'Welcome',minutes:2,interaction:true},ctx),ctx,iconData),paras=paragraphsOf(pkg);
  const texts=paras.map(p=>p.text.replace(/^[^A-Za-z]+/,'').replace(/\t/g,' | '));
  assert.deepEqual(texts,['Screen title Welcome (2 min)',`Visual direction ${PROMPTS.visual}`,`On-screen text ${PROMPTS.onScreen}`,`Narration ${PROMPTS.narration}`,`Interaction ${PROMPTS.interaction}`,`Feedback ${PROMPTS.feedback}`]);
  assert.deepEqual(paras.map(p=>p.style),['OSG Cue','OSG Cue Facilitator','OSG Cue','OSG Cue Facilitator','OSG Cue','OSG Cue']);
  const c2=makeContext();
  assert.equal(paragraphsOf(buildPackage(screenBlocks({title:'Plain'},c2),c2,iconData)).length,4);
  assert.throws(()=>screenBlocks({title:'  '},makeContext()),/Type a title for the screen first/);
});

test('the script has the lesson, each screen and its narration, and reads as a Word file with real headings',async()=>{
  const screens=screensOf([MOD,LES(1,'Welcome'),cue('Screen title','Hello'),cue('Narration','Welcome everyone. Let us begin.'),cue('Screen title','Silent'),LES(2,'More'),cue('Screen title','Next'),cue('Narration','On we go.')]);
  const text=scriptText(screens,'My course');
  assert.match(text,/^My course - narration script\n3 screens, 8 words, about 0\.1 minutes at 150 words a minute\.\n\nLesson 1\.1: Welcome\n\nScreen 1: Hello\nWelcome everyone\. Let us begin\.\n\nScreen 2: Silent\n\(no narration yet\)\n\nLesson 1\.2: More\n\nScreen 1: Next\nOn we go\.\n$/);
  const bytes=await zipWorkbook(scriptFiles(screens,'My course'),JSZip,'uint8array');
  const result=await checkDocx(bytes);
  assert.equal(result.counts.fail,0,JSON.stringify(result.findings.filter(f=>f.level==='fail')));
  const zip=await JSZip.loadAsync(bytes),doc=await zip.file('word/document.xml').async('string');
  assert.match(doc,/<w:pStyle w:val="Heading1"\/>.*Lesson 1\.1: Welcome/);assert.match(doc,/<w:pStyle w:val="Heading2"\/>.*Screen 1: Hello/);
  assert.equal(scriptFileName('Handling: conflict?','docx'),'Handling - conflict - Narration script.docx');
});

test('the storyboard look is one of the guide\'s looks, with its own accent',()=>{
  assert.ok(GUIDE_LOOKS.some(l=>l.id==='storyboard'));
  const r=resolveGuideLook({id:'storyboard'});
  assert.equal(r.ok,true);assert.equal(r.theme.navy,'0B5D5D');
  assert.equal(resolveGuideLook({id:'storyboard',accent:'#000080'}).theme.navy,'000080','the author\'s own accent wins');
});

async function addScreen(a,title,{minutes='',interaction=false}={}){
  a.type('guide-el-title',title);a.type('guide-el-minutes',minutes);a.n('guide-el-interaction').checked=interaction;
  await click(a,'guide-el-add');
  assert.equal(a.err(),'');
}
const rowIndex=(a,re)=>a.rows().findIndex(r=>re.test(r));

test('the pane adds a screen, counts the narration once it is written, and the review notices what is not',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await addScreen(a,'Welcome',{minutes:'2',interaction:true});
  const rows=guideRows(a);
  assert.deepEqual(rows.filter(r=>/^OSG Cue/.test(r)).map(r=>r.replace(/^OSG Cue( Facilitator)?: [^A-Za-z]*/,'').split(/[\s]/)[0]),['Screen','Visual','On-screen','Narration','Interaction','Feedback','Transition']);
  assert.match(a.n('guide-status').textContent,/Screen "Welcome" added\. Type over the sample sentence under each part\./);
  assert.match(a.n('guide-el-summary').textContent,/^1 screen\. Narration: 0 words, about 0 minutes at 150 words a minute\./);
  // The sample sentences count for nothing; the review lists each one.
  await click(a,'guide-open-review');
  const text=a.n('guide-review-look').children.map(li=>li.textContent).join('\n');
  assert.match(text,/The screen "Welcome" has no narration yet\./);
  assert.match(text,/a block still has the sample sentence/);
  // Type over the narration: it is counted and timed.
  const at=rowIndex(a,/^OSG Cue Facilitator: [^A-Za-z]*Narration/);
  a.adapter.typeOver(at,'Narration\t'+'word '.repeat(75).trim());
  await a.p.refresh();
  assert.match(a.n('guide-el-summary').textContent,/^1 screen\. Narration: 75 words, about 0\.5 minutes at 150 words a minute\./);
});

test('the pane needs a screen with narration before it makes the script, then opens it in a new window or downloads it as text',async()=>{
  const a=await started('My course');
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await click(a,'guide-el-script');
  assert.match(a.err(),/There are no screens yet\. Add a screen first\./);
  await addScreen(a,'Welcome');
  await click(a,'guide-el-script');
  assert.match(a.err(),/No screen has narration yet/);
  a.adapter.typeOver(rowIndex(a,/^OSG Cue Facilitator: [^A-Za-z]*Narration/),'Narration\tWelcome to the course.');
  await a.p.refresh();
  await click(a,'guide-el-script');
  assert.equal(a.err(),'');
  assert.equal(a.adapter.openedDocuments().at(-1).fileName,'My course - Narration script.docx');
  assert.match(a.n('guide-status').textContent,/The narration script is open in a new Word window/);
  await click(a,'guide-el-text');
  const d=a.adapter.downloads().at(-1);
  assert.equal(d.fileName,'My course - Narration script.txt');assert.equal(d.type,'text/plain');
  assert.match(Buffer.from(d.base64,'base64').toString('utf8'),/Screen 1: Welcome\nWelcome to the course\./);
});

test('choosing to make an E-Learning storyboard sets its look, leaves out the facilitator\'s pages and puts the E-Learning cues first',async()=>{
  const a=pane();
  await a.p.refresh();
  assert.deepEqual(a.n('guide-start-kind').children.map(o=>o.value),['','storyboard']);
  a.type('guide-start-kind','storyboard');a.n('guide-start-kind').onchange();
  assert.equal(a.n('guide-standard').checked,false);assert.equal(a.n('guide-start-look').value,'storyboard');
  a.type('guide-title','Safety basics');
  await submit(a,'guide-start-form');
  assert.equal(a.err(),'');
  assert.equal(JSON.parse(a.adapter.lookState().lookText).id,'storyboard');
  assert.equal(a.adapter.lookState().theme.navy,'0B5D5D');
  assert.equal(a.n('guide-cue-group').value,'elearning');
  assert.equal(a.n('guide-cues').children.filter(l=>!l.hidden).length,12);
  assert.ok(!a.rows().some(r=>/Getting Started|Training at a Glance/.test(r)),'no facilitator standard pages');
  // And back: choosing the facilitator's guide restores the usual choices.
  a.type('guide-start-kind','');a.n('guide-start-kind').onchange();
  assert.equal(a.n('guide-standard').checked,true);assert.equal(a.n('guide-start-look').value,'');
});
