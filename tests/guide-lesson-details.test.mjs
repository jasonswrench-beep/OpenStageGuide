// Lesson details (kit slices K1 and K2): the facts under a lesson's time line, written by the Lesson details tool, read back, checked by the Review, kept out of
// the participant workbook, and never in the way of the lesson's blocks. It drives the real pane and the real taskpane.html ids in memory.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {pane,started,submit,addModule,addLesson,addBlock,moveTo,startsWith,click} from './helpers/guide-harness.mjs';
import {readFacts,checkFacts,factsBlock,factsLines,factIssues,METHODS,MODES,emptyFacts,hasFacts} from '../src/guide-facts.mjs';
import {readGuide,updatePlan,guideIssues,lessonMinutes,kindOf} from '../src/guide-model.mjs';
import {readSource,buildWorkbook} from '../src/guide-workbook.mjs';
import {readSettings} from '../src/guide-forms.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {STYLE_NAMES} from '../src/guide-styles.mjs';

const lessonRecord=async(a,text='Lesson 1.1')=>readGuide(await a.adapter.paragraphs()).modules.flatMap(m=>m.lessons).find(l=>l.text.startsWith(text));
const factTexts=async a=>(await a.adapter.paragraphs()).filter(p=>p.style===STYLE_NAMES.fact).map(p=>p.text.trim());
async function lessonGuide(){
  const a=await started('Details guide');
  await addModule(a,'One');await addLesson(a,'First','10');await addLesson(a,'Second','5');
  return a;
}
const open=async(a,lessonText='OSG Lesson: Lesson 1.1')=>{a.p.useTool('lessonDetails');await moveTo(a,startsWith(lessonText));await a.p.refresh();};
const fill=(a,values)=>{for(const [id,v] of Object.entries(values)){const n=a.n(`guide-details-${id}`);if(id==='optional')n.checked=!!v;else n.value=v;}};
const save=async a=>{await submit(a,'guide-details-form');};

test('the facts are written in a fixed order, read back the way they were written, and checked in plain words',()=>{
  const facts={...emptyFacts(),purpose:'Name the feeling',materials:'Flip chart; markers',method:'Discussion',range:{min:15,max:20},mode:'Small groups',group:'4 to 6 people',optional:true,choose:{n:2,m:5},chooseNote:'Pick the ones that fit.'};
  assert.deepEqual(factsLines(facts),['Purpose: Name the feeling','Materials: Flip chart; markers','Method: Discussion','Range: 15 to 20 minutes','Interaction: Small groups','Group size: 4 to 6 people','Optional: This lesson may be left out.','Choose: 2 of 5. Pick the ones that fit.']);
  assert.equal(factsLines(emptyFacts()).length,0);assert.equal(hasFacts(emptyFacts()),false);assert.equal(hasFacts(facts),true);
  assert.throws(()=>checkFacts({method:'Magic'}),/Choose a method from the list/);
  assert.throws(()=>checkFacts({mode:'Crowd'}),/Choose how people work together from the list/);
  assert.throws(()=>checkFacts({range:{min:30,max:20}}),/The most minutes must be more than the least/);
  assert.throws(()=>checkFacts({range:{min:0,max:20}}),/two whole numbers from 1 to 1440/);
  assert.throws(()=>checkFacts({choose:{n:3,m:2}}),/cannot choose more than there are/);
  assert.throws(()=>checkFacts({purpose:'x'.repeat(401)}),/too long/);
  assert.ok(METHODS.includes('Break')&&MODES.includes('Pairs'));
});

test('the Lesson details tool writes the facts under the time line, and opens with them again',async()=>{
  const a=await lessonGuide();
  await open(a);
  assert.match(a.n('guide-details-where').textContent,/^Details of Lesson 1\.1: First\./);
  assert.equal(a.n('guide-sec-details').hidden,false);
  assert.deepEqual(a.n('guide-details-method').children.map(o=>o.value),['',...METHODS]);
  fill(a,{purpose:'Name the feeling',materials:'Flip chart',method:'Discussion',mode:'Pairs',group:'pairs',min:'8',max:'12',optional:true,'choose-n':'1','choose-m':'3','choose-note':'Pick one.'});
  a.n('guide-details-method').value='Discussion';
  await save(a);
  assert.equal(a.err(),'');
  assert.deepEqual(await factTexts(a),['Purpose: Name the feeling','Materials: Flip chart','Method: Discussion','Range: 8 to 12 minutes','Interaction: Pairs','Group size: pairs','Optional: This lesson may be left out.','Choose: 1 of 3. Pick one.']);
  const texts=(await a.adapter.paragraphs()).map(p=>p.text.trim());
  const time=texts.findIndex(t=>/^Time: 10 minutes/.test(t));
  assert.equal(texts[time+1],'Purpose: Name the feeling','the facts follow the time line directly');
  // another lesson, and back: the fields show what each holds
  await open(a,'OSG Lesson: Lesson 1.2');
  assert.match(a.n('guide-details-where').textContent,/Lesson 1\.2: Second/);
  assert.equal(a.n('guide-details-purpose').value,'');
  await open(a,'OSG Lesson: Lesson 1.1');
  assert.equal(a.n('guide-details-purpose').value,'Name the feeling');
  assert.equal(a.n('guide-details-method').value,'Discussion');
  assert.equal(a.n('guide-details-min').value,'8');assert.equal(a.n('guide-details-choose-m').value,'3');
  assert.equal(a.n('guide-details-optional').checked,true);
});

test('saving again replaces the facts instead of adding to them, and saving none takes them away',async()=>{
  const a=await lessonGuide();
  await open(a);
  fill(a,{purpose:'First words',materials:'',method:'',mode:'',group:'',min:'',max:'',optional:false,'choose-n':'','choose-m':'','choose-note':''});
  await save(a);
  assert.deepEqual(await factTexts(a),['Purpose: First words']);
  a.p.useTool('lessonDetails');await a.p.refresh();
  fill(a,{purpose:'Other words',method:'Lecture'});
  a.n('guide-details-method').value='Lecture';
  await save(a);
  assert.deepEqual(await factTexts(a),['Purpose: Other words','Method: Lecture'],'replaced, not stacked');
  await open(a);
  fill(a,{purpose:'',materials:''});a.n('guide-details-method').value='';a.n('guide-details-mode').value='';
  await save(a);
  assert.deepEqual(await factTexts(a),[],'no details, no facts');
  assert.equal(a.err(),'');
});

test('what cannot be written is refused with its reason and nothing changes',async()=>{
  const a=await lessonGuide();
  await open(a);
  fill(a,{purpose:'x',min:'10',max:''});await save(a);
  assert.match(a.err(),/Give both the least and the most minutes/);
  fill(a,{min:'20',max:'10'});await save(a);
  assert.match(a.err(),/The most minutes must be more than the least/);
  fill(a,{min:'',max:'','choose-n':'3','choose-m':'2'});await save(a);
  assert.match(a.err(),/cannot choose more than there are/);
  fill(a,{'choose-n':'2.5','choose-m':'4'});await save(a);
  assert.match(a.err(),/The number to choose must be a whole number/);
  assert.deepEqual(await factTexts(a),[]);
});

test('with the cursor outside a lesson the tool says so and cannot save',async()=>{
  const a=await lessonGuide();
  a.p.useTool('lessonDetails');await moveTo(a,startsWith('OSG Module'));await a.p.refresh();
  assert.match(a.n('guide-details-where').textContent,/Put your cursor in a lesson/);
  assert.equal(a.n('guide-details-save').disabled,true);
});

test('a block added with the cursor on the lesson\'s time line, or on one of its facts, goes after the last fact, and Lesson details still finds the facts',async()=>{
  for(const row of [startsWith('OSG Time'),startsWith('OSG Fact')]){
    const a=await lessonGuide();
    await open(a);fill(a,{purpose:'Why',materials:'Flip chart',method:'Lecture'});a.n('guide-details-method').value='Lecture';await save(a);
    a.p.useTool('instructionalCues');await moveTo(a,row);await a.p.refresh();
    await addBlock(a,'say','Hello everyone.');
    const rows=a.rows(),last=rows.findLastIndex(r=>/OSG Fact/.test(r)),say=rows.findIndex(r=>/Cue( Facilitator)?: .*Hello everyone/.test(r)),time=rows.findIndex(r=>/Time:/.test(r));
    assert.ok(time>=0&&last>time&&say>last,'the block comes after the last fact: '+rows.slice(time,say+1).join(' | '));
    assert.deepEqual(await factTexts(a),['Purpose: Why','Materials: Flip chart','Method: Lecture'],'and the facts are all still read');
    const lesson=await lessonRecord(a);
    assert.equal(readFacts(await a.adapter.paragraphs(),lesson).facts.method,'Lecture','Lesson details finds them');
  }
  // with no facts the Time line is the end of the header, so the block goes straight after it
  const b=await lessonGuide();
  b.p.useTool('instructionalCues');await moveTo(b,startsWith('OSG Time'));await b.p.refresh();
  await addBlock(b,'say','Hello everyone.');
  const rows=b.rows(),time=rows.findIndex(r=>/Time:/.test(r));
  assert.match(rows[time+1],/Hello everyone/);
});

test('the Review notes an Expected answer that has no Ask before it, and says nothing when an Ask comes first',async()=>{
  const a=await lessonGuide();
  a.p.useTool('instructionalCues');await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));await a.p.refresh();
  await addBlock(a,'expected-answer','Look for feelings named.');
  const notes=async()=>guideIssues(updatePlan(await a.adapter.paragraphs())).filter(i=>i.rule==='CUE-EXPECTED');
  let found=await notes();
  assert.equal(found.length,1);assert.match(found[0].message,/Expected answer in Lesson 1\.1: First has no Ask before it/);
  assert.equal(found[0].level,'warn');assert.ok(Number.isInteger(found[0].paragraph),'it points at the paragraph');
  // an Ask added at the top of the lesson comes before it: the note goes
  await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));await a.p.refresh();
  await addBlock(a,'ask','What makes it hard?');
  found=await notes();
  assert.equal(found.length,0,'with the Ask before it, there is nothing to note');
  // an Ask and then two Expected answers: both are answers to that Ask
  const b=await lessonGuide();
  b.p.useTool('instructionalCues');await moveTo(b,startsWith('OSG Lesson: Lesson 1.1'));await b.p.refresh();
  await addBlock(b,'ask','What makes it hard?');
  await addBlock(b,'expected-answer','Look for feelings named.');
  await addBlock(b,'expected-answer','Or a story.');
  assert.equal(guideIssues(updatePlan(await b.adapter.paragraphs())).filter(i=>i.rule==='CUE-EXPECTED').length,0,'an Ask and then its Expected answers: nothing to say');
});

test('blocks added to a lesson go after its facts, and a lesson moves and is deleted with them',async()=>{
  const a=await lessonGuide();
  await open(a);fill(a,{purpose:'Why',method:'Lecture'});a.n('guide-details-method').value='Lecture';await save(a);
  a.p.useTool('instructionalCues');await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));await a.p.refresh();
  await addBlock(a,'say','Hello everyone.');
  const rows=a.rows(),purpose=rows.findIndex(r=>/Fact: Purpose: Why/.test(r)),say=rows.findIndex(r=>/Cue( Facilitator)?: .*Hello everyone/.test(r));
  assert.ok(purpose>0&&say>purpose,'the block comes after the facts, never between the time line and them: '+rows.slice(purpose-2,say+1).join(' | '));
  const guide=readGuide(await a.adapter.paragraphs()),lesson=guide.modules[0].lessons[0];
  assert.ok(lesson.headerEnd>lesson.timeIndex,'the header runs from the time line to the last fact');
  assert.equal(lesson.end>=lesson.headerEnd,true);
  assert.equal(lessonMinutes(lesson),10,'the facts do not change the minutes');
});

test('the Review says what is wrong with a method, a way of working together, a choice and a range, and nothing for a lesson with no facts',async()=>{
  const a=await lessonGuide();
  const list=await a.adapter.paragraphs(),time=list.findIndex(p=>p.style===STYLE_NAMES.time);
  const P=(text,style=STYLE_NAMES.fact)=>({style,text});
  const paras=[...list.slice(0,time+1),P('Method: Magic'),P('Interaction: Crowd'),P('Choose: 4 of 2.'),P('Range: 15 to 20 minutes'),...list.slice(time+1)];
  const issues=guideIssues(updatePlan(paras)).filter(i=>/^LESSON-(METHOD|MODE|CHOOSE|RANGE)$/.test(i.rule));
  assert.deepEqual(issues.map(i=>i.rule).sort(),['LESSON-CHOOSE','LESSON-METHOD','LESSON-MODE','LESSON-RANGE']);
  assert.match(issues.find(i=>i.rule==='LESSON-RANGE').message,/is planned for 10 minutes, outside its range of 15 to 20/);
  assert.equal(guideIssues(updatePlan(list)).filter(i=>/^LESSON-(METHOD|MODE|CHOOSE|RANGE)$/.test(i.rule)).length,0);
  assert.equal(factIssues(list,readGuide(list),lessonMinutes).length,0);
});

test('the facts never reach the participant workbook, and the leak check watches for them',async()=>{
  const a=await lessonGuide();
  await open(a);fill(a,{purpose:'Facilitator purpose words',materials:'Secret flip chart',method:'Lecture'});a.n('guide-details-method').value='Lecture';await save(a);
  a.p.useTool('instructionalCues');await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));await a.p.refresh();
  await addBlock(a,'ask','What makes it hard?');
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},{DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const doc=built.files.find(f=>f.name==='word/document.xml').text;
  assert.ok(doc.includes('What makes it hard?'));
  assert.ok(!doc.includes('Facilitator purpose words')&&!doc.includes('Secret flip chart')&&!/OSGFact/.test(doc),'none of the facts is in the file');
  assert.ok(built.facilitatorTexts.some(t=>t.includes('Facilitator purpose words')),'the leak check looks for them');
  const leaked=built.files.map(f=>f.name==='word/document.xml'?{...f,text:f.text.replace('</w:body>','<w:p><w:r><w:t>Materials: Secret flip chart</w:t></w:r></w:p></w:body>')}:f);
  assert.equal(leakCheck({files:leaked,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser}).passed,false,'a fact that reached the workbook is caught');
});

test('a fact label the guide does not know is kept when the facts are saved again',async()=>{
  const a=await lessonGuide();
  const list=await a.adapter.paragraphs(),time=list.findIndex(p=>p.style===STYLE_NAMES.time);
  const lesson=readGuide(list).modules[0].lessons[0];
  const withExtra=[...list.slice(0,time+1),{style:STYLE_NAMES.fact,text:'Purpose: Mine'},{style:STYLE_NAMES.fact,text:'Source: a book of mine'},...list.slice(time+1)];
  const guide=readGuide(withExtra),l=guide.modules[0].lessons[0];
  const {facts}=readFacts(withExtra,l);
  assert.equal(facts.purpose,'Mine');assert.deepEqual(facts.extra,['Source: a book of mine']);
  assert.match(factsBlock(facts),/Source: a book of mine/);
  void lesson;void kindOf;void click;
});

// ---- the pages that show the facts
test('Training at a Glance has a Method column once a lesson says its method, and says it is out of date when a method changes',async()=>{
  const a=await lessonGuide();
  await click(a,'guide-build-glance');
  const head=()=>{const ps=a.adapter.snapshot(),at=ps.findIndex(p=>p.style==='OSG Table Head'&&p.text.trim()==='Module or lesson');const cells=[];for(let i=at;i<ps.length&&/^OSG Table (Head|Cell)$/.test(ps[i].style);i++)cells.push(ps[i].text.trim());return cells;};
  assert.deepEqual(head().slice(0,3),['Module or lesson','Minutes','Starts at minute']);
  assert.ok(!head().includes('Method'),'no lesson has a method yet');
  await open(a);fill(a,{method:'Discussion'});a.n('guide-details-method').value='Discussion';await save(a);
  await a.p.refresh();
  assert.match(a.n('guide-glance-note').textContent,/out of date/,'the page does not have the column the guide now needs');
  await click(a,'guide-build-glance');
  const cells=head();
  assert.deepEqual(cells.slice(0,4),['Module or lesson','Minutes','Starts at minute','Method']);
  const rows=[];for(let i=4;i<cells.length;i+=4)rows.push(cells.slice(i,i+4));
  assert.equal(rows.find(r=>r[0].startsWith('Lesson 1.1'))[3],'Discussion');
  assert.equal(rows.find(r=>r[0].startsWith('Lesson 1.2'))[3],'','a lesson that says nothing has an empty cell');
  assert.match(a.n('guide-glance-note').textContent,/matches your guide/);
  await open(a);a.n('guide-details-method').value='Lecture';await save(a);
  await a.p.refresh();
  assert.match(a.n('guide-glance-note').textContent,/out of date/,'a changed method is a changed page');
});

test('a module overview has a Method column, a key message and teaching aids to type over, and counts its slides',async()=>{
  const a=await lessonGuide();
  await open(a);fill(a,{method:'Game'});a.n('guide-details-method').value='Game';await save(a);
  a.p.useTool('instructionalCues');await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));await a.p.refresh();
  await addBlock(a,'slide','Opening slide');
  await moveTo(a,startsWith('OSG Module: Module 1'));
  await click(a,'guide-build-overview');assert.equal(a.err(),'');
  const rows=a.rows();
  assert.ok(rows.some(r=>/^OSG Overview: Key message:/.test(r)),'a line for the key message');
  assert.ok(rows.some(r=>/^OSG Overview: Teaching aids:/.test(r)),'and for the teaching aids');
  assert.ok(rows.includes('OSG Table Head: Method'),'the table has a Method column');
  assert.ok(rows.some(r=>/^OSG Overview: Slides: 1 slide\.$/.test(r)),'and the slides of the module are counted: '+rows.filter(r=>/Overview/.test(r)).join(' | '));
  assert.match(a.n('guide-overview-note').textContent,/matches your guide/);
});

// ---- days, the stated length and breaks (slice K2)
const agendaTexts=a=>a.n('guide-agenda-body').children.map(tr=>tr.children.map(c=>c.textContent).join(' | '));
test('a lesson that starts a new day sends the agenda clock back to that time, and the agenda says which day each lesson is in',async()=>{
  const a=await lessonGuide();
  await open(a,'OSG Lesson: Lesson 1.2');fill(a,{newday:'08:30',mode:'Small groups'});a.n('guide-details-mode').value='Small groups';await save(a);
  assert.equal(a.err(),'');
  assert.deepEqual(await factTexts(a),['Interaction: Small groups','New day: 8:30']);
  a.p.useTool('timingAndAgenda');await a.p.refresh();
  const rows=agendaTexts(a);
  assert.ok(rows[0].startsWith('Day 1'),rows.join(' / '));
  const i1=rows.findIndex(r=>r.startsWith('09:00 to 09:10 | Lesson 1.1'));
  assert.ok(i1>0,'day 1 runs from the start time: '+rows.join(' / '));
  assert.ok(rows.some(r=>r.startsWith('Day 2')),'a second day');
  assert.ok(rows.some(r=>r.startsWith('08:30 to 08:35 | Lesson 1.2: Second (Small groups)')),'day 2 starts at its own time, and the way people work together is shown: '+rows.join(' / '));
  assert.match(a.n('guide-agenda-summary').textContent,/2 lessons in 2 days, 15 minutes, the first day from 09:00\./);
  // a time that is not a clock time is refused
  await open(a,'OSG Lesson: Lesson 1.2');
  assert.throws(()=>checkFacts({newDay:'25:99'}),/clock time/);
});

test('what the author says the programme lasts is kept in the guide, and the Review says when the lessons add up to something else',async()=>{
  const a=await lessonGuide();
  a.p.useTool('timingAndAgenda');await a.p.refresh();
  a.type('guide-length-hours','0');a.type('guide-length-minutes','15');await a.n('guide-length-minutes').onchange();
  assert.equal(a.err(),'');
  const review=async()=>guideIssues(updatePlan(await a.adapter.paragraphs(),{forms:readSettings(await a.adapter.formSettings())})).filter(i=>i.rule==='PROGRAM-LENGTH');
  assert.equal((await review()).length,0,'10 and 5 minutes are the 15 the author said');
  assert.match(a.n('guide-length-note').textContent,/The lessons add up to 15 minutes, as you say\./);
  a.type('guide-length-hours','1');a.type('guide-length-minutes','0');await a.n('guide-length-minutes').onchange();
  const issues=await review();
  assert.equal(issues.length,1);
  assert.match(issues[0].message,/The guide says the programme is 1 hour, but its lessons add up to 15 minutes\./);
  assert.match(a.n('guide-length-note').textContent,/You say 1 hour, but the lessons add up to 15 minutes\./);
  a.type('guide-length-hours','');a.type('guide-length-minutes','');await a.n('guide-length-minutes').onchange();
  assert.equal((await review()).length,0,'no length said, nothing to compare');
  a.type('guide-length-minutes','75');await a.n('guide-length-minutes').onchange();
  assert.match(a.err(),/minutes as a whole number from 0 to 59/);
});

test('a day of more than three hours with no break in it is a warning, and a Break block or a Break lesson clears it',async()=>{
  const a=await started('Long guide');
  await addModule(a,'One');await addLesson(a,'Long one','200');
  const rules=async()=>guideIssues(updatePlan(await a.adapter.paragraphs())).filter(i=>i.rule==='DAY-BREAK');
  const found=await rules();
  assert.equal(found.length,1);
  assert.match(found[0].message,/^The guide is 3 hours 20 minutes long with no break in it\./);
  a.p.useTool('instructionalCues');await moveTo(a,startsWith('OSG Lesson: Lesson 1.1'));await a.p.refresh();
  await addBlock(a,'break','Coffee',{minutes:15});
  assert.equal((await rules()).length,0,'a Break block is a break');
  // the same by the lesson's method
  const b=await started('Method guide');
  await addModule(b,'One');await addLesson(b,'Long one','200');
  assert.equal(guideIssues(updatePlan(await b.adapter.paragraphs())).filter(i=>i.rule==='DAY-BREAK').length,1);
  await open(b);fill(b,{});b.n('guide-details-method').value='Break';await save(b);
  assert.equal(guideIssues(updatePlan(await b.adapter.paragraphs())).filter(i=>i.rule==='DAY-BREAK').length,0,'a lesson whose method is Break counts');
});

test('with several days each is judged on its own, and a short day needs no break',async()=>{
  const a=await started('Two days');
  await addModule(a,'One');await addLesson(a,'Day one','100');await addLesson(a,'Day two','200');
  await open(a,'OSG Lesson: Lesson 1.2');fill(a,{newday:'09:00'});await save(a);
  const found=guideIssues(updatePlan(await a.adapter.paragraphs())).filter(i=>i.rule==='DAY-BREAK');
  assert.equal(found.length,1,'only the 200 minute day');
  assert.match(found[0].message,/^Day 2 is 3 hours 20 minutes long/);
});
