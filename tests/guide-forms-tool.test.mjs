// The Forms tool: the guide's forms with their numbers and print lines, how they are numbered, and the copies to make. It drives the real pane
// module and the real taskpane.html ids in memory. It is not a browser (scripts/check-program-browser.mjs opens the tool in one).
import test from 'node:test';
import assert from 'node:assert/strict';
import {pane,submit,addModule,addLesson,click,moveTo,startsWith} from './helpers/guide-harness.mjs';
import {formTitleBlock,printLine,formText,citationControl} from '../src/guide-forms-ooxml.mjs';
import {para,run} from '../src/guide-ooxml.mjs';
import {readSettings,writeSettings} from '../src/guide-forms.mjs';
import {guideIssues,updatePlan} from '../src/guide-model.mjs';

async function guideWithForms(){
  const a=pane();
  await a.p.refresh();
  a.type('guide-title','Forms guide');await submit(a,'guide-start-form');
  await addModule(a,'One');await addLesson(a,'First');
  await addModule(a,'Two');await addLesson(a,'Second');
  return a;
}
const index=(a,text)=>a.adapter.snapshot().findIndex(p=>p.text.includes(text));
const insertAfter=(a,after,build)=>a.adapter.insert(build,{where:'after',after,spare:false});
const listed=a=>a.n('guide-forms-list').children.map(li=>li.textContent);
const rows=a=>a.n('guide-forms-copies-body').children.map(tr=>tr.children.map(td=>td.textContent));
const change=async(a,id,value)=>{a.type(id,value);await a.n(id).onchange();};

async function withForms(){
  const a=await guideWithForms();
  await insertAfter(a,index(a,'Lesson 1.1')+1,()=>
    formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})+printLine({copies:{n:1,per:'person'},paper:'Letter',orientation:'portrait'})+formText('Read this first.')
    +formTitleBlock({id:'f2',label:'Worksheet',number:'1.2',title:'Plans',facilitator:true})+formText('Answers.'));
  await insertAfter(a,index(a,'Lesson 2.1')+1,()=>
    formTitleBlock({id:'f3',label:'Checklist',number:'2.1',title:'Room'})+printLine({copies:{n:2,per:'team'}})+formText('Tick them.'));
  a.p.useTool('forms');
  await a.p.refresh();
  return a;
}

test('the Forms tool lists the forms with their numbers, who they are for and their print line, and says how many there are',async()=>{
  const a=await withForms();
  assert.equal(a.n('guide-sec-forms').hidden,false,'its panel is showing');
  assert.match(a.n('guide-forms-summary').textContent,/^3 forms: 2 for everyone, 1 for the facilitator only/);
  const items=listed(a);
  assert.equal(items.length,3);
  assert.match(items[0],/^Handout 1\.1: Goals \(for everyone; Print: 1 per person, Letter, portrait\)/);
  assert.match(items[1],/^Worksheet 1\.2: Plans \(for the facilitator; no print line\)/);
  assert.match(items[2],/^Checklist 2\.1: Room \(for everyone; Print: 2 per team\)/);
  assert.equal(a.n('guide-forms-list').children[0].children[2].textContent,'Go there','each form has a link that goes to it');
});

test('a guide with no forms says so and points to New form',async()=>{
  const a=await guideWithForms();
  a.p.useTool('forms');await a.p.refresh();
  assert.match(a.n('guide-forms-summary').textContent,/^There are no forms in this guide yet\. Choose New form/);
  assert.equal(a.n('guide-forms-list').children.length,0);
  assert.equal(a.n('guide-forms-copies-table').hidden,true);
  assert.match(a.n('guide-forms-copies-summary').textContent,/No form has a print line yet/);
});

test('the numbers of people, teams and tables are kept with the guide, add the print lines up, and clear the Review warning that asked for them',async()=>{
  const a=await withForms();
  const review=async()=>guideIssues(updatePlan(await a.adapter.paragraphs(),{forms:readSettings(await a.adapter.formSettings())})).filter(i=>i.rule==='FORM-PRINT-GROUP').map(i=>i.message);
  assert.equal((await review()).length,2,'one warning for people and one for teams, and no way yet to answer them');
  assert.match((await review())[0],/Say how many in Forms, in the Activities menu\./,'the warning says where');
  assert.match(a.n('guide-forms-copies-summary').textContent,/2 forms cannot be added up yet/);
  assert.deepEqual(a.n('guide-forms-copies-missing').children.map(li=>li.textContent),['Handout 1.1: Goals: The guide does not say how many people.','Checklist 2.1: Room: The guide does not say how many teams.']);

  await change(a,'guide-forms-people','24');
  assert.equal(a.err(),'');
  assert.equal(readSettings(await a.adapter.formSettings()).group.people,24,'kept in the guide\'s own settings');
  assert.match(a.n('guide-forms-copies-summary').textContent,/1 form added up so far, 24 copies\. 1 form cannot be added up yet\./);
  assert.deepEqual(rows(a)[0],['Handout 1.1: Goals','1 per person','24']);
  assert.deepEqual(rows(a)[1],['Checklist 2.1: Room','2 per team','Not known yet']);
  assert.equal((await review()).length,1,'the people warning is gone');

  await change(a,'guide-forms-teams','6');
  assert.match(a.n('guide-forms-copies-summary').textContent,/^36 copies to make in all\./);
  assert.deepEqual(rows(a)[1],['Checklist 2.1: Room','2 per team','12']);
  assert.equal(a.n('guide-forms-copies-missing').children.length,0);
  assert.equal((await review()).length,0,'and the teams warning too');

  await change(a,'guide-forms-people','');
  assert.equal(readSettings(await a.adapter.formSettings()).group.people??null,null,'an empty box takes the number away');
  assert.equal(readSettings(await a.adapter.formSettings()).group.teams,6,'and leaves the others');
});

test('a number that is not a whole number from 1 to 9999 is refused and nothing is kept',async()=>{
  const a=await withForms();
  for(const bad of ['0','-3','2.5','abc','10000']){
    await change(a,'guide-forms-people',bad);
    assert.match(a.err(),/whole number from 1 to 9999/,bad);
    assert.equal(readSettings(await a.adapter.formSettings()).group?.people??null,null,`${bad} is not kept`);
  }
});

test('the numbering scheme is chosen in the tool, kept with the guide, and Update guide uses it',async()=>{
  const a=await withForms();
  assert.equal(a.n('guide-forms-scheme').value,'auto');
  assert.match(a.n('guide-forms-scheme-note').textContent,/numbered by unit, each form taking the number of its module/);
  a.type('guide-forms-scheme','letter');await a.n('guide-forms-scheme').onchange();
  assert.equal(a.err(),'');
  assert.equal(readSettings(await a.adapter.formSettings()).scheme,'letter');
  assert.match(a.n('guide-forms-scheme-note').textContent,/numbered by letter \(A, B, C\)/);
  assert.match(a.n('guide-status').textContent,/Choose Update guide to renumber/);
  await click(a,'guide-update');
  const texts=a.adapter.snapshot().map(p=>p.text.trim());
  assert.ok(texts.includes('Handout A: Goals'));assert.ok(texts.includes('Worksheet B: Plans'));assert.ok(texts.includes('Checklist C: Room'));
  a.p.useTool('forms');await a.p.refresh();
  assert.match(listed(a)[0],/^Handout A: Goals/,'the list shows the new numbers');
  a.type('guide-forms-scheme','none');await a.n('guide-forms-scheme').onchange();
  assert.equal(a.n('guide-forms-scheme-note').textContent,'Forms are not numbered.');
});

test('the settings written by the tool read back the way the model reads them, and both are kept in the project',async()=>{
  const text=writeSettings({scheme:'sequence',group:{people:12,teams:null,tables:3}});
  assert.deepEqual(readSettings(text),{scheme:'sequence',group:{people:12,teams:null,tables:3}});
  assert.deepEqual(readSettings(writeSettings({scheme:'auto'})),{scheme:'auto'});
});

test('New form asks the program page to run its New form dialog, and says what to do when there is no page around the tools',async()=>{
  const a=await withForms();
  await click(a,'guide-forms-new');
  assert.match(a.n('guide-status').textContent,/Choose Insert, then Form/,'with no program page around it the tool says how to start a form');
});

// ---- the Handout cue: Which form
const optionsOf=a=>a.n('guide-block-which').children.map(o=>[o.value,o.textContent]);
const citationsIn=async a=>(await a.adapter.paragraphs()).filter(p=>p.marks?.refs?.length).map(p=>({text:p.text.trim(),refs:p.marks.refs}));
async function handoutGuide(){
  const a=await withForms();
  a.p.useTool('instructionalCues');await a.p.refresh();
  await moveTo(a,r=>/Lesson 1\.1/.test(r));
  a.pick('handout');
  return a;
}

test('the Handout cue offers the forms of the guide to cite, and the other cues do not',async()=>{
  const a=await handoutGuide();
  assert.equal(a.n('guide-block-which-part').hidden,false);
  assert.deepEqual(optionsOf(a),[['','No form (type the words yourself)'],['f1','Handout 1.1: Goals'],['f2','Worksheet 1.2: Plans (for the facilitator only)'],['f3','Checklist 2.1: Room']]);
  a.pick('say');
  assert.equal(a.n('guide-block-which-part').hidden,true,'only a Handout block cites a form');
  const empty=await guideWithForms();
  empty.p.useTool('instructionalCues');await empty.p.refresh();empty.pick('handout');
  assert.deepEqual(optionsOf(empty),[['','There are no forms in the guide yet']]);
});

test('a Handout block cites the form chosen: its words, the number of the form, and a citation Update guide can keep right',async()=>{
  const a=await handoutGuide();
  a.type('guide-block-which','f1');a.type('guide-block-text','Give out');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  const cites=await citationsIn(a);
  assert.equal(cites.length,1);
  assert.match(cites[0].text,/Handout\s+Give out Handout 1\.1: Goals$/,'the words typed come before the name of the form');
  assert.deepEqual(cites[0].refs,[{id:'f1',text:'Handout 1.1: Goals'}]);
  assert.equal(a.n('guide-block-which').value,'','the form is cleared for the next block');
  // The form is renumbered: Update guide corrects the name in the block.
  a.p.useTool('forms');await a.p.refresh();
  a.type('guide-forms-scheme','letter');await a.n('guide-forms-scheme').onchange();
  await click(a,'guide-update');
  const after=await citationsIn(a);
  assert.match(after[0].text,/Give out Handout A: Goals$/,'the block follows the form to its new number');
  assert.deepEqual(after[0].refs,[{id:'f1',text:'Handout A: Goals'}]);
});

test('a form alone is enough for a Handout block, with minutes after it; with no form and no words it is refused',async()=>{
  const a=await handoutGuide();
  a.type('guide-block-which','f3');a.type('guide-block-minutes','5');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  const cites=await citationsIn(a);
  assert.match(cites[0].text,/Handout\s+Checklist 2\.1: Room \(5 min\)$/);
  a.pick('handout');
  await submit(a,'guide-block-form');
  assert.match(a.err(),/Type the text for this block first/);
});

test('Edit this block on a Handout block that cites a form shows the form and the words before it, and Update block keeps the citation',async()=>{
  const a=await handoutGuide();
  a.type('guide-block-which','f2');a.type('guide-block-text','Give out');
  await submit(a,'guide-block-form');
  await moveTo(a,startsWith('OSG Cue: Handout'));
  await click(a,'guide-edit-selected');
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-block-which-part').hidden,false);
  assert.equal(a.n('guide-block-which').value,'f2');
  assert.equal(a.n('guide-block-text').value,'Give out','the name of the form is not in the words');
  a.type('guide-block-text','Hand out the');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  const cites=await citationsIn(a);
  assert.equal(cites.length,1,'still one citation, not two and not none');
  assert.match(cites[0].text,/Hand out the Worksheet 1\.2: Plans$/);
  assert.deepEqual(cites[0].refs,[{id:'f2',text:'Worksheet 1.2: Plans'}]);
  // Changing to another form, and then to none.
  await moveTo(a,startsWith('OSG Cue: Handout'));await click(a,'guide-edit-selected');
  a.type('guide-block-which','f1');await submit(a,'guide-block-form');
  assert.deepEqual((await citationsIn(a))[0].refs,[{id:'f1',text:'Handout 1.1: Goals'}]);
  await moveTo(a,startsWith('OSG Cue: Handout'));await click(a,'guide-edit-selected');
  a.type('guide-block-which','');await submit(a,'guide-block-form');
  assert.equal((await citationsIn(a)).length,0,'No form takes the citation away');
});

test('a block that cites a form in the middle of its words is not rewritten by the pane, so the link is not lost',async()=>{
  const a=await handoutGuide();
  await a.adapter.insert(ctx=>para('OSG Cue',run('Handout')+run(' Give out ')+citationControl({id:'f1',text:'Handout 1.1: Goals'},ctx)+run(' first.')),{where:'after',after:(await a.adapter.cursorIndex()),spare:false});
  await moveTo(a,startsWith('OSG Cue: Handout'));
  await click(a,'guide-edit-selected');
  assert.match(a.err(),/cites a form in the middle of its words/);
  assert.notEqual(a.n('guide-block-go').textContent,'Update block','the form was not put into edit mode');
  assert.equal((await citationsIn(a)).length,1,'and the citation is still there');
});

test('Update guide numbers a form put in on the page a moment before, without waiting for the tools to read the guide again',async()=>{
  const a=await guideWithForms();
  await click(a,'guide-update');
  // Insert > Form on the page: the pane has not read the guide since.
  // After the lesson's time line (the last paragraph that names the lesson; the first is its line in the contents table).
  const lesson=a.adapter.snapshot().findLastIndex(p=>p.text.startsWith('Lesson 2.1'));
  await insertAfter(a,lesson+1,()=>formTitleBlock({id:'f9',label:'Template',number:'',title:'Next steps'})+formText('Plan it.'));
  await click(a,'guide-update');
  assert.ok(a.adapter.snapshot().some(p=>p.text.trim()==='Template 2.1: Next steps'),'the new form is numbered at once');
  assert.match(a.n('guide-status').textContent,/^Updated /);
});
