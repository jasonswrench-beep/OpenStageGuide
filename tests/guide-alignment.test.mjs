// Learning Alignment on the guide (rebuilt 2026-11-07 so it needs no older course record): which objectives have an activity and an
// assessment item, tied together by what they say. Drives the real pane module in memory (not native Word).
import test from 'node:test';
import assert from 'node:assert/strict';
import {alignment,alignmentSummary,likeness} from '../src/guide-alignment.mjs';
import {started,click,addModule,addLesson,addBlock,submit} from './helpers/guide-harness.mjs';

const typeItem=(a,type,fields)=>{a.itemRadio(type).checked=true;a.itemRadio(type).onchange();for(const [id,v] of Object.entries(fields))a.n(id).value=v;};
async function addActivity(a,title,{objective='',instructions='Do the task together.'}={}){
  a.pick('activity');
  a.n('guide-activity-parts').checked=true;a.n('guide-activity-parts').onchange();
  a.type('guide-block-text',title);a.type('guide-act-objective',objective);a.type('guide-act-instructions',instructions);
  await submit(a,'guide-block-form');assert.equal(a.err(),'',title);
}
async function aligned(){
  const a=await started('Alignment course');
  await addModule(a,'Recognize');await addLesson(a,'Reframe','20');
  await addBlock(a,'objective','Identify two ways to reframe a disagreement.');
  await addBlock(a,'objective','Explain why tone matters in conflict.');
  await addActivity(a,'Reframe in pairs',{objective:'identify two ways to reframe a disagreement'});
  await addActivity(a,'Tone practice',{objective:'Explain why tone matters in a conflict situation'});
  await addActivity(a,'Name the feelings',{objective:'Name three feelings'});
  a.pick('activity');
  a.n('guide-activity-parts').checked=false;a.n('guide-activity-parts').onchange();   // an Activity written as plain lines
  a.type('guide-block-text','Plain warm-up\nTalk to your neighbour.');a.type('guide-act-objective','');await submit(a,'guide-block-form');assert.equal(a.err(),'');
  await click(a,'guide-tab-assess');await click(a,'guide-sec-level-two');
  typeItem(a,'true-false',{'guide-item-stem':'Reframing helps?','guide-item-objective':'Identify two ways to reframe a disagreement.'});a.n('guide-item-true').checked=true;
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  typeItem(a,'true-false',{'guide-item-stem':'Tone is not important?','guide-item-objective':''});a.n('guide-item-true').checked=false;a.n('guide-item-false').checked=true;
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  typeItem(a,'true-false',{'guide-item-stem':'Unrelated?','guide-item-objective':'A totally different aim'});a.n('guide-item-true').checked=true;
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  return a;
}
const open=async a=>{a.p.useTool('openDesigner');await a.p.openTool('learningAlignment');};

test('"the same" is exact words, capitals and punctuation aside; "similar" is most of the words; anything else is neither',()=>{
  assert.equal(likeness('Identify two ways.','identify TWO ways'),'same');
  assert.equal(likeness('Explain why tone matters in conflict.','Explain why tone matters in a conflict situation'),'similar');
  assert.equal(likeness('Name three feelings','Explain why tone matters'),null);
  assert.equal(likeness('','anything'),null);assert.equal(likeness('anything',''),null);assert.equal(likeness(null,undefined),null);
  assert.equal(likeness('a b','c d'),null,'words of one or two letters are not enough to match on');
});

test('objectives are tied to the activities and items that say the same thing, and what is tied to nothing is listed with why',async()=>{
  const a=await aligned();
  const al=alignment(await a.adapter.paragraphs());
  assert.deepEqual(al.objectives.map(o=>o.text),['Identify two ways to reframe a disagreement.','Explain why tone matters in conflict.']);
  const [o1,o2]=al.objectives;
  assert.deepEqual(o1.activities.map(x=>[x.title,x.similar]),[['Reframe in pairs',false]]);
  assert.deepEqual(o1.items.map(x=>x.similar),[false]);
  assert.match(o1.items[0].label,/in Level Two Assessment/);
  assert.deepEqual(o2.activities.map(x=>[x.title,x.similar]),[['Tone practice',true]],'a small rewording is still a link, and is marked');
  assert.deepEqual(o2.items,[]);
  assert.deepEqual(al.summary,{objectives:2,practised:2,checked:1,both:1});
  assert.equal(alignmentSummary(al),'2 objectives: 2 with an activity, 1 with an assessment item, 1 with both.');
  assert.deepEqual(al.loose.map(l=>l.kind),['activity','activity','item','item']);
  assert.equal(al.loose[0].text,'Name the feelings');
  assert.equal(al.loose[0].why,'its objective, "Name three feelings", matches none of your lesson objectives.');
  assert.equal(al.loose[1].text,'Plain warm-up');
  assert.equal(al.loose[1].why,'it has no Objective part.');
  assert.match(al.loose[2].text,/^Question 2 .*in Level Two Assessment$/);
  assert.equal(al.loose[2].why,'it has no learning objective.');
  assert.match(al.loose[3].text,/^Question 3 /);
  assert.equal(al.loose[3].why,'its learning objective, "A totally different aim", matches none of your lesson objectives.');
});

test('a guide with no objectives says what to write, and a survey item is not measured against objectives',async()=>{
  const a=await started('None');
  const empty=alignment(await a.adapter.paragraphs());
  assert.deepEqual(empty.summary,{objectives:0,practised:0,checked:0,both:0});
  assert.match(alignmentSummary(empty),/There are no objectives yet\. Add an Objective block to a lesson/);
  await click(a,'guide-tab-assess');await click(a,'guide-sec-level-one');
  typeItem(a,'open-comment',{'guide-item-stem':'Any comments?'});
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  assert.deepEqual(alignment(await a.adapter.paragraphs()).loose,[],'a survey question has no right answer and no objective to check');
});

test('the tool opens on its own panel and lists each objective with its activity and assessment, and what is still missing',async()=>{
  const a=await aligned();await open(a);
  assert.equal(a.n('guide-sec-align').hidden,false);assert.equal(a.n('guide-sec-module').hidden,true);
  assert.equal(a.n('guide-align-summary').textContent,'2 objectives: 2 with an activity, 1 with an assessment item, 1 with both.');
  const items=a.n('guide-align-list').children;
  assert.equal(items.length,2);
  assert.equal(items[0].children[0].textContent,'Lesson 1.1: Reframe: Identify two ways to reframe a disagreement.'.replace('Lesson 1.1: Reframe: ','Lesson 1.1: Reframe: '));
  assert.equal(items[0].children[1].textContent,'Activity: Reframe in pairs');
  assert.match(items[0].children[2].textContent,/^Assessment: Question 1 .* in Level Two Assessment$/);
  assert.equal(items[1].children[1].textContent,'Activity: Tone practice (similar words)');
  assert.equal(items[1].children[2].textContent,'No assessment item yet.');
  assert.equal(items[1].children[2].className,'guide-align-gap','what is missing is marked');
  assert.equal(items[0].children[1].className??'','');
  assert.equal(a.n('guide-align-loose-heading').hidden,false);
  assert.equal(a.n('guide-align-loose').children.length,4);
  assert.match(a.n('guide-align-loose').children[0].textContent,/^Name the feelings: its objective, "Name three feelings", matches none of your lesson objectives\. /);
  assert.equal(a.err(),'');
});

test('Go there moves the cursor to the objective, and the panel is drawn again when the guide changes',async()=>{
  const a=await aligned();await open(a);
  const go=a.n('guide-align-list').children[1].children[3];
  assert.equal(go.attributes['aria-label'],'Go to the objective in Lesson 1.1: Reframe');
  await go.onclick();
  const paras=await a.adapter.paragraphs(),at=await a.adapter.cursorIndex();
  assert.match(paras[at].text,/Explain why tone matters in conflict/);
  assert.match(a.n('guide-status').textContent,/^Moved to the objective in Lesson 1\.1: Reframe\./);
  await addBlock(a,'objective','Practise a calm opening line.');
  await click(a,'guide-align-refresh');
  assert.match(a.n('guide-align-summary').textContent,/^3 objectives: /);
});
