// One tool at a time (planned 2026-11-06, approved 2026-11-07): a ribbon button opens the pane on that tool's panel alone, not on the
// whole Guide page. This drives the real pane module and the real taskpane.html ids in memory. It is not native Word.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {TOOLS,GUIDE_SECTIONS,relatedTools} from '../src/guide-tools.mjs';
import {pane,started,click,addModule,addLesson,moveTo,startsWith} from './helpers/guide-harness.mjs';
import {focusedId} from './helpers/app-dom.mjs';

const PARTS=['add','slides','elearning','look','organize'];
const PART_OF={'guide-sec-module':'add','guide-sec-block':'add','guide-sec-owncues':'add','guide-sec-import':'slides','guide-sec-update':'slides','guide-sec-look':'look','guide-sec-standard':'look','guide-sec-built':'look','guide-sec-move':'organize','guide-sec-keeptidy':'organize','guide-sec-contents':'organize','guide-sec-library':'add','guide-sec-agenda':'organize','guide-sec-align':'organize','guide-sec-versions':'organize','guide-sec-backup':'organize','guide-sec-restore':'organize'};
const guideTools=Object.values(TOOLS).filter(t=>t.kind==='guide');

// What is on screen in the Guide screen: the one part showing, and the pieces of it that are not hidden.
const onScreen=a=>{
  const parts=PARTS.filter(p=>!a.n(`guide-part-${p}`).hidden);
  const sections=GUIDE_SECTIONS.filter(s=>!a.n(s).hidden&&PART_OF[s]===parts[0]);
  return {parts,sections};
};
const withTools=async(opts={})=>{
  const seen=[],a=pane(undefined,{onTool:id=>seen.push(id)});
  await a.p.refresh();
  a.type('guide-title','Reframing conflict');await a.n('guide-start-form').onsubmit({preventDefault(){}});
  a.p.useTool('openDesigner');
  return {a,seen};
};

test('every section the tools name is in the page, and the page lists no other piece the tools do not know',()=>{
  const html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');
  for(const id of GUIDE_SECTIONS)assert.equal(html.split(`id="${id}"`).length-1,1,`${id} is in taskpane.html once`);
  assert.deepEqual(Object.keys(PART_OF).sort(),[...GUIDE_SECTIONS].sort(),'this test knows every section');
  for(const id of ['tool-strip','tool-title','tool-note','tool-more','tool-more-list','guide-help','guide-help-heading','guide-help-sample','guide-help-keys','guide-sec-titlebar','guide-sec-selected'])assert.equal(html.split(`id="${id}"`).length-1,1,id);
});

test('with no tool open the pane shows every part and both tab strips, as the first pane did',async()=>{
  const a=await started();
  assert.equal(a.n('guide-parts').hidden,false);
  assert.equal(a.n('guide-switch').hidden,false);
  for(const s of GUIDE_SECTIONS)assert.equal(a.n(s).hidden,false,`${s} shows`);
});

for(const t of guideTools){
  test(`${t.label}: the Guide screen shows its part and only the pieces it names`,async()=>{
    const {a,seen}=await withTools();
    await a.p.openTool(t.id);
    assert.equal(a.n('guide-build').hidden,false);assert.equal(a.n('guide-start').hidden,true);
    const {parts,sections}=onScreen(a);
    assert.deepEqual(parts,[t.part],'one part shows');
    assert.deepEqual([...sections].sort(),[...t.sections].sort(),'only the pieces this tool names');
    assert.equal(a.n('guide-parts').hidden,true,'the five tabs are not shown');
    assert.equal(a.n('guide-switch').hidden,true,'nor the switch between Guide, Assessments and Workbook');
    assert.equal(a.n('guide-sec-titlebar').hidden,!t.titlebar);
    assert.equal(a.n('guide-sec-selected').hidden,!t.selected);
    assert.equal(a.p.state().tool,t.id);
    assert.ok(seen.includes(t.id)||t.id==='openDesigner','the strip is told');
    assert.equal(a.err(),'');
  });
}

test('the module part stays with the writing tools: Add Module, Add Activity and Instructional Cues each show where the cursor is, and each offers the others',async()=>{
  const {a}=await withTools();
  for(const id of ['addModule','addActivity','instructionalCues']){
    await a.p.openTool(id);
    assert.equal(a.n('guide-sec-selected').hidden,false,`${id} shows the "In your document" line`);
    const more=relatedTools(id).map(t=>t.id);
    assert.ok(['addModule','addActivity','instructionalCues'].filter(x=>x!==id).every(x=>more.includes(x)),`${id} offers the other writing tools`);
  }
});

test('Add Activity opens the block form on Activity written as parts; Instructional Cues takes it back to an ordinary cue',async()=>{
  const {a}=await withTools();
  await a.p.openTool('addActivity');
  assert.equal(a.n('guide-cue-current').textContent,'Activity');
  assert.equal(a.n('guide-activity-parts').checked,true);
  assert.equal(a.n('guide-activity-fields').hidden,false,'the labelled parts are showing');
  await a.p.openTool('instructionalCues');
  assert.equal(a.n('guide-activity-parts').checked,false);
  assert.equal(a.n('guide-activity-fields').hidden,true);
  assert.notEqual(a.n('guide-cue-current').textContent,'Activity');
  await a.p.openTool('instructionalCues');
  assert.equal(a.n('guide-sec-owncues').hidden,false,'the cues of your own are with the cues');
});

test('Add Module shows only the module and lesson form: no block form, no cue list',async()=>{
  const {a}=await withTools();
  await a.p.openTool('addModule');
  assert.equal(a.n('guide-sec-module').hidden,false);
  assert.equal(a.n('guide-sec-block').hidden,true);
  assert.equal(a.n('guide-sec-owncues').hidden,true);
});

test('adding a lesson goes on to the block form, so the next thing the author does is write in it',async()=>{
  const {a,seen}=await withTools();
  await a.p.openTool('addModule');
  await addModule(a,'Recognize the conflict');
  await addLesson(a,'Notice the signs');
  assert.equal(a.p.state().tool,'instructionalCues');
  assert.equal(seen.at(-1),'instructionalCues');
  assert.equal(a.n('guide-sec-block').hidden,false);assert.equal(a.n('guide-sec-module').hidden,true);
  assert.equal(focusedId(),'guide-block-text');
  // writing a block keeps the author in the tool they are in
  await a.p.openTool('addActivity');
  a.type('guide-block-text','Role play');await a.n('guide-block-form').onsubmit({preventDefault(){}});
  assert.equal(a.p.state().tool,'addActivity');
});

test('Edit this module, pressed in Open Designer, opens the module form with its edit box',async()=>{
  const {a}=await withTools();
  await addModule(a,'Recognize the conflict');
  await moveTo(a,startsWith('OSG Module'));
  await a.p.openTool('openDesigner');
  assert.equal(a.n('guide-selected').hidden,false);
  await click(a,'guide-edit-selected');
  assert.equal(a.p.state().tool,'addModule');
  assert.equal(a.n('guide-edit-form').hidden,false);
  assert.equal(a.n('guide-edit-heading').textContent,'Edit this module');
  assert.equal(a.n('guide-sec-module').hidden,false);
});

test('Course Outline shows the contents and Move or delete; Open Designer shows Keep it tidy, and neither shows the other',async()=>{
  const {a}=await withTools();
  await a.p.openTool('courseOutline');
  assert.equal(a.n('guide-sec-contents').hidden,false);assert.equal(a.n('guide-sec-keeptidy').hidden,true);
  await a.p.openTool('openDesigner');
  assert.equal(a.n('guide-sec-keeptidy').hidden,false);assert.equal(a.n('guide-sec-contents').hidden,true);
});

test('Assessments, the workbook and Review each open their own screen alone, and the cursor moving does not change the tool',async()=>{
  const {a,seen}=await withTools();
  await a.p.openTool('assessments');
  assert.equal(a.n('guide-assess').hidden,false);assert.equal(a.n('guide-build').hidden,true);
  assert.equal(a.n('guide-switch').hidden,true);
  assert.equal(a.n('guide-assess-parts').hidden,false,'Assessments keeps its own three parts: Items, Paste questions, Question bank');
  await a.p.quietRefresh();
  assert.equal(a.p.state().tool,'assessments');
  await a.p.openTool('participantWorkbook');
  assert.equal(a.n('guide-workbook').hidden,false);assert.equal(a.n('guide-assess').hidden,true);
  await a.p.openTool('reviewGuide');
  assert.equal(a.n('guide-review').hidden,false);assert.equal(a.n('guide-workbook').hidden,true);
  assert.deepEqual(seen.slice(-3),['assessments','participantWorkbook','reviewGuide']);
  assert.equal(a.err(),'');
});

test('Review guide, pressed in Open Designer, becomes the Review guide tool, and Back goes to Open Designer',async()=>{
  const {a,seen}=await withTools();
  await a.p.openTool('openDesigner');
  await click(a,'guide-open-review');
  assert.equal(a.n('guide-review').hidden,false);
  assert.equal(a.p.state().tool,'reviewGuide');
  await click(a,'guide-review-back');
  assert.equal(a.n('guide-build').hidden,false);
  assert.equal(a.p.state().tool,'openDesigner');
  assert.deepEqual(seen.slice(-2),['reviewGuide','openDesigner']);
});

test('in a document with no guide, a guide tool shows Start; Help and Review do not need one',async()=>{
  const seen=[],a=pane(undefined,{onTool:id=>seen.push(id)});await a.p.refresh();a.p.useTool('openDesigner');
  await a.p.openTool('addModule');
  assert.equal(a.n('guide-start').hidden,false);assert.equal(a.n('guide-build').hidden,true);
  await a.p.openTool('help');
  assert.equal(a.n('guide-help').hidden,false);assert.equal(a.n('guide-start').hidden,true);
  await a.p.quietRefresh();
  assert.equal(a.n('guide-help').hidden,false,'the cursor moving does not take Help away');
  await a.p.openTool('reviewGuide');
  assert.equal(a.n('guide-review').hidden,false);
  // starting the guide from Start carries on in the tool that was asked for
  await a.p.openTool('addModule');
  a.type('guide-title','Reframing conflict');await a.n('guide-start-form').onsubmit({preventDefault(){}});
  assert.equal(a.n('guide-build').hidden,false);
  assert.equal(a.p.state().tool,'addModule');
  assert.equal(a.n('guide-sec-module').hidden,false);
});

test('Help: the sample guide button starts the sample and opens Open Designer, and the instructions link is there',async()=>{
  const a=pane(undefined,{onTool:()=>{}});await a.p.refresh();a.p.useTool('help');
  await a.p.openTool('help');
  await click(a,'guide-help-sample');
  assert.equal(a.n('guide-build').hidden,false);
  assert.equal(a.p.state().tool,'openDesigner');
  assert.match(a.n('guide-total').textContent,/20 min/);
  const html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');
  assert.match(html,/<a id="guide-help-read" href="help\.html"/);
});

test('a piece whose heading says what the strip already says leaves it to the strip; one that says something else keeps its own',async()=>{
  const {a}=await withTools();
  a.n('guide-module-heading').textContent='Add a module or lesson';
  a.n('guide-el-heading').textContent='E-Learning screens';
  a.n('guide-upd-heading').textContent='Update from slides';
  await a.p.openTool('addModule');
  assert.equal(a.n('guide-module-heading').hidden,true);
  assert.equal(a.n('guide-el-heading').hidden,false,'the heading of a tool that is not open is left alone');
  await a.p.openTool('eLearningScreens');
  assert.equal(a.n('guide-el-heading').hidden,true);
  assert.equal(a.n('guide-module-heading').hidden,false,'and the module form\'s own heading is back for the other tools');
  await a.p.openTool('addActivity');
  assert.equal(a.n('guide-module-heading').hidden,false);
  await a.p.openTool('updateFromSlides');
  assert.equal(a.n('guide-upd-heading').hidden,true);
});

test('a tool asked for while the pane is still busy opens when the work in hand is done, and the last one asked for wins (found in a real Word: a click a moment after Review Guide opened did nothing)',async()=>{
  const {a}=await withTools();
  const first=a.p.openTool('reviewGuide');
  assert.equal(a.p.state().busy,true,'the pane is busy opening Review Guide');
  const second=a.p.openTool('addModule');
  const third=a.p.openTool('timingAndAgenda');
  await Promise.all([first,second,third]);
  assert.equal(a.p.state().tool,'timingAndAgenda','the last tool asked for is the one that opened');
  assert.equal(a.p.state().busy,false);
  assert.equal(a.n('guide-sec-agenda').hidden,false,'and its panel is showing');
  assert.equal(a.err(),'');
});

test('a tool chosen while a file window is open is not carried out long after the window closes',async()=>{
  const {a}=await withTools();
  await a.p.openTool('importSlides');
  let release;
  a.adapter.pickFiles=()=>new Promise(resolve=>{release=resolve;});
  const choosing=a.n('guide-slides-choose').onclick();
  assert.equal(a.p.state().busy,true,'the pane waits for the file window');
  await a.p.openTool('addModule');
  release({cancelled:true});
  await choosing;
  assert.equal(a.p.state().tool,'importSlides','the pane stays where the author was');
  assert.equal(a.p.state().busy,false);
});
