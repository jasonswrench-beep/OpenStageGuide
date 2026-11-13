import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {started,click,startsWith,moveTo,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

// The Guide screen is five short parts, one at a time, and the long cue list is put away behind a button. The author did a lot of
// scrolling (a recording of 2026-11-06: about six screens of one page). The in-memory document stands in for Word; it is not Word.
const PARTS=['add','slides','elearning','look','organize'];
const shown=a=>PARTS.filter(id=>!a.n(`guide-part-${id}`).hidden);
const selected=a=>PARTS.filter(id=>a.n(`guide-part-tab-${id}`).attributes?.get?.('aria-selected')==='true'||a.n(`guide-part-tab-${id}`).getAttribute?.('aria-selected')==='true');
const key=(a,id,k)=>{let stopped=false;a.n(id).onkeydown({key:k,preventDefault(){stopped=true;}});return stopped;};

test('the Guide screen opens on Add; choosing a tab shows that part alone and marks the tab',async()=>{
  const a=await started();
  assert.deepEqual(shown(a),['add']);
  assert.deepEqual(selected(a),['add']);
  for(const id of PARTS){
    a.n(`guide-part-tab-${id}`).onclick();
    assert.deepEqual(shown(a),[id],id);
    assert.deepEqual(selected(a),[id],`${id} is the selected tab`);
    assert.equal(a.n(`guide-part-tab-${id}`).tabIndex,0,'the selected tab is the one in the tab order');
    assert.ok(PARTS.filter(x=>x!==id).every(x=>a.n(`guide-part-tab-${x}`).tabIndex===-1));
  }
});

test('arrow keys, Home and End move along the tabs and wrap round',async()=>{
  const a=await started();
  assert.equal(key(a,'guide-part-tab-add','ArrowRight'),true);assert.deepEqual(shown(a),['slides']);
  assert.equal(key(a,'guide-part-tab-slides','ArrowRight'),true);assert.deepEqual(shown(a),['elearning']);
  key(a,'guide-part-tab-elearning','End');assert.deepEqual(shown(a),['organize']);
  key(a,'guide-part-tab-organize','ArrowRight');assert.deepEqual(shown(a),['add'],'after the last tab comes the first');
  key(a,'guide-part-tab-add','ArrowLeft');assert.deepEqual(shown(a),['organize'],'before the first tab comes the last');
  key(a,'guide-part-tab-organize','Home');assert.deepEqual(shown(a),['add']);
  assert.equal(key(a,'guide-part-tab-add','a'),false,'other keys are left alone');
});

test('anything that sends the author to the block form or the edit form shows the Add part first',async()=>{
  const a=await started('Editing');
  await addModule(a,'One');await addLesson(a,'First','10');
  a.n('guide-part-tab-organize').onclick();
  assert.deepEqual(shown(a),['organize']);
  // inserting a block (done from Add, but the pane returns there so the next block can be typed)
  a.n('guide-part-tab-look').onclick();
  await addBlock(a,'ask','What makes it hard?');
  assert.deepEqual(shown(a),['add'],'after a block is added the pane is on Add, ready for the next one');
  // Edit this block, pressed from another part, brings the form into view
  await moveTo(a,startsWith('OSG Cue: Ask'));
  a.n('guide-part-tab-organize').onclick();
  await click(a,'guide-edit-selected');
  assert.deepEqual(shown(a),['add']);
  assert.equal(a.n('guide-block-text').value,'What makes it hard?');
});

test('every control of the Guide screen sits in one of its five parts, and the ones used every day are in Add',()=>{
  const html=readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8').replace(/\r\n/g,'\n');
  const at=id=>{const i=html.indexOf(`id="${id}"`);assert.ok(i>0,`${id} is in the page`);return i;};
  const range=id=>{const start=html.indexOf(`<div id="guide-part-${id}"`);assert.ok(start>0,id);const next=PARTS.map(x=>html.indexOf(`<div id="guide-part-${x}"`)).filter(i=>i>start).sort((x,y)=>x-y)[0]??html.indexOf('<div id="guide-assess"');return [start,next];};
  const within=(id,part)=>{const [s,e]=range(part),i=at(id);assert.ok(i>s&&i<e,`${id} is in the ${part} part`);};
  for(const id of ['guide-edit-form','guide-new-title','guide-add-module','guide-add-lesson','guide-cue-toggle','guide-cue-panel','guide-cues','guide-block-text','guide-block-notes','guide-block-minutes','guide-block-go','guide-pic-choose','guide-own-save'])within(id,'add');
  for(const id of ['guide-slides-choose','guide-slides-go','guide-slides-list'])within(id,'slides');
  for(const id of ['guide-el-add','guide-el-script','guide-el-text','guide-el-author'])within(id,'elearning');
  for(const id of ['guide-look-apply','guide-looks','guide-add-standard','guide-add-appendix','guide-build-glance','guide-build-cues','guide-build-overview'])within(id,'look');
  for(const id of ['guide-tidy','guide-move-up','guide-delete','guide-update','guide-open-review','guide-refresh','guide-contents'])within(id,'organize');
  // The title, the total and the "In your document" bar are above the tabs, so they show whichever part is open.
  const tabs=at('guide-parts');
  for(const id of ['guide-build-heading','guide-total','guide-selected','guide-edit-selected'])assert.ok(at(id)<tabs,`${id} is above the tabs`);
  // The tabs are a real tab list.
  assert.match(html,/<div id="guide-parts"[^>]*role="tablist"/);
  for(const id of PARTS)assert.match(html,new RegExp(`<button type="button" role="tab" id="guide-part-tab-${id}" aria-controls="guide-part-${id}"`));
});

test('the cue is chosen from a short bar: its name and a Change button, with the list behind it',async()=>{
  const a=await started();
  assert.equal(a.n('guide-cue-panel').hidden,true,'the list is put away');
  assert.equal(a.n('guide-cue-toggle').getAttribute('aria-expanded'),'false');
  assert.equal(a.n('guide-cue-current').textContent,'Say');
  assert.equal(a.n('guide-cue-toggle-action').textContent,'Change');
  // choosing a cue changes the name on the bar and its picture
  a.pick('ask');
  assert.equal(a.n('guide-cue-current').textContent,'Ask');
  assert.match(a.n('guide-cue-current-icon').src,/icons-doc\/.*\.png$/);
  assert.match(a.n('guide-cue-help').textContent,/./,'what the cue is for shows under the bar');
});

test('opening the list starts on the chosen cue\'s own kind; Enter on a cue puts the list away and moves to the text',async()=>{
  const a=await started();
  a.pick('polling');
  click(a,'guide-cue-toggle');
  assert.equal(a.n('guide-cue-panel').hidden,false);
  assert.equal(a.n('guide-cue-toggle').getAttribute('aria-expanded'),'true');
  assert.equal(a.n('guide-cue-toggle-action').textContent,'Close');
  assert.equal(a.n('guide-cue-group').value,'virtual','the list shows the kind Polling belongs to');
  assert.match(a.n('guide-cue-count').textContent,/^13 cues match\./);
  // arrow keys through the radios choose without closing; Enter closes
  a.pick('vote');
  assert.equal(a.n('guide-cue-panel').hidden,false,'choosing by arrow key leaves the list open');
  let stopped=false;a.radio('vote').onkeydown({key:'Enter',preventDefault(){stopped=true;}});
  assert.equal(stopped,true);
  assert.equal(a.n('guide-cue-panel').hidden,true);
  assert.equal(a.n('guide-cue-toggle').getAttribute('aria-expanded'),'false');
  assert.equal(a.n('guide-cue-current').textContent,'Vote');
  // the bar's button opens it again, and closes it
  click(a,'guide-cue-toggle');assert.equal(a.n('guide-cue-panel').hidden,false);
  click(a,'guide-cue-toggle');assert.equal(a.n('guide-cue-panel').hidden,true);
});

test('a kind or a search the author already chose is not overwritten when the list opens',async()=>{
  const a=await started();
  a.type('guide-cue-find','hand');a.n('guide-cue-find').oninput();
  click(a,'guide-cue-toggle');
  assert.equal(a.n('guide-cue-group').value,'','the search stays as it was typed');
  assert.match(a.n('guide-cue-count').textContent,/^2 cues match\./);
  click(a,'guide-cue-toggle');
  a.type('guide-cue-find','');a.n('guide-cue-find').oninput();
  a.type('guide-cue-group','materials');a.n('guide-cue-group').onchange();
  click(a,'guide-cue-toggle');
  assert.equal(a.n('guide-cue-group').value,'materials','a kind chosen by hand stays');
});

const AP=['items','paste','bank'];
const ashown=a=>AP.filter(id=>!a.n(`guide-apart-${id}`).hidden);

test('the Assessments screen is three short parts: items, pasted questions, a question bank',async()=>{
  const a=await started();
  await click(a,'guide-tab-assess');
  assert.deepEqual(ashown(a),['items']);
  for(const id of AP){a.n(`guide-apart-tab-${id}`).onclick();assert.deepEqual(ashown(a),[id],id);assert.equal(a.n(`guide-apart-tab-${id}`).getAttribute('aria-selected'),'true');}
  a.n('guide-apart-tab-items').onclick();
  assert.equal(key(a,'guide-apart-tab-items','ArrowRight'),true);assert.deepEqual(ashown(a),['paste']);
  key(a,'guide-apart-tab-paste','End');assert.deepEqual(ashown(a),['bank']);
  key(a,'guide-apart-tab-bank','ArrowRight');assert.deepEqual(ashown(a),['items'],'wraps round');
});

test('starting a section brings the Items part back, with the question box ready',async()=>{
  const a=await started();
  await click(a,'guide-tab-assess');
  a.n('guide-apart-tab-paste').onclick();
  assert.deepEqual(ashown(a),['paste']);
  await click(a,'guide-sec-level-two');
  assert.deepEqual(ashown(a),['items']);
});

test('the type of item is chosen from a short bar like the cue, with the list behind it',async()=>{
  const a=await started();
  assert.equal(a.n('guide-type-panel').hidden,true);
  assert.equal(a.n('guide-type-current').textContent,'Multiple choice');
  click(a,'guide-type-toggle');
  assert.equal(a.n('guide-type-panel').hidden,false);assert.equal(a.n('guide-type-toggle').getAttribute('aria-expanded'),'true');assert.equal(a.n('guide-type-toggle-action').textContent,'Close');
  a.itemRadio('true-false').checked=true;a.itemRadio('true-false').onchange();
  assert.equal(a.n('guide-type-current').textContent,'True or false');
  assert.equal(a.n('guide-type-panel').hidden,false,'an arrow key through the radios leaves the list open');
  let stopped=false;a.itemRadio('true-false').onkeydown({key:'Enter',preventDefault(){stopped=true;}});
  assert.equal(stopped,true);assert.equal(a.n('guide-type-panel').hidden,true);assert.equal(a.n('guide-type-toggle-action').textContent,'Change');
  assert.equal(a.n('guide-truefalse')?.hidden??a.n('guide-item-truefalse').hidden,false,'the form changed with the type');
});

test('every control of the Assessments screen sits in one of its three parts',()=>{
  const html=readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8').replace(/\r\n/g,'\n');
  const at=id=>{const i=html.indexOf(`id="${id}"`);assert.ok(i>0,id);return i;};
  const range=part=>{const start=html.indexOf(`<div id="guide-apart-${part}"`);assert.ok(start>0,part);const next=AP.map(x=>html.indexOf(`<div id="guide-apart-${x}"`)).filter(i=>i>start).sort((x,y)=>x-y)[0]??html.indexOf('<div id="guide-workbook"');return [start,next];};
  const within=(id,part)=>{const [s,e]=range(part),i=at(id);assert.ok(i>s&&i<e,`${id} is in the ${part} part`);};
  for(const id of ['guide-sec-level-one','guide-sec-level-two','guide-sec-level-three','guide-item-form','guide-type-toggle','guide-item-types','guide-item-stem','guide-item-opt-6','guide-item-go'])within(id,'items');
  for(const id of ['guide-paste-form','guide-paste-text','guide-paste-go'])within(id,'paste');
  for(const id of ['guide-bank-form','guide-bank-file','guide-bank-go'])within(id,'bank');
  assert.ok(at('guide-assess-heading')<at('guide-assess-parts')&&at('guide-assess-selected')<at('guide-assess-parts'));
});