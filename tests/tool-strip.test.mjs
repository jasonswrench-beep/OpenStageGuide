// The strip above the open tool: its name, its note, and "More tools". Real page ids, in-memory DOM; not a browser.
import test from 'node:test';
import assert from 'node:assert/strict';
import {dom} from './helpers/app-dom.mjs';
import {mountToolStrip,namesItself} from '../src/tool-strip.mjs';
import {TOOLS,TOOL_IDS,relatedTools} from '../src/guide-tools.mjs';

const make=()=>{const ui=dom(),opened=[],strip=mountToolStrip(ui.document,{open:id=>opened.push(id)});return {ui,strip,opened,n:id=>ui.nodes.get(id)};};
const labels=n=>n('tool-more-list').children.map(li=>li.children[0].textContent);

test('a tool without a heading of its own gets its name and note in the strip, and the short list of tools that belong with it',()=>{
  const {strip,n}=make();
  strip.show('addActivity');
  assert.equal(n('tool-strip').hidden,false);
  assert.equal(n('tool-title').hidden,false);assert.equal(n('tool-title').textContent,'Add an activity');
  assert.equal(n('tool-note').hidden,false);assert.equal(n('tool-note').textContent,TOOLS.addActivity.note);
  assert.deepEqual(labels(n),['Add Module','Instructional Cues','Activity Library']);
  assert.equal(n('tool-more').hidden,false);assert.equal(n('tool-more').open,false,'the list is folded until it is wanted');
  assert.equal(strip.current(),'addActivity');
});

test('a tool whose screen names itself leaves the naming to the screen, but keeps More tools',()=>{
  const {strip,n}=make();
  for(const id of ['assessments','reviewGuide','participantWorkbook','help','openDesigner']){
    strip.show(id);
    assert.ok(namesItself(TOOLS[id]),id);
    assert.equal(n('tool-title').hidden,true,`${id}: no second heading`);
    assert.equal(n('tool-note').hidden,true);
    assert.equal(n('tool-more').hidden,false);assert.ok(n('tool-more-list').children.length>=1);
  }
});

test('More tools lists only the tools named for the open tool, as buttons that open them, with what each is for',()=>{
  const {strip,n,opened}=make();
  for(const id of TOOL_IDS){
    strip.show(id);
    const buttons=n('tool-more-list').children.map(li=>li.children[0]);
    assert.deepEqual(buttons.map(b=>b.dataset.tool),relatedTools(id).map(t=>t.id),id);
    assert.ok(buttons.length<TOOL_IDS.length-10,`${id} does not list the whole ribbon`);
    for(const b of buttons){
      assert.equal(b.type,'button');
      assert.match(b.getAttribute('aria-label'),new RegExp(`^${TOOLS[b.dataset.tool].label}: `));
    }
  }
  strip.show('assessments');
  n('tool-more-list').children[0].children[0].onclick();
  assert.deepEqual(opened,['learningAlignment']);
});

test('showing another tool folds the list again and replaces it, and an unknown tool hides the strip',()=>{
  const {strip,n}=make();
  strip.show('addModule');n('tool-more').open=true;
  strip.show('appearance');
  assert.equal(n('tool-more').open,false);
  assert.deepEqual(labels(n),['Generate Materials','Course Settings','Participant Workbook']);
  strip.show('nonsense');
  assert.equal(n('tool-strip').hidden,true);assert.equal(strip.current(),null);
  strip.show('help');strip.hide();
  assert.equal(n('tool-strip').hidden,true);
});
