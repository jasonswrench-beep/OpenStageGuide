// The three older panels (Timing & Agenda, Learning Alignment, Delivery Versions) still need the older course record: they say so plainly,
// say the ribbon's tool of the same name now works from the guide, and never send the author to a tab that is not there
// (reported 2026-11-07: "Create your course on the Course tab first" - but there is no Course tab in the new pane).
import test from 'node:test';
import assert from 'node:assert/strict';
import {makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';

let sequence=0;
async function app(project=null){
  const h=makeHost(project),ui=dom(),n=id=>ui.nodes.get(id);
  Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},fetch:async()=>{throw new Error('no network');},
    localStorage:{getItem(){return null;},setItem(){}}});
  await import(`../src/app.mjs?legacyCourse=${++sequence}`);await settle();
  const go=async name=>{const route=ui.routes.find(b=>b.dataset.view===name);assert.ok(route,name);await route.onclick();await settle();};
  return {n,go};
}

test('the older Timing & Agenda, Learning Alignment and Delivery Versions panels say they need the older course, and that the ribbon tool works from the guide',async()=>{
  const a=await app(null);
  for(const [view,name] of [['agenda','Timing & Agenda'],['edit-activity','Learning Alignment'],['pack','Delivery Versions']]){
    await a.go(view);
    const said=a.n('status').textContent;
    assert.ok(said.startsWith(`${name} on the ribbon now works from your guide.`),view+': '+said);
    assert.match(said,/older Course tool/);
    assert.doesNotMatch(said,/Course tab/,'there is no Course tab to go to');
  }
});

test('an older panel that needs a course tells where the Course tool is',async()=>{
  const a=await app(null);
  await a.go('outline');
  assert.match(a.n('status').textContent,/Open Legacy tools at the bottom of the pane, choose Course, and create it\./);
  assert.doesNotMatch(a.n('status').textContent,/Course tab/);
});
