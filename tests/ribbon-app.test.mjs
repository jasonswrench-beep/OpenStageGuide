// The ribbon, end to end in memory: the real app module, the real page ids and a stand-in for Word. A ribbon button (or "More tools")
// opens one tool: the guide's own tools are the guide pane's work (every tool is one now; the older panels remain behind Legacy tools).
// Not native Word.
import test from 'node:test';
import assert from 'node:assert/strict';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {TOOLS,TOOL_IDS,PROGRAM_NAMES} from '../src/guide-tools.mjs';

async function boot(tag){
  const h=makeHost(fixture()),ui=dom(),ribbon=new Map();
  h.Office.actions.associate=(name,handler)=>ribbon.set(name,handler);
  h.Office.addin.showAsTaskpane=async()=>{};
  Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''}});
  await import(`../src/app.mjs?${tag}`);await settle();
  const n=id=>ui.nodes.get(id),press=async id=>{let done=0;await ribbon.get(id)({completed(){done++;}});await settle();return done;};
  const moreButton=id=>n('tool-more-list').children.map(li=>li.children[0]).find(b=>b.dataset.tool===id);
  const shown=()=>[...ui.nodes.values()].filter(x=>x.id.startsWith('view-')&&!x.hidden).map(x=>x.id);
  return {h,ui,ribbon,n,press,shown,moreButton};
}

test('the pane starts on one tool, Open Designer, with its More tools list; every ribbon button has a function',async()=>{
  const {ribbon,n,shown}=await boot('ribbon-tools-start');
  for(const id of TOOL_IDS)assert.equal(typeof ribbon.get(id),'function',id);
  assert.equal(ribbon.size,TOOL_IDS.length+6,'the twenty-one buttons and the six keyboard navigation actions');
  assert.equal(n('tool-strip').hidden,false);
  assert.deepEqual(shown(),['view-guide']);
  assert.ok(n('tool-more-list').children.length>=1);
});

test('the two backup buttons are guide tools now (rebuilt 2026-11-08): the guide pane with their own heading, not the older Backups page; no button opens an older panel any more',async()=>{
  const {n,press,shown}=await boot('ribbon-tools-backups');
  assert.deepEqual(Object.values(TOOLS).filter(t=>t.kind==='legacy').map(t=>t.id),[]);
  for(const id of ['saveBackup','restoreBackup']){
    assert.equal(await press(id),1,`${id} completes its command`);
    assert.deepEqual(shown(),['view-guide'],id);
    assert.equal(n('tool-strip').hidden,false);assert.equal(n('tool-title').textContent,PROGRAM_NAMES[id]);
    assert.ok(n('tool-more-list').children.length>=1,id);
  }
});

test('from one tool to another and back to the same tool, the strip follows (found in a real Word, 2026-11-07, when one of them was an older panel)',async()=>{
  const {n,press,moreButton}=await boot('ribbon-tools-return');
  const more=()=>n('tool-more-list').children.map(li=>li.children[0].dataset.tool);
  await press('help');
  assert.deepEqual(more(),TOOLS.help.related);
  moreButton('saveBackup').onclick();await settle();
  assert.deepEqual(more(),TOOLS.saveBackup.related,'the other tool has its own strip');
  moreButton('help').onclick();await settle();
  assert.deepEqual(more(),TOOLS.help.related,'and back on Help the strip says so, though the pane was already on it');
  // the same for a guide tool
  await press('addModule');
  moreButton('courseOutline').onclick();await settle();
  moreButton('addModule').onclick();await settle();
  assert.equal(n('tool-title').textContent,PROGRAM_NAMES.addModule);
  assert.deepEqual(more(),TOOLS.addModule.related);
});

test('the Activity Library is a guide tool now: the guide pane with its own panel, not the older library page',async()=>{
  const {n,press,shown}=await boot('ribbon-tools-library');
  assert.equal(await press('activityLibrary'),1);
  assert.deepEqual(shown(),['view-guide']);
  assert.equal(n('tool-title').textContent,PROGRAM_NAMES.activityLibrary);
  // (this document has no guide yet, so the pane shows Start; the library's own panel is tested in activity-library.test.mjs)
  assert.ok(n('guide-lib-list').children.length>=5,'with ready-made activities to choose from');
});

test('a guide tool shows the guide pane with the tool\'s own heading, and More tools moves on to the next tool',async()=>{
  const {n,press,shown,moreButton}=await boot('ribbon-tools-guide');
  await press('timingAndAgenda');
  await press('addModule');
  assert.deepEqual(shown(),['view-guide']);
  assert.equal(n('tool-title').textContent,PROGRAM_NAMES.addModule);assert.equal(n('tool-title').hidden,false);
  const more=n('tool-more-list').children.map(li=>li.children[0].dataset.tool);
  assert.deepEqual(more,TOOLS.addModule.related);
  // from a guide tool to another tool, by More tools (Save Backup used to open the older panel; it is a guide tool now)
  await press('help');
  moreButton('saveBackup').onclick();await settle();
  assert.deepEqual(shown(),['view-guide']);assert.equal(n('tool-title').textContent,PROGRAM_NAMES.saveBackup);
  // and back to a guide tool the same way
  await press('addActivity');
  assert.equal(n('tool-title').textContent,PROGRAM_NAMES.addActivity);
  moreButton('addModule').onclick();await settle();
  assert.equal(n('tool-title').textContent,PROGRAM_NAMES.addModule);
  assert.deepEqual(shown(),['view-guide']);
});

test('the older navigation buttons still work, and show their panel without a tool\'s strip over it',async()=>{
  const {ui,n,press,shown}=await boot('ribbon-tools-old-nav');
  await press('addModule');
  assert.equal(n('tool-strip').hidden,false);
  await ui.routes.find(x=>x.dataset.view==='course').onclick();await settle();
  assert.deepEqual(shown(),['view-course']);
  assert.equal(n('tool-strip').hidden,true);
  // the Guide button of the old navigation goes to Open Designer
  await ui.routes.find(x=>x.dataset.view==='guide').onclick();await settle();
  assert.deepEqual(shown(),['view-guide']);assert.equal(n('tool-strip').hidden,false);
});
