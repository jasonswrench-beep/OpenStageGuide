// The "Your cursor is in ..." box of the pane. In a block it holds a line of text, the Edit button, a note and the Write below this
// block button. Found in a real Word pane 320 pixels wide (2026-11-09, pictures for the Quick Start): the row did not wrap, so the text
// was squeezed into a column a few letters wide and the Write below this block button was pushed off the right edge.
// This test only guards the rules that make it wrap; the look itself was checked in a real Word.
import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const css=fs.readFileSync(new URL('../public/style.css',import.meta.url),'utf8');
const html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');

test('the cursor box wraps, and its note and Write below button each take a whole row',()=>{
  const bar=/#view-guide \.guide-selected-bar\{([^}]*)\}/.exec(css);
  assert.ok(bar,'the bar rule is there');
  assert.match(bar[1],/display:flex/);
  assert.match(bar[1],/flex-wrap:wrap/,'the row wraps instead of squeezing');
  assert.match(css,/#guide-selected-hint\{[^}]*flex:1 1 100%/,'the note takes its own row');
  assert.match(css,/#guide-leave-block\{[^}]*flex:1 1 100%/,'the Write below button takes its own row');
});

test('reading order matches the look: cursor line, Edit, the note, then Write below this block',()=>{
  const at=id=>html.indexOf('id="'+id+'"');
  const order=['guide-selected-text','guide-edit-selected','guide-selected-hint','guide-leave-block'].map(at);
  assert.ok(order.every(n=>n>0),'all four parts are in the page');
  assert.deepEqual([...order].sort((a,b)=>a-b),order,'they are in that order in the page');
});
