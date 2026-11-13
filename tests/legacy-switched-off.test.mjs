import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

// The author's word, 2026-11-08: the older panels are not needed, and a "Legacy tools" label makes the first version sound like a revision.
// The section is switched off (hidden), not deleted: the older panels' code and tests remain until someone deletes them on purpose.
const html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');

test('the Legacy tools section cannot be seen or reached: its whole block is hidden',()=>{
  const open=/<aside class="legacy-foot"[^>]*>/.exec(html);
  assert.ok(open,'the older section is still in the page (switched off, not deleted)');
  assert.match(open[0],/\shidden[\s>]/,'the section is hidden');
  const css=fs.readFileSync(new URL('../public/style.css',import.meta.url),'utf8');
  assert.match(css,/\[hidden\]\{display:none!important\}/,'hidden wins over every other rule');
});

test('no word "Legacy" shows anywhere in the page outside that hidden block',()=>{
  const open=html.indexOf('<aside class="legacy-foot"'),close=html.indexOf('</aside>',open);
  const outside=html.slice(0,open)+html.slice(close);
  assert.doesNotMatch(outside,/legacy/i);
});
