import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {checkParts} from '../src/guide-check.mjs';
import {NS,stylesPart} from '../src/guide-styles.mjs';

// Review's "two or more empty paragraphs in a row" rule. In the author's Test 3 guide it fired at the very end of the
// document: Word's own last paragraph, plus the one its body export adds after it.
const run=text=>`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const empty='<w:p/>';
const sect='<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
const check=body=>checkParts({
  root:new DOMParser().parseFromString(`<w:document ${NS}><w:body>${body}</w:body></w:document>`,'application/xml'),
  stylesRoot:new DOMParser().parseFromString(stylesPart(),'application/xml'),
  title:'A guide',language:'en-US'
});
const emptyRule=body=>check(body).findings.find(f=>f.rule==='EMPTY-PARAGRAPHS');

test('empty paragraphs at the very end of the document are not counted',()=>{
  assert.equal(emptyRule(`${run('one')}${run('two')}${empty}${empty}${sect}`),undefined);
  assert.equal(emptyRule(`${run('one')}${empty}${empty}${empty}`),undefined,'with no section properties after them either');
});

test('two empty paragraphs in the middle of the document are still reported, with where they are',()=>{
  const f=emptyRule(`${run('one')}${empty}${empty}${run('two')}${empty}${sect}`);
  assert.match(f.message,/1 place with two or more empty paragraphs in a row/);
  assert.equal(f.paragraph,1);
});

test('a single empty paragraph in the middle is fine, and so is one at the end',()=>{
  assert.equal(emptyRule(`${run('one')}${empty}${run('two')}${empty}${sect}`),undefined);
});

test('two places in the middle are counted as two',()=>{
  const f=emptyRule(`${run('a')}${empty}${empty}${run('b')}${empty}${empty}${empty}${run('c')}${empty}${empty}${sect}`);
  assert.match(f.message,/2 places with two or more empty paragraphs in a row/);
});

test('empty paragraphs inside a table cell are still counted there',()=>{
  const cell=`<w:tbl><w:tr><w:tc>${run('x')}${empty}${empty}</w:tc></w:tr></w:tbl>`;
  assert.match(emptyRule(`${run('one')}${cell}${run('two')}${sect}`).message,/1 place/);
});
