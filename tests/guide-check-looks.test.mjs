// The Review's checks of how words look: size, alignment, the contrast of colored words, what stands in for a heading, and what leans on color
// or underlining alone. A guide the program made itself raises none of them.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {startGuide} from '../src/guide-ooxml.mjs';
import {checkParts} from '../src/guide-check.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RULES=['TEXT-SIZE','JUSTIFIED','FAUX-HEADING','COLOR-ONLY','UNDERLINE','CONTRAST'];
const p=(runs,pPr='')=>`<w:p xmlns:w="${W}">${pPr?`<w:pPr>${pPr}</w:pPr>`:''}${runs}</w:p>`;
const r=(t,rPr='')=>`<w:r>${rPr?`<w:rPr>${rPr}</w:rPr>`:''}<w:t xml:space="preserve">${t}</w:t></w:r>`;

async function review(bodies,edits){
  const a=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  if(edits)await a.setStyleEdits(edits);
  await a.insert(()=>bodies.join(''),{where:'start',spare:false});
  return checkParts(await a.reviewParts());
}
const of=(result,rule,level)=>result.findings.filter(f=>f.rule===rule&&(!level||f.level===level));

test('a guide the program made itself raises none of the look checks, and the ones that can pass say so',async()=>{
  const a=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  await a.insert(c=>startGuide({title:'A guide',standard:true},c),{where:'start',select:'last'});
  const result=checkParts(await a.reviewParts());
  const bad=result.findings.filter(f=>RULES.includes(f.rule)&&f.level!=='pass');
  assert.deepEqual(bad,[]);
  assert.ok(of(result,'TEXT-SIZE','pass').length&&of(result,'JUSTIFIED','pass').length);
});

test('colored words that are hard to read fail, by the contrast they need: more for small words, less for large ones',async()=>{
  const result=await review([p(r('Pale words','<w:color w:val="BBBBBB"/>')),p(r('Plenty dark','<w:color w:val="222222"/>'))]);
  const fail=of(result,'CONTRAST','fail');
  assert.equal(fail.length,1);
  assert.match(fail[0].message,/Colored words are hard to read: "Pale words" in BBBBBB on FFFFFF is 1\.9:1; it needs 4\.5:1\./);
  assert.equal(fail[0].paragraph,0);
  const large=await review([p(r('Big and mid grey','<w:color w:val="767676"/><w:sz w:val="40"/>')),p(r('Small and mid grey','<w:color w:val="949494"/><w:sz w:val="22"/>'))]);
  assert.equal(of(large,'CONTRAST','fail').length,1,'the same family of grey passes as a large word and fails as a small one');
  assert.match(of(large,'CONTRAST','fail')[0].message,/Small and mid grey/);
  const mid=await review([p(r('Mid','<w:color w:val="8E8E8E"/><w:sz w:val="40"/>'))]);
  assert.equal(of(mid,'CONTRAST','fail').length,0,'3:1 is enough at 20 points');
  const fine=await review([p(r('Dark','<w:color w:val="1F3A5F"/>'))]);
  assert.equal(of(fine,'CONTRAST','pass').filter(f=>/colored words meet/.test(f.message)).length,1);
});

test('colored words on a shaded paragraph are measured against the shading, not the page',async()=>{
  const result=await review([p(r('Yellow on cream','<w:color w:val="E6D88A"/>'),'<w:shd w:val="clear" w:color="auto" w:fill="FFF4CE"/>')]);
  const fail=of(result,'CONTRAST','fail');
  assert.equal(fail.length,1);
  assert.match(fail[0].message,/on FFF4CE/);
});

test('words under 10 points are flagged, whether they are small by their own size or their style\'s',async()=>{
  const direct=await review([p(r('Fine print','<w:sz w:val="16"/>')),p(r('Fine'))]);
  assert.equal(of(direct,'TEXT-SIZE','warn').length,1);
  assert.match(of(direct,'TEXT-SIZE','warn')[0].message,/One paragraph has words under 10 points/);
  const viaStyle=await review([p(r('Notes in a small style'),'<w:pStyle w:val="OSGNotes"/>')],{overrides:{OSGNotes:{font:{size:8}}},custom:[]});
  assert.equal(of(viaStyle,'TEXT-SIZE','warn').length,1);
  const two=await review([p(r('a','<w:sz w:val="16"/>')),p(r('b','<w:sz w:val="14"/>'))]);
  assert.match(of(two,'TEXT-SIZE','warn')[0].message,/2 paragraphs have words under 10 points/);
  const none=await review([p(r('Ten points exactly','<w:sz w:val="20"/>'))]);
  assert.equal(of(none,'TEXT-SIZE','warn').length,0);
});

test('justified paragraphs are flagged, whether justified by hand or by their style',async()=>{
  const direct=await review([p(r('Spread to both edges'),'<w:jc w:val="both"/>'),p(r('Left'))]);
  assert.equal(of(direct,'JUSTIFIED','warn').length,1);
  assert.match(of(direct,'JUSTIFIED','warn')[0].message,/dyslexia/);
  const viaStyle=await review([p(r('In a justified style'),'<w:pStyle w:val="OSGCueBody"/>')],{overrides:{OSGCueBody:{paragraph:{alignment:'justify'}}},custom:[]});
  assert.equal(of(viaStyle,'JUSTIFIED','warn').length,1);
  assert.equal(of(await review([p(r('Centered'),'<w:jc w:val="center"/>')]),'JUSTIFIED','warn').length,0);
});

test('a short paragraph that is all bold or all large, with no style of a heading, is flagged as a heading in everything but its style',async()=>{
  const bold=await review([p(r('Safety first','<w:b/>')),p(r('A whole sentence in bold, which is emphasis, not a heading.','<w:b/>')),p(r('Some ','<w:b/>')+r('mixed words')),p(r('Safety again','<w:b/>'),'<w:pStyle w:val="Heading2"/>')]);
  const f=of(bold,'FAUX-HEADING','warn');
  assert.equal(f.length,1);
  assert.equal(f[0].paragraph,0);
  assert.match(f[0].message,/looks like a heading/);
  const big=await review([p(r('Big words','<w:sz w:val="36"/>'))]);
  assert.equal(of(big,'FAUX-HEADING','warn').length,1);
  const list=await review([p(r('Bold list item','<w:b/>'),'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')]);
  assert.equal(of(list,'FAUX-HEADING','warn').length,0,'a list item is not a heading');
});

test('color that is the only thing setting words apart is flagged; bold, italic, a size or a character style alongside it, or a grey, is not',async()=>{
  const result=await review([
    p(r('Take care of ')+r('the red words','<w:color w:val="C00000"/>')),
    p(r('Take care of ')+r('these bold red words','<w:b/><w:color w:val="C00000"/>')),
    p(r('Take care of ')+r('these italic red words','<w:i/><w:color w:val="C00000"/>')),
    p(r('Take care of ')+r('these dark grey words','<w:color w:val="444444"/>')),
    p(r('All of it is red','<w:color w:val="C00000"/>')),
  ]);
  const f=of(result,'COLOR-ONLY','warn');
  assert.equal(f.length,1);
  assert.equal(f[0].paragraph,0);
  assert.match(f[0].message,/color is the only thing/);
});

test('underlined words that are not a link are flagged',async()=>{
  const result=await review([p(r('See ')+r('this one','<w:u w:val="single"/>')),p(r('Plain')),p(r('Not underlined','<w:u w:val="none"/>'))]);
  const f=of(result,'UNDERLINE','warn');
  assert.equal(f.length,1);
  assert.equal(f[0].paragraph,0);
  assert.match(f[0].message,/look like links/);
});
