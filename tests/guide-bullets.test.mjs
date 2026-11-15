// A paragraph made a bullet, and the numbering the guide's bullets share (Format > Bullets, Ctrl+Shift+L). The paragraph's XML changes only by a list mark; the guide's Word
// package defines the one bulleted list they belong to, so a bullet is a bullet in the Word file the program makes (and an activity's steps no longer come out lettered).
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {isBulleted,canBullet,withBullet} from '../src/guide-edit.mjs';
import {BULLET_NUM_ID,numberingPart} from '../src/guide-styles.mjs';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {docxFromPackage} from '../src/guide-docx.mjs';
import {para,run,startGuide} from '../src/guide-ooxml.mjs';
import {describeGuide} from '../program/ui/preview.js';
import {guideToHtml} from '../src/guide-html.mjs';
import {rowKey} from '../program/ui/editor-model.js';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const P=inner=>`<w:p xmlns:w="${W}">${inner}</w:p>`;
const plain=P('<w:r><w:t>First point</w:t></w:r>');

test('a bullet is put on and taken off a paragraph without touching anything else in it',()=>{
  const on=withBullet(plain,true,io);
  assert.equal(isBulleted(on,io),true);assert.equal(isBulleted(plain,io),false);
  assert.match(on,new RegExp(`<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${BULLET_NUM_ID}"/></w:numPr></w:pPr>`));
  assert.ok(on.includes('First point'),'its words are as they were');
  assert.equal(withBullet(on,true,io),on,'on twice is on once');
  const off=withBullet(on,false,io);
  assert.equal(isBulleted(off,io),false);assert.ok(!off.includes('<w:pPr'),'a paragraph with nothing else in its properties has none again');
  assert.equal(withBullet(plain,false,io),plain,'off for a paragraph with no bullet changes nothing');
});

test('the list mark goes where Word keeps it: after the style and what follows the style, before the rest',()=>{
  const styled=P('<w:pPr><w:pStyle w:val="OSGCueBody"/><w:keepNext/><w:spacing w:after="60"/><w:jc w:val="center"/></w:pPr><w:r><w:t>x</w:t></w:r>');
  const on=withBullet(styled,true,io);
  assert.match(on,/<w:pStyle w:val="OSGCueBody"\/><w:keepNext\/><w:numPr>.*<\/w:numPr><w:spacing w:after="60"\/><w:jc w:val="center"\/>/);
  const off=withBullet(on,false,io);
  assert.equal(off.replace(/ xmlns:w="[^"]+"/,''),styled.replace(/ xmlns:w="[^"]+"/,''),'taking it off gives back what was there');
  const other=P(`<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>a</w:t></w:r>`);
  assert.equal(isBulleted(other,io),false,'a lettered list is not the bulleted one');
  assert.equal(isBulleted(withBullet(other,true,io),io),true,'making it a bullet replaces its list');
  assert.equal(withBullet(other,false,io),other,'and taking bullets off leaves a lettered list alone');
});

test('only paragraphs of words can be bullets: not headings, a block\'s own line, the time line, a form line or a table',()=>{
  assert.equal(canBullet(plain,'Normal',io),true);
  assert.equal(canBullet(plain,'OSG Cue Body',io),true);
  for(const style of ['OSG Module','OSG Lesson','OSG Guide Title','heading 1','OSG Cue','OSG Cue Facilitator','OSG Time','OSG Form Line','TOC 1'])assert.equal(canBullet(plain,style,io),false,style);
  assert.equal(canBullet(`<w:tbl xmlns:w="${W}"/>`,'OSG Table Cell',io),false);
  assert.equal(canBullet('',`Normal`,io),false);
});

test('the guide\'s bullets share one list in its Word package, which defines it as a bullet; lettered lists are as they were',async()=>{
  const part=numberingPart(2,new Set(),true);
  assert.match(part,new RegExp(`<w:num w:numId="${BULLET_NUM_ID}"><w:abstractNumId w:val="1"/></w:num>`));
  assert.match(part,/<w:numFmt w:val="bullet"\/>/);assert.match(part,/<w:numFmt w:val="lowerLetter"\/>/);
  assert.ok(!numberingPart(2).includes('numFmt w:val="bullet"'),'with no bullets there is no bulleted list');
  const a=demo();
  await a.insert(c=>startGuide({title:'Bullets',standard:false},c)+para('OSGCueBody',run('Point one'))+para('OSGCueBody',run('Point two')),{where:'start',select:'last'});
  const ps=a.exportState().paragraphs,at=ps.findIndex(p=>p.text==='Point one');
  await a.replaceParagraphs({start:at,count:1,guard:'Point one',xmls:[withBullet(ps[at].xml,true,io)],select:at});
  const pkg=await a.sourcePackage();
  assert.match(pkg,/\/word\/numbering\.xml/);assert.match(pkg,new RegExp(`<w:numId w:val="${BULLET_NUM_ID}"/>`));
  const zip=await JSZip.loadAsync(await docxFromPackage(pkg,JSZip,io,{entries:a.exportState().paragraphs}));
  const numbering=await zip.file('word/numbering.xml').async('string');
  assert.match(numbering,new RegExp(`<w:num w:numId="${BULLET_NUM_ID}"><w:abstractNumId w:val="1"/></w:num>`));
  assert.match(await zip.file('word/_rels/document.xml.rels').async('string'),/relationships\/numbering/);
  assert.match(await zip.file('word/document.xml').async('string'),new RegExp(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${BULLET_NUM_ID}"/></w:numPr>`));
});

test('an activity\'s steps, which are a bulleted list, join the shared bulleted list when they are put in the program\'s guide',async()=>{
  const a=demo();
  await a.insert(ctx=>{ctx.lists=1;ctx.bullets.add(1);return para('OSGCueBody',run('Step one'),'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');},{where:'start',select:'last'});
  const xml=a.exportState().paragraphs.find(p=>p.text==='Step one').xml;
  assert.match(xml,new RegExp(`<w:numId w:val="${BULLET_NUM_ID}"/>`));assert.doesNotMatch(xml,/<w:numId w:val="1"\/>/);
  assert.match(await a.sourcePackage(),/<w:numFmt w:val="bullet"\/>/);
});

test('a bullet is drawn on the page and in the web page, and the key and the menu name it',()=>{
  const entries=[{xml:withBullet(plain,true,io),text:'First point',style:'Normal'},{xml:plain,text:'First point',style:'Normal'}];
  const nodes=describeGuide(entries,x=>new DOMParser().parseFromString(x,'application/xml'));
  assert.match(nodes[0].cls,/\bg-bullet\b/);assert.doesNotMatch(nodes[1].cls,/g-bullet/);
  const html=guideToHtml({nodes,title:'T'});
  assert.match(html,/<p class="g-para g-bullet"><span>First point<\/span><\/p>/);
  assert.match(html,/\.g-bullet::before\{content:"•"/);
  const key=(over)=>rowKey({key:'L',ctrl:true,meta:false,alt:false,shift:true,...over});
  assert.equal(key(),'bullets');assert.equal(key({shift:false,key:'l'}),'alignLeft','Ctrl+L is still align left');
  assert.equal(key({ctrl:false}),null);assert.equal(key({alt:true}),null);
});
