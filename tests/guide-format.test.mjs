// Formatting a paragraph by hand (src/guide-format.mjs): alignment, space, line spacing and indents, kept in the order Word's schema gives.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {withParagraphFormat,readParagraphFormat,canFormat,formatOf,ALIGNMENTS,FORMAT_LIMITS} from '../src/guide-format.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const plain=`<w:p xmlns:w="${W}"><w:r><w:t>Hello</w:t></w:r></w:p>`;
const styled=`<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="OSGCueBody"/><w:keepNext/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Step</w:t></w:r></w:p>`;
const EMPTY={alignment:null,spaceBefore:null,spaceAfter:null,lineSpacing:null,leftIndent:null,firstLineIndent:null};
const order=xml=>{const p=Array.from(io.parse(xml).documentElement.childNodes).find(n=>n.localName==='pPr');return p?Array.from(p.childNodes).filter(n=>n.nodeType===1).map(n=>n.localName):[];};

test('a paragraph with no formatting of its own reads as following its style',()=>{
  assert.deepEqual(readParagraphFormat(plain,io),EMPTY);
  assert.deepEqual(readParagraphFormat(styled,io),EMPTY);
});

test('alignment, space, line spacing and indents are set and read back, in points and in multiples of a line',()=>{
  const out=withParagraphFormat(plain,'Normal',{alignment:'center',spaceBefore:6,spaceAfter:12,lineSpacing:1.5,leftIndent:36,firstLineIndent:18},io);
  assert.deepEqual(readParagraphFormat(out,io),{alignment:'center',spaceBefore:6,spaceAfter:12,lineSpacing:1.5,leftIndent:36,firstLineIndent:18});
  assert.match(out,/<w:jc w:val="center"\/>/);
  assert.match(out,/w:before="120"/);assert.match(out,/w:after="240"/);assert.match(out,/w:line="360" w:lineRule="auto"/);
  assert.match(out,/w:left="720"/);assert.match(out,/w:firstLine="360"/);
  const hang=withParagraphFormat(plain,'Normal',{leftIndent:36,firstLineIndent:-18},io);
  assert.match(hang,/w:hanging="360"/);assert.equal(readParagraphFormat(hang,io).firstLineIndent,-18);
  for(const [name,word] of Object.entries(ALIGNMENTS)){
    const a=withParagraphFormat(plain,'Normal',{alignment:name},io);
    assert.match(a,new RegExp(`<w:jc w:val="${word}"/>`));
    assert.equal(readParagraphFormat(a,io).alignment,name);
  }
});

test('a value left out is left as it is, and null takes the paragraph\'s own setting away; with none left the properties go too',()=>{
  const a=withParagraphFormat(plain,'Normal',{alignment:'right',spaceAfter:8},io);
  const b=withParagraphFormat(a,'Normal',{spaceBefore:4},io);
  assert.deepEqual(readParagraphFormat(b,io),{...EMPTY,alignment:'right',spaceBefore:4,spaceAfter:8},'the others stayed');
  const c=withParagraphFormat(b,'Normal',{spaceAfter:null},io);
  assert.deepEqual(readParagraphFormat(c,io),{...EMPTY,alignment:'right',spaceBefore:4});
  assert.equal(withParagraphFormat(b,'Normal',EMPTY,io),plain,'back to what it was, with no empty properties left');
  assert.equal(withParagraphFormat(withParagraphFormat(styled,'OSG Cue Body',{alignment:'center',spaceAfter:6},io),'OSG Cue Body',EMPTY,io),styled,'a paragraph with a style and a list goes back exactly');
});

test('the properties stay in the order Word\'s schema gives, whatever order they were set in, and a style, a list and other properties are kept',()=>{
  let x=withParagraphFormat(styled,'OSG Cue Body',{alignment:'justify'},io);
  x=withParagraphFormat(x,'OSG Cue Body',{leftIndent:18},io);
  x=withParagraphFormat(x,'OSG Cue Body',{spaceBefore:3,lineSpacing:1.15},io);
  assert.deepEqual(order(x),['pStyle','keepNext','numPr','spacing','ind','jc']);
  assert.match(x,/<w:numId w:val="2"\/>/);
  const y=withParagraphFormat(plain,'Normal',{alignment:'left'},io);
  assert.deepEqual(order(y),['jc']);
});

test('values outside what Word and a page can use are refused, and nothing changes',()=>{
  for(const bad of [{spaceBefore:-1},{spaceBefore:145},{spaceAfter:'6'},{spaceAfter:NaN},{lineSpacing:0.1},{lineSpacing:5},{leftIndent:-1},{leftIndent:999},{firstLineIndent:-100},{firstLineIndent:100},{alignment:'diagonal'}])
    assert.throws(()=>withParagraphFormat(plain,'Normal',bad,io),/must be|Alignment/,JSON.stringify(bad));
  assert.deepEqual(FORMAT_LIMITS.line,[0.8,3]);
});

test('a paragraph whose spacing the program sets (a line to write on, a box, a row of fields, the time line) or that is not a paragraph is not formatted by hand',async()=>{
  const a=await started('F');
  await addModule(a,'M');await addLesson(a,'L','5');
  const time=a.adapter.exportState().paragraphs.find(p=>p.style==='OSG Time');
  for(const style of ['OSG Form Line','OSG Form Box','OSG Form Field','OSG Answer Line','OSG Time','TOC 1'])assert.equal(canFormat(plain,style,io).ok,false,style);
  assert.throws(()=>withParagraphFormat(time.xml,'OSG Time',{alignment:'center'},io),/program sets the spacing/);
  assert.equal(canFormat('',  'Normal',io).ok,false);
  assert.equal(canFormat(`<w:tbl xmlns:w="${W}"/>`,'Normal',io).ok,false);
  assert.equal(canFormat(plain,'Normal',io).ok,true);
  const exact=`<w:p xmlns:w="${W}"><w:pPr><w:spacing w:before="120" w:after="120" w:line="1080" w:lineRule="exact"/></w:pPr></w:p>`;
  assert.throws(()=>withParagraphFormat(exact,'Normal',{lineSpacing:1.5},io),/line height of a set size/);
  assert.match(withParagraphFormat(exact,'Normal',{alignment:'center'},io),/w:lineRule="exact"/,'its set height is left alone when only the alignment changes');
});

test('every paragraph a guide has that can be formatted can be, without changing its words, its style or what else is in it',async()=>{
  const a=await started('F');
  await addModule(a,'Naming it');await addLesson(a,'First','10');await addBlock(a,'ask','What makes it hard?',{notes:'Listen.',minutes:'5'});
  let done=0;
  for(const p of a.adapter.exportState().paragraphs){
    if(!p.xml||!canFormat(p.xml,p.style,io).ok)continue;
    const out=withParagraphFormat(p.xml,p.style,{alignment:'center',spaceAfter:9},io);
    const back=withParagraphFormat(out,p.style,{alignment:null,spaceAfter:null},io);
    assert.equal(back,p.xml,`${p.style}: formatting and then clearing it is the identity`);
    assert.equal(io.parse(out).documentElement.textContent,io.parse(p.xml).documentElement.textContent,'the words are the same');
    done++;
  }
  assert.ok(done>=10,'a good many paragraphs: '+done);
});

test('formatOf reads an element the same way, for the page',()=>{
  const el=io.parse(withParagraphFormat(plain,'Normal',{alignment:'right',spaceBefore:2},io)).documentElement;
  assert.deepEqual(formatOf(el),{...EMPTY,alignment:'right',spaceBefore:2});
});
