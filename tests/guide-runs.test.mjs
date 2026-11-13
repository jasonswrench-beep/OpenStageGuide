// Formatting inside a sentence: the settings of a run, the spans a paragraph's words are held as, and that editing the words, splitting and
// joining paragraphs keeps each stretch's look and everything Word wrote that the program does not model.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {cleanChange,fmtOfRpr,withFmt,runCss,normalize,replaceText,applyChange,slice,rprAt,summarize,textOf,sameSpans} from '../src/guide-runs.mjs';
import {readParagraph,setSpans,setWords,splitWords,joinWords,paragraphWithSpans,nextParagraphSpec} from '../src/guide-edit.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run=(text,rPr='')=>`<w:r>${rPr?`<w:rPr>${rPr}</w:rPr>`:''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para=(runs,pPr='')=>`<w:p xmlns:w="${W}">${pPr?`<w:pPr>${pPr}</w:pPr>`:''}${runs}</w:p>`;
const B='<w:rPr><w:b/><w:bCs/></w:rPr>';

test('the settings of a run are read, and a change is laid over them in the order Word\'s schema gives, whatever order the changes came in',()=>{
  let rpr='';
  for(const change of [{size:14},{color:'#c00000'},{script:'superscript'},{underline:true},{italic:true},{bold:true},{font:'Georgia'},{strike:true}])rpr=withFmt(rpr,change,io);
  const order=[...rpr.matchAll(/<w:(\w+)[ />]/g)].map(m=>m[1]).filter(n=>n!=='rPr');
  assert.deepEqual(order,['rFonts','b','bCs','i','iCs','strike','color','sz','szCs','u','vertAlign'],'rFonts, b, i, strike, color, sz, u, vertAlign: the order Word wants');
  assert.deepEqual(fmtOfRpr(rpr,io),{bold:true,italic:true,strike:true,underline:true,script:'superscript',color:'#C00000',font:'Georgia',size:14});
  assert.ok(!/xmlns:w=/.test(rpr),'the text of a run\'s properties does not repeat the namespace');
});

test('off is written as off, so a bold style can be turned off for some words; null takes the setting away so the style decides again',()=>{
  const off=withFmt('',{bold:false},io);
  assert.match(off,/<w:b w:val="0"\/>/);
  assert.equal(fmtOfRpr(off,io).bold,false);
  assert.equal(withFmt(off,{bold:null},io),'','nothing left, nothing written');
  assert.equal(withFmt(withFmt('',{italic:true,color:'#112233'},io),{italic:null},io),'<w:rPr><w:color w:val="112233"/></w:rPr>');
  assert.equal(fmtOfRpr(withFmt('',{underline:false},io),io).underline,false);
});

test('what the program does not model goes through a change untouched: the language, a theme font, a ligature setting',()=>{
  const w14='http://schemas.microsoft.com/office/word/2010/wordml';
  const rpr=`<w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:lang w:val="fr-CA"/><w14:ligatures xmlns:w14="${w14}" w14:val="standard"/></w:rPr>`;
  const out=withFmt(rpr,{bold:true},io);
  assert.match(out,/w:lang w:val="fr-CA"/);
  assert.match(out,/w:asciiTheme="minorHAnsi"/);
  assert.match(out,/w14:ligatures/);
  assert.match(out,/<w:b\/>/);
  assert.equal(fmtOfRpr(rpr,io).font,undefined,'a theme font is not a name the program can show, so it is not claimed');
  // And putting a font on replaces the theme's, so the two do not disagree.
  const named=withFmt(rpr,{font:'Arial'},io);
  assert.match(named,/w:ascii="Arial"/);
  assert.doesNotMatch(named,/asciiTheme/);
});

test('a change that is not valid is explained in words, and nothing is changed',()=>{
  assert.throws(()=>cleanChange({color:'red'}),/six hexadecimal digits/);
  assert.throws(()=>cleanChange({size:2}),/from 6 to 96/);
  assert.throws(()=>cleanChange({size:'big'}),/from 6 to 96/);
  assert.throws(()=>cleanChange({bold:'yes'}),/bold is on or off/);
  assert.throws(()=>cleanChange({script:'up'}),/superscript, subscript or neither/);
  assert.throws(()=>cleanChange({font:'<b>'}),/A font name can have/);
  assert.throws(()=>cleanChange({style:'a b'}),/not known/);
  assert.throws(()=>cleanChange({shadow:true}),/not a setting/);
  assert.throws(()=>cleanChange(null),/not in the right form/);
  assert.deepEqual(cleanChange({size:12.3,color:'#abcdef'}),{size:12.5,color:'#ABCDEF'},'half points, and colors in capitals');
});

test('spans: neighbours with one look are one, the words after a caret take the look before it, and a change splits only where it must',()=>{
  const spans=[{text:'Say ',rpr:''},{text:'this',rpr:B.replace('<w:rPr>','<w:rPr>')},{text:' now',rpr:''}];
  const rpr=spans[1].rpr;
  assert.equal(textOf(spans),'Say this now');
  assert.deepEqual(normalize([{text:'a',rpr:''},{text:'b',rpr:''},{text:'',rpr:rpr},{text:'c',rpr}]),[{text:'ab',rpr:''},{text:'c',rpr}]);
  assert.equal(rprAt(spans,6),rpr,'inside "this"');
  assert.equal(rprAt(spans,8),rpr,'at its end: the look of the character before');
  assert.equal(rprAt(spans,0),'','at the very start: the first character\'s');
  // typing inside bold words is bold; at the start of the paragraph it takes the first look; after bold it is bold
  assert.deepEqual(replaceText(spans,6,6,'XX'),[{text:'Say ',rpr:''},{text:'thXXis',rpr},{text:' now',rpr:''}]);
  assert.deepEqual(replaceText(spans,0,0,'Oh, '),[{text:'Oh, Say ',rpr:''},{text:'this',rpr},{text:' now',rpr:''}]);
  // deleting the bold words leaves one plain stretch
  assert.deepEqual(replaceText(spans,4,8,''),[{text:'Say  now',rpr:''}]);
  // making "y t" italic splits the spans on both sides and leaves the rest alone
  const italic=applyChange(spans,2,5,{italic:true},io);
  assert.deepEqual(italic.map(s=>s.text),['Sa','y ','t','his',' now']);
  assert.equal(fmtOfRpr(italic[1].rpr,io).italic,true);
  assert.equal(fmtOfRpr(italic[2].rpr,io).italic,true);
  assert.equal(fmtOfRpr(italic[2].rpr,io).bold,true,'bold stays under the new italic');
  assert.equal(fmtOfRpr(italic[3].rpr,io).italic,undefined);
  assert.deepEqual(slice(spans,2,6).map(s=>s.text),['y ','th']);
});

test('what a stretch is like, setting by setting: the same for all of it, or mixed, or not said',()=>{
  const spans=[{text:'Say ',rpr:''},{text:'this',rpr:withFmt('',{bold:true},io)},{text:' now',rpr:''}];
  assert.equal(summarize(spans,4,8,io).bold,true);
  assert.equal(summarize(spans,0,4,io).bold,undefined,'a plain stretch does not say');
  assert.equal(summarize(spans,0,8,io).bold,'mixed');
  assert.equal(summarize(spans,5,5,io).bold,true,'with no characters chosen, the character before the caret');
  assert.equal(summarize(spans,0,12,io).italic,undefined);
});

test('what a run looks like on a page: weight, slant, rules through, color, and raised or lowered',()=>{
  const css=runCss({bold:true,italic:true,underline:true,strike:true,color:'#C00000',font:'Georgia',size:14,script:'superscript'});
  assert.equal(css.fontWeight,'700');assert.equal(css.fontStyle,'italic');
  assert.match(css.textDecoration,/underline/);assert.match(css.textDecoration,/line-through/);
  assert.equal(css.color,'#C00000');assert.match(css.fontFamily,/Georgia/);
  assert.equal(css.verticalAlign,'super');
  assert.equal(runCss({script:'subscript'}).verticalAlign,'sub');
  assert.equal(runCss({script:'subscript'}).fontSize,'smaller','smaller, unless the size was chosen');
  assert.notEqual(runCss({script:'subscript',size:10}).fontSize,'smaller');
  assert.deepEqual(runCss({}),{});
});

// ---- in a paragraph

const mixed=para(`${run('Say ')}${run('this','<w:b/><w:bCs/>')}${run(' now','<w:i/><w:lang w:val="fr-CA"/>')}`,'<w:pStyle w:val="OSGCueBody"/>');

test('a paragraph\'s words are read as spans with the look each run had',()=>{
  const info=readParagraph(mixed,'OSG Cue Body',io);
  assert.equal(info.editable,true);
  assert.equal(info.text,'Say this now');
  assert.deepEqual(info.spans.map(s=>s.text),['Say ','this',' now']);
  assert.equal(info.spans[0].rpr,'');
  assert.equal(fmtOfRpr(info.spans[1].rpr,io).bold,true);
  assert.match(info.spans[2].rpr,/w:lang w:val="fr-CA"/);
});

test('the same spans give back the very same paragraph, and a change to them is only the runs that changed',()=>{
  const info=readParagraph(mixed,'OSG Cue Body',io);
  assert.equal(setSpans(mixed,'OSG Cue Body',info.spans,io),mixed,'nothing changed, nothing rewritten');
  assert.equal(setWords(mixed,'OSG Cue Body','Say this now',io),mixed);
  const next=setSpans(mixed,'OSG Cue Body',applyChange(info.spans,0,3,{bold:true},io),io);
  assert.match(next,/<w:pPr><w:pStyle w:val="OSGCueBody"\/><\/w:pPr>/,'the paragraph\'s properties stay');
  assert.deepEqual(readParagraph(next,'OSG Cue Body',io).spans.map(s=>[s.text,fmtOfRpr(s.rpr,io).bold]),[['Say',true],[' ',undefined],['this',true],[' now',undefined]]);
});

test('typing in the middle of formatted words keeps the formatting of the words around it',()=>{
  const out=setWords(mixed,'OSG Cue Body','Say this very now',io);
  const spans=readParagraph(out,'OSG Cue Body',io).spans;
  assert.equal(textOf(spans),'Say this very now');
  const at=t=>spans.find(s=>s.text.includes(t));
  assert.equal(fmtOfRpr(at('this').rpr,io).bold,true);
  assert.equal(fmtOfRpr(at('very').rpr,io).italic,true,'new words after "this " take the look of the character before them');
  assert.match(at('now').rpr,/w:lang w:val="fr-CA"/,'the language is still there');
  const deleted=setWords(mixed,'OSG Cue Body','Say  now',io);
  assert.equal(readParagraph(deleted,'OSG Cue Body',io).spans.length,2,'the bold word is gone, and with it its run');
});

test('a heading\'s number keeps its look when its words change, and the words are the ones typed',()=>{
  const heading=para(`${run('Module 1: ','<w:color w:val="1F3A5F"/>')}${run('Naming','<w:color w:val="1F3A5F"/>')}`,'<w:pStyle w:val="OSGModule"/>');
  const info=readParagraph(heading,'OSG Module',io);
  assert.equal(info.prefix,'Module 1: ');
  assert.equal(info.text,'Naming');
  assert.deepEqual(info.spans.map(s=>s.text),['Naming']);
  const out=setWords(heading,'OSG Module','Naming feelings',io);
  const again=readParagraph(out,'OSG Module',io);
  assert.equal(again.prefix,'Module 1: ');
  assert.equal(again.text,'Naming feelings');
  assert.match(out,/Module 1: /);
  assert.equal((out.match(/1F3A5F/g)||[]).length,1,'one run, one color: the prefix and the words are one stretch');
});

test('Enter in the middle of bold words makes two paragraphs, each with its half still bold',()=>{
  const p=para(run('Hello brave new world','<w:b/>'),'<w:pStyle w:val="OSGCueBody"/>');
  const cut=splitWords(p,'OSG Cue Body',11,io);
  assert.equal(textOf(readParagraph(cut.first,'OSG Cue Body',io).spans),'Hello brave');
  assert.equal(textOf(readParagraph(cut.second,'OSG Cue Body',io).spans),' new world');
  assert.equal(fmtOfRpr(readParagraph(cut.first,'OSG Cue Body',io).spans[0].rpr,io).bold,true);
  assert.equal(fmtOfRpr(readParagraph(cut.second,'OSG Cue Body',io).spans[0].rpr,io).bold,true);
  assert.match(cut.second,/OSGCueBody/,'the new paragraph is in the same style');
});

test('Backspace at the start joins two paragraphs and each keeps its own look',()=>{
  const a={xml:para(run('Plain '),'<w:pStyle w:val="OSGCueBody"/>'),style:'OSG Cue Body'};
  const b={xml:para(run('bold','<w:b/>'),'<w:pStyle w:val="OSGCueBody"/>'),style:'OSG Cue Body'};
  const joined=joinWords(a,b,io);
  assert.equal(joined.offset,6);
  const spans=readParagraph(joined.xml,'OSG Cue Body',io).spans;
  assert.deepEqual(spans.map(s=>s.text),['Plain ','bold']);
  assert.equal(fmtOfRpr(spans[1].rpr,io).bold,true);
  assert.equal(spans[0].rpr,'');
});

test('a new paragraph made from spans is well formed, in the style and list a spec names, and has the runs it should',()=>{
  const spec=nextParagraphSpec(para(run('x'),'<w:pStyle w:val="OSGNotes"/>'),'OSG Facilitator Notes',io);
  const xml=paragraphWithSpans(spec,[{text:'A & B ',rpr:''},{text:'<bold>',rpr:withFmt('',{bold:true,color:'#C00000'},io)}],io);
  const d=io.parse(xml);
  assert.equal(d.getElementsByTagName('parsererror').length,0);
  const info=readParagraph(xml,'OSG Facilitator Notes',io);
  assert.equal(info.text,'A & B <bold>');
  assert.deepEqual(fmtOfRpr(info.spans[1].rpr,io),{bold:true,color:'#C00000'});
  assert.equal(paragraphWithSpans({styleId:null,numPr:null},[],io).includes('<w:r>'),false,'no words, no run');
});

test('what cannot be formatted by the editor is refused as before, and the spans of such a paragraph are not offered',()=>{
  const withBox=para(`${run('A ')}<w:r><w:drawing/></w:r>`);
  const info=readParagraph(withBox,'Normal',io);
  assert.equal(info.editable,false);
  assert.deepEqual(info.spans,[]);
  assert.throws(()=>setSpans(withBox,'Normal',[{text:'x',rpr:''}],io),/more than words/);
});

test('words with characters XML cannot hold are cleaned, and too many are refused',()=>{
  const p=para(run('x'));
  const out=setSpans(p,'Normal',[{text:'a\u0000b\nc',rpr:''}],io);
  assert.equal(readParagraph(out,'Normal',io).text,'ab c');
  assert.throws(()=>setSpans(p,'Normal',[{text:'x'.repeat(100001),rpr:''}],io),/up to 100,000 characters/);
  assert.ok(sameSpans([{text:'a',rpr:''},{text:'b',rpr:''}],[{text:'ab',rpr:''}]));
});
