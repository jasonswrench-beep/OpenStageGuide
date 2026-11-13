// The guide as a Word file, its styles as a Word template, and the styles of a Word file read back. The files are checked as zips: the parts a
// Word file must have, the content types, the relationships, and that what is inside is what the guide holds.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {startGuide,para,run} from '../src/guide-ooxml.mjs';
import {docxFromPackage,readDocxStyles,bodyWithContents,withContentsStyles,withSection,DOCX_LIMITS} from '../src/guide-docx.mjs';
import {listStyles} from '../src/guide-styles-edit.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function standardGuide(){
  const a=demo();
  await a.insert(c=>startGuide({title:'A guide',standard:true},c),{where:'start',select:'last'});
  return a;
}
const open=async bytes=>JSZip.loadAsync(bytes);
const text=async(zip,name)=>zip.file(name).async('string');
const wellFormed=xml=>{const d=io.parse(xml);assert.equal(d.getElementsByTagName('parsererror').length,0);return d;};

test('a Word file is a zip with the parts Word needs, each well formed and each named in the content types',async()=>{
  const a=await standardGuide();
  const bytes=await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:a.exportState().paragraphs});
  const zip=await open(bytes);
  for(const name of ['[Content_Types].xml','_rels/.rels','word/document.xml','word/_rels/document.xml.rels','word/styles.xml','word/settings.xml'])assert.ok(zip.file(name),`${name} is in the file`);
  for(const name of Object.keys(zip.files))if(/\.(xml|rels)$/.test(name))wellFormed(await text(zip,name));
  const types=await text(zip,'[Content_Types].xml');
  assert.match(types,/PartName="\/word\/document.xml" ContentType="application\/vnd.openxmlformats-officedocument.wordprocessingml.document.main\+xml"/);
  assert.match(types,/PartName="\/word\/styles.xml"/);
  assert.match(types,/PartName="\/word\/settings.xml"/);
  assert.match(await text(zip,'_rels/.rels'),/Target="word\/document.xml"/);
  assert.match(await text(zip,'word/_rels/document.xml.rels'),/relationships\/settings" Target="settings.xml"/,'the document names its settings');
  const doc=wellFormed(await text(zip,'word/document.xml'));
  const body=Array.from(doc.getElementsByTagNameNS(W,'body')[0].childNodes).filter(n=>n.nodeType===1);
  assert.equal(body.at(-1).localName,'sectPr','the body ends with the page setup Word needs');
  assert.equal(doc.getElementsByTagNameNS(W,'sectPr').length,1);
  assert.equal(Object.values(zip.files).filter(f=>f.dir).length,0,'no entries for folders, which some services that open Word files refuse');
});

test('the title of the guide is the title of the Word file, written so it cannot break the file',async()=>{
  const a=demo();
  const zip=await open(await docxFromPackage(await a.sourcePackage(),JSZip,io,{title:'Reframing <conflict> & "trust"'+String.fromCharCode(10)}));
  const core=wellFormed(await text(zip,'docProps/core.xml'));
  assert.equal(core.getElementsByTagName('dc:title')[0].textContent,'Reframing <conflict> & "trust"');
  assert.match(await text(zip,'_rels/.rels'),/core-properties" Target="docProps\/core.xml"/);
  assert.match(await text(zip,'[Content_Types].xml'),/PartName="\/docProps\/core.xml"/);
  const without=await open(await docxFromPackage(await a.sourcePackage(),JSZip,io));
  assert.equal(without.file('docProps/core.xml'),null,'no title, no property');
});

test('the styles in the file are the styles the guide holds, with the author\'s changes and styles of their own',async()=>{
  const a=await standardGuide();
  await a.setStyleEdits({overrides:{OSGCueBody:{font:{name:'Georgia',size:14,color:'#112233'}}},custom:[{id:'Callout',name:'Callout',type:'paragraph',basedOn:'OSGCueBody',spec:{shading:'#FFF4CE'}}]});
  const zip=await open(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:a.exportState().paragraphs}));
  const styles=await text(zip,'word/styles.xml');
  assert.match(styles,/w:styleId="Callout"/);
  assert.match(styles,/w:ascii="Georgia"/);
  const names=listStyles(styles,io,new Set(['Callout']));
  assert.equal(names.find(s=>s.id==='Callout').name,'Callout');
  assert.equal(names.find(s=>s.id==='OSGCueBody').effective.font.name,'Georgia');
});

test('the contents of the guide become one real contents field, so Word can fill in the page numbers',async()=>{
  const a=await standardGuide();
  a.seed([]);
  await a.insert(()=>[
    para('Heading1',run('Contents')),
    para('Heading1',run('Module one')),
    para('Heading2',run('Lesson A')),
    para('Heading1',run('Module two')),
  ].join(''),{where:'start',spare:false});
  // The table of contents the program shows: plain lines with the styles TOC 1 and TOC 2.
  const entries=[
    {xml:'',text:'Module one\t1',style:'TOC 1'},
    {xml:'',text:'Lesson A\t1',style:'TOC 2'},
    {xml:'',text:'Module two\t1',style:'TOC 1'},
    {xml:para('Normal',run('After the table')),text:'After the table',style:'Normal'},
  ];
  const body=bodyWithContents(entries);
  assert.equal((body.match(/fldCharType="begin"/g)||[]).length,1,'one field');
  assert.equal((body.match(/fldCharType="separate"/g)||[]).length,1);
  assert.equal((body.match(/fldCharType="end"/g)||[]).length,1);
  assert.match(body,/TOC \\o "1-2" \\h \\z \\u/);
  assert.match(body,/w:pStyle w:val="TOC1"/);
  assert.match(body,/w:pStyle w:val="TOC2"/);
  assert.match(body,/Module one</);
  assert.doesNotMatch(body,/Module one\t/,'the cached lines carry no page number the program made up');
  assert.ok(body.indexOf('Module two')<body.indexOf('After the table'),'the paragraph after the table stays after it');
  assert.equal(bodyWithContents([{xml:para('Normal',run('x')),text:'x',style:'Normal'}]),para('Normal',run('x')),'a guide with no contents lines is left as it is');
});

test('the contents levels the body uses are defined in the styles, once',()=>{
  const styles='<w:styles xmlns:w="'+W+'"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>';
  const once=withContentsStyles(styles,'<w:pStyle w:val="TOC1"/><w:pStyle w:val="TOC2"/>');
  assert.equal((once.match(/w:styleId="TOC1"/g)||[]).length,1);
  assert.equal((once.match(/w:styleId="TOC2"/g)||[]).length,1);
  assert.equal(withContentsStyles(once,'<w:pStyle w:val="TOC1"/>'),once,'a style already there is not added again');
  assert.equal(withContentsStyles(styles,'<w:p/>'),styles,'and none is added when the body has no contents');
});

test('page setup is added once',()=>{
  assert.match(withSection('<w:p/>'),/<w:sectPr>.*<\/w:sectPr>$/);
  assert.equal(withSection(withSection('<w:p/>')),withSection('<w:p/>'));
});

test('a Word template holds the styles and one empty paragraph, in the content type Word gives templates, and none of the guide\'s pictures',async()=>{
  const a=await standardGuide();
  const pkg=(await a.sourcePackage()).replace('</pkg:package>',`<pkg:part pkg:name="/word/media/image9.png" pkg:contentType="image/png" pkg:compression="store"><pkg:binaryData>${PNG}</pkg:binaryData></pkg:part></pkg:package>`);
  const zip=await open(await docxFromPackage(pkg,JSZip,io,{template:true}));
  assert.match(await text(zip,'[Content_Types].xml'),/wordprocessingml\.template\.main\+xml/);
  assert.doesNotMatch(await text(zip,'[Content_Types].xml'),/document\.main\+xml/);
  const doc=wellFormed(await text(zip,'word/document.xml'));
  const body=Array.from(doc.getElementsByTagNameNS(W,'body')[0].childNodes).filter(n=>n.nodeType===1).map(n=>n.localName);
  assert.deepEqual(body,['p','sectPr']);
  assert.equal(Object.keys(zip.files).filter(n=>n.startsWith('word/media/')).length,0,'no pictures');
  assert.doesNotMatch(await text(zip,'word/_rels/document.xml.rels'),/relationships\/image/,'and no relationship to one');
  assert.match(await text(zip,'word/styles.xml'),/w:styleId="OSGFacilitatorNotes"|w:styleId="OSGNotes"/);
});

test('the pictures of a guide go into the file with their types named',async()=>{
  const a=demo();
  const pkg=(await a.sourcePackage()).replace('</pkg:package>',`<pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png" pkg:compression="store"><pkg:binaryData>${PNG}</pkg:binaryData></pkg:part><pkg:part pkg:name="/word/media/image2.jpeg" pkg:contentType="image/jpeg" pkg:compression="store"><pkg:binaryData>${PNG}</pkg:binaryData></pkg:part></pkg:package>`);
  const zip=await open(await docxFromPackage(pkg,JSZip,io));
  assert.ok(zip.file('word/media/image1.png'));
  assert.equal((await zip.file('word/media/image1.png').async('base64')),PNG);
  const types=await text(zip,'[Content_Types].xml');
  assert.match(types,/Extension="png" ContentType="image\/png"/);
  assert.match(types,/Extension="jpeg" ContentType="image\/jpeg"/);
});

test('a part with an unsafe name is refused, and a package that is not a package is explained',async()=>{
  const a=demo();
  const evil=(await a.sourcePackage()).replace('</pkg:package>','<pkg:part pkg:name="/word/../../evil.xml" pkg:contentType="application/xml"><pkg:xmlData><x/></pkg:xmlData></pkg:part></pkg:package>');
  await assert.rejects(docxFromPackage(evil,JSZip,io),/name a Word file cannot hold/);
  await assert.rejects(docxFromPackage('<notapackage/>',JSZip,io),/not a Word package/);
});

test('a file made here can be read back: its styles and the same styles the guide had',async()=>{
  const a=await standardGuide();
  await a.setStyleEdits({overrides:{OSGCueBody:{font:{name:'Georgia',size:14}}},custom:[]});
  const bytes=await docxFromPackage(await a.sourcePackage(),JSZip,io,{template:true});
  const {stylesXml,themeFonts}=await readDocxStyles(bytes,JSZip,io);
  assert.deepEqual(themeFonts,{},'a file without a theme names no theme fonts');
  const read=listStyles(stylesXml,io,new Set());
  assert.equal(read.find(s=>s.id==='OSGCueBody').effective.font.name,'Georgia');
  assert.equal(read.length,a.listStyles().length,'every style the guide holds');
});

test('the theme\'s fonts are read, so a style that names "the body font" has one',async()=>{
  const zip=new JSZip();
  zip.file('word/styles.xml',`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
  zip.file('word/theme/theme1.xml','<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="x"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
  const {themeFonts}=await readDocxStyles(await zip.generateAsync({type:'uint8array'}),JSZip,io);
  assert.deepEqual(themeFonts,{major:'Calibri Light',minor:'Calibri'});
  const odd=new JSZip();
  odd.file('word/styles.xml',`<w:styles xmlns:w="${W}"/>`);
  odd.file('word/theme/theme1.xml','<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="x"><a:minorFont><a:latin typeface="&lt;script&gt;"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
  assert.deepEqual((await readDocxStyles(await odd.generateAsync({type:'uint8array'}),JSZip,io)).themeFonts,{},'a font name with markup in it is not taken');
});

test('a file that cannot be read is explained in words, and nothing in it is run',async()=>{
  await assert.rejects(readDocxStyles(new Uint8Array([1,2,3,4]),JSZip,io),/not a Word file/);
  await assert.rejects(readDocxStyles(new Uint8Array(0),JSZip,io),/Choose a Word file/);
  const none=new JSZip();none.file('hello.txt','hi');
  await assert.rejects(readDocxStyles(await none.generateAsync({type:'uint8array'}),JSZip,io),/no styles in it/);
  const wrong=new JSZip();wrong.file('word/styles.xml','<html/>');
  await assert.rejects(readDocxStyles(await wrong.generateAsync({type:'uint8array'}),JSZip,io),/not in a form/);
  const big=new JSZip();big.file('word/styles.xml',`<w:styles xmlns:w="${W}">${' '.repeat(DOCX_LIMITS.partBytes+10)}</w:styles>`);
  await assert.rejects(readDocxStyles(await big.generateAsync({type:'uint8array',compression:'DEFLATE'}),JSZip,io),/too large/);
});
