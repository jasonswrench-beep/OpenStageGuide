// Making a new guide from a Word document or from pasted lines (kit slice K7). The Word files are made by the real exporter and read back, and hostile files are
// made by hand: what could reach outside the file is gone, what the guide does not have is left out and said so, and nothing is opened, run or written over.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {docxFromPackage,readDocxParts} from '../src/guide-docx.mjs';
import {guideFromText,guideFromWord,styleFor,IMPORT_LIMITS} from '../src/guide-import.mjs';
import {readGuide} from '../src/guide-model.mjs';
import {parseProject,serializeProject} from '../src/guide-project.mjs';
import {STYLE_NAMES} from '../src/guide-styles.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

test('pasted lines become a guide: the first line is the title, Module and Lesson lines are numbered in order, a lesson may end in its minutes, the rest is plain',async()=>{
  const {guide,notes,modules,lessons}=guideFromText('Reframing conflict\n\nModule 7: Listening\nSome words about listening.\nLesson 9.9: Reflect back (10 min)\nLesson 1.2: Ask well - 5 minutes\nModule 2: Speaking\nLesson 2.1: Say it plainly');
  assert.equal(guide.title,'Reframing conflict');
  assert.deepEqual(guide.paragraphs.map(p=>[p.style,p.text]),[
    [STYLE_NAMES.guideTitle,'Reframing conflict'],
    [STYLE_NAMES.module,'Module 1: Listening'],
    ['Normal','Some words about listening.'],
    [STYLE_NAMES.lesson,'Lesson 1.1: Reflect back'],
    [STYLE_NAMES.time,'Time: 10 minutes. Cumulative time: 10 minutes.'],
    [STYLE_NAMES.lesson,'Lesson 1.2: Ask well'],
    [STYLE_NAMES.time,'Time: 5 minutes. Cumulative time: 15 minutes.'],
    [STYLE_NAMES.module,'Module 2: Speaking'],
    [STYLE_NAMES.lesson,'Lesson 2.1: Say it plainly']
  ]);
  assert.equal(modules,2);assert.equal(lessons,3);
  assert.match(notes.join(' '),/2 modules and 3 lessons/);assert.match(notes.join(' '),/1 other line was added as a plain paragraph/);assert.match(notes.join(' '),/Update guide/);
  // it opens as a guide: the program reads it back as modules and lessons with their minutes
  const a=demo();a.importState(guide);
  const read=readGuide(await a.paragraphs());
  assert.equal(read.modules.filter(m=>m.index>=0).length,2);
  assert.deepEqual(read.modules.flatMap(m=>m.lessons.map(l=>l.text)),['Lesson 1.1: Reflect back','Lesson 1.2: Ask well','Lesson 2.1: Say it plainly']);
  // and it is a project file the program can save and open again
  assert.equal(parseProject(serializeProject(guide)).guide.paragraphs.length,9);
});

test('pasted text that starts with a module has a title of its own, and a title given is used and the first line kept',()=>{
  assert.equal(guideFromText('Module 1: One\nLesson 1.1: A').guide.title,'Pasted guide');
  const g=guideFromText('A first line\nModule 1: One',{title:'Mine'}).guide;
  assert.equal(g.title,'Mine');
  assert.deepEqual(g.paragraphs.slice(0,2).map(p=>p.text),['Mine','A first line'],'the first line is not lost');
});

test('pasted text that cannot be made into a guide says what to do',()=>{
  assert.throws(()=>guideFromText('   \n  '),/Paste some text first/);
  assert.throws(()=>guideFromText('Title\nLesson 1.1: Early'),/Line 2 is a lesson before any module/);
  assert.throws(()=>guideFromText('Module 1: A\nLesson 1.1:  (5 min)'),/Line 2 gives a lesson no title/);
  assert.throws(()=>guideFromText(`Title\n${'x'.repeat(IMPORT_LIMITS.pasteLine+1)}`),/Line 2 is longer than/);
  assert.throws(()=>guideFromText(Array.from({length:IMPORT_LIMITS.pasteLines+1},(_,i)=>`Line ${i}`).join('\n')),/more than 5000 lines/);
});

test('a Word file the program made is read back into a guide with the same modules, lessons, blocks and table, and its pictures',async()=>{
  const a=await started('Round trip');
  await addModule(a,'One');await addLesson(a,'First','10');await addBlock(a,'ask','What makes it hard?');
  await addModule(a,'Two');await addLesson(a,'Second','5');
  const bytes=await docxFromPackage(await a.adapter.sourcePackage(),JSZip,io,{entries:a.adapter.exportState().paragraphs,title:'Round trip'});
  const parts=await readDocxParts(bytes,JSZip);
  assert.ok(parts.document&&parts.styles&&parts.rels);
  assert.ok(Object.keys(parts.media).length>0&&Object.values(parts.media).every(m=>m.mime==='image/png'),'the cue icons are pictures in the file');
  const {guide,notes,modules,lessons}=guideFromWord(parts,io,{fallbackTitle:'x'});
  assert.equal(guide.title,'Round trip');
  assert.equal(modules,2);assert.equal(lessons,2);
  const before=a.adapter.exportState().paragraphs.filter(p=>p.style!=='TOC 1'&&p.style!=='TOC 2'&&(p.text.trim()||p.xml));
  const after=guide.paragraphs.filter(p=>p.style!=='TOC 1'&&p.style!=='TOC 2'&&(p.text.trim()||p.xml));
  const own=p=>/^OSG /.test(p.style);
  assert.deepEqual(after.filter(own).map(p=>[p.style,p.text.trim()]),before.filter(own).map(p=>[p.style,p.text.trim()]),'every paragraph in the guide\'s own styles, in order, with its words');
  assert.ok(Object.keys(guide.media).length>0&&Object.keys(guide.media).every(id=>/^rIdPicImp\d+$/.test(id)),'the pictures came with it, under the ids a project names them by');
  assert.ok(guide.paragraphs.some(p=>p.xml.includes('r:embed="rIdPicImp1"')||/r:embed="rIdPicImp\d+"/.test(p.xml)));
  assert.match(notes.join(' '),/2 modules and 2 lessons/);
  assert.match(notes.join(' '),/Headers, footers, the look/);
  // it opens as a guide, with the same modules and lessons and the same minutes
  const b=demo();b.importState(guide);
  const was=readGuide(await a.adapter.paragraphs()),now=readGuide(await b.paragraphs());
  assert.deepEqual(now.modules.map(m=>[m.text,m.lessons.map(l=>l.text)]),was.modules.map(m=>[m.text,m.lessons.map(l=>l.text)]));
});

test('a Word file that is not a guide still comes in as plain paragraphs, with the styles the guide does not have named and left out',async()=>{
  const document=`<w:document xmlns:w="${W}"><w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Chapter</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Quote9"/></w:pPr><w:r><w:t>A styled line</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>A bullet</w:t></w:r></w:p>
    <w:sectPr/></w:body></w:document>`;
  const styles=`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Quote9"><w:name w:val="My Fancy Quote"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style></w:styles>`;
  const {guide,notes,modules}=guideFromWord({document,styles},io,{fallbackTitle:'A plain file'});
  assert.equal(modules,0);
  assert.deepEqual(guide.paragraphs.map(p=>[p.style,p.text]),[[STYLE_NAMES.guideTitle,'A plain file'],['heading 1','Chapter'],['Normal','A styled line'],['Normal','A bullet']],'a title is made when the file has none');
  assert.ok(!/Quote9/.test(guide.paragraphs[2].xml)&&!/numPr/.test(guide.paragraphs[3].xml));
  const said=notes.join(' ');
  assert.match(said,/no modules or lessons/);assert.match(said,/"My Fancy Quote"/);assert.match(said,/1 paragraph lost their bullets or numbers/);
});

test('a hostile Word file is read without harm: fields that reach out, links, tracked deletions, comments, notes and pictures that are not pictures are gone',async()=>{
  const document=`<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>
    <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "\\\\\\\\evil\\\\share\\\\x.png" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>after the field</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    <w:p><w:ins w:id="1" w:author="a"><w:r><w:t>kept </w:t></w:r></w:ins><w:del w:id="2" w:author="a"><w:r><w:delText>deleted</w:delText></w:r></w:del><w:commentRangeStart w:id="5"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
    <w:p><w:bookmarkStart w:id="9" w:name="_Toc123"/><w:bookmarkStart w:id="10" w:name="osg_form_f1"/><w:r><w:t>marked</w:t></w:r><w:bookmarkEnd w:id="10"/><w:bookmarkEnd w:id="9"/></w:p>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="10" cy="10"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdGood"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="10" cy="10"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdFake"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t>text beside a bad picture</w:t></w:r></w:p>
    <w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr></w:pPr></w:p>
    </w:body></w:document>`;
  const rels=`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdGood" Type="x" Target="media/good.png"/><Relationship Id="rIdFake" Type="x" Target="media/fake.png"/><Relationship Id="rIdExt" Type="x" Target="http://evil/x.png" TargetMode="External"/></Relationships>`;
  const media={'word/media/good.png':{mime:'image/png',base64:PNG},'word/media/fake.png':{mime:'image/png',base64:Buffer.from('this is not a png at all').toString('base64')}};
  const {guide,notes}=guideFromWord({document,rels,media},io,{fallbackTitle:'Hostile'});
  const a=demo();a.importState(guide);                       // the guide's own cleaning runs here too
  const xml=a.exportState().paragraphs.map(p=>p.xml).join('\n');
  assert.ok(xml.length>0&&xml.includes('after the field'),'the paragraphs are there to be checked');
  assert.ok(!/INCLUDEPICTURE|evil|delText|commentReference|footnoteReference|_Toc123|a:blip r:embed="rIdFake"|sectPr/.test(xml),'none of it is left');
  assert.ok(xml.includes('osg_form_f1'),'the guide\'s own form bookmark is kept');
  const texts=(await a.paragraphs()).map(p=>p.text.trim()).filter(Boolean);
  assert.ok(texts.includes('after the field')&&texts.includes('kept commented')&&texts.includes('text beside a bad picture'));
  assert.ok(!texts.some(t=>/deleted/.test(t)));
  assert.deepEqual(Object.keys(guide.media),['rIdPicImp1'],'only the real picture came');
  assert.equal(guide.paragraphs.filter(p=>/<w:drawing/.test(p.xml)).length,1);
  assert.match(notes.join(' '),/Tracked changes were accepted/);assert.match(notes.join(' '),/Comments and notes/);assert.match(notes.join(' '),/could not be kept/);
  assert.ok(!guide.paragraphs.some(p=>!p.text.trim()&&!p.xml.includes('drawing')),'the paragraph that only ended a section is gone');
});

test('a file that is not a Word document, or that inflates too large, is refused with its reason',async()=>{
  await assert.rejects(readDocxParts(new Uint8Array([1,2,3,4]),JSZip),/not a Word file/);
  await assert.rejects(readDocxParts(new Uint8Array(),JSZip),/Choose a Word file under/);
  const noDoc=new JSZip();noDoc.file('hello.txt','x');
  await assert.rejects(readDocxParts(await noDoc.generateAsync({type:'uint8array'}),JSZip),/has no document/);
  const bomb=new JSZip();bomb.file('word/document.xml',' '.repeat(7*1024*1024));
  await assert.rejects(readDocxParts(await bomb.generateAsync({type:'uint8array',compression:'DEFLATE'}),JSZip),/document in that file is too large/);
  assert.throws(()=>guideFromWord({document:'<w:document xmlns:w="urn:x"><w:body/></w:document>'},io),/not a Word document/);
  assert.throws(()=>guideFromWord({document:`<w:document xmlns:w="${W}"><w:body><w:p/></w:body></w:document>`},io),/no text in it/);
  assert.throws(()=>guideFromWord({},io),/no document in it/);
});

test('the styles of a file are read by name: the guide\'s own, the built-in ones it uses, and nothing else',()=>{
  assert.deepEqual(styleFor('OSG Module'),{name:'OSG Module',id:'OSGModule'});
  assert.deepEqual(styleFor('Heading 2'),{name:'heading 2',id:'Heading2'});
  assert.equal(styleFor('toc 1').toc,1);
  assert.deepEqual(styleFor(''),{name:'Normal',id:null});
  assert.equal(styleFor('Body Text'),null);
});
