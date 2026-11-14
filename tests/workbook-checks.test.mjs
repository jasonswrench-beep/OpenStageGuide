// Stage 4: the structure check and the print check on a workbook, and the reviewed PDF route.
// These are not Word's Accessibility Checker or a screen-reader test, and a PDF is not proven tagged until a person looks at it.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readSource,buildWorkbook,pdfFileName} from '../src/guide-workbook.mjs';
import {workbookStructure,printCheck,checkPdfTags,WORKBOOK_HUMAN_CHECKS} from '../src/workbook-checks.mjs';
import {resolveWorkbookBrand} from '../src/workbook-presentation.mjs';
import {createWordAdapter,createDemoAdapter} from '../src/guide-word.mjs';
import {createProjectAdapter,NO_PDF} from '../src/guide-project-adapter.mjs';
import {started,pane,WORD_QUIRKS,click,submit,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const xml={DOMParserClass:DOMParser,SerializerClass:XMLSerializer};
const typeItem=(a,type,fields)=>{a.itemRadio(type).checked=true;a.itemRadio(type).onchange();for(const [id,v] of Object.entries(fields))a.n(id).value=v;};
async function guide(quirks=WORD_QUIRKS,extra={}){
  const a=await started('Reframing conflict',{quirks,...extra});
  await addModule(a,'Recognize the conflict');await addLesson(a,'Reframe the conversation','20');
  await addBlock(a,'objective','Identify two ways to reframe a disagreement.');
  await addBlock(a,'ask','What makes a disagreement hard to reframe?');
  await click(a,'guide-tab-assess');await click(a,'guide-sec-level-one');
  typeItem(a,'rating-statements',{'guide-item-stem':'Trainer.','guide-item-statements':'The trainer was prepared.'});
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  return a;
}
const built=async(a,options={})=>buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),options,xml);
const effective=(familyId='clearline',brand={})=>resolveWorkbookBrand(familyId,brand).effective;
const edit=(files,name,fn)=>files.map(f=>f.name===name?{...f,text:fn(f.text)}:f);

// ---- the structure check
test('a workbook made from a good guide has nothing to fix and nothing to look at',async()=>{
  const a=await guide(),b=await built(a);
  const r=workbookStructure(b.files,{DOMParserClass:DOMParser});
  assert.deepEqual([r.fix,r.look],[[],[]]);
  assert.ok(r.counts.pass>=4,'title, language, headings and the table header all pass');
});

test('the structure check finds a missing title, a table with no header row, and a missing language',async()=>{
  const b=await built(await guide());
  const check=files=>workbookStructure(files,{DOMParserClass:DOMParser});
  assert.deepEqual(check(edit(b.files,'docProps/core.xml',t=>t.replace(/<dc:title>[^<]*<\/dc:title>/,''))).fix.map(f=>f.rule),['TITLE']);
  const noHeader=check(edit(b.files,'word/document.xml',t=>t.replaceAll('<w:tblHeader/>','')));
  assert.deepEqual(noHeader.fix.map(f=>f.rule),['TABLE-HEADER']);
  assert.match(noHeader.fix[0].message,/Table 1 has no header row marked/);
  const noLanguage=check(edit(edit(b.files,'word/settings.xml',t=>t.replace(/<w:themeFontLang[^>]*\/>/,'')),'word/styles.xml',t=>t.replace(/<w:lang [^>]*\/>/,'')));
  assert.deepEqual(noLanguage.fix.map(f=>f.rule),['LANGUAGE']);
});

// ---- the print check
test('the print check says how the look prints, in colour and in black and white',()=>{
  const colour=printCheck(effective(),{lines:3}),gray=printCheck(effective('clearline',{printMode:'grayscale'}),{lines:3});
  assert.deepEqual(colour.filter(f=>f.level==='warn'),[]);
  assert.match(colour[0].message,/^Body text prints as a gray at \d+\.\d to 1 against the page in black and white\.$/);
  assert.ok(colour.some(f=>/The accent color prints as a gray at/.test(f.message)));
  assert.ok(colour.some(f=>f.level==='info'&&/color version.*choose Black and white/.test(f.message)));
  assert.ok(gray.some(f=>f.level==='info'&&/black-and-white version.*icons are still colored pictures/.test(f.message)));
  assert.ok(colour.some(f=>/Body text is 12 pt/.test(f.message))&&colour.some(f=>/Margins are 0\.8 in/.test(f.message))&&colour.some(f=>/Lines to write on are about 0\.35 in apart/.test(f.message)));
  assert.equal(printCheck(effective(),{lines:0}).some(f=>/Lines to write on/.test(f.message)),false,'no lines, nothing to say about them');
});

test('the print check warns about a gray that is too light, small text, tight margins and close lines',()=>{
  const base=effective();
  const warns=values=>printCheck({...base,...values},{lines:3}).filter(f=>f.level==='warn').map(f=>f.message);
  assert.match(warns({accent:'#C8C8C8'})[0],/The accent color prints as a gray that is too light in black and white \(1\.\d to 1; it needs 4\.5 to 1\)/);
  assert.match(warns({bodyPt:10})[0],/Body text is 10 pt, which is small for a printed page/);
  assert.match(warns({marginIn:0.3})[0],/Margins are 0\.3 in, which some printers will cut off/);
  assert.match(warns({responseLineIn:0.2})[0],/closer together than 0\.25 in/);
});

// ---- the screen
const finish=async a=>{await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');};
const items=id=>a=>a.n(id).children.map(li=>li.textContent);

test('the Workbook screen shows what was checked: nothing to fix, the print check, and what a person still checks',async()=>{
  const a=await guide();await finish(a);
  assert.equal(a.adapter.openedDocuments().length,1);
  assert.equal(a.n('guide-workbook-checks-box').hidden,false);
  assert.deepEqual(items('guide-workbook-fix')(a),['Nothing to fix.']);assert.deepEqual(items('guide-workbook-look-at')(a),['Nothing to look at.']);
  assert.ok(items('guide-workbook-print')(a).some(t=>/prints as a gray at/.test(t)));
  assert.deepEqual(items('guide-workbook-human')(a),WORKBOOK_HUMAN_CHECKS);
  assert.match(a.n('guide-workbook-ok').textContent,/Structure check: 0 to fix, 0 to look at\.$/);
});

test('a structure problem is listed but does not stop the workbook, which can be fixed in the copy',async()=>{
  const a=await guide();
  const original=a.adapter.sourcePackage.bind(a.adapter);
  a.adapter.sourcePackage=async()=>(await original()).replaceAll('<w:tblHeader/>','');
  await finish(a);
  assert.equal(a.adapter.openedDocuments().length,1,'the workbook was still made');
  assert.match(items('guide-workbook-fix')(a)[0],/Table 1 has no header row marked/);
  assert.match(a.n('guide-workbook-ok').textContent,/Structure check: 1 to fix, 0 to look at\.$/);
});

// ---- the reviewed PDF route
test('a tagged PDF is saved after a clean review, with what was checked in it',async()=>{
  const a=await guide();
  await click(a,'guide-open-review');
  assert.equal(a.n('guide-review').hidden,false);
  await click(a,'guide-review-pdf');
  assert.equal(a.adapter.downloads().length,1);
  const d=a.adapter.downloads()[0];
  assert.equal(d.fileName,'Reframing conflict.pdf');assert.equal(d.type,'application/pdf');
  assert.match(Buffer.from(d.base64,'base64').toString('latin1'),/^%PDF-/);
  assert.equal(a.n('guide-review-pdf-ok').hidden,false);assert.equal(a.n('guide-review-pdf-problem').hidden,true);
  assert.match(a.n('guide-review-pdf-ok').textContent,/The PDF was saved as "Reframing conflict\.pdf" in your downloads\. The PDF is tagged and has a language and a title\./);
});

test('with anything to fix no PDF is made, and the review is shown again',async()=>{
  const a=await guide();
  await a.adapter.setTitle('');
  await click(a,'guide-open-review');
  await click(a,'guide-review-pdf');
  assert.equal(a.adapter.downloads().length,0);
  assert.match(a.n('guide-review-pdf-problem').textContent,/^No PDF was made: 1 thing to fix first, listed above\./);
  assert.ok(a.n('guide-review-fix').children.some(li=>/no title/.test(li.textContent)),'the title is on the list');
  await a.adapter.setTitle('Reframing conflict');
  await click(a,'guide-review-pdf');
  assert.equal(a.adapter.downloads().length,1,'once it is fixed the PDF is made');
});

test('a PDF Word did not tag is still saved, with a plain warning of how to get a tagged one',async()=>{
  const a=await guide({...WORD_QUIRKS,untaggedPdf:true});
  await click(a,'guide-open-review');await click(a,'guide-review-pdf');
  assert.equal(a.adapter.downloads().length,1);
  assert.equal(a.n('guide-review-pdf-problem').hidden,false);
  assert.match(a.n('guide-review-pdf-problem').textContent,/It was saved, but: Word did not tag this PDF.*File > Export > Create PDF\/XPS > Options > Document structure tags for accessibility/);
});

test('if Word cannot make the PDF the screen says so and how to make one by hand',async()=>{
  const a=await guide({...WORD_QUIRKS,cannotExportPdf:true});
  await click(a,'guide-open-review');await click(a,'guide-review-pdf');
  assert.equal(a.adapter.downloads().length,0);
  assert.match(a.n('guide-review-pdf-problem').textContent,/^Word could not make the PDF\. .*You can make one yourself: File > Export > Create PDF\/XPS/);
});

test('in the program, with no Word behind it, Save as PDF says so and how to get one, and saves nothing (it once saved an empty file and called it tagged)',async()=>{
  const document=createProjectAdapter(createDemoAdapter({...xml,quirks:WORD_QUIRKS}),{read:async()=>null,write:async()=>{},soon:()=>0,cancel:()=>{}});
  await document.ready;
  await assert.rejects(document.exportPdf(),{message:NO_PDF},'the document itself refuses');
  const a=await guide(WORD_QUIRKS,{adapter:document});
  await click(a,'guide-open-review');await click(a,'guide-review-pdf');
  assert.equal(a.adapter.downloads().length,0,'no file is offered');
  assert.equal(a.n('guide-review-pdf-problem').textContent,NO_PDF,'the screen says what the program says, whole');
  assert.equal(a.n('guide-review-pdf-ok').hidden,true,'and does not say a PDF was saved');
  assert.doesNotMatch(NO_PDF,/Word could not/);
  assert.match(NO_PDF,/Save as Word file/);
  assert.match(NO_PDF,/Create PDF\/XPS/);
});

test('the tag check reads a PDF\'s catalogue for the marks of a tagged one, and says so when it cannot tell',()=>{
  const pdf=catalogue=>`%PDF-1.7\n1 0 obj ${catalogue} endobj\n`;
  assert.equal(checkPdfTags(pdf('<< /Type /Catalog /MarkInfo << /Marked true >> /StructTreeRoot 5 0 R /Lang (en-US) >> /Title (T)')).tagged,'yes');
  const noLang=checkPdfTags(pdf('<< /MarkInfo << /Marked true >> /StructTreeRoot 5 0 R >>'));
  assert.equal(noLang.tagged,'yes');assert.match(noLang.message,/no language is set.*no title is set/);
  assert.equal(checkPdfTags(pdf('<< /Type /Catalog >>')).tagged,'no');
  assert.equal(checkPdfTags(pdf('<< /Type /ObjStm >> stream ... endstream')).tagged,'unknown');
  assert.equal(checkPdfTags('not a pdf').tagged,'no');assert.match(checkPdfTags('not a pdf').message,/not a PDF file/);
  assert.equal(checkPdfTags(Uint8Array.from(Buffer.from(pdf('<< /MarkInfo << /Marked true >> /StructTreeRoot 1 0 R >>')))).tagged,'yes','bytes as well as text');
  assert.equal(pdfFileName('Sample: Handling conflict'),'Sample - Handling conflict.pdf');
});

test('the sample guide has a document title, unless the document already had one, so its review is clean and a PDF can be saved',async()=>{
  const a=pane(undefined,{quirks:WORD_QUIRKS});await a.p.refresh();
  await click(a,'guide-sample');
  assert.equal(await a.adapter.title(),'Sample guide: Handling workplace conflict');
  await click(a,'guide-open-review');
  assert.deepEqual(a.n('guide-review-fix').children.map(li=>li.textContent),['Nothing to fix.']);
  await click(a,'guide-review-pdf');
  assert.equal(a.adapter.downloads().length,1);
  assert.equal(a.adapter.downloads()[0].fileName,'Sample guide - Handling workplace conflict.pdf');
  const b=pane(undefined,{quirks:WORD_QUIRKS});await b.adapter.setTitle('My own title');await b.p.refresh();
  await click(b,'guide-sample');
  assert.equal(await b.adapter.title(),'My own title','a title the document already had is kept');
});

// ---- Word's own PDF, read a slice at a time (a stand-in for Office.js, not Word)
const fakeOffice=({slices=[[37,80,68,70],[45,49]],status='succeeded',closeFails=false}={})=>{
  const closed=[];
  const file={size:slices.flat().length,sliceCount:slices.length,
    getSliceAsync:(i,cb)=>cb({status:'succeeded',value:{index:i,data:slices[i]}}),
    closeAsync:cb=>{closed.push(true);cb({status:closeFails?'failed':'succeeded',error:{message:'close failed'}});}};
  return {office:{FileType:{Pdf:'pdf'},context:{document:{getFileAsync:(type,opts,cb)=>cb(status==='succeeded'?{status,value:file}:{status,error:{message:'Word refused'}})}}},closed};
};
test('the real adapter reads Word\'s PDF slice by slice, closes the file, and says plainly when Word refuses',async()=>{
  const good=fakeOffice();
  const adapter=createWordAdapter({Office:good.office,DOMParserClass:DOMParser});
  assert.equal(Buffer.from(await adapter.exportPdf(),'base64').toString('latin1'),'%PDF-1');
  assert.equal(good.closed.length,1);
  await assert.rejects(createWordAdapter({Office:fakeOffice({status:'failed'}).office,DOMParserClass:DOMParser}).exportPdf(),/Word refused/);
  await assert.rejects(createWordAdapter({Office:fakeOffice({slices:[]}).office,DOMParserClass:DOMParser}).exportPdf(),/this add-in cannot read/);
  await assert.rejects(createWordAdapter({Office:{context:{}},DOMParserClass:DOMParser}).exportPdf(),/Word is not ready yet/);
  assert.equal(Buffer.from(await createWordAdapter({Office:fakeOffice({closeFails:true}).office,DOMParserClass:DOMParser}).exportPdf(),'base64').length,6,'a failure to close does not lose the PDF');
});
