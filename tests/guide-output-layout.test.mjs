// Layout choices for the Word file (kit slice K6): each module on a new page, a running header, a footer with the page number (on the right or on the outside of two-sided
// pages), and the one-page summary on a sheet of its own. They are kept with the guide's settings and applied when the guide is saved as a Word file.
// The file is checked as a zip, as tests/guide-docx.test.mjs does; the Layout tool is driven through the real pane and the real taskpane.html ids in memory.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {started,submit,addModule,addLesson,addBlock,click} from './helpers/guide-harness.mjs';
import {docxFromPackage} from '../src/guide-docx.mjs';
import {cleanLayout,isDefaultLayout,describeLayout,layoutParts,layoutBody,bodyWithLayoutRefs,withEvenAndOdd,withPageBreakBefore,LAYOUT_DEFAULTS,SUMMARY_HEADING} from '../src/guide-output-layout.mjs';
import {readSettings,writeSettings} from '../src/guide-forms.mjs';
import {summaryRows,summaryBody} from '../src/guide-pages.mjs';
import {readGuide} from '../src/guide-model.mjs';
import {readSource,buildWorkbook} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
async function guide(){
  const a=await started('Layout guide');
  await addModule(a,'One');await addLesson(a,'First','10');
  await addModule(a,'Two');await addLesson(a,'Second','5');
  return a;
}
const wordFile=async(a,layout)=>JSZip.loadAsync(await docxFromPackage(await a.adapter.sourcePackage(),JSZip,io,{entries:a.adapter.exportState().paragraphs,title:'Layout guide',...(layout?{layout}:{})}));
const text=(zip,name)=>zip.file(name).async('string');
const wellFormed=xml=>{const d=io.parse(xml);assert.equal(d.getElementsByTagName('parsererror').length,0);return d;};
const paragraphsOf=xml=>xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g)||[];
const sectPrs=xml=>xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)||[];

test('a layout is read from anything: unknown words fall back to the defaults, and the default layout is not kept',()=>{
  assert.deepEqual(cleanLayout(null),LAYOUT_DEFAULTS);
  assert.deepEqual(cleanLayout({dividers:'yes',header:'bold',footer:'sideways',footerText:'  A  \n  line  '}),{dividers:false,header:'none',footer:'none',footerText:'A line'});
  assert.equal(cleanLayout({footerText:'x'.repeat(500)}).footerText.length,120);
  assert.equal(isDefaultLayout({}),true);assert.equal(isDefaultLayout({dividers:true}),false);
  assert.equal(describeLayout({}),'');
  assert.match(describeLayout({dividers:true,header:'module',footer:'mirrored',footerText:'Acme'}),/each module starts on a new page; the name of the module runs along the top; the page number is on the outside of each page.*footer words/);
  const kept=readSettings(writeSettings({scheme:'auto',layout:{dividers:true,footer:'plain',footerText:'Acme'}}));
  assert.deepEqual(kept.layout,{dividers:true,header:'none',footer:'plain',footerText:'Acme'});
  assert.equal('layout' in readSettings(writeSettings({scheme:'auto',layout:{}})),false,'the default is left out');
  assert.equal('layout' in readSettings('{"scheme":"auto","layout":{"header":"nonsense"}}'),false);
});

test('without a layout the Word file has no header, no footer and no page breaks of its own: it is made as it always was',async()=>{
  const a=await guide();
  const plain=await wordFile(a),same=await wordFile(a,{});
  assert.equal(await text(plain,'word/document.xml'),await text(same,'word/document.xml'));
  assert.ok(!Object.keys(plain.files).some(n=>/Layout\d/.test(n)),'no layout parts');
  const doc=await text(plain,'word/document.xml');
  assert.ok(!/headerReference|footerReference/.test(doc));
  assert.equal(paragraphsOf(doc).filter(p=>/<w:pageBreakBefore\/>/.test(p)).length,3,'only the three front headings the guide always starts on a page');
  assert.ok(!/evenAndOddHeaders/.test(await text(plain,'word/settings.xml')));
});

test('each module and assessment starts a page when asked, and nothing else does',async()=>{
  const a=await guide();
  const doc=await text(await wordFile(a,{dividers:true}),'word/document.xml');
  const paras=paragraphsOf(doc),modules=paras.filter(p=>/<w:pStyle w:val="OSGModule"\/>/.test(p));
  assert.equal(modules.length,2);
  assert.ok(modules.every(p=>/<w:pageBreakBefore\/>/.test(p)));
  assert.equal(paras.filter(p=>/<w:pageBreakBefore\/>/.test(p)).length,3+2,'the three front headings it always had, and the two modules');
  wellFormed(doc);
  assert.ok(/<w:pStyle [^>]*\/><w:pageBreakBefore\/>/.test(modules[0]),'in the place the schema keeps it, straight after the style');
  assert.equal(withPageBreakBefore(modules[0]),modules[0],'asking twice changes nothing');
  assert.match(withPageBreakBefore('<w:p><w:pPr><w:pStyle w:val="X"/><w:keepNext/><w:keepLines/><w:spacing w:after="0"/></w:pPr></w:p>'),/<w:keepNext\/><w:keepLines\/><w:pageBreakBefore\/><w:spacing/);
  assert.match(withPageBreakBefore('<w:p><w:r><w:t>x</w:t></w:r></w:p>'),/^<w:p><w:pPr><w:pageBreakBefore\/><\/w:pPr>/);
});

test('a plain header and footer are two parts, linked from every section, with the page number and the author\'s words',async()=>{
  const a=await guide();
  const zip=await wordFile(a,{header:'title',footer:'plain',footerText:'Acme & Sons'});
  for(const n of ['word/headerLayout1.xml','word/footerLayout1.xml'])assert.ok(zip.file(n),`${n} is in the file`);
  assert.ok(!zip.file('word/headerLayout2.xml')&&!zip.file('word/footerLayout2.xml'),'one of each: nothing is mirrored');
  for(const name of Object.keys(zip.files))if(/\.(xml|rels)$/.test(name))wellFormed(await text(zip,name));
  const header=await text(zip,'word/headerLayout1.xml'),footer=await text(zip,'word/footerLayout1.xml');
  assert.match(header,/Layout guide/);assert.match(header,/w:jc w:val="right"/);
  assert.match(footer,/Acme &amp; Sons/);assert.match(footer,/ PAGE /);assert.match(footer,/w:ptab w:relativeTo="margin" w:alignment="right"/);
  const types=await text(zip,'[Content_Types].xml'),rels=await text(zip,'word/_rels/document.xml.rels');
  assert.match(types,/PartName="\/word\/headerLayout1.xml" ContentType="application\/vnd.openxmlformats-officedocument.wordprocessingml.header\+xml"/);
  assert.match(types,/PartName="\/word\/footerLayout1.xml" ContentType="application\/vnd.openxmlformats-officedocument.wordprocessingml.footer\+xml"/);
  assert.match(rels,/Id="rIdOsgLayoutH1"[^>]*Target="headerLayout1.xml"/);assert.match(rels,/Id="rIdOsgLayoutF1"[^>]*Target="footerLayout1.xml"/);
  const doc=await text(zip,'word/document.xml'),sections=sectPrs(doc);
  assert.ok(sections.length>=1);
  assert.ok(sections.every(s=>/<w:headerReference w:type="default" r:id="rIdOsgLayoutH1"\/><w:footerReference w:type="default" r:id="rIdOsgLayoutF1"\/><w:pgSz/.test(s)),'every section points at them, before its page size');
  assert.ok(!/evenAndOddHeaders/.test(await text(zip,'word/settings.xml')));
});

test('mirrored footers put the number on the outside of each page, and turn on separate odd and even pages',async()=>{
  const a=await guide();
  const zip=await wordFile(a,{footer:'mirrored',footerText:'Acme'});
  const odd=await text(zip,'word/footerLayout1.xml'),even=await text(zip,'word/footerLayout2.xml');
  assert.ok(odd.indexOf('Acme')<odd.indexOf('PAGE'),'odd pages: the words, then the number at the right');
  assert.ok(even.indexOf('PAGE')<even.indexOf('Acme'),'even pages: the number at the left, then the words');
  assert.match(await text(zip,'word/settings.xml'),/<w:evenAndOddHeaders\/>/);
  const doc=await text(zip,'word/document.xml');
  assert.ok(sectPrs(doc).every(s=>/<w:footerReference w:type="default"[^>]*\/><w:footerReference w:type="even"/.test(s)));
  wellFormed(await text(zip,'word/settings.xml'));
  assert.equal(withEvenAndOdd('<w:settings><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="x"/></w:settings>'),'<w:settings><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:evenAndOddHeaders/><w:characterSpacingControl w:val="x"/></w:settings>','after the tab stop, before what follows it');
  assert.equal(withEvenAndOdd('<w:settings><w:evenAndOddHeaders/></w:settings>'),'<w:settings><w:evenAndOddHeaders/></w:settings>');
  assert.match(withEvenAndOdd('<w:settings><w:compat></w:compat></w:settings>'),/<w:evenAndOddHeaders\/><w:compat>/);
});

test('a header that names the module leaves the front pages without one, and the pages after the first module carry the name',async()=>{
  const a=await guide();
  const zip=await wordFile(a,{header:'module',footer:'mirrored'});
  const header=await text(zip,'word/headerLayout1.xml'),even=await text(zip,'word/headerLayout2.xml'),front=await text(zip,'word/headerLayout3.xml');
  assert.match(header,/STYLEREF "OSG Module"/);assert.match(header,/w:jc w:val="right"/);assert.match(even,/w:jc w:val="left"/);
  assert.ok(!/STYLEREF|<w:t/.test(front),'the front header is empty');
  const doc=await text(zip,'word/document.xml'),first=doc.search(/<w:pStyle w:val="OSGModule"\/>/);
  const before=sectPrs(doc.slice(0,first)),after=sectPrs(doc.slice(first));
  assert.equal(before.length,1,'the front pages end in a section break of their own');
  assert.match(before[0],/rIdOsgLayoutH3/);assert.doesNotMatch(before[0],/rIdOsgLayoutH1/);
  assert.ok(after.length>=1&&after.every(s=>/rIdOsgLayoutH1/.test(s)&&!/rIdOsgLayoutH3/.test(s)));
  wellFormed(doc);
});

test('a layout replaces the header and footer a package already had, and a template carries none of it',async()=>{
  const body='<w:p><w:pPr><w:pStyle w:val="OSGModule"/></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId9"/><w:footerReference w:type="default" r:id="rId8"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
  const spec=layoutParts({header:'title',footer:'plain'},{title:'T'});
  const out=bodyWithLayoutRefs(body,spec);
  assert.doesNotMatch(out,/rId9|rId8/);assert.match(out,/rIdOsgLayoutH1/);
  const onlyFooter=bodyWithLayoutRefs(body,layoutParts({footer:'plain'},{title:'T'}));
  assert.match(onlyFooter,/rId9/,'the header the guide had is kept when the layout does not touch the header');
  assert.doesNotMatch(onlyFooter,/rId8/);
  const a=await guide();
  const t=await JSZip.loadAsync(await docxFromPackage(await a.adapter.sourcePackage(),JSZip,io,{template:true,layout:{dividers:true,header:'title',footer:'mirrored'}}));
  assert.ok(!Object.keys(t.files).some(n=>/Layout\d/.test(n)));
  assert.ok(!/evenAndOddHeaders|pageBreakBefore/.test((await text(t,'word/document.xml'))+(await text(t,'word/settings.xml'))));
  assert.equal(layoutParts({},{title:'T'}).parts.length,0);
  assert.equal(bodyWithLayoutRefs('<w:body/>',layoutParts({},{title:'T'})),'<w:body/>');
});

test('the Layout tool keeps its choices with the guide, says what they do, and disables the footer words when there is no footer',async()=>{
  const a=await guide();
  a.p.useTool('layout');await a.p.refresh();
  assert.equal(a.n('guide-sec-layout').hidden,false,'its panel is showing');
  assert.deepEqual(a.n('guide-layout-header').children.map(o=>o.value),['none','title','module']);
  assert.deepEqual(a.n('guide-layout-footer').children.map(o=>o.value),['none','plain','mirrored']);
  assert.match(a.n('guide-layout-says').textContent,/No layout is chosen/);
  assert.equal(a.n('guide-layout-footertext').disabled,true);
  a.n('guide-layout-dividers').checked=true;await a.n('guide-layout-dividers').onchange();assert.equal(a.err(),'');
  a.type('guide-layout-footer','mirrored');await a.n('guide-layout-footer').onchange();
  a.type('guide-layout-header','module');await a.n('guide-layout-header').onchange();
  assert.equal(a.n('guide-layout-footertext').disabled,false);
  a.type('guide-layout-footertext','Acme');await a.n('guide-layout-footertext').onchange();
  assert.deepEqual(readSettings(await a.adapter.formSettings()).layout,{dividers:true,header:'module',footer:'mirrored',footerText:'Acme'});
  assert.match(a.n('guide-layout-says').textContent,/each module starts on a new page.*name of the module.*outside of each page/);
  assert.match(a.n('guide-status').textContent,/kept with the guide/);
  // it is read back from the guide, not from the page
  a.p.useTool('forms');await a.p.refresh();a.p.useTool('layout');await a.p.refresh();
  assert.equal(a.n('guide-layout-header').value,'module');assert.equal(a.n('guide-layout-dividers').checked,true);
  // and the Word file made from this guide uses it
  const zip=await wordFile(a,readSettings(await a.adapter.formSettings()).layout);
  assert.ok(zip.file('word/headerLayout1.xml')&&zip.file('word/footerLayout2.xml'));
  a.n('guide-layout-dividers').checked=false;a.type('guide-layout-header','none');a.type('guide-layout-footer','none');a.type('guide-layout-footertext','');
  for(const id of ['guide-layout-dividers','guide-layout-header','guide-layout-footer','guide-layout-footertext'])await a.n(id).onchange();
  assert.equal('layout' in readSettings(await a.adapter.formSettings()),false,'back to the default: nothing is kept');
});

test('the one-page summary lists each module with its minutes, its lessons and its aims, is built again over itself, and falls out of date',async()=>{
  const a=await guide();
  a.p.useTool('layout');await a.p.refresh();
  assert.match(a.n('guide-summary-note').textContent,/has not been built yet/);
  await click(a,'guide-build-summary');assert.equal(a.err(),'');
  const rows=a.adapter.snapshot().map(p=>p.text.trim());
  assert.equal(rows.filter(t=>t===SUMMARY_HEADING).length,1);
  assert.deepEqual(rows.slice(rows.indexOf('Module'),rows.indexOf('Module')+4),['Module','Minutes','Lessons','What learners will do']);
  assert.ok(rows.includes('10')&&rows.includes('5'),'each module\'s minutes');
  assert.ok(rows.includes('Total: 15 minutes in 2 modules.'));
  assert.match(a.n('guide-summary-note').textContent,/matches your guide/);
  assert.match(a.n('guide-status').textContent,/built with 2 modules/);
  await addModule(a,'Three');await addLesson(a,'Third','20');
  a.p.useTool('layout');await a.p.refresh();
  assert.match(a.n('guide-summary-note').textContent,/out of date/);
  await click(a,'guide-build-summary');assert.equal(a.err(),'');
  assert.equal(a.adapter.snapshot().filter(p=>p.text.trim()===SUMMARY_HEADING).length,1,'replaced, not added again');
  assert.match(a.n('guide-status').textContent,/The earlier one was replaced/);
  assert.ok(a.adapter.snapshot().some(p=>p.text.trim()==='Total: 35 minutes in 3 modules.'));
});

test('the Word file puts the one-page summary on a sheet of its own: a break before its heading and another after its table',async()=>{
  const a=await guide();
  a.p.useTool('layout');await a.p.refresh();await click(a,'guide-build-summary');assert.equal(a.err(),'');
  const doc=await text(await wordFile(a),'word/document.xml');
  const paras=paragraphsOf(doc),at=paras.findIndex(p=>p.includes('>One-page summary<'));
  assert.ok(at>=0);assert.match(paras[at],/<w:pageBreakBefore\/>/);
  const tableEnd=doc.indexOf('</w:tbl>',doc.indexOf('>One-page summary<'));
  const afterTable=/<w:p[ >][\s\S]*?<\/w:p>/.exec(doc.slice(tableEnd))[0];
  assert.match(afterTable,/pageBreakBefore/);
  assert.match(afterTable,/Total: 15 minutes in 2 modules/,'the line under the table is the first thing on the next page');
  wellFormed(doc);
  // a guide that has no summary page is not touched
  assert.equal(layoutBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>',{},{summaryHeading:SUMMARY_HEADING}),'<w:p><w:r><w:t>x</w:t></w:r></w:p>');
});

test('the pure summary says what to do with no modules, and counts what a long list of aims leaves out',()=>{
  assert.throws(()=>summaryBody([]),/Add a module first/);
  const body=summaryBody([{name:'Big',minutes:10,lessons:2,goals:'One; Two; Three; Four; and 3 more'}],{withHeading:true});
  assert.ok(body.includes(SUMMARY_HEADING)&&body.includes('and 3 more')&&body.includes('Total: 10 minutes in 1 module.'));
  const paras=[{style:'OSG Module',text:'Module 1: Big'},{style:'OSG Lesson',text:'Lesson 1.1: L'}];
  const rows=summaryRows(paras,readGuide(paras));
  assert.equal(rows.length,1);assert.equal(rows[0].lessons,1);assert.equal(rows[0].goals,'No objective yet.');
});

test('the one-page summary stays out of the participant workbook, and the leak check still passes',async()=>{
  const a=await guide();await addBlock(a,'ask','What makes it hard?');
  a.p.useTool('layout');await a.p.refresh();await click(a,'guide-build-summary');assert.equal(a.err(),'');
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},{DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const doc=built.files.find(f=>f.name==='word/document.xml').text;
  assert.ok(!doc.includes('One-page summary')&&!doc.includes('The whole guide on one page'));
  const leak=leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  assert.equal(leak.passed,true,JSON.stringify(leak.findings.filter(f=>f.level!=='pass')));
});
