// The learner workbook (Stage 3, slice 3a): what is kept, the file that is made, the leak check, the screen.
// Everything here runs against stand-ins for Word. The file opening in Word, and Word's own checker, are not proven here.
import test from 'node:test';
import assert from 'node:assert/strict';
import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs/promises';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {fileURLToPath} from 'node:url';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {classify,summarize,readSource,buildWorkbook,zipWorkbook,workbookFileName,lookTheme,leakInputsOf} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {resolveWorkbookBrand} from '../src/workbook-presentation.mjs';
import {readStyles,textOf} from '../src/guide-check.mjs';
import {stylesPart,styleList} from '../src/guide-styles.mjs';
import {para,run} from '../src/guide-ooxml.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {readDocxFiles,checkWorkbookFile,facilitatorTextsFromDocx} from '../scripts/workbook-leak-check.mjs';
import {packageToDocx} from './helpers/guide-package.mjs';
import {started,WORD_QUIRKS,click,submit,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml={DOMParserClass:DOMParser,SerializerClass:XMLSerializer};
const run_=promisify(execFile);

// ---- who is a paragraph for
const P=(style,text='x')=>({style,text});
const PARAS=[
  P('OSG Guide Title','Guide'),P('Normal','Facilitator guide'),P('heading 1','Contents'),P('TOC 1','Module 1: A'),P('heading 1','Getting Started'),P('Normal','Say what this guide is for.'),
  P('OSG Module','Module 1: A'),P('OSG Lesson','Lesson 1.1: B'),P('OSG Time','Time: 10 minutes. Cumulative time: 0 minutes.'),
  P('OSG Cue',' Ask Why is it hard? (5 min)'),P('OSG Cue Body','a second learner line'),P('OSG Facilitator Notes','Facilitator notes: listen for blame'),
  P('OSG Cue Facilitator',' Say Welcome everyone'),P('OSG Cue Body','a second facilitator line'),P('Normal','my own remark about the room'),
  P('OSG Assessment','Level Two Assessment'),P('Normal','Answer each question.'),P('OSG Question','Question 1. Which?'),P('OSG Option','one'),P('OSG Answer Line',''),
  P('OSG Answer Key','Answer key (multiple choice): a.'),P('Normal',''),
  P('OSG Module','Module 2: C'),P('OSG Question','Question 9. stray')
];
const reasons=()=>classify(PARAS).decisions.map(d=>d.reason);

test('blocks for everyone are kept, and so are their extra lines; facilitator-only blocks take their extra lines with them',()=>{
  const r=reasons();
  assert.equal(r[9],'block');assert.equal(r[10],'block','a learner cue\'s extra line stays');
  assert.equal(r[12],'facilitator-block');assert.equal(r[13],'facilitator-block','a facilitator cue\'s extra line goes with it');
  assert.equal(r[11],'notes');assert.equal(r[8],'time');
});

test('the front pages are left out, the headings are kept, and the title is read, not copied',()=>{
  const r=reasons(),c=classify(PARAS);
  assert.equal(c.title,'Guide');assert.equal(r[0],'title');
  assert.deepEqual(r.slice(1,6),['front','front','front','front','front']);
  assert.deepEqual([r[6],r[7],r[15],r[22]],['heading','heading','heading','heading']);
});

test('an assessment section is kept except its keys, and anything of the author\'s own outside one is left out',()=>{
  const r=reasons();
  assert.deepEqual(r.slice(15,22),['heading','section','item','item','item','key','blank']);
  assert.equal(r[14],'own','a typed paragraph in a lesson is left out');
  assert.equal(r[23],'own','an item outside any section is left out too');
});

test('a heading inside a section ends it, so what follows is not part of the section',()=>{
  const c=classify([P('OSG Guide Title','G'),P('OSG Module','Module 1: A'),P('OSG Assessment','Level One Assessment'),P('Normal','kept note'),P('heading 1','A heading'),P('Normal','after the heading')]);
  assert.deepEqual(c.decisions.map(d=>d.reason),['title','heading','heading','section','own','own']);
});

test('the pane\'s summary counts what is for learners and what is left out',()=>{
  const s=summarize(PARAS);
  assert.deepEqual([s.title,s.modules,s.lessons,s.blocks,s.items],['Guide',2,1,1,1]);
  assert.deepEqual(s.left,{facilitatorBlocks:1,facilitatorForms:0,notes:1,keys:1,timeBlocks:1,front:5,overview:0,own:2});
  assert.deepEqual(s.leftOut.map(o=>o.text),['my own remark about the room','Question 9. stray']);
  assert.equal(summarize([P('OSG Guide Title','G'),P('OSG Cue Facilitator',' Say hello')]).kept,0);
});

// ---- the whole thing, through the pane
const typeItem=(a,type,fields)=>{a.itemRadio(type).checked=true;a.itemRadio(type).onchange();for(const [id,v] of Object.entries(fields))a.n(id).value=v;};
async function richGuide({own=true}={}){
  const a=await started('Reframing conflict',{quirks:WORD_QUIRKS});
  await addModule(a,'Recognize the conflict');await addLesson(a,'Reframe the conversation','20');
  await addBlock(a,'objective','Identify two ways to reframe a disagreement.',{minutes:'2'});
  await addBlock(a,'say','Welcome everyone to the session.\nExplain why this matters today.',{notes:'Pause after the example.',minutes:'5'});
  await addBlock(a,'ask','What makes a disagreement hard to reframe?',{notes:'Listen for blame and tone.',minutes:'5'});
  await addBlock(a,'activity','In pairs, write two neutral questions.',{notes:'Allow eight minutes.'});
  await addBlock(a,'key-point','A neutral question describes the situation without blame.');
  // The module already has its own Transition block, which is for the facilitator.
  await click(a,'guide-tab-assess');
  await click(a,'guide-sec-level-two');
  typeItem(a,'multiple-choice',{'guide-item-stem':'Which step comes first?','guide-item-opt-1':'Name the feeling','guide-item-opt-2':'Describe the facts','guide-item-why':'Facts keep the first words neutral.','guide-item-points':'2'});
  a.n('guide-item-correct-2').checked=true;await submit(a,'guide-item-form');
  typeItem(a,'matching',{'guide-item-stem':'Match each term with its meaning.','guide-item-pairs':'Active listening | Fully attending to the speaker\nProbing | Asking follow-up questions','guide-item-extras':'Interrupting to advise'});
  await submit(a,'guide-item-form');
  await click(a,'guide-sec-level-one');
  typeItem(a,'rating-statements',{'guide-item-stem':'Trainer. Mark one box for each statement.','guide-item-statements':'The trainer was prepared.\nThe trainer was clear.'});
  await submit(a,'guide-item-form');
  assert.equal(a.err(),'');
  if(own){
    const at=a.rows().findIndex(r=>r.startsWith('OSG Lesson'));
    await a.adapter.insert(()=>para(null,run('A reminder I typed myself about the room.')),{where:'after',after:at,select:'last'});
  }
  return a;
}
const makeWorkbook=async a=>{await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');};
async function filesOf(a,which=0){
  const opened=a.adapter.openedDocuments()[which];
  assert.ok(opened,'a workbook was opened');
  return readDocxFiles(Buffer.from(opened.base64,'base64'));
}
const text=(files,name)=>files.find(f=>f.name===name)?.text;
const parse=t=>new DOMParser().parseFromString(t,'application/xml');
const outParas=files=>{
  const styles=readStyles(parse(text(files,'word/styles.xml')));
  return Array.from(parse(text(files,'word/document.xml')).getElementsByTagNameNS(W,'p')).map(p=>{
    const id=Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val')||'Normal';
    return {style:styles.get(id)?.name||id,text:textOf(p).trim()};
  });
};
const rowsOf=files=>outParas(files).map(p=>`${p.style}: ${p.text}`);

test('a workbook is made from a whole guide and opened in a new window, named for the guide',async()=>{
  const a=await richGuide();
  await makeWorkbook(a);
  assert.equal(a.err(),'');
  assert.equal(a.adapter.openedDocuments().length,1);assert.equal(a.adapter.downloads().length,0);
  assert.equal(a.adapter.openedDocuments()[0].fileName,'Reframing conflict - Participant workbook.docx');
  assert.match(a.n('guide-workbook-ok').textContent,/The workbook is open in a new Word window\. Save it as "Reframing conflict - Participant workbook\.docx"/);
  assert.match(a.n('guide-workbook-ok').textContent,/Leak check passed: nothing meant for the facilitator is in the file\. Left out for learners: 2 facilitator-only blocks, 3 sets of facilitator notes, 2 answer keys, 1 time block, the contents and Getting Started pages and 1 paragraph of your own text\. Structure check: 0 to fix, 0 to look at\.$/);
  assert.equal(a.n('guide-workbook-problem').hidden,true);
});

test('the workbook holds the learner content and a fresh title page, contents and learning objectives, and nothing for the facilitator',async()=>{
  const a=await richGuide();await makeWorkbook(a);
  const rows=rowsOf(await filesOf(a));
  const has=r=>assert.ok(rows.includes(r),`${r} is in the workbook`);
  const lacks=s=>assert.ok(!rows.some(r=>r.includes(s)),`${s} is not in the workbook`);
  assert.deepEqual(rows.slice(0,3),['OSG Guide Title: Reframing conflict','OSG Workbook Subtitle: Participant workbook','OSG Answer Line: Name:']);
  assert.equal(rows[3],'heading 1: Contents');
  // the contents table's entries, from the headings that are kept
  has('toc 1: Module 1: Recognize the conflict');has('toc 2: Lesson 1.1: Reframe the conversation');has('toc 1: Level Two Assessment');has('toc 1: Level One Assessment');
  // the learning objectives, collected from the Objective block
  has('heading 1: Learning objectives');has('OSG Item Note: Lesson 1.1: Reframe the conversation');has('OSG Workbook Bullet: Identify two ways to reframe a disagreement.');
  // the guide's learner blocks, with their minutes taken off
  has('OSG Module: Module 1: Recognize the conflict');has('OSG Lesson: Lesson 1.1: Reframe the conversation');
  assert.ok(rows.includes('OSG Cue: Ask  What makes a disagreement hard to reframe?')||rows.some(r=>/^OSG Cue: .*What makes a disagreement hard to reframe\?$/.test(r)));
  assert.ok(rows.some(r=>/^OSG Cue: .*In pairs, write two neutral questions\.$/.test(r)));
  assert.ok(rows.some(r=>/^OSG Cue: .*A neutral question describes the situation without blame\.$/.test(r)));
  assert.ok(!rows.some(r=>/\(\d+ min\)/.test(r)),'a block\'s minutes are the facilitator\'s');
  // the assessments
  has('OSG Assessment: Level Two Assessment');has('OSG Question: Question 1. Which step comes first?');has('OSG Table Head: Your match');has('OSG Table Cell: Active listening');
  assert.equal(rows.filter(r=>r==='OSG Rating Box: ☐').length,10);
  // what must not be there
  for(const s of ['Welcome everyone','Explain why this matters','Pause after the example','Listen for blame','Allow eight minutes','Say how this module leads','Answer key','Facilitator notes','Time:','Why:','Points:','Getting Started','Update this table','A reminder I typed myself'])lacks(s);
  assert.ok(!rows.some(r=>/^OSG (Cue Facilitator|Facilitator Notes|Answer Key|Time):/.test(r)));
});

test('the file is a small, clean package: only the expected parts, every picture and list it uses is there',async()=>{
  const a=await richGuide();await makeWorkbook(a);
  const files=await filesOf(a),names=files.map(f=>f.name);
  for(const n of ['[Content_Types].xml','_rels/.rels','word/document.xml','word/_rels/document.xml.rels','word/styles.xml','word/settings.xml','word/numbering.xml','docProps/core.xml'])assert.ok(names.includes(n),n);
  assert.ok(names.every(n=>/^word\/media\/[^/]+\.png$/.test(n)||!n.startsWith('word/media/')));
  const doc=parse(text(files,'word/document.xml')),rels=parse(text(files,'word/_rels/document.xml.rels'));
  const relIds=new Set(Array.from(rels.getElementsByTagName('Relationship')).map(r=>r.getAttribute('Id')));
  const embeds=Array.from(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main','blip')).map(b=>b.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','embed'));
  assert.ok(embeds.length>0,'the icons came with the blocks');
  assert.ok(embeds.every(id=>relIds.has(id)),'every picture has its relationship');
  for(const r of Array.from(rels.getElementsByTagName('Relationship')))if(/image$/.test(r.getAttribute('Type')))assert.ok(names.includes(`word/${r.getAttribute('Target')}`),`${r.getAttribute('Target')} is in the file`);
  const numbering=parse(text(files,'word/numbering.xml')),numIds=new Set(Array.from(numbering.getElementsByTagNameNS(W,'num')).map(n=>n.getAttributeNS(W,'numId')));
  for(const n of Array.from(doc.getElementsByTagNameNS(W,'numId')))assert.ok(numIds.has(n.getAttributeNS(W,'val')),'every list the content uses is defined');
  assert.ok(text(files,'word/numbering.xml').includes('lowerLetter')&&text(files,'word/numbering.xml').includes('w:val="bullet"'));
  const styles=text(files,'word/styles.xml');
  for(const gone of ['OSGCueFacilitator','OSGNotes','OSGAnswerKey','OSGTime'])assert.ok(!styles.includes(`w:styleId="${gone}"`),`${gone} is not defined`);
  assert.ok(!styles.includes('w:val="OSGTime"'),'no style points at the time block that was left out');
});

test('the lessons start on a new page when a module or an assessment begins, and the page is Letter in the Clearline look',async()=>{
  const a=await richGuide();await makeWorkbook(a);
  const files=await filesOf(a),doc=parse(text(files,'word/document.xml'));
  const withBreak=Array.from(doc.getElementsByTagNameNS(W,'p')).filter(p=>p.getElementsByTagNameNS(W,'pageBreakBefore').length).map(p=>textOf(p).trim());
  assert.ok(withBreak.includes('Module 1: Recognize the conflict')&&withBreak.includes('Level Two Assessment')&&withBreak.includes('Contents')&&withBreak.includes('Learning objectives'));
  assert.ok(!withBreak.includes('Lesson 1.1: Reframe the conversation'));
  assert.match(text(files,'word/document.xml'),/<w:pgSz w:w="12240" w:h="15840"\/>/);
  const styles=text(files,'word/styles.xml');
  assert.match(styles,/w:ascii="Aptos"/);assert.match(styles,/w:ascii="Aptos Display"/,'headings use the look\'s heading font');
  assert.match(styles,/<w:color w:val="164B6B"\/>/,'the look\'s accent');
  assert.match(text(files,'word/settings.xml'),/<w:themeFontLang w:val="en-US"\/>/);
});

test('the workbook passes the structure check, with its title and language, headings and table headers',async()=>{
  const a=await richGuide();await makeWorkbook(a);
  const opened=a.adapter.openedDocuments()[0],result=await checkDocx(Buffer.from(opened.base64,'base64'));
  assert.deepEqual(result.findings.filter(f=>f.level==='fail'),[],'no failures');
  assert.equal(result.findings.filter(f=>f.rule==='TABLE-HEADER'&&f.level==='pass').length,2);
  assert.deepEqual(result.findings.filter(f=>['TABLE-BLANK','TABLE-MERGED','EMPTY-PARAGRAPHS','TYPED-LIST'].includes(f.rule)),[]);
});

test('the leak check passes on what was made, and so does the command-line check against the guide',async()=>{
  const a=await richGuide();await makeWorkbook(a);
  const opened=a.adapter.openedDocuments()[0],bytes=Buffer.from(opened.base64,'base64');
  const guide=await packageToDocx(await a.adapter.sourcePackage(),'Reframing conflict');
  const texts=await facilitatorTextsFromDocx(guide);
  assert.ok(texts.includes('Facilitator notes: Pause after the example.'));
  assert.ok(texts.some(t=>t.includes('Welcome everyone to the session.')));
  assert.ok(texts.some(t=>t.startsWith('Answer key (multiple choice)')));
  const result=await checkWorkbookFile(bytes,guide);
  assert.equal(result.passed,true,result.summary);
  assert.ok(result.searchedSource);
});

test('the command-line check runs on saved files and gives an exit code',async()=>{
  const a=await richGuide();await makeWorkbook(a);
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'osg-leak-'));
  try{
    const book=path.join(dir,'workbook.docx'),guide=path.join(dir,'guide.docx'),script=fileURLToPath(new URL('../scripts/workbook-leak-check.mjs',import.meta.url));
    await fs.writeFile(book,Buffer.from(a.adapter.openedDocuments()[0].base64,'base64'));
    await fs.writeFile(guide,await packageToDocx(await a.adapter.sourcePackage(),'Reframing conflict'));
    const ok=await run_('node',[script,book,'--source',guide]);
    assert.match(ok.stdout,/Leak check passed/);assert.match(ok.stdout,/PASS FACILITATOR-TEXT/);
    // the guide itself is not a workbook: the same check, run on it, fails
    await assert.rejects(run_('node',[script,guide,'--source',guide]),e=>e.code===1&&/FAIL/.test(e.stdout));
  }finally{await fs.rm(dir,{recursive:true,force:true});}
});

// ---- the leak check refuses a file that has anything in it that should not be there
async function madeFiles(){const a=await richGuide();await makeWorkbook(a);return {files:await filesOf(a),texts:leakInputsOf(await sourceOf(a))};}
async function sourceOf(a){const s=readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser});return s;}
const withFile=(files,name,fn)=>files.map(f=>f.name===name?{...f,text:fn(f.text)}:f);
// `texts` is either a list of facilitator words or what leakInputsOf gives (the words and what the file may say anyway).
const check=(files,texts,expectedTitle='Reframing conflict - Participant workbook',saved=false)=>leakCheck({files,...(Array.isArray(texts)?{facilitatorTexts:texts}:texts),expectedTitle,saved},{DOMParserClass:DOMParser});
const failed=r=>r.failed.map(f=>f.rule);

test('the leak check refuses a workbook with a facilitator paragraph\'s words in it, in the body or in a screen reader\'s description',async()=>{
  const {files,texts}=await madeFiles();
  assert.equal(check(files,texts).passed,true);
  const body=withFile(files,'word/document.xml',t=>t.replace('</w:body>','<w:p><w:r><w:t>Remember to pause after the example.  Pause after the example.</w:t></w:r></w:p></w:body>'));
  assert.deepEqual(failed(check(body,texts)),['FACILITATOR-TEXT']);
  const described=withFile(files,'word/document.xml',t=>t.replace('<w:tblCaption w:val="Matching items"/>','<w:tblCaption w:val="Welcome everyone to the session."/>'));
  assert.deepEqual(failed(check(described,texts)),['FACILITATOR-TEXT']);
});

test('the leak check refuses answer keys and facilitator notes by their labels and by their styles',async()=>{
  const {files,texts}=await madeFiles();
  const key=withFile(files,'word/document.xml',t=>t.replace('</w:body>','<w:p><w:r><w:t>Answer key (multiple choice): zz.</w:t></w:r></w:p></w:body>'));
  assert.ok(failed(check(key,[])).includes('LABELS'),'caught by its label even when there is no guide to compare with');
  const styled=withFile(files,'word/document.xml',t=>t.replace('</w:body>','<w:p><w:pPr><w:pStyle w:val="OSGNotes"/></w:pPr><w:r><w:t>Anything.</w:t></w:r></w:p></w:body>'));
  assert.deepEqual(failed(check(styled,texts)),['STYLES']);
  const defined=withFile(files,'word/styles.xml',t=>t.replace('</w:styles>','<w:style w:type="paragraph" w:styleId="OSGAnswerKey"><w:name w:val="OSG Answer Key"/></w:style></w:styles>'));
  assert.deepEqual(failed(check(defined,texts)),['STYLES']);
});

test('the leak check refuses comments, tracked changes, hidden text, notes, links, fields and custom content',async()=>{
  const {files,texts}=await madeFiles();
  const body=extra=>withFile(files,'word/document.xml',t=>t.replace('</w:body>',`<w:p>${extra}</w:p></w:body>`));
  const cases={
    'tracked insertion':'<w:ins w:id="1" w:author="a"><w:r><w:t>new</w:t></w:r></w:ins>',
    'tracked deletion':'<w:del w:id="2" w:author="a"><w:r><w:delText>old</w:delText></w:r></w:del>',
    'hidden text':'<w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r>',
    'a comment mark':'<w:commentRangeStart w:id="3"/>',
    'a footnote':'<w:r><w:footnoteReference w:id="4"/></w:r>',
    'a link':'<w:hyperlink w:anchor="x"><w:r><w:t>l</w:t></w:r></w:hyperlink>',
    'a field':'<w:r><w:instrText> HYPERLINK "http://example.com" </w:instrText></w:r>',
    'a content control that is not a check box':'<w:sdt><w:sdtPr><w:alias w:val="x"/></w:sdtPr><w:sdtContent><w:r><w:t>y</w:t></w:r></w:sdtContent></w:sdt>'
  };
  for(const [what,xmlText] of Object.entries(cases))assert.deepEqual(failed(check(body(xmlText),texts)),['HIDDEN'],what);
  assert.equal(check(body('<w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>shown</w:t></w:r>'),texts).passed,true,'text that is not hidden is fine');
});

test('the leak check refuses stray parts, outside links and document properties beyond the title',async()=>{
  const {files,texts}=await madeFiles();
  assert.deepEqual(failed(check([...files,{name:'word/comments.xml',text:`<w:comments xmlns:w="${W}"/>`}],texts)),['PARTS','HIDDEN']);
  assert.deepEqual(failed(check([...files,{name:'customXml/item1.xml',text:'<a/>'}],texts)),['PARTS']);
  assert.deepEqual(failed(check([...files,{name:'word/media/x.emf',base64:'AA=='}],texts)),['PARTS']);
  const link=withFile(files,'word/_rels/document.xml.rels',t=>t.replace('</Relationships>','<Relationship Id="rIdX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com" TargetMode="External"/></Relationships>'));
  assert.deepEqual(failed(check(link,texts)),['RELATIONSHIPS']);
  const author=withFile(files,'docProps/core.xml',t=>t.replace('</cp:coreProperties>','<dc:creator>Someone</dc:creator></cp:coreProperties>'));
  assert.deepEqual(failed(check(author,texts)),['PROPERTIES']);
  assert.deepEqual(failed(check(files,texts,'Another title')),['PROPERTIES']);
  assert.deepEqual(failed(check(withFile(files,'word/document.xml',t=>t.slice(0,-20)),texts)).includes('XML'),true);
});

test('a facilitator paragraph too short to search for is matched as a whole paragraph',()=>{
  const files=[{name:'word/document.xml',text:`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Go on.</w:t></w:r></w:p><w:p><w:r><w:t>Go on, then, carefully.</w:t></w:r></w:p></w:body></w:document>`}];
  assert.deepEqual(failed(check(files,['Go on'],'')),['FACILITATOR-TEXT'],'the whole paragraph is the same words');
  assert.equal(check([{name:'word/document.xml',text:`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Go on, then, carefully.</w:t></w:r></w:p></w:body></w:document>`}],['Go on'],'').passed,true,'a short phrase inside a longer one is not a leak');
});

// ---- what the builder takes out of kept paragraphs
test('tracked changes, comments, hidden text, links and the author\'s own content controls never reach the workbook',async()=>{
  const a=await richGuide({own:false});
  let pkg=await a.adapter.sourcePackage();
  const dirt='<w:ins w:id="1" w:author="x"><w:r><w:t> INSERTED</w:t></w:r></w:ins><w:del w:id="2" w:author="x"><w:r><w:delText> DELETEDWORDS</w:delText></w:r></w:del><w:r><w:rPr><w:vanish/></w:rPr><w:t> HIDDENWORDS</w:t></w:r><w:commentRangeStart w:id="5"/><w:r><w:commentReference w:id="5"/></w:r><w:hyperlink r:id="rId99"><w:r><w:t> LINKED</w:t></w:r></w:hyperlink><w:sdt><w:sdtPr><w:alias w:val="mine"/></w:sdtPr><w:sdtContent><w:r><w:t> CONTROLLED</w:t></w:r></w:sdtContent></w:sdt>';
  const at=pkg.indexOf('What makes a disagreement hard to reframe?');
  const end=pkg.indexOf('</w:p>',at);
  pkg=pkg.slice(0,end)+dirt+pkg.slice(end);
  const built=buildWorkbook(readSource(pkg,{DOMParserClass:DOMParser}),{},xml);
  const doc=text(built.files,'word/document.xml'),paragraph=outParas(built.files).find(p=>p.text.includes('What makes a disagreement'));
  assert.match(paragraph.text,/INSERTED/);assert.match(paragraph.text,/LINKED/);assert.match(paragraph.text,/CONTROLLED/,'an insertion, a link and a control keep their words, as plain text');
  for(const gone of ['DELETEDWORDS','HIDDENWORDS','commentRange','commentReference','w:hyperlink','w:val="mine"','w:ins ','w:del '])assert.ok(!doc.includes(gone),`${gone} is not in the workbook`);
  assert.ok(doc.includes('w14:checkbox'),'the rating check boxes are still real check boxes');
  assert.equal(leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser}).passed,true);
});

test('a Say block\'s extra lines and a Trainer notes block never reach the workbook, even when set to another audience',async()=>{
  const a=await started('Audience',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await addBlock(a,'ask','Learner question?');
  await addBlock(a,'trainer-notes','Private advice line one.\nPrivate advice line two.');
  // A cue that is for everyone by default, set to Facilitator only for this block.
  a.pick('discussion');a.type('guide-block-text','Hidden by audience.\nSecond hidden line.');
  a.n('guide-audience-facilitator').checked=true;a.n('guide-audience-everyone').checked=false;
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  await makeWorkbook(a);
  const rows=rowsOf(await filesOf(a)).join('\n');
  assert.match(rows,/Learner question\?/);
  for(const s of ['Private advice','Second hidden line'])assert.ok(!rows.includes(s),s);
});

// ---- the screen
test('the Workbook screen says what is for learners and what is left out, and lists the author\'s own text with a way to it',async()=>{
  const a=await richGuide();
  await click(a,'guide-tab-workbook');
  assert.equal(a.n('guide-workbook').hidden,false);assert.equal(a.n('guide-build').hidden,true);
  assert.equal(a.n('guide-tab-workbook').attributes['aria-pressed'],'true');
  assert.match(a.n('guide-workbook-summary').textContent,/^1 module, 1 lesson, 4 blocks for learners and 3 assessment items\.$/);
  assert.match(a.n('guide-workbook-left').textContent,/^Left out for learners: 2 facilitator-only blocks, 3 sets of facilitator notes, 2 answer keys, 1 time block, the contents and Getting Started pages and 1 paragraph of your own text\.$/);
  assert.match(a.n('guide-workbook-added').textContent,/Learning objectives page from 1 Objective block/);
  assert.equal(a.n('guide-workbook-leftout-box').hidden,false);
  assert.equal(a.n('guide-workbook-leftout').children.length,1);
  assert.match(a.n('guide-workbook-leftout').children[0].textContent,/A reminder I typed myself/);
  await a.n('guide-workbook-leftout').children[0].children[0].onclick();
  assert.equal(a.adapter.snapshot()[a.adapter.cursor()].text,'A reminder I typed myself about the room.');
});

test('with nothing for learners the button is off and the screen says why',async()=>{
  const a=await started('Empty',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await addBlock(a,'say','Only for the facilitator.');
  await click(a,'guide-tab-workbook');
  assert.match(a.n('guide-workbook-summary').textContent,/Nothing in this guide is for learners yet/);
  assert.equal(a.n('guide-workbook-go').disabled,true,'headings alone are not enough to make a workbook from');
  await click(a,'guide-tab-guide');await addBlock(a,'ask','Now something for learners?');await click(a,'guide-tab-workbook');
  assert.equal(a.n('guide-workbook-go').disabled,false,'a block for everyone is');
  const bare=await started('Bare',{quirks:WORD_QUIRKS});
  await click(bare,'guide-tab-workbook');
  assert.match(bare.n('guide-workbook-summary').textContent,/Nothing in this guide is for learners yet/);
  assert.equal(bare.n('guide-workbook-go').disabled,true);
});

test('if Word cannot open the file, it is offered as a download and the screen says so',async()=>{
  const a=await started('Fallback',{quirks:{...WORD_QUIRKS,cannotOpenDocument:true}});
  await addModule(a,'One');await addLesson(a,'First');await addBlock(a,'ask','A question for learners?');
  await makeWorkbook(a);
  assert.equal(a.adapter.openedDocuments().length,0);assert.equal(a.adapter.downloads().length,1);
  assert.equal(a.adapter.downloads()[0].fileName,'Fallback - Participant workbook.docx');
  assert.match(a.n('guide-workbook-ok').textContent,/Word could not open it for you, so it was downloaded as "Fallback - Participant workbook\.docx"/);
});

test('a workbook that fails the leak check is not offered, and the screen says which words leaked',async()=>{
  const a=await started('Leaky',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await addBlock(a,'say','Open the session with the agenda for today.');
  await addBlock(a,'ask','A question for the learners?');
  // A Word that put the facilitator's words into the description of every icon (alt text), which the builder keeps.
  const original=a.adapter.sourcePackage.bind(a.adapter);
  a.adapter.sourcePackage=async()=>(await original()).replaceAll('descr=""','descr="Open the session with the agenda for today."');
  await makeWorkbook(a);
  assert.equal(a.adapter.openedDocuments().length+a.adapter.downloads().length,0,'nothing was offered');
  assert.equal(a.n('guide-workbook-problem').hidden,false);
  assert.match(a.n('guide-workbook-problem').textContent,/The workbook was not made, because its check found a problem\. Words from a facilitator paragraph are in the workbook: "Open the session with the agenda for today\."/);
});

test('switching screens drops the workbook message, and making the workbook leaves the guide as it was',async()=>{
  const a=await richGuide();
  const before=a.rows();
  await makeWorkbook(a);
  assert.deepEqual(a.rows(),before,'the guide is untouched');
  assert.equal(a.n('guide-workbook-msg-box').hidden,false);
  await click(a,'guide-tab-guide');await click(a,'guide-tab-workbook');
  assert.equal(a.n('guide-workbook-msg-box').hidden,true);
});

// ---- the look, the file name, the styles
test('the file name and title come from the guide, without characters Windows will not allow',()=>{
  assert.equal(workbookFileName('Reframing: conflict / 2026?'),'Reframing - conflict 2026 - Participant workbook.docx');
  assert.equal(workbookFileName('Sample guide: Handling workplace conflict'),'Sample guide - Handling workplace conflict - Participant workbook.docx');
  assert.equal(workbookFileName(''),'Guide - Participant workbook.docx');
});

test('the Clearline look becomes real style values, and the facilitator guide\'s own styles are unchanged',()=>{
  const look=resolveWorkbookBrand('clearline',{});
  assert.equal(look.eligible,true,look.issues.join(' '));
  const theme=lookTheme(look.effective);
  assert.deepEqual([theme.navy,theme.ink,theme.bodyHalfPoints,theme.line,theme.sizes],['164B6B','1A1A1A',24,300,{title:52,h1:32,h2:28}]);
  const plain=stylesPart();
  assert.match(plain,/w:ascii="Calibri"/);assert.match(plain,/<w:sz w:val="24"\/>/);assert.match(plain,/w:line="276"/);assert.match(plain,/<w:color w:val="102037"\/>/);
  assert.ok(styleList().length>=18);
  assert.ok(stylesPart({},new Set(['Normal'])).includes('w:styleId="Normal"')&&!stylesPart({},new Set(['Normal'])).includes('OSGModule'));
});

test('a look that is not allowed stops the workbook, with the reason',async()=>{
  const a=await richGuide({own:false});
  const source=readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser});
  assert.throws(()=>buildWorkbook(source,{familyId:'clearline',brand:{text:'#999999',background:'#FFFFFF'}},xml),/This look cannot be used: Body text\/background contrast is below 4\.5:1\./);
  assert.throws(()=>buildWorkbook(source,{familyId:'nope'},xml),/family ID is unavailable/);
});

test('a guide with nothing for learners cannot be made into a workbook',async()=>{
  const a=await started('Nothing',{quirks:WORD_QUIRKS});
  const source=readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser});
  assert.throws(()=>buildWorkbook(source,{},xml),/Nothing in this guide is for learners yet/);
});

test('the file can be zipped, read back and opened as a package',async()=>{
  const a=await richGuide();
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},xml);
  const bytes=await zipWorkbook(built.files,JSZip,'nodebuffer'),zip=await JSZip.loadAsync(bytes);
  assert.deepEqual(Object.keys(zip.files).sort(),built.files.map(f=>f.name).sort());
  assert.equal(built.summary.title,'Reframing conflict');
});

// ---- what an independent review found, each kept as a test
const buildFrom=async(a,edit=p=>p)=>buildWorkbook(readSource(edit(await a.adapter.sourcePackage()),{DOMParserClass:DOMParser}),{},xml);
const passes=built=>leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
const allText=files=>files.filter(f=>f.text!==undefined).map(f=>f.text).join('\n');
// Puts `snippet` at the end of the first paragraph whose text contains `word`.
const inject=(pkg,word,snippet)=>{const at=pkg.indexOf(word);assert.ok(at>=0,`${word} is in the source`);const end=pkg.indexOf('</w:p>',at);return pkg.slice(0,end)+snippet+pkg.slice(end);};

test('an assessment section keeps its note and information lines and its items, and not what is typed among or after them',()=>{
  const c=classify([P('OSG Guide Title','G'),P('OSG Module','Module 1: A'),P('OSG Assessment','Level Two Assessment'),P('Normal','The note to the learner.'),P('Normal','Your name:'),
    P('OSG Question','Question 1. Q?'),P('OSG Option','one'),P('Normal','A facilitator reminder typed between items.'),P('OSG Question','Question 2. R?'),P('Normal','A reminder typed after the last item.')]);
  assert.deepEqual(c.decisions.map(d=>d.reason),['title','heading','heading','section','section','item','item','own','item','own']);
});

test('a facilitator cue inside a section ends its lead, and a pasted key in a style Word renamed is not part of the section',()=>{
  const cue=classify([P('OSG Guide Title','G'),P('OSG Module','Module 1: A'),P('OSG Assessment','Level One Assessment'),P('OSG Cue Facilitator',' Say Watch the clock.'),P('Normal','typed under the cue')]);
  assert.deepEqual(cue.decisions.slice(3).map(d=>d.reason),['facilitator-block','own']);
  const renamed=classify([P('OSG Guide Title','G'),P('OSG Module','Module 1: A'),P('OSG Assessment','Level One Assessment'),P('OSG Answer Key1','Answer key (multiple choice): b.'),P('Normal','Answer each question.')]);
  assert.deepEqual(renamed.decisions.slice(3).map(d=>d.reason),['own','own']);
  const many=classify([P('OSG Module','Module 1: A'),P('OSG Assessment','Level One Assessment'),...Array.from({length:10},(_,i)=>P('Normal',`line ${i}`))]);
  assert.equal(many.decisions.filter(d=>d.reason==='section').length,8,'the lead has a limit, so a long run of typed lines cannot all pass for it');
});

test('text the author types after the last item, in an assessment, stays out of the workbook and is listed',async()=>{
  const a=await richGuide();
  await a.adapter.insert(()=>para(null,run('Facilitator reminder: the pass mark is 80%.')),{where:'after',after:a.rows().length-1,select:'last'});
  await makeWorkbook(a);
  assert.ok(!allText(await filesOf(a)).includes('pass mark'));
  await click(a,'guide-tab-guide');await click(a,'guide-tab-workbook');
  assert.ok(Array.from(a.n('guide-workbook-leftout').children).some(li=>/Facilitator reminder: the pass mark/.test(li.textContent)),'it is in the left-out list');
});

test('hidden words never reach the title page, the contents, the learning objectives or the document title',async()=>{
  const a=await richGuide({own:false});
  const hidden='<w:r><w:rPr><w:vanish/></w:rPr><w:t> SECRET-WORDS</w:t></w:r>';
  const built=await buildFrom(a,p=>inject(inject(inject(p,'Reframing conflict',hidden),'Lesson 1.1',hidden),'Identify two ways',hidden));
  assert.ok(!allText(built.files).includes('SECRET-WORDS'));
  assert.equal(built.title,'Reframing conflict');
  assert.equal(passes(built).passed,true);
});

test('text boxes, floating pictures, alternate content and formatting revisions do not come through, and the people named in them do not either',async()=>{
  const a=await richGuide({own:false});
  const box='<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="77" name="Sticky note" descr="NOTE-ALT"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:txbxContent><w:p><w:r><w:t>BOXED-SECRET</w:t></w:r></w:p></w:txbxContent></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    +'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:r><w:t>ALTERNATE-SECRET</w:t></w:r></mc:Choice></mc:AlternateContent>'
    +'<w:pPrChange w:id="5" w:author="Jason Q Trainer"><w:pPr/></w:pPrChange><w:r><w:rPr><w:rPrChange w:id="6" w:author="Jason Q Trainer"><w:rPr/></w:rPrChange></w:rPr><w:t> shown</w:t></w:r>'
    +'<w:smartTag w:uri="x" w:element="y"><w:smartTagPr><w:attr w:name="n" w:val="SMART-SECRET"/></w:smartTagPr><w:r><w:t> tagged</w:t></w:r></w:smartTag>';
  const built=await buildFrom(a,p=>inject(p,'Reframe the conversation',box));
  const doc=allText(built.files);
  for(const gone of ['BOXED-SECRET','ALTERNATE-SECRET','NOTE-ALT','Sticky note','Jason Q Trainer','pPrChange','rPrChange','SMART-SECRET','txbxContent','AlternateContent'])assert.ok(!doc.includes(gone),`${gone} is not in the workbook`);
  const lesson=outParas(built.files).find(p=>p.text.startsWith('Lesson 1.1'));
  assert.match(lesson.text,/shown tagged$/,'the words around them stay');
  assert.equal(passes(built).passed,true);
});

test('a block\'s minutes are taken off however the runs are split',async()=>{
  const a=await richGuide({own:false});
  const built=await buildFrom(a,p=>p.replaceAll(' (5 min)</w:t></w:r>',' (5</w:t></w:r><w:r><w:t xml:space="preserve"> min)</w:t></w:r><w:r><w:t xml:space="preserve"></w:t></w:r>'));
  const rows=rowsOf(built.files);
  assert.ok(rows.some(r=>r.includes('What makes a disagreement hard to reframe?')));
  assert.ok(!rows.some(r=>/min\)|\(\d/.test(r)),rows.filter(r=>/min\)/.test(r)).join('|'));
});

test('moved text is kept where it was moved to and dropped where it was moved from; a SEQ field keeps its result',async()=>{
  const a=await richGuide({own:false});
  const edits='<w:moveTo w:id="9" w:author="x"><w:r><w:t> Moved-in sentence.</w:t></w:r></w:moveTo><w:moveFrom w:id="8" w:author="x"><w:r><w:t> Moved-out sentence.</w:t></w:r></w:moveFrom>'
    +'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t> Figure 1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const built=await buildFrom(a,p=>inject(p,'What makes a disagreement hard',edits));
  const ask=outParas(built.files).find(p=>p.text.includes('What makes a disagreement'));
  assert.match(ask.text,/Moved-in sentence\./);assert.match(ask.text,/Figure 1/);
  assert.ok(!allText(built.files).includes('Moved-out sentence'));assert.ok(!allText(built.files).includes('SEQ Figure'));
  assert.equal(passes(built).passed,true,'a field in learner text does not stop the workbook');
});

test('text hidden by a character style, even one based on another, does not come through',async()=>{
  const a=await richGuide({own:false});
  const styles='<w:style w:type="character" w:styleId="Secret"><w:name w:val="Secret"/><w:rPr><w:vanish/></w:rPr></w:style><w:style w:type="character" w:styleId="Secret2"><w:name w:val="Secret2"/><w:basedOn w:val="Secret"/></w:style></w:styles>';
  const built=await buildFrom(a,p=>inject(p.replace('</w:styles>',styles),'What makes a disagreement hard','<w:r><w:rPr><w:rStyle w:val="Secret2"/></w:rPr><w:t> STYLE-HIDDEN</w:t></w:r>'));
  assert.ok(!allText(built.files).includes('STYLE-HIDDEN'));
});

// Things that are not leaks, which a first version refused.
const madeFor=async(build)=>{const a=await started('Handling workplace conflict',{quirks:WORD_QUIRKS});await addModule(a,'One');await addLesson(a,'First');await build(a);await makeWorkbook(a);return a;};
test('a learner block that repeats a facilitator line, or one word of it, does not stop the workbook',async()=>{
  for(const [say,notes,ask] of [['Open the session with the agenda for today.','','Open the session with the agenda for today.'],['Conflict.','','What is your view?'],['Say it.','Brainstorm.','Brainstorm three ways to reframe it.']]){
    const a=await madeFor(async a=>{await addBlock(a,'say',say,{notes});await addBlock(a,'ask',ask);});
    assert.equal(a.adapter.openedDocuments().length,1,`${say} | ${notes} | ${ask}: ${a.n('guide-workbook-problem').textContent}`);
  }
});

test('copying an Objective block\'s sentence into an item\'s Objective field does not stop the workbook, and the key stays out',async()=>{
  const a=await started('Objectives',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await addBlock(a,'objective','Identify two ways to reframe a disagreement.');
  await click(a,'guide-tab-assess');await click(a,'guide-sec-level-two');
  typeItem(a,'true-false',{'guide-item-stem':'Reframing helps?','guide-item-objective':'Identify two ways to reframe a disagreement.'});a.n('guide-item-true').checked=true;
  await submit(a,'guide-item-form');assert.equal(a.err(),'');
  await makeWorkbook(a);
  assert.equal(a.adapter.openedDocuments().length,1,a.n('guide-workbook-problem').textContent);
  const rows=rowsOf(await filesOf(a));
  assert.ok(!rows.some(r=>r.includes('Answer key')||r.includes('Objective:')));
});

test('the leak check still refuses what is added to the file: a facilitator paragraph, the author\'s own typed text, alt text and containers',async()=>{
  const {files,texts}=await madeFiles();
  assert.equal(check(files,texts).passed,true);
  const add=xmlText=>withFile(files,'word/document.xml',t=>t.replace('</w:body>',`${xmlText}</w:body>`));
  assert.deepEqual(failed(check(add('<w:p><w:r><w:t>A reminder I typed myself about the room.</w:t></w:r></w:p>'),texts)),['FACILITATOR-TEXT'],'the author\'s own left-out text');
  assert.deepEqual(failed(check(add('<w:p><w:r><w:t>Fine.</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr id="9" name="Pause after the example" descr=""/></wp:inline></w:drawing></w:r></w:p>'),texts)),['FACILITATOR-TEXT'],'a picture\'s name');
  assert.deepEqual(failed(check(add('<w:p><w:pPrChange w:id="1" w:author="a"><w:pPr/></w:pPrChange></w:p>'),texts)),['HIDDEN']);
  assert.deepEqual(failed(check(add('<w:p><w:r><w:txbxContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:txbxContent></w:r></w:p>'),texts)),['HIDDEN']);
  assert.deepEqual(failed(check(add('<w:p><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/></w:p>'),texts)),['HIDDEN']);
});

test('a workbook Word has saved again passes the saved-copy check, which accepts what Word adds and still searches it',async()=>{
  const {files,texts}=await madeFiles();
  const word=[
    ...withFile(withFile(withFile(files,'docProps/core.xml',t=>t.replace('</cp:coreProperties>','<dc:creator>Jason</dc:creator><cp:lastModifiedBy>Jason</cp:lastModifiedBy><cp:revision>2</cp:revision><dc:description></dc:description></cp:coreProperties>')),
      '_rels/.rels',t=>t.replace('</Relationships>','<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')),
      'word/_rels/document.xml.rels',t=>t.replace('</Relationships>','<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rIdW" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/></Relationships>')),
    {name:'docProps/app.xml',text:'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>'},
    {name:'word/theme/theme1.xml',text:'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"/>'},
    {name:'word/fontTable.xml',text:`<w:fonts xmlns:w="${W}"/>`},{name:'word/webSettings.xml',text:`<w:webSettings xmlns:w="${W}"/>`}
  ];
  const toc='<w:p><w:hyperlink w:anchor="_Toc100"><w:r><w:t>Module 1</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc100 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>';
  const saved=withFile(word,'word/document.xml',t=>t.replace('</w:body>',`${toc}</w:body>`));
  assert.equal(check(saved,texts,'Reframing conflict - Participant workbook',true).passed,true,check(saved,texts,'Reframing conflict - Participant workbook',true).summary);
  assert.ok(failed(check(saved,texts)).includes('PARTS'),'the strict check, used before a workbook is offered, does not accept them');
  // what Word adds is still searched
  const leaky=[...saved.filter(f=>f.name!=='docProps/app.xml'),{name:'docProps/app.xml',text:'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>Welcome everyone to the session.</Company></Properties>'}];
  assert.deepEqual(failed(check(leaky,texts,'Reframing conflict - Participant workbook',true)),['FACILITATOR-TEXT']);
  const description=withFile(saved,'docProps/core.xml',t=>t.replace('<dc:description></dc:description>','<dc:description>Notes for the facilitator</dc:description>'));
  assert.deepEqual(failed(check(description,texts,'Reframing conflict - Participant workbook',true)),['PROPERTIES']);
  const external=withFile(saved,'word/document.xml',t=>t.replace('<w:hyperlink w:anchor="_Toc100">','<w:hyperlink r:id="rIdX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'));
  assert.ok(failed(check(external,texts,'Reframing conflict - Participant workbook',true)).includes('HIDDEN'),'a link out is still refused');
  // Word writes an empty footnotes part and endnotes part (separator lines only) into every copy it saves; found in real Word.
  const noteRels=t=>t.replace('</Relationships>','<Relationship Id="rIdN1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdN2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>');
  const sep=(root,one)=>`<w:${root} xmlns:w="${W}"><w:${one} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${one}><w:${one} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${one}></w:${root}>`;
  const withNotes=(foot=sep('footnotes','footnote'))=>[...withFile(saved,'word/_rels/document.xml.rels',noteRels),{name:'word/footnotes.xml',text:foot},{name:'word/endnotes.xml',text:sep('endnotes','endnote')}];
  assert.equal(check(withNotes(),texts,'Reframing conflict - Participant workbook',true).passed,true,check(withNotes(),texts,'Reframing conflict - Participant workbook',true).summary);
  assert.ok(failed(check(withNotes(),texts)).includes('PARTS'),'the strict check still refuses a notes part');
  const realNote=sep('footnotes','footnote').replace('</w:footnotes>','<w:footnote w:id="1"><w:p><w:r><w:t>Say this quietly</w:t></w:r></w:p></w:footnote></w:footnotes>');
  assert.ok(failed(check(withNotes(realNote),texts,'Reframing conflict - Participant workbook',true)).includes('HIDDEN'),'a real footnote is still refused');
  const textInSeparator=sep('footnotes','footnote').replace('<w:separator/>','<w:separator/><w:t>hidden words</w:t>');
  assert.ok(failed(check(withNotes(textInSeparator),texts,'Reframing conflict - Participant workbook',true)).includes('HIDDEN'),'text inside a separator is still refused');
  // and through the command-line path, as a saved .docx
  const a=await richGuide();await makeWorkbook(a);
  const guide=await packageToDocx(await a.adapter.sourcePackage(),'Reframing conflict');
  const bytes=await zipWorkbook(saved,JSZip,'nodebuffer');
  const result=await checkWorkbookFile(bytes,guide);
  assert.equal(result.passed,true,result.summary);
});
