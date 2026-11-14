// The learner workbook, slice 3b: the looks, the brand settings, the footer, the page, black and white, and the lines to write on.
// Everything here runs against stand-ins for Word, not Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readSource,buildWorkbook,WORKBOOK_LOOKS,MAX_WRITING_LINES} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {readStyles,textOf} from '../src/guide-check.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {zipWorkbook} from '../src/guide-workbook.mjs';
import {readDocxFiles} from '../scripts/workbook-leak-check.mjs';
import {started,pane,WORD_QUIRKS,click,submit,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml={DOMParserClass:DOMParser,SerializerClass:XMLSerializer};
const parse=t=>new DOMParser().parseFromString(t,'application/xml');
const text=(files,name)=>files.find(f=>f.name===name)?.text;
const typeItem=(a,type,fields)=>{a.itemRadio(type).checked=true;a.itemRadio(type).onchange();for(const [id,v] of Object.entries(fields))a.n(id).value=v;};

// A guide with a block of each kind that matters here, and two assessment sections with tables in them.
async function guide({tables=true}={}){
  const a=await started('Reframing conflict',{quirks:WORD_QUIRKS});
  await addModule(a,'Recognize the conflict');await addLesson(a,'Reframe the conversation','30');
  await addBlock(a,'objective','Identify two ways to reframe a disagreement.');
  await addBlock(a,'say','Welcome everyone to the session.',{notes:'Pause after the example.'});
  await addBlock(a,'ask','What makes a disagreement hard to reframe?\nA second learner line.');
  await addBlock(a,'key-point','A neutral question describes the situation without blame.');
  await addBlock(a,'discussion','Talk about the last disagreement you saw.');
  await addBlock(a,'activity','In pairs, write two neutral questions.',{notes:'Allow eight minutes.'});
  await addBlock(a,'workbook','Write your own example on this page.');
  if(tables){
    await click(a,'guide-tab-assess');await click(a,'guide-sec-level-two');
    typeItem(a,'matching',{'guide-item-stem':'Match each term with its meaning.','guide-item-pairs':'Active listening | Fully attending to the speaker\nProbing | Asking follow-up questions'});
    await submit(a,'guide-item-form');
    await click(a,'guide-sec-level-one');
    typeItem(a,'rating-statements',{'guide-item-stem':'Trainer.','guide-item-statements':'The trainer was prepared.'});
    await submit(a,'guide-item-form');
    assert.equal(a.err(),'');
  }
  return a;
}
const sourceOf=async a=>readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser});
const build=async(a,options={})=>buildWorkbook(await sourceOf(a),options,xml);
const strict=built=>leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
const rowsOf=files=>{
  const styles=readStyles(parse(text(files,'word/styles.xml')));
  return Array.from(parse(text(files,'word/document.xml')).getElementsByTagNameNS(W,'p')).map(p=>{
    const id=Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val')||'Normal';
    return `${styles.get(id)?.name||id}: ${textOf(p)}`.trim();
  });
};
const makeWorkbook=async a=>{await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');};
const opened=async a=>{const d=a.adapter.openedDocuments()[0];assert.ok(d,a.n('guide-workbook-problem').textContent);return readDocxFiles(Buffer.from(d.base64,'base64'));};

test('all five looks are offered, each with its own fonts and accent, and each makes a workbook that passes both checks',async()=>{
  assert.deepEqual(WORKBOOK_LOOKS.map(l=>l.id),['clearline','editorial-study','practice-lab','field-guide','reflection-journal']);
  assert.ok(WORKBOOK_LOOKS.every(l=>l.name&&l.note));
  const a=await guide();
  for(const family of WORKBOOK_FAMILIES){
    const built=await build(a,{familyId:family.id}),styles=text(built.files,'word/styles.xml');
    assert.ok(styles.includes(`w:ascii="${family.defaults.bodyFont}"`),`${family.name} body font`);
    assert.ok(styles.includes(`<w:rFonts w:ascii="${family.defaults.headingFont}"`),`${family.name} heading font`);
    assert.ok(styles.includes(`<w:color w:val="${family.defaults.accent.slice(1)}"/>`),`${family.name} accent`);
    assert.equal(strict(built).passed,true,`${family.name}: ${strict(built).summary}`);
    const result=await checkDocx(await zipWorkbook(built.files,JSZip,'nodebuffer'));
    assert.deepEqual(result.findings.filter(f=>f.level==='fail'),[],`${family.name} passes the structure check`);
  }
});

test('the footer carries the organization and the title and "Page X of Y", on every page but the first',async()=>{
  const a=await guide();
  const built=await build(a,{brand:{organizationName:'Smith & Co'}});
  const footer=parse(text(built.files,'word/footer1.xml')),doc=text(built.files,'word/document.xml');
  assert.equal(textOf(footer.documentElement).replace(/\s+/g,' ').trim(),'Smith & Co · Reframing conflict Page 1 of 1','what it says before the fields are updated');
  assert.deepEqual(Array.from(footer.getElementsByTagNameNS(W,'instrText')).map(i=>i.textContent.trim()),['PAGE','NUMPAGES']);
  assert.match(doc,/<w:footerReference w:type="default" r:id="rIdFooter"\/>/);assert.match(doc,/<w:titlePg\/>/);
  assert.match(text(built.files,'word/_rels/document.xml.rels'),/Id="rIdFooter" Type="[^"]+\/footer" Target="footer1\.xml"/);
  assert.match(text(built.files,'[Content_Types].xml'),/PartName="\/word\/footer1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.footer\+xml"/);
  assert.equal(strict(built).passed,true);
  assert.equal(textOf((parse(text((await build(a)).files,'word/footer1.xml'))).documentElement).replace(/\s+/g,' ').trim(),'Reframing conflict Page 1 of 1','with no organization it is the title alone');
});

test('A4 and Letter change the page and the footer\'s right edge',async()=>{
  const a=await guide({tables:false});
  const letter=await build(a),a4=await build(a,{brand:{pageSize:'A4'}});
  assert.match(text(letter.files,'word/document.xml'),/<w:pgSz w:w="12240" w:h="15840"\/><w:pgMar w:top="1152"/);
  assert.match(text(a4.files,'word/document.xml'),/<w:pgSz w:w="11906" w:h="16838"\/>/);
  assert.match(text(letter.files,'word/footer1.xml'),/w:pos="9936"/,'12240 less two margins of 1152');
  assert.match(text(a4.files,'word/footer1.xml'),/w:pos="9602"/,'11906 less two margins of 1152');
});

// Colours that are not grey: a fill or colour whose red, green and blue differ.
const colours=files=>{
  const out=[];
  for(const f of files.filter(f=>/^word\/(document|styles|footer1)\.xml$/.test(f.name)))
    for(const el of Array.from(parse(f.text).getElementsByTagName('*')))
      for(const [name,value] of [['fill',el.getAttributeNS(W,'fill')],['color',el.getAttributeNS(W,'color')],['val',el.localName==='color'?el.getAttributeNS(W,'val'):'']])
        if(/^[0-9a-f]{6}$/i.test(value)&&!(value.slice(0,2)===value.slice(2,4)&&value.slice(2,4)===value.slice(4)))out.push(`${f.name} ${el.nodeName} ${name}=${value}`);
  return out;
};
test('black and white leaves no colour: the accent is dark grey and the tables\' shading and borders are grey too',async()=>{
  const a=await guide();
  assert.ok((await build(a)).files&&colours((await build(a)).files).length>0,'in colour there are colours (so the next check means something)');
  const gray=await build(a,{brand:{printMode:'grayscale'}});
  assert.deepEqual(colours(gray.files),[]);
  assert.match(text(gray.files,'word/styles.xml'),/<w:color w:val="333333"\/>/);
  assert.equal(strict(gray).passed,true);
  // a bright accent typed in is not used in black and white
  assert.deepEqual(colours((await build(a,{brand:{printMode:'grayscale',accent:'#D4AF37'}})).files),[]);
});

test('an accent color sets the headings, and one too pale to read as text is refused with the reason',async()=>{
  const a=await guide({tables:false});
  const red=await build(a,{brand:{accent:'#7A1F2B'}});
  assert.match(text(red.files,'word/styles.xml'),/<w:color w:val="7A1F2B"\/>/);
  await assert.rejects(async()=>build(a,{brand:{accent:'#F2C94C'}}),/This look cannot be used: the accent color is too light to read as heading text \(its contrast with the page is 1\.\d to 1, and it needs at least 4\.5 to 1\)/);
  await assert.rejects(async()=>build(a,{brand:{accent:'red'}}),/This look cannot be used: accent must be an opaque #RRGGBB color/);
  assert.equal((await build(a,{brand:{printMode:'grayscale',accent:'#F2C94C'}})).files.length>5,true,'black and white does not use it, so it is not refused');
});

test('questions and activities are followed by lines to write on, and other blocks are not',async()=>{
  const a=await guide({tables:false});
  const rows=rowsOf((await build(a)).files);
  const at=s=>rows.findIndex(r=>r.includes(s));
  const notesAfter=s=>{const i=at(s);let j=i+1;while(rows[j]?.startsWith('OSG Cue Body'))j++;return rows.slice(j,j+3);};
  const three=['OSG Answer Line: Your notes:','OSG Answer Line:','OSG Answer Line:'];
  assert.deepEqual(notesAfter('What makes a disagreement hard'),three,'after Ask, and after its extra line');
  assert.ok(rows[at('What makes a disagreement hard')+1].startsWith('OSG Cue Body: A second learner line.'),'the extra line stays with its block');
  assert.deepEqual(notesAfter('Talk about the last disagreement'),three,'after Discussion');
  assert.deepEqual(notesAfter('In pairs, write two neutral questions'),three,'after Activity');
  assert.deepEqual(notesAfter('Write your own example'),three,'after Workbook');
  assert.ok(!rows[at('Identify two ways')+1].startsWith('OSG Answer Line'),'not after an Objective');
  assert.ok(!rows[at('A neutral question describes')+1].startsWith('OSG Answer Line'),'not after a Key point');
  assert.equal(rows.filter(r=>r==='OSG Answer Line: Your notes:').length,4);
});

test('the number of lines is the author\'s to choose, from none to six, and anything else is refused',async()=>{
  const a=await guide({tables:false});
  const count=async lines=>rowsOf((await build(a,{lines})).files).filter(r=>r.startsWith('OSG Answer Line')&&r!=='OSG Answer Line: Name:');
  assert.equal((await count(0)).length,0);
  assert.equal((await count(1)).length,4,'one line each, which is the label');
  assert.equal((await count(MAX_WRITING_LINES)).length,4*6);
  for(const bad of [7,-1,1.5,'3',NaN])await assert.rejects(async()=>build(a,{lines:bad}),/Lines to write on must be a whole number from 0 to 6/);
});

test('the lines to write on do not trip the structure check or the leak check',async()=>{
  const a=await guide();
  const built=await build(a,{lines:6});
  assert.equal(strict(built).passed,true);
  const result=await checkDocx(await zipWorkbook(built.files,JSZip,'nodebuffer'));
  assert.deepEqual(result.findings.filter(f=>f.level==='fail'||f.rule==='EMPTY-PARAGRAPHS'),[]);
});

test('the leak check also reads the footer, and refuses a header or a footer that links anywhere',async()=>{
  const a=await guide();
  const built=await build(a),withFooter=t=>built.files.map(f=>f.name==='word/footer1.xml'?{...f,text:t(f.text)}:f);
  const check=files=>leakCheck({files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  const leaked=check(withFooter(t=>t.replace('</w:p>','<w:r><w:t> Welcome everyone to the session.</w:t></w:r></w:p>')));
  assert.deepEqual(leaked.failed.map(f=>f.rule),['FACILITATOR-TEXT']);
  const link=check(withFooter(t=>t.replace('</w:p>','<w:r><w:instrText> HYPERLINK "http://example.com" </w:instrText></w:r></w:p>')));
  assert.deepEqual(link.failed.map(f=>f.rule),['HIDDEN']);
  const header=check([...built.files,{name:'word/header1.xml',text:`<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`}]);
  assert.deepEqual(header.failed.map(f=>f.rule),['PARTS']);
});

// ---- the screen
test('the screen offers the looks and the brand settings, with the first look chosen and three lines to write on',async()=>{
  const a=await guide({tables:false});
  await click(a,'guide-tab-workbook');
  const radios=a.n('guide-wb-looks').children.map(l=>l.children[0]);
  assert.deepEqual(radios.map(r=>r.id),WORKBOOK_LOOKS.map(l=>`guide-wb-look-${l.id}`));
  assert.equal(radios[0].checked,true);assert.equal(radios.filter(r=>r.checked).length,1);
  assert.match(a.n('guide-wb-looks').children[2].children[1].textContent,/^Practice Lab: Numbered responses\.$/);
  await click(a,'guide-workbook-go');
  const files=await opened(a);
  assert.ok(text(files,'word/styles.xml').includes('w:ascii="Aptos Display"'),'Clearline');
  assert.equal(rowsOf(files).filter(r=>r==='OSG Answer Line: Your notes:').length,4,'three lines is the default');
});

test('what the author chooses is used, and a bad choice is refused above the button without making anything',async()=>{
  const a=await guide({tables:false});
  await click(a,'guide-tab-workbook');
  const practice=a.n('guide-wb-looks').children[2].children[0];
  for(const r of a.n('guide-wb-looks').children)r.children[0].checked=false;
  practice.checked=true;
  a.n('guide-wb-org').value='Wrench Training';a.n('guide-wb-accent').value='7A1F2B';a.n('guide-wb-page-a4').checked=true;a.n('guide-wb-page-letter').checked=false;a.n('guide-wb-lines').value='2';
  await click(a,'guide-workbook-go');
  assert.equal(a.err(),'');
  const files=await opened(a),styles=text(files,'word/styles.xml');
  assert.match(text(files,'word/document.xml'),/<w:pgSz w:w="11906" w:h="16838"\/>/);
  assert.match(text(files,'word/footer1.xml'),/Wrench Training · Reframing conflict/);
  assert.match(styles,/<w:color w:val="7A1F2B"\/>/,'a typed accent without the # works');
  assert.equal(rowsOf(files).filter(r=>r==='OSG Answer Line: Your notes:').length,4);
  assert.equal(rowsOf(files).filter(r=>r==='OSG Answer Line:').length,4,'two lines each: the label and one more');
  // refusals
  const before=a.adapter.openedDocuments().length;
  for(const [field,value,pattern] of [['guide-wb-accent','blue',/Type the accent color as a # and six letters or digits/],['guide-wb-accent','#F2C94C',/too light to read as heading text/],['guide-wb-lines','x',/Lines to write on must be a whole number from 0 to 6/],['guide-wb-lines','7',/Lines to write on must be a whole number from 0 to 6/]]){
    a.n('guide-wb-accent').value='';a.n('guide-wb-lines').value='3';
    a.n(field).value=value;
    await click(a,'guide-workbook-go');
    assert.match(a.n('guide-workbook-problem').textContent,pattern,`${field} ${value}`);
    assert.equal(a.n('guide-workbook-problem').hidden,false);
  }
  assert.equal(a.adapter.openedDocuments().length,before,'nothing was made');
});

test('the look is remembered between visits, and a pane with nothing remembered starts on the first look',async()=>{
  const store=new Map(),storage={getItem:k=>store.has(k)?store.get(k):null,setItem:(k,v)=>store.set(k,v)};
  const a=pane(undefined,{quirks:WORD_QUIRKS,storage});await a.p.refresh();
  a.type('guide-title','Remember');await submit(a,'guide-start-form');
  await addModule(a,'One');await addLesson(a,'First');await addBlock(a,'ask','A question?');
  await click(a,'guide-tab-workbook');
  for(const r of a.n('guide-wb-looks').children)r.children[0].checked=false;
  a.n('guide-wb-looks').children[4].children[0].checked=true;
  a.n('guide-wb-org').value='My Studio';a.n('guide-wb-print-gray').checked=true;a.n('guide-wb-print-color').checked=false;a.n('guide-wb-lines').value='5';
  await click(a,'guide-workbook-go');
  assert.equal(a.adapter.openedDocuments().length,1);
  const b=pane(undefined,{quirks:WORD_QUIRKS,storage});
  const chosen=b.n('guide-wb-looks').children.map(l=>l.children[0]).find(r=>r.checked);
  assert.equal(chosen.id,'guide-wb-look-reflection-journal');
  assert.equal(b.n('guide-wb-org').value,'My Studio');assert.equal(b.n('guide-wb-print-gray').checked,true);assert.equal(b.n('guide-wb-lines').value,'5');
  const broken=pane(undefined,{quirks:WORD_QUIRKS,storage:{getItem:()=>'not json {',setItem(){throw new Error('refused');}}});
  assert.equal(broken.n('guide-wb-looks').children.map(l=>l.children[0]).find(r=>r.checked).id,'guide-wb-look-clearline','a damaged or refused store changes nothing');
  assert.equal(pane(undefined,{quirks:WORD_QUIRKS}).n('guide-wb-looks').children.length,5);
});
