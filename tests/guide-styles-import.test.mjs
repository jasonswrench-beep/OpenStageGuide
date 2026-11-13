// Taking the styles of a Word file into the guide: which styles of the file match the guide's, what the file's styles come to (with the theme's
// fonts and the document's defaults), what is brought inside the guide's limits, and the edits that result, undone with one Undo.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {readDocxStyles} from '../src/guide-docx.mjs';
import {listStyles} from '../src/guide-styles-edit.mjs';
import {candidates,fitSpec,planImport,importSummary,changeBetween,AS_NEW,IMPORT_LIMITS} from '../src/guide-styles-import.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// What a template made in Word looks like: fonts from the theme, defaults for the spacing, a few styles with the usual Word names, and one of the company's own.
const STYLES=`<w:styles xmlns:w="${W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Heading1"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="CompanyNote"><w:name w:val="Company Note"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:pBdr><w:left w:val="single" w:sz="24" w:space="4" w:color="C00000"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:ind w:left="360"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Huge"><w:name w:val="Huge"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Bad&lt;Font" w:hAnsi="Bad&lt;Font"/><w:sz w:val="400"/></w:rPr><w:pPr><w:spacing w:before="9000"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;
const THEME='<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="x"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>';

async function readTemplate(styles=STYLES){
  const zip=new JSZip();
  zip.file('word/styles.xml',styles);
  zip.file('word/theme/theme1.xml',THEME);
  const {stylesXml,themeFonts}=await readDocxStyles(await zip.generateAsync({type:'uint8array'}),JSZip,io);
  return listStyles(stylesXml,io,new Set(),{themeFonts,paragraphDefaults:true});
}
const byName=(list,name)=>list.find(s=>s.name===name);

test('a style of a Word file comes to what Word makes of it: the theme\'s fonts, the document\'s defaults and the styles it is based on',async()=>{
  const file=await readTemplate();
  const h1=byName(file,'heading 1');
  assert.equal(h1.effective.font.name,'Calibri Light','the heading font of the theme');
  assert.equal(h1.effective.font.size,16);
  assert.equal(h1.effective.font.color,'#2F5496');
  assert.equal(h1.effective.paragraph.keepWithNext,true);
  assert.equal(h1.effective.paragraph.spaceBefore,12);
  assert.equal(h1.effective.paragraph.spaceAfter,0);
  const h2=byName(file,'heading 2');
  assert.equal(h2.effective.font.size,13,'its own size');
  assert.equal(h2.effective.font.bold,true);
  assert.equal(h2.effective.font.color,'#2F5496','and the rest from heading 1');
  const normal=byName(file,'Normal');
  assert.equal(normal.effective.font.name,'Calibri','the body font of the theme');
  assert.equal(normal.effective.font.size,11);
  assert.equal(normal.effective.paragraph.spaceAfter,8,'the space after Word gives every paragraph');
  assert.equal(normal.effective.paragraph.lineSpacing,1.08);
  assert.deepEqual(byName(file,'Company Note').effective.borders,[{location:'left',color:'#C00000',width:3}]);
});

test('a style of the file that has the name of one of the guide\'s changes that one; the rest can be added as the author\'s own; character styles are not offered',async()=>{
  const a=demo();
  const file=await readTemplate();
  const list=candidates(file,a.listStyles());
  assert.deepEqual(list.map(c=>c.name),['Normal','heading 1','heading 2','Company Note','Huge'],'the matching styles come first, in the order of the file');
  assert.equal(list.find(c=>c.name==='heading 1').to,'Heading1');
  assert.equal(list.find(c=>c.name==='Normal').to,'Normal');
  assert.equal(list.find(c=>c.name==='Company Note').to,null);
  assert.ok(!list.some(c=>c.name==='Strong'));
});

test('what the guide cannot hold is brought inside what it can, or left out, and said',async()=>{
  const file=await readTemplate();
  const huge=fitSpec(byName(file,'Huge').effective);
  assert.equal(huge.spec.font.size,96);
  assert.equal(huge.spec.paragraph.spaceBefore,144);
  assert.equal(huge.spec.font.name,undefined,'a font name with < in it is not taken');
  assert.ok(huge.notes.some(n=>/font size was brought to 96/.test(n)));
  assert.ok(huge.notes.some(n=>/font name/.test(n)));
  assert.ok(huge.notes.some(n=>/space before was brought to 144/.test(n)));
  assert.deepEqual(fitSpec({font:{size:12},paragraph:{alignment:'sideways'}}).spec,{font:{size:12}},'an alignment the guide does not know is left out');
});

test('importing a style that matches changes the guide\'s style by only what differs, and every paragraph in it follows',async()=>{
  const a=demo();
  const file=await readTemplate();
  const guide=a.listStyles().filter(s=>s.type==='paragraph');
  const plan=planImport({file,guide,edits:a.styleEdits(),picks:[{fileId:'Heading1',to:'Heading1'}]});
  assert.deepEqual(plan.changed,['Heading 1']);
  const o=plan.edits.overrides.Heading1;
  assert.equal(o.font.name,'Calibri Light');
  assert.equal(o.font.size,16);
  assert.equal(o.font.color,'#2F5496');
  assert.equal(o.paragraph.spaceBefore,12);
  await a.setStyleEdits(plan.edits,{group:'import'});
  const now=a.listStyles().find(s=>s.id==='Heading1');
  assert.equal(now.effective.font.name,'Calibri Light');
  assert.equal(now.effective.font.size,16);
  assert.match(importSummary(plan),/1 style changed \(Heading 1\)\. Ctrl\+Z takes it back\./);
  // The same styles again change nothing.
  const again=planImport({file,guide:a.listStyles().filter(s=>s.type==='paragraph'),edits:a.styleEdits(),picks:[{fileId:'Heading1',to:'Heading1'}]});
  assert.deepEqual(again.changed,[]);
  assert.deepEqual(again.unchanged,['Heading 1']);
  assert.match(importSummary(again),/already look the way the file has them/);
});

test('a style of the file can be added as one of the author\'s own, which starts from the ordinary text and holds only what differs from it',async()=>{
  const a=demo();
  const file=await readTemplate();
  const guide=a.listStyles().filter(s=>s.type==='paragraph');
  const plan=planImport({file,guide,edits:a.styleEdits(),picks:[{fileId:'CompanyNote',to:AS_NEW}]});
  assert.deepEqual(plan.added,['Company Note']);
  const made=plan.edits.custom[0];
  assert.equal(made.name,'Company Note');
  assert.equal(made.basedOn,'Normal');
  assert.equal(made.spec.shading,'#FFF2CC');
  assert.equal(made.spec.font.italic,true);
  assert.equal(made.spec.paragraph.alignment,'center');
  assert.equal(made.spec.paragraph.leftIndent,18);
  assert.deepEqual(made.spec.borders,[{location:'left',color:'#C00000',width:3}]);
  await a.setStyleEdits(plan.edits);
  assert.equal(a.listStyles().find(s=>s.name==='Company Note').custom,true,'it is one of the author\'s own');
  assert.match(importSummary(plan),/1 style added as your own \(Company Note\)/);
});

test('a name the guide already has gets a number when a style is added under it, so no two styles share a name',async()=>{
  const a=demo();
  const file=await readTemplate();
  const guide=a.listStyles().filter(s=>s.type==='paragraph');
  const plan=planImport({file,guide,edits:a.styleEdits(),picks:[{fileId:'Normal',to:AS_NEW},{fileId:'CompanyNote',to:AS_NEW}]});
  assert.deepEqual(plan.added,['Normal (2)','Company Note']);
  await a.setStyleEdits(plan.edits);
  const second=planImport({file,guide:a.listStyles().filter(s=>s.type==='paragraph'),edits:a.styleEdits(),picks:[{fileId:'CompanyNote',to:AS_NEW}]});
  assert.deepEqual(second.added,['Company Note (2)']);
});

test('a style of the author\'s own can be the one a file style goes into; the change is kept in that style',async()=>{
  const a=demo();
  await a.setStyleEdits({overrides:{},custom:[{id:'Callout',name:'Callout',type:'paragraph',basedOn:'Normal',spec:{font:{bold:true}}}]});
  const file=await readTemplate();
  const plan=planImport({file,guide:a.listStyles().filter(s=>s.type==='paragraph'),edits:a.styleEdits(),picks:[{fileId:'CompanyNote',to:'Callout'}]});
  assert.deepEqual(plan.changed,['Callout']);
  assert.equal(plan.edits.custom[0].spec.shading,'#FFF2CC');
  assert.equal(plan.edits.custom[0].spec.font.bold,false,'the file\'s style is not bold, so the style is no longer');
  assert.equal(plan.edits.custom[0].spec.font.italic,true);
  assert.deepEqual(plan.edits.overrides,{});
});

test('a file style that says nothing about bold takes bold away from the guide\'s style it replaces',async()=>{
  const a=demo();
  const file=await readTemplate();
  const target=a.listStyles().find(s=>s.id==='Heading1');
  assert.equal(target.effective.font.bold,true,'the guide\'s heading is bold');
  const change=changeBetween(target.effective,fitSpec(byName(file,'heading 1').effective).spec);
  assert.equal(change.font.bold,false);
});

test('a pick that names a style the guide does not have is refused, and picks of nothing are fine',async()=>{
  const a=demo();
  const file=await readTemplate();
  const guide=a.listStyles().filter(s=>s.type==='paragraph');
  assert.throws(()=>planImport({file,guide,edits:a.styleEdits(),picks:[{fileId:'Heading1',to:'NoSuchStyle'}]}),/no style to put Heading 1 in/);
  const none=planImport({file,guide,edits:a.styleEdits(),picks:[]});
  assert.deepEqual(none.edits,{overrides:{},custom:[]});
  assert.equal(importSummary(none),'No style was chosen.');
});

test('too many styles of the author\'s own are refused in words, and the guide is left as it was',async()=>{
  const a=demo();
  const file=await readTemplate();
  const custom=Array.from({length:60},(_,i)=>({id:`Mine${i}`,name:`Mine ${i}`,type:'paragraph',basedOn:'Normal',spec:{}}));
  await a.setStyleEdits({overrides:{},custom});
  assert.throws(()=>planImport({file,guide:a.listStyles().filter(s=>s.type==='paragraph'),edits:a.styleEdits(),picks:[{fileId:'CompanyNote',to:AS_NEW}]}),/up to 60 styles of your own/);
  assert.equal(a.styleEdits().custom.length,60);
});

test('only the first styles of a very large file are offered',()=>{
  const many=Array.from({length:IMPORT_LIMITS.styles+40},(_,i)=>({id:`S${i}`,name:`Style ${i}`,type:'paragraph',effective:{font:{},paragraph:{}}}));
  assert.equal(candidates(many,[]).length,IMPORT_LIMITS.styles);
});

test('one Undo takes back everything a style import did',async()=>{
  const a=demo();
  const file=await readTemplate();
  const guide=a.listStyles().filter(s=>s.type==='paragraph');
  const plan=planImport({file,guide,edits:a.styleEdits(),picks:[{fileId:'Heading1',to:'Heading1'},{fileId:'Normal',to:'Normal'},{fileId:'CompanyNote',to:AS_NEW}]});
  await a.setStyleEdits(plan.edits,{group:'import:1'});
  assert.ok(Object.keys(a.styleEdits().overrides).length>=2&&a.styleEdits().custom.length===1);
  await a.undo();
  assert.deepEqual(a.styleEdits(),{overrides:{},custom:[]});
});
