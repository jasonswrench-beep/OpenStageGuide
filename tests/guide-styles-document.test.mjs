// The guide's styles in the program's document: edits are kept with the guide, laid over the program's definitions in what Word is given and
// what the Review reads, undone with Undo, saved in the project, and a look applied through the Appearance tool becomes edits of the styles.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {parseProject,serializeProject} from '../src/guide-project.mjs';
import {withStyle,readParagraph} from '../src/guide-edit.mjs';
import {lookSpecs,resolveGuideLook} from '../src/guide-looks.mjs';
import {para,run} from '../src/guide-ooxml.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const stylesIn=pkg=>/pkg:name="\/word\/styles.xml"[\s\S]*?<pkg:xmlData>([\s\S]*?)<\/pkg:xmlData>/.exec(pkg)[1];
const style=(a,id)=>a.listStyles().find(s=>s.id===id);

test('a change to a style is kept with the guide and is in the styles of the Word file and of the Review, which the style\'s paragraphs then follow',async()=>{
  const a=demo();
  await a.insert(()=>para('OSGCueBody',run('Body words')),{where:'start',spare:false});
  assert.deepEqual(a.styleEdits(),{overrides:{},custom:[]});
  await a.setStyleEdits({overrides:{OSGCueBody:{font:{name:'Georgia',size:14,color:'#112233'},paragraph:{alignment:'center'}}},custom:[]});
  assert.deepEqual(style(a,'OSGCueBody').own.font,{name:'Georgia',size:14,color:'#112233'});
  assert.match(stylesIn(await a.sourcePackage()),/w:ascii="Georgia"/,'the Word file is given the style as changed');
  const {stylesRoot}=await a.reviewParts();
  assert.match(io.serialize(stylesRoot),/w:ascii="Georgia"/,'and so is the Review');
  assert.equal(a.exportState().styles.overrides.OSGCueBody.font.size,14,'and the project');
});

test('a style of the author\'s own can be given to a paragraph, which then reads with its name; the style stays until no paragraph uses it',async()=>{
  const a=demo();
  await a.insert(()=>para('OSGCueBody',run('Plain words')),{where:'start',spare:false});
  await a.setStyleEdits({overrides:{},custom:[{id:'Callout',name:'Callout',type:'paragraph',basedOn:'OSGCueBody',spec:{shading:'#FFF4CE',font:{italic:true}}}]});
  assert.equal(style(a,'Callout').name,'Callout');
  const p=a.exportState().paragraphs[0];
  await a.replaceParagraphs({start:0,count:1,guard:p.text,xmls:[withStyle(p.xml,p.style,'Callout',io)]});
  assert.equal(a.exportState().paragraphs[0].style,'Callout','the paragraph reads with the style\'s name');
  assert.equal(readParagraph(a.exportState().paragraphs[0].xml,'Callout',io).text,'Plain words');
  assert.deepEqual(a.styleUse(),{Callout:1});
  await assert.rejects(a.setStyleEdits({overrides:{},custom:[]}),/Callout is used by 1 paragraph, so it stays/);
  assert.equal(a.styleEdits().custom.length,1,'nothing was taken away');
  const q=a.exportState().paragraphs[0];
  await a.replaceParagraphs({start:0,count:1,guard:q.text,xmls:[withStyle(q.xml,q.style,'OSGCueBody',io)]});
  await a.setStyleEdits({overrides:{},custom:[]});
  assert.equal(a.styleEdits().custom.length,0,'with no paragraph using it, it can go');
});

test('a style that is not in the right form is refused and the styles are as they were',async()=>{
  const a=demo();
  await a.setStyleEdits({overrides:{OSGCueBody:{font:{size:13}}},custom:[]});
  for(const bad of [{overrides:{OSGCueBody:{font:{size:1}}},custom:[]},{overrides:{},custom:[{id:'X',name:'X',type:'paragraph',basedOn:'Nothing',spec:{}}]},'nope',{overrides:{NoSuchStyle:{font:{size:12}}},custom:'x'}])
    await assert.rejects(a.setStyleEdits(bad),/./,JSON.stringify(bad));
  assert.equal(a.styleEdits().overrides.OSGCueBody.font.size,13);
});

test('Undo takes a change to styles back, and Redo brings it again, with the paragraphs: one history',async()=>{
  const a=demo();
  await a.insert(()=>para('OSGCueBody',run('x')),{where:'start',spare:false});
  await a.setStyleEdits({overrides:{OSGCueBody:{font:{size:20}}},custom:[]});
  assert.equal(style(a,'OSGCueBody').own.font.size,20);
  a.undo();
  assert.notEqual(style(a,'OSGCueBody').own.font?.size,20,'the style is as it was');
  assert.deepEqual(a.styleEdits().overrides,{});
  a.redo();
  assert.equal(style(a,'OSGCueBody').own.font.size,20);
  await a.setStyleEdits({overrides:{OSGCueBody:{font:{size:21}}},custom:[]},{group:'panel'});
  await a.setStyleEdits({overrides:{OSGCueBody:{font:{size:22}}},custom:[]},{group:'panel'});
  const steps=a.historyState().undo;
  a.undo();
  assert.equal(a.historyState().undo,steps-1);
  assert.equal(style(a,'OSGCueBody').own.font.size,20,'two edits in a group are one step');
});

test('the styles are saved in the project and come back with it; a project whose styles are bad is not opened',async()=>{
  const a=demo();
  await a.setStyleEdits({overrides:{OSGNotes:{shading:'#EEEEEE'}},custom:[{id:'Callout',name:'Callout',type:'paragraph',basedOn:'Normal',spec:{font:{bold:true}}}]});
  const text=serializeProject(a.exportState());
  const b=demo();
  b.importState(parseProject(text).guide);
  assert.equal(style(b,'OSGNotes').own.shading,'#EEEEEE');
  assert.equal(style(b,'Callout').name,'Callout');
  assert.equal(b.historyState().undo,0,'a project that is opened starts with no history');
  assert.throws(()=>b.importState({...parseProject(text).guide,styles:{overrides:{OSGNotes:{font:{size:1}}},custom:[]}}),/./);
});

test('a look applied through the Appearance tool becomes edits of the styles, with the look\'s fonts, sizes, colours, spacing, shading and borders, as one step of Undo',async()=>{
  const a=demo();
  const r=await a.applyLook([{name:'OSG Cue Body',font:{name:'Aptos',size:11,color:'#112233',bold:false},paragraph:{alignment:'Centered',spaceAfter:6,lineSpacing:14.4,keepWithNext:true},shading:'#EEEEEE',borders:[{location:'Left',color:'#112233',width:'Pt300'}]},{name:'No Such Style',font:{size:9}}]);
  assert.deepEqual(r.applied,['OSG Cue Body']);assert.deepEqual(r.skipped,['No Such Style']);assert.deepEqual(r.refused,[]);
  const own=style(a,'OSGCueBody').own;
  assert.deepEqual(own.font,{name:'Aptos',size:11,color:'#112233',bold:false});
  assert.deepEqual(own.paragraph,{alignment:'center',spaceAfter:6,lineSpacing:1.2,keepWithNext:true,leftIndent:120});
  assert.equal(own.shading,'#EEEEEE');
  assert.deepEqual(own.borders,[{location:'left',color:'#112233',width:3}]);
  a.undo();
  assert.deepEqual(a.styleEdits().overrides,{},'the whole look is one step');
});

test('the program\'s own look specs apply as they are: every style the look names that the guide has is changed, and the look\'s heading font reaches the headings',async()=>{
  const a=demo();
  const resolved=resolveGuideLook({id:'clearline'});
  assert.equal(resolved.ok,true);
  const specs=lookSpecs(resolved.theme,DOMParser);
  const r=await a.applyLook(specs);
  assert.ok(r.applied.length>=15,'many styles: '+r.applied.length);
  assert.deepEqual(r.refused,[]);
  const heading=style(a,'OSGModule')?.effective.font||{};
  assert.equal(heading.name,resolved.theme.headingFont||resolved.theme.bodyFont,'the module heading has the look\'s font');
  const lookState=a.lookState();
  assert.equal(lookState.applied.length,1,'the look is also what the Appearance tool recorded');
});

test('a character style of the author\'s own is in the Word file, counts the paragraphs whose words use it, and stays while any does',async()=>{
  const a=demo();
  await a.setStyleEdits({overrides:{},custom:[{id:'Warning',name:'Warning words',type:'character',basedOn:'DefaultParagraphFont',spec:{font:{bold:true,color:'#C00000'}}}]});
  assert.match(stylesIn(await a.sourcePackage()),/<w:style w:type="character" w:customStyle="1" w:styleId="Warning">/);
  assert.doesNotMatch(stylesIn(await a.sourcePackage()).match(/<w:style w:type="character"[^>]*w:styleId="Warning"[\s\S]*?<\/w:style>/)[0],/<w:next|<w:pPr>/,'a character style has no next style and no paragraph settings');
  assert.equal(style(a,'Warning').type,'character');
  assert.equal(style(a,'Warning').custom,true);
  const rich=`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t xml:space="preserve">Careful: </w:t></w:r><w:r><w:rPr><w:rStyle w:val="Warning"/></w:rPr><w:t>hot</w:t></w:r></w:p>`;
  await a.insert(()=>rich,{where:'start',spare:false});
  assert.deepEqual(a.styleUse(),{Warning:1,...(a.styleUse().Normal?{Normal:a.styleUse().Normal}:{})},'the paragraph counts once, for the style its words use');
  await assert.rejects(a.setStyleEdits({overrides:{},custom:[]}),/Warning words is used by 1 paragraph, so it stays/);
  assert.equal(a.styleEdits().custom.length,1,'nothing was taken away');
});
