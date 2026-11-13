import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {kindOf} from '../src/guide-model.mjs';
import {STYLE_NAMES,STYLE_IDS,NS,stylesPart} from '../src/guide-styles.mjs';
import {makeContext} from '../src/guide-ooxml.mjs';
import {readForms} from '../src/guide-forms.mjs';
import {formTitleBlock,formText,writingLines,writingBox,fieldRow,sourceLine,checkList,formTable,ratingGrid,FORM_LIMITS} from '../src/guide-forms-ooxml.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',W14='http://schemas.microsoft.com/office/word/2010/wordml';
// Strict: a damaged piece of XML must fail the test, not be quietly repaired.
const strict=new DOMParser({errorHandler:{warning(){},error(e){throw new Error(e);},fatalError(e){throw new Error(e);}}});
const parse=body=>strict.parseFromString(`<w:document ${NS}><w:body>${body}</w:body></w:document>`,'application/xml');
const all=(doc,name)=>Array.from(doc.getElementsByTagNameNS(name==='checkbox'?W14:W,name));
const count=(body,re)=>(body.match(re)||[]).length;
const styleCount=(body,id)=>count(body,new RegExp(`<w:pStyle w:val="${id}"/>`,'g'));
const ctx=()=>makeContext();

test('every piece a form is made of is well-formed Word XML',()=>{
  const c=ctx();
  const pieces=[
    formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'}),formText('One line.\nAnother line.'),writingLines({label:'Your notes',count:3}),writingBox({label:'Draw it',lines:5}),
    fieldRow([{label:'Name',width:'full'},{label:'Date',width:'short'}]),sourceLine('Adapted with permission.'),
    checkList({items:['Book the room','Print the handouts'],columns:['Who','When'],ctx:c}),
    formTable({columns:[{name:'Goal'},{name:'Done',kind:'check'}],rows:[['Read']],blankRows:1,sample:['Example goal','' ],totals:{label:'Total',columns:[1]},ctx:c}),
    ratingGrid({statements:['I plan.','I check.'],evidence:true,ctx:c})
  ];
  for(const p of pieces)assert.doesNotThrow(()=>parse(p));
});

test('lines to write on are bordered paragraphs, one for each line, with the label above and no typed underscores',()=>{
  const x=writingLines({label:'Your notes',count:4});
  assert.equal(styleCount(x,STYLE_IDS.formLine),4);
  assert.equal(styleCount(x,STYLE_IDS.formText),1);
  assert.ok(x.indexOf('Your notes')<x.indexOf(STYLE_IDS.formLine));
  assert.doesNotMatch(x,/_{2,}/);
  assert.equal(styleCount(writingLines(),STYLE_IDS.formLine),3,'three by default');
  for(const bad of [0,13,2.5,'3'])assert.throws(()=>writingLines({count:bad}),/whole number from 1 to 12/);
  assert.throws(()=>writingLines({label:'x'.repeat(201)}),/longer than 200/);
});

test('a box to write in is one bordered paragraph of a set height',()=>{
  const x=writingBox({label:'Draw it',lines:5});
  assert.match(x,/<w:spacing w:before="120" w:after="120" w:line="1800" w:lineRule="exact"\/>/);
  assert.equal(styleCount(x,STYLE_IDS.formBox),1);
  assert.equal(styleCount(x,STYLE_IDS.formText),1);
  assert.throws(()=>writingBox({lines:0}),/whole number from 1 to 12/);
  assert.throws(()=>writingBox({lines:13}),/whole number from 1 to 12/);
});

test('fields on a row each end in a line made by a tab leader, and the last one reaches the right margin',()=>{
  const x=fieldRow([{label:'Name',width:'full'},{label:'Date',width:'short'},{label:'Signature',width:'bogus'}]);
  const stops=all(parse(x),'tab').filter(t=>t.getAttributeNS(W,'leader')==='underscore');
  assert.equal(stops.length,3);
  assert.equal(Number(stops.at(-1).getAttributeNS(W,'pos')),9360);
  const positions=all(parse(x),'tab').filter(t=>t.parentNode.localName==='tabs').map(t=>Number(t.getAttributeNS(W,'pos')));
  assert.deepEqual([...positions].sort((a,b)=>a-b),positions,'tab stops run left to right');
  assert.doesNotMatch(x,/_{2,}/);
  assert.match(x,/Name/);assert.match(x,/Signature/);
  assert.throws(()=>fieldRow([]),/1 to 4 fields/);
  assert.throws(()=>fieldRow(Array.from({length:5},(_,i)=>({label:`F${i}`}))),/1 to 4 fields/);
  assert.throws(()=>fieldRow([{label:' '}]),/needs a label/);
});

test('a checklist has a header row Word knows about, a real check box for each item, and the extra columns to write in',()=>{
  const c=ctx();
  const x=checkList({items:['Book the room','Print the handouts','Test the projector'],columns:['Who','When','Notes'],ctx:c});
  const doc=parse(x);
  assert.equal(all(doc,'tblHeader').length,1);
  assert.equal(all(doc,'cantSplit').length,4,'a row is never split across pages');
  assert.equal(all(doc,'checkbox').length,3);
  assert.equal(all(doc,'tr').length,4);
  assert.equal(all(doc,'gridCol').length,5);
  assert.match(x,/w:alias w:val="Book the room: done"/);
  assert.match(x,/<w:tblCaption w:val="Checklist"\/>/);
  assert.match(x,/space for Who, When, Notes/);
  assert.equal(c.sdt,3,'each check box has a number of its own');
  const widths=all(doc,'gridCol').map(g=>Number(g.getAttributeNS(W,'w')));
  assert.equal(widths.reduce((a,b)=>a+b,0),9360);
  assert.throws(()=>checkList({items:[],ctx:c}),/at least one item/);
  assert.throws(()=>checkList({items:['a'],columns:['1','2','3','4'],ctx:c}),/up to 3 extra columns/);
  assert.throws(()=>checkList({items:Array.from({length:61},(_,i)=>`i${i}`),ctx:c}),/up to 60 items/);
});

test('a table to fill in has named columns, a marked header row, an example row, blank rows, check boxes where asked, and a row for totals',()=>{
  const c=ctx();
  const x=formTable({columns:[{name:'Activity',weight:3},{name:'Person',kind:'blank'},{name:'Days',kind:'number'},{name:'Done',kind:'check'}],rows:[['Book the room','Sam'],['Print']],blankRows:2,sample:['Order lunch','Pat','2'],totals:{label:'Total',columns:[2]},ctx:c,caption:'Action plan'});
  const doc=parse(x);
  assert.equal(all(doc,'tblHeader').length,1);
  assert.equal(all(doc,'tr').length,1+1+2+2+1,'header, example, two filled rows, two blank rows, totals');
  assert.equal(all(doc,'checkbox').length,4,'a check box in each of the four body rows that are not the example or totals');
  assert.match(x,/Example: Order lunch/);
  assert.match(x,/w:alias w:val="Done: row 2"/);
  assert.match(x,/<w:tblCaption w:val="Action plan"\/>/);
  assert.match(x,/columns: Activity, Person, Days, Done\. The last row is for totals\./);
  const widths=all(doc,'gridCol').map(g=>Number(g.getAttributeNS(W,'w')));
  assert.equal(widths.reduce((a,b)=>a+b,0),9360);
  assert.ok(widths[0]>widths[1],'a heavier column is wider');
  assert.throws(()=>formTable({columns:[],ctx:c}),/1 to 8 columns/);
  assert.throws(()=>formTable({columns:[{name:'A'},{name:' '}],ctx:c}),/needs a name/);
  assert.throws(()=>formTable({columns:Array.from({length:9},(_,i)=>({name:`c${i}`})),ctx:c}),/1 to 8 columns/);
  assert.throws(()=>formTable({columns:[{name:'A'}],blankRows:41,ctx:c}),/from 0 to 40/);
  assert.throws(()=>formTable({columns:[{name:'A'}],rows:Array.from({length:40},()=>['x']),blankRows:1,ctx:c}),/up to 40 rows/);
});

test('a rating grid has a check box for every statement at every point, and an evidence column when asked',()=>{
  const c=ctx();
  const x=ratingGrid({statements:['I plan ahead.','I check my work.'],labels:['1 (low)','2','3','4','5 (high)'],evidence:true,ctx:c});
  const doc=parse(x);
  assert.equal(all(doc,'checkbox').length,10);
  assert.equal(all(doc,'tblHeader').length,1);
  assert.equal(all(doc,'gridCol').length,1+5+1);
  assert.match(x,/w:alias w:val="I plan ahead\.: 1 \(low\)"/);
  assert.match(x,/Evidence/);
  assert.match(x,/from 1 \(low\) to 5 \(high\)/);
  const plain=ratingGrid({statements:['A'],ctx:c});
  assert.equal(all(parse(plain),'gridCol').length,1+5,'no evidence column unless asked');
  assert.throws(()=>ratingGrid({statements:[],ctx:c}),/at least one statement/);
  assert.throws(()=>ratingGrid({statements:['a'],labels:['1'],ctx:c}),/2 to 7 points/);
  assert.throws(()=>ratingGrid({statements:['a'],labels:Array.from({length:8},(_,i)=>`${i}`),ctx:c}),/2 to 7 points/);
});

test('text and credit lines are paragraphs in their own styles, and need words',()=>{
  assert.equal(styleCount(formText('One\nTwo\n\nThree'),STYLE_IDS.formText),3);
  assert.equal(styleCount(sourceLine('Adapted with permission.'),STYLE_IDS.formSource),1);
  assert.throws(()=>formText('  \n '),/Type the words/);
  assert.throws(()=>sourceLine(''),/credit line/);
});

test('a form for the facilitator has its own title style, and is still a form: it is numbered with the others',()=>{
  const facilitator=formTitleBlock({id:'f2',label:'Checklist',number:'',title:'Planning',facilitator:true});
  assert.match(facilitator,new RegExp(`<w:pStyle w:val="${STYLE_IDS.formTitleFacilitator}"/>`));
  assert.equal(kindOf({style:STYLE_NAMES.formTitleFacilitator}),'form-title');
  const read=readForms([{style:STYLE_NAMES.formTitle,text:'Handout: A',marks:{bookmarks:['osg_form_f1']}},{style:STYLE_NAMES.formTitleFacilitator,text:'Checklist: B',marks:{bookmarks:['osg_form_f2']}}]);
  assert.deepEqual(read.forms.map(f=>f.audience),['everyone','facilitator']);
});

test('the styles part defines every style the elements use, and the model knows each by its kind',()=>{
  const part=stylesPart();
  for(const key of ['formTitle','formTitleFacilitator','formText','formLine','formField','formBox','formSource']){
    assert.ok(part.includes(`w:styleId="${STYLE_IDS[key]}"`),`${key} has a definition`);
    assert.ok(part.includes(`<w:name w:val="${STYLE_NAMES[key]}"/>`),`${key} has its name`);
  }
  assert.deepEqual(['formText','formLine','formField','formBox','formSource'].map(k=>kindOf({style:STYLE_NAMES[k]})),['form-text','form-line','form-field','form-box','form-source']);
});

test('put into a document, the elements come back in their styles, one paragraph for each, with a table read as its cells',async()=>{
  const adapter=createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const c=ctx();
  await adapter.insert(()=>formTitleBlock({id:'f1',label:'Handout',number:'A',title:'Goals'})+formText('Fill this in.')+fieldRow([{label:'Name'}])+writingLines({count:2})+writingBox({lines:3})
    +checkList({items:['One','Two'],ctx:c})+sourceLine('Credit.'),{where:'start',spare:false});
  const styles=adapter.snapshot().map(p=>p.style);
  assert.deepEqual(styles.slice(0,8),['OSG Form Title','OSG Form Text','OSG Form Field','OSG Form Line','OSG Form Line','OSG Form Box','OSG Table Head','OSG Table Head']);
  assert.equal(styles.at(-2),'OSG Form Source','the credit line is last; the document ends in the plain empty paragraph an insert always leaves');
  assert.equal(styles.at(-1),'Normal');
  assert.equal(styles.filter(s=>s==='OSG Rating Box').length,2);
  const forms=readForms(await adapter.paragraphs());
  assert.deepEqual(forms.forms.map(f=>[f.id,f.title,f.audience]),[['f1','Goals','everyone']]);
});

test('limits are written down',()=>{
  assert.deepEqual(Object.keys(FORM_LIMITS).sort(),['bands','boxLines','cardWords','cards','cell','columns','extraColumns','fields','groups','items','label','lines','rows','statements']);
});
