// Forms, slices F4 and F5: card sets to cut apart, and the self-scored instrument.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {updatePlan,guideIssues,kindOf} from '../src/guide-model.mjs';
import {readForms} from '../src/guide-forms.mjs';
import {formTitleBlock,cardSet,scoreInstrument,formText,FORM_LIMITS} from '../src/guide-forms-ooxml.mjs';
import {readSource,buildWorkbook,classify} from '../src/guide-workbook.mjs';
import {leakCheck} from '../src/workbook-leak.mjs';
import {STYLE_NAMES,STYLE_IDS,NS,stylesPart} from '../src/guide-styles.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const strict=new DOMParser({errorHandler:{warning(){},error(e){throw new Error(e);},fatalError(e){throw new Error(e);}}});
const parse=body=>strict.parseFromString(`<w:document ${NS}><w:body>${body}</w:body></w:document>`,'application/xml');
const all=(doc,name)=>Array.from(doc.getElementsByTagNameNS(W,name));
const texts=doc=>all(doc,'tc').map(c=>Array.from(c.getElementsByTagNameNS(W,'t')).map(t=>t.textContent).join(''));

// ---- F5: cards
test('a set of cards is one bordered paragraph for each card, with a new page after every N, and the label above',()=>{
  const x=cardSet({cards:['Role A\nYou are the manager.','Role B\nYou are the report.','Role C','Role D','Role E'],perPage:2,label:'Cut these apart.'});
  const doc=parse(x);
  const cards=all(doc,'p').filter(p=>Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val')===STYLE_IDS.formCard);
  assert.equal(cards.length,5);
  assert.equal(all(doc,'br').length,2,'a card\'s lines are separated by line breaks, inside the card');
  assert.equal(all(doc,'pageBreakBefore').length,2,'a page break before the third and the fifth card when 2 go to a page');
  assert.ok(x.indexOf('Cut these apart.')<x.indexOf(STYLE_IDS.formCard));
  assert.equal(all(parse(cardSet({cards:['a','b','c','d','e'],perPage:4})),'pageBreakBefore').length,1);
  assert.equal(all(parse(cardSet({cards:['a','b']})),'pageBreakBefore').length,0);
  assert.throws(()=>cardSet({cards:[]}),/at least one card/);
  assert.throws(()=>cardSet({cards:['a'],perPage:3}),/2, 4, 6 or 8/);
  assert.throws(()=>cardSet({cards:Array.from({length:61},(_,i)=>`c${i}`)}),/up to 60 cards/);
  assert.throws(()=>cardSet({cards:['x'.repeat(301)]}),/up to 300 letters/);
});

test('the card style has a dashed border and a dashed line between cards, which is the line to cut along',()=>{
  const part=stylesPart();
  const def=part.slice(part.indexOf(`w:styleId="${STYLE_IDS.formCard}"`),part.indexOf('</w:style>',part.indexOf(`w:styleId="${STYLE_IDS.formCard}"`)));
  assert.match(def,/<w:between w:val="dashed"/);
  assert.equal((def.match(/w:val="dashed"/g)||[]).length,5);
  assert.equal(kindOf({style:STYLE_NAMES.formCard}),'form-card');
});

// ---- F4: the instrument
const STATEMENTS=['I say what I think.','I listen first.','I keep my promises.','I admit mistakes.'];
const GROUPS=[{name:'Trust',items:[1,3]},{name:'Conflict',items:[2,4]}];

test('an instrument has a table to rate in, a grid to add up in, and a table of meanings, each with a header row, a caption and a description',()=>{
  const x=scoreInstrument({statements:STATEMENTS,groups:GROUPS,divide:true,scale:'1 to 5',bands:[{group:'Trust',from:2,to:5,meaning:'Low trust.'},{group:'Trust',from:6,to:10,meaning:'High trust.'},{group:'Conflict',from:2,to:10,meaning:'Some conflict.'}]});
  const doc=parse(x);
  const tables=all(doc,'tbl');
  assert.equal(tables.length,3);
  assert.equal(all(doc,'tblHeader').length,3);
  assert.deepEqual(all(doc,'tblCaption').map(c=>c.getAttributeNS(W,'val')),['Self-scored instrument','Self-scored instrument: adding up','Self-scored instrument: what the scores mean']);
  assert.ok(all(doc,'tblDescription').every(d=>d.getAttributeNS(W,'val').length>20));
  const cells=texts(doc);
  assert.ok(cells.includes('Score (1 to 5)'));
  assert.ok(cells.includes('1, 3'),'the grid lists the statements to add for each scale');
  assert.ok(cells.includes('Divide by')&&cells.includes('Average'));
  assert.ok(cells.includes('2 to 5')&&cells.includes('High trust.')&&cells.includes('Every scale')===false,'a range names its scale, so no "every scale" row');
  assert.equal(all(doc,'checkbox').length,0,'scores are written, not ticked');
  assert.doesNotThrow(()=>parse(x));
});

test('without division or meanings, the instrument is the two tables a person needs, and a range for every scale says so',()=>{
  const plain=scoreInstrument({statements:STATEMENTS,groups:GROUPS});
  assert.equal(all(parse(plain),'tbl').length,2);
  assert.ok(!texts(parse(plain)).includes('Divide by'));
  const shared=scoreInstrument({statements:STATEMENTS,groups:GROUPS,bands:[{from:0,to:4,meaning:'Low.'},{from:5,to:9,meaning:'High.'}]});
  const cells=texts(parse(shared));
  assert.equal(cells.filter(c=>c==='Scale').length,1,'a Scale column in the grid, and none in the meanings when no range names a scale');
  assert.ok(cells.includes('0 to 4'));
  const single=scoreInstrument({statements:STATEMENTS,groups:GROUPS,bands:[{from:3,to:3,meaning:'Exactly three.'}]});
  assert.ok(texts(parse(single)).includes('3'));
});

test('an instrument cannot be built with a statement that is not added up, added twice, or named wrongly',()=>{
  const bad=(over,pattern)=>assert.throws(()=>scoreInstrument({statements:STATEMENTS,groups:GROUPS,...over}),pattern);
  bad({groups:[{name:'Trust',items:[1,3]},{name:'Conflict',items:[2]}]},/Statement 4 is not in any scale/);
  bad({groups:[{name:'Trust',items:[1,3]},{name:'Conflict',items:[2]}]},/would not be added up/);
  bad({groups:[{name:'Trust',items:[1]}]},/Statements 2, 3, 4 are not in any scale/);
  bad({groups:[{name:'Trust',items:[1,2,3]},{name:'Conflict',items:[3,4]}]},/Statement 3 is in two scales: Trust and Conflict/);
  bad({groups:[{name:'Trust',items:[1,2,3,4,5]}]},/statement 5, which is not one of the 4/);
  bad({groups:[{name:'Trust',items:[1,2]},{name:'trust',items:[3,4]}]},/same name/);
  bad({groups:[{name:'',items:[1,2,3,4]}]},/needs a name/);
  bad({groups:[{name:'Trust',items:[]},{name:'Conflict',items:[1,2,3,4]}]},/has no statements/);
  bad({groups:[]},/1 to 8 scales/);
  bad({bands:[{from:5,to:2,meaning:'x'}]},/lower number to a higher/);
  bad({bands:[{from:1,to:2,meaning:' '}]},/words that say what it means/);
  bad({bands:[{group:'Nope',from:1,to:2,meaning:'x'}]},/not one of the scales/);
  assert.throws(()=>scoreInstrument({statements:[],groups:GROUPS}),/at least one statement/);
  assert.equal(FORM_LIMITS.groups,8);
});

// ---- both, in a guide, in the workbook, and in what the Review says about them
async function guideWithBoth(){
  const a=await started('Instruments');
  await addModule(a,'One');await addLesson(a,'First');await addBlock(a,'ask','What do you notice?');
  const at=a.adapter.snapshot().findIndex(p=>p.text.includes('What do you notice?'));
  await a.adapter.insert(()=>
    formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'How I work'})+formText('Rate yourself, then add up.')+scoreInstrument({statements:STATEMENTS,groups:GROUPS,divide:true,bands:[{from:2,to:10,meaning:'Keep going.'}]})
    +formTitleBlock({id:'f2',label:'Cards',number:'1.2',title:'Role cards'})+cardSet({cards:['Role A\nThe manager.','Role B\nThe report.'],perPage:2})
    +formTitleBlock({id:'f3',label:'Cards',number:'1.3',title:'Facilitator keys',facilitator:true})+cardSet({cards:['SECRETKEY one'],perPage:2}),
    {where:'after',after:at,spare:false});
  return a;
}

test('the tables of an instrument and the cards read as parts of their forms, so the Review finds nothing wrong with them',async()=>{
  const a=await guideWithBoth();
  const list=await a.adapter.paragraphs();
  const read=readForms(list);
  assert.deepEqual(read.forms.map(f=>[f.id,f.audience,f.parts>0,f.headerless.length]),[['f1','everyone',true,0],['f2','everyone',true,0],['f3','facilitator',true,0]]);
  assert.deepEqual(guideIssues(updatePlan(list)).filter(i=>/^FORM-/.test(i.rule)),[]);
});

test('the learner workbook carries the instrument and the cards whole, and leaves out the facilitator\'s cards',async()=>{
  const a=await guideWithBoth();
  const paras=await a.adapter.paragraphs();
  const kinds=classify(paras);
  assert.ok(paras.some((p,i)=>p.style==='OSG Form Card'&&kinds.decisions[i].keep));
  assert.ok(paras.some((p,i)=>p.style==='OSG Form Card'&&!kinds.decisions[i].keep&&kinds.decisions[i].reason==='facilitator-form'));
  const built=buildWorkbook(readSource(await a.adapter.sourcePackage(),{DOMParserClass:DOMParser}),{},{DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
  const doc=built.files.find(f=>f.name==='word/document.xml').text,styles=built.files.find(f=>f.name==='word/styles.xml').text;
  assert.ok(doc.includes('I listen first.')&&doc.includes('Keep going.')&&doc.includes('The manager.'));
  assert.equal((doc.match(/<w:tblHeader\b/g)||[]).length,3,'the instrument\'s three tables keep their header rows');
  assert.ok(!doc.includes('SECRETKEY'));
  assert.ok(styles.includes(STYLE_IDS.formCard),'the workbook defines the card style it uses');
  assert.ok(built.facilitatorTexts.some(t=>t.includes('SECRETKEY')));
  const result=leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:DOMParser});
  assert.equal(result.passed,true,JSON.stringify(result.findings.filter(f=>f.level!=='pass')));
});
