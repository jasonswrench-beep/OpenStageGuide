// Importing questions from an OpenAssess bank file (.openassess): a zip of manifest.json, questions/q_*.json and folders.json, read by the
// shared core's bank reader (src/openstage-bank). Proven against stand-ins for Word, banks written the way OpenAssess's model describes, and
// one bank OpenAssess's own code wrote.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import JSZip from 'jszip';
import {questionToItem,readBankFile,MAX_BANK_ADD} from '../src/guide-bank.mjs';
import {readQuestion} from '../src/openstage-bank/index.mjs';
import {readItem,readAssessments} from '../src/guide-model.mjs';
import {started,WORD_QUIRKS,click,submit,addModule,addLesson} from './helpers/guide-harness.mjs';

const rt=html=>({format:'rich',content:html});
const opt=(id,text,correct=false)=>({id,content:rt(`<p>${text}</p>`),correct});
let counter=0;
const Q=(type,body,o={})=>{
  const n=++counter,id=`q_${String(n).padStart(26,'0')}`;
  return {id,version:1,type,title:o.title??`Question ${n}`,stem:rt(o.stem??`<p>Stem ${n}?</p>`),points:o.points??1,body,feedback:o.feedback??{},metadata:{tags:[],status:'draft',...(o.metadata||{})},media:o.media??[],math:o.math??[],findings:[],history:[]};
};
async function bankBytes(questions,{manifest={},folders}={}){
  const zip=new JSZip();
  zip.file('manifest.json',JSON.stringify({format:'openassess-bank',formatVersion:1,bankId:'b_1',title:'Light and radio',created:'x',modified:'x',questionCount:questions.length,generator:'test',...manifest}));
  for(const q of questions)zip.file(`questions/${q.id}.json`,typeof q==='string'?q:JSON.stringify(q));
  if(folders)zip.file('folders.json',JSON.stringify(folders));
  return zip.generateAsync({type:'arraybuffer'});
}
const MC=(o={})=>Q('multipleChoice',{shuffle:true,options:[opt('a','Einstein'),opt('b','Michelson',true),opt('c','Edison')]},o);

// ---- each type (a question file as the bank reader gives it)
const read=raw=>{const r=readQuestion(raw);assert.ok(r.question,r.problem);return r.question;};
const toItem=raw=>questionToItem(read(raw));
const item=q=>{const r=toItem(q);assert.ok(r.item,JSON.stringify(r.problem));return r.item;};
const problem=q=>{const r=toItem(q);assert.ok(r.problem,'expected a problem');return r.problem.message;};

test('the words come as the reader gives them: the stem keeps its lines, an option or a reason is one line, entities decoded, tags dropped',()=>{
  const mc=item(MC({stem:'<p>One &amp; two</p><p>Three<br>four</p><ul><li>a</li><li>b</li></ul>',feedback:{correct:rt('<p><strong>Bold</strong> and <em>em</em></p><p>&lt;b&gt; &#65; &#x42;</p>')}}));
  assert.equal(mc.stem,'One & two\nThree\nfour\na\nb');
  assert.equal(mc.why,'Bold and em <b> A B');
  assert.equal(item({...Q('trueFalse',{correct:true}),stem:{format:'plain',content:'Line one\n\nLine two'}}).stem,'Line one\nLine two');
  assert.equal(item(Q('trueFalse',{correct:true},{stem:'<table><tr><td>x</td><td>y</td></tr></table>'})).stem,'x y');
});

test('multiple choice, multiple response and true or false carry their options and keys, points, objective and why',()=>{
  const mc=item(MC({points:2,metadata:{objective:'Name the scientist.'},feedback:{correct:rt('<p>He measured it in 1887.</p>')}}));
  assert.deepEqual([mc.type,mc.options,mc.correct,mc.points,mc.objective,mc.why],['multiple-choice',['Einstein','Michelson','Edison'],[1],'2','Name the scientist.','He measured it in 1887.']);
  const mr=item(Q('multipleResponse',{shuffle:true,partialCredit:false,options:[opt('a','Neon',true),opt('b','Argon',true),opt('c','Oxygen')]}));
  assert.deepEqual([mr.type,mr.correct],['multiple-response',[0,1]]);
  assert.equal(item(Q('trueFalse',{correct:false})).truth,false);
  assert.equal(item(Q('trueFalse',{correct:true})).truth,true);
});

test('fill in the blank, short answer, ordering and matching carry what the key needs',()=>{
  assert.deepEqual([item(Q('fillInBlank',{caseSensitive:false,blanks:[{id:'x',accepted:['Shakespeare','William Shakespeare']}]})).answer],['Shakespeare; William Shakespeare']);
  assert.equal(item(Q('shortAnswer',{caseSensitive:false,accepted:['Na']})).answer,'Na');
  const ord=item(Q('ordering',{items:[{id:'3',content:rt('Third'),position:3},{id:'1',content:rt('First'),position:1},{id:'2',content:rt('Second'),position:2}]}));
  assert.deepEqual(ord.orderItems,['First','Second','Third']);
  const m=item(Q('matching',{shuffle:true,distractors:['Telephone'],pairs:[{id:'p1',left:rt('Marconi'),right:'Radio'},{id:'p2',left:rt('Edison'),right:'Light bulb'}]}));
  assert.deepEqual([m.pairs,m.extras],[[{item:'Marconi',match:'Radio'},{item:'Edison',match:'Light bulb'}],['Telephone']]);
});

test('numeric: exact, tolerance and range, with units; a range is written as its middle and half its width, and says so',()=>{
  const exact=item(Q('numeric',{mode:'exact',value:60,units:'km per hour'}));
  assert.deepEqual([exact.answer,exact.tolerance,exact.units],['60','','km per hour']);
  const tol=item(Q('numeric',{mode:'tolerance',value:60.5,tolerance:0.5}));
  assert.deepEqual([tol.answer,tol.tolerance],['60.5','0.5']);
  const r=toItem(Q('numeric',{mode:'range',min:59,max:61.5},{title:'Speed'}));
  assert.deepEqual([r.item.answer,r.item.tolerance],['60.25','1.25']);
  assert.match(r.notes[0],/Speed: the range 59 to 61\.5 is written as 60\.25 plus or minus 1\.25/);
});

test('an essay becomes a reflection: its model answer and rubric in the key, or a placeholder that is reported',()=>{
  const full=item(Q('essay',{modelAnswer:rt('<p>A good answer names the ether wind.</p>'),rubric:[{criterion:'Names the ether wind',points:2,description:'Says what it was'},{criterion:'Clear',points:1}]}));
  assert.equal(full.type,'reflection');
  assert.equal(full.answer,'A good answer names the ether wind.\nRubric:\nNames the ether wind: Says what it was (2 points)\nClear (1 point)');
  const bare=toItem(Q('essay',{},{title:'Discuss'}));
  assert.equal(bare.item.answer,'Add a model answer, or what to look for.');
  assert.match(bare.notes[0],/Discuss: an essay with no model answer gets a placeholder/);
});

test('what the guide cannot hold is refused with the reason, and nothing is guessed',()=>{
  assert.match(problem(MC({stem:'<p>Look: <oa-media ref="m_1"></oa-media></p>'})),/picture or an equation/);
  assert.match(problem(MC({math:[{id:'e1',latex:'x',mathml:'x'}]})),/picture or an equation/);
  assert.match(problem(Q('hotspot',{mediaRef:'m',regions:[],alternative:rt('x')})),/The type "hotspot" is not read here/);
  assert.match(problem(Q('calculated',{formula:'a+b',variables:[],tolerance:1,toleranceIsPercent:false})),/"calculated" is not read here/);
  assert.match(problem(Q('fillInBlank',{caseSensitive:false,blanks:[{id:'1',accepted:['a']},{id:'2',accepted:['b']}]})),/2 blanks, and the guide's fill in the blank has one/);
  assert.match(problem(Q('trueFalse',{})),/No answer is marked/);
  assert.match(problem(Q('multipleChoice',{shuffle:true,options:[opt('a','One'),opt('b','Two')]})),/No correct option is marked/);
  assert.match(problem(Q('multipleChoice',{shuffle:true,options:[opt('a','One',true),opt('b','Two',true)]})),/More than one option is marked correct/);
  assert.match(problem(Q('multipleChoice',{shuffle:true,options:Array.from({length:7},(_,i)=>opt(`o${i}`,`Option ${i}`,i===0))})),/7 options and a guide question can have up to 6/);
  assert.match(problem(Q('matching',{shuffle:true,distractors:[],pairs:Array.from({length:11},(_,i)=>({id:`p${i}`,left:rt(`L${i}`),right:`R${i}`}))})),/11 pairs and a guide matching question can have up to 10/);
  assert.match(problem(Q('ordering',{items:[{id:'1',content:rt('A'),position:1},{id:'2',content:rt('B'),position:2}]})),/2 items and a guide ordering question has 3 to 8/);
  assert.match(questionToItem(null).problem.message,/could not be read/);
});

// ---- the file
test('a file that is not a bank is refused with the reason, in the guide\'s words',async()=>{
  await assert.rejects(readBankFile(new TextEncoder().encode('not a zip').buffer),/^Error: This file could not be opened as an OpenAssess bank\. It may be damaged, or it may not be a bank\.$/);
  await assert.rejects(readBankFile(new ArrayBuffer(0)),/could not be opened as an OpenAssess bank/);
  const noManifest=new JSZip();noManifest.file('questions/q_1.json','{}');
  await assert.rejects(readBankFile(await noManifest.generateAsync({type:'arraybuffer'})),/no manifest\.json/);
  const badManifest=new JSZip();badManifest.file('manifest.json','{ not json');
  await assert.rejects(readBankFile(await badManifest.generateAsync({type:'arraybuffer'})),/The bank's manifest could not be read, so the file may be damaged\./);
  await assert.rejects(readBankFile(await bankBytes([],{manifest:{format:'something-else'}})),/This file says it is "something-else", not an OpenAssess bank\./);
  await assert.rejects(readBankFile(await bankBytes([],{manifest:{formatVersion:undefined}})),/does not say which version of the bank format/);
  await assert.rejects(readBankFile(await bankBytes([],{manifest:{formatVersion:2}})),/newer OpenAssess than this guide can read \(it reads bank format 1\)\. Update OpenStage Guide/);
  // The bytes may come as an ArrayBuffer (from a file) or a Uint8Array.
  assert.equal((await readBankFile(new Uint8Array(await bankBytes([MC()])))).all.length,1);
});

test('a bank is read in the author\'s folder order, then the rest; retired, damaged and unreadable questions are handled and named',async()=>{
  const a=MC({title:'A'}),b=MC({title:'B'}),c=MC({title:'C'}),d=MC({title:'D',metadata:{status:'retired'}}),e=MC({title:'E'}),empty=MC({title:'F',stem:'<p></p>'});
  const bytes=await bankBytes([a,b,c,d,e,empty],{folders:{filters:[],root:[{id:'f1',name:'Chapter 1',questionIds:[c.id,empty.id],children:[{id:'f2',name:'Quiz',questionIds:[a.id,d.id],children:[]}]},{id:'f3',name:'Empty',questionIds:[],children:[]}]}});
  // the damaged file needs a name
  const zip=await JSZip.loadAsync(bytes);zip.file('questions/q_bad.json','{ not json');
  const bank=await readBankFile(await zip.generateAsync({type:'arraybuffer'}));
  assert.equal(bank.title,'Light and radio');
  assert.deepEqual(bank.all.map(q=>q.title),['C','A','B','E',empty.id,'q_bad'],'folder order, then the rest in id order, then the files that could not be read; the retired one is left out');
  assert.match(bank.all.at(-2).problem.message,/^The question has no stem \(no question text\)\.$/);
  assert.match(bank.all.at(-1).problem.message,/^The question file could not be read/);
  assert.deepEqual(bank.folders.map(f=>[f.name,f.depth,f.ids.length]),[['Chapter 1',0,2],['Quiz',1,1]],'a folder holds the ones under it, not the retired or unreadable; empty folders are not offered');
});

test('a bank written by OpenAssess\'s own code (a real file, not one built here) reads: seven questions come across, the hotspot is named',async()=>{
  const bytes=fs.readFileSync(new URL('./fixtures/openassess-written-by-openassess.openassess',import.meta.url));
  const bank=await readBankFile(bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength));
  assert.equal(bank.title,'Cross-check bank');
  assert.deepEqual(bank.all.filter(q=>q.item).map(q=>q.item.type).sort(),['fill-blank','matching','multiple-choice','numeric','ordering','reflection','true-false']);
  assert.match(bank.all.find(q=>q.problem).problem.message,/"hotspot" is not read here/);
  const mc=bank.all.find(q=>q.item?.type==='multiple-choice').item;
  assert.deepEqual([mc.options,mc.correct,mc.points,mc.why],[['Einstein','Michelson','Edison'],[1],'2','1887.']);
});

// ---- the pane
const fileOf=async(questions,opts)=>{const buf=await bankBytes(questions,opts);return {name:'bank.openassess',arrayBuffer:async()=>buf};};
async function inSection(){
  const a=await started('Bank',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await click(a,'guide-tab-assess');await click(a,'guide-sec-level-two');
  assert.equal(a.err(),'');return a;
}
const choose=async(a,questions,opts)=>{a.n('guide-bank-file').files=[await fileOf(questions,opts)];await a.n('guide-bank-file').onchange();};
const questionRows=a=>a.rows().filter(r=>r.startsWith('OSG Question'));

test('choosing a bank file shows what it holds and what cannot be added, and nothing is inserted yet',async()=>{
  const a=await inSection();
  const before=a.rows();
  await choose(a,[MC({title:'Speed of light'}),Q('trueFalse',{correct:true}),Q('hotspot',{mediaRef:'m',regions:[],alternative:rt('x')},{title:'Map'}),Q('trueFalse',{},{title:'Sound'})]);
  assert.equal(a.n('guide-bank-box').hidden,false);
  assert.match(a.n('guide-bank-summary').textContent,/^Light and radio: 4 questions here\. 2 questions can be added \(1 multiple choice and 1 true or false\); 2 cannot, and are listed below\./);
  assert.deepEqual(a.n('guide-bank-skipped').children.map(li=>li.textContent).sort(),['Map: The type "hotspot" is not read here. Add this question with Add an item.','Sound: No answer is marked: it is not said whether the statement is true or false.']);
  assert.equal(a.n('guide-bank-window').textContent,'Next to add: questions 1 to 2 of 2.');
  assert.deepEqual(a.rows(),before);
});

test('Add questions from the bank puts them in the section, numbered on, as items that read back, in one step',async()=>{
  const a=await inSection();
  const before=a.rows();
  await choose(a,[MC({title:'One',stem:'<p>Who measured it?</p>',points:2,feedback:{correct:rt('<p>1887.</p>')}}),Q('numeric',{mode:'exact',value:42,units:'kg'},{stem:'<p>How heavy?</p>'}),Q('hotspot',{mediaRef:'m',regions:[],alternative:rt('x')})]);
  await submit(a,'guide-bank-form');
  assert.equal(a.err(),'');assert.equal(a.n('guide-bank-problem').hidden,true);
  assert.deepEqual(questionRows(a),['OSG Question: Question 1. Who measured it?','OSG Question: Question 2. How heavy?']);
  assert.match(a.n('guide-bank-ok').textContent,/^2 questions added from the bank, numbered 1 to 2 \(questions 1 to 2 of 2 that can be added\)\. That is all of them\. 1 question could not be added and is listed above\./);
  const paras=await a.adapter.paragraphs();
  const it=readItem(paras,paras.findIndex(p=>/Question 1\./.test(p.text)));
  assert.deepEqual([it.type,it.correct,it.points,it.why],['multiple-choice',[1],'2','1887.']);
  assert.ok(a.rows().includes('OSG Answer Key: Answer key (numeric): 42 kg'));
  assert.ok(a.adapter.undo());
  assert.deepEqual(a.rows(),before,'one Ctrl+Z takes all of them out');
});

test('more than 50 are added 50 at a time, and the start moves on by itself',async()=>{
  const a=await inSection();
  await choose(a,Array.from({length:60},(_,i)=>MC({stem:`<p>Number ${i+1}?</p>`})));
  assert.equal(a.n('guide-bank-window').textContent,'Next to add: questions 1 to 50 of 60.');
  await submit(a,'guide-bank-form');
  assert.equal(questionRows(a).length,50);
  assert.equal(a.n('guide-bank-start').value,'51');
  assert.match(a.n('guide-bank-ok').textContent,/questions 1 to 50 of 60 that can be added\)\. 10 more can still be added: choose the button again\./);
  await submit(a,'guide-bank-form');
  assert.equal(questionRows(a).length,60);
  assert.match(a.n('guide-bank-ok').textContent,/numbered 51 to 60 \(questions 51 to 60 of 60 that can be added\)\. That is all of them\./);
  assert.equal(MAX_BANK_ADD,50);
  assert.deepEqual(readAssessments(await a.adapter.paragraphs()).sections[0].items.map(i=>i.number),Array.from({length:60},(_,i)=>i+1));
});

test('a folder can be chosen, and the start can be set to begin part way through',async()=>{
  const q=Array.from({length:4},(_,i)=>MC({stem:`<p>Q${i+1}?</p>`}));
  const a=await inSection();
  await choose(a,q,{folders:{filters:[],root:[{id:'f1',name:'Part one',questionIds:[q[0].id,q[1].id],children:[]},{id:'f2',name:'Part two',questionIds:[q[2].id,q[3].id],children:[]}]}});
  assert.deepEqual(a.n('guide-bank-folder').children.map(o=>o.textContent),['All questions (4)','Part one (2)','Part two (2)']);
  a.n('guide-bank-folder').value='f2';a.n('guide-bank-folder').onchange();
  assert.equal(a.n('guide-bank-window').textContent,'Next to add: questions 1 to 2 of 2.');
  a.n('guide-bank-start').value='2';a.n('guide-bank-start').oninput();
  await submit(a,'guide-bank-form');
  assert.deepEqual(questionRows(a),['OSG Question: Question 1. Q4?']);
});

test('the pane says what is wrong: no file, a bad file, a bad start, no section, nothing that can be added',async()=>{
  const a=await inSection();
  await submit(a,'guide-bank-form');
  assert.match(a.n('guide-bank-problem').textContent,/Choose a bank file first/);
  a.n('guide-bank-file').files=[{name:'x.openassess',arrayBuffer:async()=>new TextEncoder().encode('nope').buffer}];await a.n('guide-bank-file').onchange();
  assert.match(a.n('guide-bank-problem').textContent,/could not be opened as an OpenAssess bank/);
  assert.equal(a.n('guide-bank-box').hidden,true);
  await choose(a,[MC()]);
  a.n('guide-bank-start').value='9';await submit(a,'guide-bank-form');
  assert.match(a.n('guide-bank-problem').textContent,/Start at a number from 1 to 1/);
  await choose(a,[Q('hotspot',{mediaRef:'m',regions:[],alternative:rt('x')})]);
  await submit(a,'guide-bank-form');
  assert.match(a.n('guide-bank-problem').textContent,/no questions here that can be added/);
  const b=await started('No section',{quirks:WORD_QUIRKS});
  await addModule(b,'One');await addLesson(b,'First');await click(b,'guide-tab-assess');
  await choose(b,[MC()]);await submit(b,'guide-bank-form');
  assert.match(b.n('guide-bank-problem').textContent,/Items go inside an assessment section, and your cursor is not in one/);
  assert.equal(questionRows(b).length,0);
});

test('imported questions make a workbook with the questions and without their keys',async()=>{
  const a=await inSection();
  await choose(a,[MC(),Q('essay',{modelAnswer:rt('<p>Secret model answer.</p>')},{stem:'<p>Explain it.</p>'})]);
  await submit(a,'guide-bank-form');
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true,'the leak check passes');
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64'));
  const xml=await zip.file('word/document.xml').async('string');
  assert.ok(xml.includes('Explain it.')&&!xml.includes('Secret model answer'));
});
