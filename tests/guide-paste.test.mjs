// Pasting many questions at once, in the plain-text layout that test banks use. Proven against stand-ins for Word, not Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import {parseStandardFormat,readPastedQuestions,MAX_PASTE} from '../src/guide-paste.mjs';
import {readAssessments,readItem} from '../src/guide-model.mjs';
import {started,WORD_QUIRKS,click,submit,addModule,addLesson} from './helpers/guide-harness.mjs';

const read=text=>readPastedQuestions(text);
const types=r=>r.items.map(i=>i.type);

// ---- the layout, read
test('a numbered question with lettered options and a * before the correct one is multiple choice',()=>{
  const r=read('1) Who measured the speed of light?\na. Einstein\n*b) Michelson\nc. Edison\n\n2. Which is a mammal?\n*a. Bat\nb. Trout');
  assert.deepEqual(r.problems,[]);
  assert.deepEqual(r.items.map(i=>[i.type,i.stem,i.options,i.correct]),[
    ['multiple-choice','Who measured the speed of light?',['Einstein','Michelson','Edison'],[1]],
    ['multiple-choice','Which is a mammal?',['Bat','Trout'],[0]]
  ]);
  assert.equal(r.found,2);assert.deepEqual([...r.counts],[['multiple-choice',2]]);
});

test('the key can be an Answer: line, or an Answers: list at the end, in place of the *',()=>{
  const r=read('1) First?\na. One\nb. Two\nAnswer: b\n\n2) Second?\na. One\nb. Two\nc. Three\n\nAnswers:\n2. C');
  assert.deepEqual(r.problems,[]);
  assert.deepEqual(r.items.map(i=>i.correct),[[1],[2]]);
});

test('a publisher\'s "ANS: B" line with its notes after it is read, and the notes are skipped',()=>{
  const r=read('9) What is the purpose?\na. To summarise\nb. To focus\nc. To report\nANS: B  DIF: Easy  REF: 87\nOBJ: 3.2\nTOP: Design');
  assert.deepEqual(r.problems,[]);
  assert.deepEqual([r.items[0].correct,r.ignored],[[1],2]);
});

test('Type: MA or MR is multiple response and needs at least one correct option',()=>{
  const r=read('Type: MA\n1) Which are noble gases?\n*a. Neon\n*b. Argon\nc. Oxygen\n\nType: MR\n2) Which are prime?\na. 4\nb. 5\nc. 9\nd. 11\n\nAnswers:\n2. b d');
  assert.deepEqual(r.problems,[]);
  assert.deepEqual(r.items.map(i=>[i.type,i.correct]),[['multiple-response',[0,1]],['multiple-response',[1,3]]]);
});

test('true or false is read from True and False options however the key is written',()=>{
  const r=read('1) Light is a wave.\na. True\n*b. False\n\n2) Sound is faster.\n*a. True\nb. False\n\n3) Radio is EM.\na. True\nb. False\n\n4) Light is a particle.\na. False\nb. True\n\n5) Fifth.\nType: TF\na. True\nb. False\nAnswer: T\n\nAnswers:\n3. A\n4. A');
  // The Type: line inside question 5's text is part of its text, so 5 is also read as true or false from its options.
  assert.deepEqual(r.problems,[]);
  assert.deepEqual(r.items.map(i=>[i.type,i.truth]),[['true-false',false],['true-false',true],['true-false',true],['true-false',false],['true-false',true]]);
});

test('Type: E, S and F read their answers from the options or the Answers: list, and say when they used a placeholder',()=>{
  const r=read('Type: E\n1) Explain the experiment.\na) In 1887 they found no ether wind.\n\nType: S\n2) Chemical symbol for sodium?\n\nType: S\n3) Who painted the Mona Lisa?\n\nType: F\n4) Who wrote Hamlet?\na. William Shakespeare\nb. Shakespeare\n\nAnswers:\n2. Na\n3. Leonardo da Vinci\n3. Da Vinci');
  assert.deepEqual(r.problems,[]);
  assert.deepEqual(r.items.map(i=>[i.type,i.answer]),[
    ['reflection','In 1887 they found no ether wind.'],
    ['short-answer','Na'],
    ['short-answer','Leonardo da Vinci; Da Vinci'],
    ['fill-blank','William Shakespeare; Shakespeare']
  ]);
  const none=read('Type: E\n1) Discuss.');
  assert.equal(none.items[0].answer,'Add a model answer, or what to look for.');
  assert.match(none.notes[0],/placeholder in its key/);
});

test('Type: MT reads each option as an item, an equals sign and its match',()=>{
  const r=read('Type: MT\n1) Match the name to the discovery.\na. Michelson-Morley = Speed of light\nb. Einstein = Theory of relativity\nc. Marconi = Radio waves');
  assert.deepEqual(r.problems,[]);
  assert.deepEqual(r.items[0].pairs,[{item:'Michelson-Morley',match:'Speed of light'},{item:'Einstein',match:'Theory of relativity'},{item:'Marconi',match:'Radio waves'}]);
  assert.match(read('Type: MT\n1) Match.\na. one\nb. two').problems[0].message,/an item, an equals sign and its match/);
});

test('Points: before a question gives its points; feedback, titles and publisher notes are skipped',()=>{
  const r=read('Points: 2.5\n10) Which is a measure of centre?\n~ Correct. The mean is.\n@ Not quite.\na. Deviation\n@ Spread, not centre.\n*b. Mean\nc. Range');
  assert.deepEqual(r.problems,[]);
  assert.deepEqual([r.items[0].points,r.items[0].correct,r.ignored],['2.5',[1],3]);
});

test('a line that looks like a directive in the middle of a question\'s text stays in the text',()=>{
  const r=read('Points: 2\nTitle: Reading a label\n1) Read the label that says:\nPoints: 20\nType: A\nThen answer the question below.\na. Voltage\n*b. Current\n\nPoints: 3\n2) What does the second ask?\n*a. Nothing\nb. Everything');
  assert.deepEqual(r.problems,[]);
  assert.equal(r.items[0].stem,'Read the label that says: Points: 20 Type: A Then answer the question below.');
  assert.deepEqual(r.items.map(i=>i.points),['2','3']);
});

test('a question\'s text and an option can run over more than one line',()=>{
  const r=read('1) A long question\nthat wraps onto a second line?\na. A long option\nthat also wraps\n*b. Short');
  assert.equal(r.items[0].stem,'A long question that wraps onto a second line?');
  assert.deepEqual(r.items[0].options,['A long option that also wraps','Short']);
});

test('an Answers: list ends at its first blank line, so questions can follow it, and a repeated number it is keyed by is refused',()=>{
  const ok=read('1) First?\na. A\nb. B\n\nAnswers:\n1. B\n\n2) Second?\n*a. A\nb. B');
  assert.deepEqual(ok.problems,[]);assert.equal(ok.items.length,2);
  const dup=read('Chapter 1\n\n1) First?\na. A\nb. B\n\nAnswers:\n1. B\n\nChapter 2\n\n1) Again?\na. A\nb. B\n\nAnswers:\n1. A');
  assert.equal(dup.items.length,0);
  assert.match(dup.problems[0].message,/Number 1 is used 2 times and the Answers: list is keyed by it/);
  assert.equal(dup.problems.length,1,'one clear problem, not a pile of follow-on ones');
});

test('questions that cannot be read are named with their number and line, and each says what to do',()=>{
  const r=read('1) No key?\na. One\nb. Two\n\n2) Two keys?\n*a. One\n*b. Two\n\n3) Bad letter?\na. One\nb. Two\nAnswer: d\n\n4) Skipped letter?\na. One\nc. Three\n*d. Four\n\n5) Odd type?\nType: ORD\n\nType: ORD\n6) Ordering?\na. x\nb. y\n\n7) One option?\n*a. Only');
  const by=n=>r.problems.find(p=>p.number===n);
  assert.match(by(1).message,/No correct answer is marked/);
  assert.match(by(2).message,/If more than one is meant, put Type: MA/);
  assert.match(by(3).message,/The answer "d" is not one of the options/);
  assert.match(by(4).message,/skip or do not start at a/);
  assert.match(by(6).message,/Type: ORD is not read here/);
  assert.match(by(7).message,/at least two lettered options/);
  assert.equal(by(1).line,1);assert.equal(by(2).line,5);
  assert.equal(r.problems.length,7,'every question here has a problem');
});

test('a question with no options and no Type: line is refused rather than guessed at, unless it has an answer',()=>{
  assert.match(read('1) Is this a heading or a question?').problems[0].message,/cannot be told from a heading/);
  const r=read('1) Capital of Italy?\nAnswer: Rome');
  assert.deepEqual(r.problems,[]);assert.deepEqual([r.items[0].type,r.items[0].answer],['short-answer','Rome']);
});

test('more than six options is refused, with the limit',()=>{
  const opts=['a','b','c','d','e','f','g'].map((l,i)=>`${i===0?'*':''}${l}. Option ${l}`).join('\n');
  assert.match(read(`1) Too many?\n${opts}`).problems[0].message,/7 options and a guide question can have up to 6/);
});

test('the builder\'s own refusals come back as problems for that question, not as a surprise on insert',()=>{
  const r=read('Type: MT\n1) Match.\na. A = 1\nb. B = 1\nc. C = 2\nd. D = 3\ne. E = 4\nf. F = 5\ng. G = 6\nh. H = 7\ni. I = 8\nj. J = 9\nk. K = 10');
  assert.ok(r.problems.length===1&&/up to 10 pairs/.test(r.problems[0].message));
  assert.equal(MAX_PASTE,50);
});

test('nothing but text with no numbered question finds none, and a stray heading before the first question is skipped',()=>{
  assert.equal(read('Just a paragraph of text.\nAnd another.').found,0);
  const r=read('Chapter 1: Light\n\n1) A question?\n*a. Yes\nb. No');
  assert.deepEqual([r.found,r.ignored,r.problems],[1,1,[]]);
  assert.equal(parseStandardFormat('').questions.length,0);
});

// ---- the pane
const typeText=(a,text)=>{a.n('guide-paste-text').value=text;};
async function inSection(level='level-two'){
  const a=await started('Paste test',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await click(a,'guide-tab-assess');await click(a,`guide-sec-${level}`);
  assert.equal(a.err(),'');return a;
}
const BANK='Type: MA\n1) Which are noble gases?\n*a. Neon\n*b. Argon\nc. Oxygen\n\n2) Sound is faster than light.\na. True\n*b. False\n\n3) Who measured it?\na. Einstein\n*b. Michelson\nc. Edison\n\nType: S\n4) Chemical symbol for sodium?\n\nAnswers:\n4. Na';

test('Add the pasted questions puts them all in the section, numbered to follow what is there, in one step',async()=>{
  const a=await inSection();
  const before=a.rows();
  typeText(a,BANK);
  await submit(a,'guide-paste-form');
  assert.equal(a.err(),'');assert.equal(a.n('guide-paste-problem').hidden,true);
  const rows=a.rows();
  assert.deepEqual(rows.filter(r=>r.startsWith('OSG Question')),[
    'OSG Question: Question 1. Which are noble gases? Choose all that apply.',
    'OSG Question: Question 2. Sound is faster than light.',
    'OSG Question: Question 3. Who measured it?',
    'OSG Question: Question 4. Chemical symbol for sodium?'
  ]);
  assert.deepEqual(rows.filter(r=>r.startsWith('OSG Answer Key: Answer key')),[
    'OSG Answer Key: Answer key (multiple response): a, b.',
    'OSG Answer Key: Answer key (true or false): False.',
    'OSG Answer Key: Answer key (multiple choice): b.',
    'OSG Answer Key: Answer key (short answer): Na'
  ]);
  assert.equal(a.n('guide-paste-text').value,'','the box is emptied');
  assert.match(a.n('guide-paste-ok').textContent,/4 questions added, numbered 1 to 4: 1 multiple response, 1 true or false, 1 multiple choice and 1 short answer\./);
  assert.ok(a.adapter.undo());
  assert.deepEqual(a.rows(),before,'one Ctrl+Z takes out all of them');
});

test('pasted questions follow the items already in the section and are numbered on from them',async()=>{
  const a=await inSection();
  typeText(a,'1) First?\n*a. Yes\nb. No');await submit(a,'guide-paste-form');
  typeText(a,'7) Another?\na. Yes\n*b. No\n\n7) And another?\n*a. Yes\nb. No');await submit(a,'guide-paste-form');
  assert.equal(a.err(),'');
  const items=readAssessments(await a.adapter.paragraphs()).sections[0].items;
  assert.deepEqual(items.map(i=>i.number),[1,2,3],'renumbered, though the pasted numbers repeat');
});

test('what is pasted reads back as ordinary items that Edit this item can open',async()=>{
  const a=await inSection();
  typeText(a,BANK);await submit(a,'guide-paste-form');
  const paras=await a.adapter.paragraphs();
  const at=paras.findIndex(p=>/Question 2\./.test(p.text));
  const it=readItem(paras,at);
  assert.deepEqual([it.type,it.truth,it.stem],['true-false',false,'Sound is faster than light.']);
  const mr=readItem(paras,paras.findIndex(p=>/Question 1\./.test(p.text)));
  assert.deepEqual([mr.type,mr.options,mr.correct],['multiple-response',['Neon','Argon','Oxygen'],[0,1]]);
});

test('if any question cannot be read nothing is added, the paste stays, and each problem is listed',async()=>{
  const a=await inSection();
  const before=a.rows();
  const text='1) Fine?\n*a. Yes\nb. No\n\n2) No key?\na. One\nb. Two\n\nType: ORD\n3) Odd?\na. x\nb. y';
  typeText(a,text);await submit(a,'guide-paste-form');
  assert.deepEqual(a.rows(),before);
  assert.equal(a.n('guide-paste-text').value,text);
  assert.match(a.n('guide-paste-problem').textContent,/Nothing was added, because 2 problems must be fixed first/);
  const items=a.n('guide-paste-list').children.map(li=>li.textContent);
  assert.equal(a.n('guide-paste-list').hidden,false);
  assert.match(items[0],/^Question 2, line 5: No correct answer is marked/);
  assert.match(items[1],/^Question 3, line 10: Type: ORD is not read here/);
  assert.equal(a.n('guide-paste-go').disabled,false);
});

test('an empty box, a box with no questions and too many questions each say what to do',async()=>{
  const a=await inSection();
  typeText(a,'  ');await submit(a,'guide-paste-form');
  assert.match(a.n('guide-paste-problem').textContent,/Paste the questions first/);
  typeText(a,'Just some words.');await submit(a,'guide-paste-form');
  assert.match(a.n('guide-paste-problem').textContent,/No questions were found/);
  typeText(a,Array.from({length:51},(_,i)=>`${i+1}) Q${i+1}?\n*a. Yes\nb. No`).join('\n\n'));await submit(a,'guide-paste-form');
  assert.match(a.n('guide-paste-problem').textContent,/51 questions were found and up to 50 can be added at a time/);
  assert.equal(a.rows().filter(r=>r.startsWith('OSG Question')).length,0);
});

test('outside an assessment section the paste is kept and the message says where it can go',async()=>{
  const a=await started('No section',{quirks:WORD_QUIRKS});
  await addModule(a,'One');await addLesson(a,'First');
  await click(a,'guide-tab-assess');
  const before=a.rows();
  typeText(a,BANK);await submit(a,'guide-paste-form');
  assert.match(a.n('guide-paste-problem').textContent,/Items go inside an assessment section, and your cursor is not in one\. What you pasted is still here/);
  assert.equal(a.n('guide-paste-text').value,BANK);
  assert.deepEqual(a.rows(),before);
});

test('notes about placeholders and skipped lines are in the message, and 50 questions are accepted in one paste',async()=>{
  const a=await inSection();
  typeText(a,'Title: Chapter bank\nType: E\n1) Discuss the case.\n\nDIF: Easy\n2) Fine?\n*a. Yes\nb. No');await submit(a,'guide-paste-form');
  assert.match(a.n('guide-paste-ok').textContent,/placeholder in its key/);
  assert.match(a.n('guide-paste-ok').textContent,/2 lines \(titles, feedback and publisher notes\) were not used\./);
  const b=await inSection();
  typeText(b,Array.from({length:50},(_,i)=>`${i+1}) Q${i+1}?\n*a. Yes\nb. No`).join('\n\n'));await submit(b,'guide-paste-form');
  assert.equal(b.err(),'');assert.equal(b.n('guide-paste-problem').hidden,true);
  assert.equal(b.rows().filter(r=>r.startsWith('OSG Question')).length,50);
});

test('a pasted quiz makes a workbook with the questions and without their keys',async()=>{
  const a=await inSection();
  typeText(a,BANK);await submit(a,'guide-paste-form');
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true,'the leak check passes');
});
