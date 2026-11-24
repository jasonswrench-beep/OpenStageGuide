// The Decks and slides tool on its own panel (src/guide-pane-decks.mjs): start a deck from an activity, edit and preview its cards, print it as a Word sheet, name it in the activity's
// materials, and make the slides as a PowerPoint file. Drives the real pane module and the real taskpane.html ids in memory. It is not native Word.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {readSessions} from '../src/session-record.mjs';
import {readBlock} from '../src/guide-model.mjs';
import {started,click,addModule,addLesson,submit} from './helpers/guide-harness.mjs';

const withTool=async()=>{
  const a=await started('Skill day');
  await addModule(a,'Practice');await addLesson(a,'Warm up','5');
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  a.n('guide-tpl-list').children[0].children[2].children[0].onclick();
  a.type('guide-tpl-title','Skill practice');
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  a.p.useTool('sessionDecks');await a.p.openTool('sessionDecks');
  return a;
};
const decks=a=>a.n('guide-deck-list').children;
const recorded=async a=>readSessions(await a.adapter.sessions());
const startDeck=async(a,{title='',type='activity',activity=''}={})=>{a.type('guide-deck-title',title);a.type('guide-deck-type',type);a.type('guide-deck-activity',activity);await submit(a,'guide-deck-new-form');};
const cardTitles=a=>a.n('guide-deck-cards').children.map(li=>li.children[0].textContent);

test('the tool is in the Make menu, offers the three kinds of deck and the activities of the guide, and starts with no decks',async()=>{
  const a=await withTool();
  assert.equal(a.n('guide-sec-decks').hidden,false);assert.equal(a.n('guide-sec-materials').hidden,true);
  assert.deepEqual(a.n('guide-deck-type').children.map(o=>o.value),['activity','toolkit','facilitator']);
  const options=a.n('guide-deck-activity').children;
  assert.equal(options[0].value,'');
  assert.ok(options.some(o=>o.value==='scenario rehearsal#1'&&/^Scenario rehearsal \(Skill practice\)$/.test(o.textContent)));
  assert.equal(decks(a).length,0);
  assert.equal(a.n('guide-deck-editor').hidden,true);
  assert.equal(a.err(),'');
});

test('a deck started from an activity has a card for each step, opens for editing, and is kept in the project record',async()=>{
  const a=await withTool();
  await startDeck(a,{activity:'scenario rehearsal#1'});
  assert.equal(a.err(),'');
  assert.equal(decks(a).length,1);
  assert.match(a.n('guide-status').textContent,/Scenario rehearsal: cards is started with 3 cards\./);
  assert.equal(a.n('guide-deck-editor').hidden,false);
  assert.match(a.n('guide-deck-editor-title').textContent,/^Scenario rehearsal: cards \(3 cards\)$/);
  assert.equal(cardTitles(a).length,3);
  assert.match(cardTitles(a)[0],/Take the roles of performer/);
  assert.match(cardTitles(a)[0],/\(has a back\)/);
  assert.equal(a.n('guide-deck-paper').value,'Letter');assert.equal(a.n('guide-deck-perpage').value,'4');
  assert.match(a.n('guide-deck-preview').children[0].textContent,/^Front: Take the roles/);
  assert.match(a.n('guide-deck-preview').children[1].textContent,/^Back: Follow up: What changed/);
  const rec=await recorded(a);
  assert.equal(rec.decks.length,1);assert.equal(rec.decks[0].activity,'Scenario rehearsal');
  assert.equal(a.n('guide-deck-fit').children[0].textContent,'Every card fits at this size.');
});

test('cards are added, changed, moved and removed, and every change is kept',async()=>{
  const a=await withTool();
  await startDeck(a,{title:'Feedback cards',type:'activity'});
  assert.equal(cardTitles(a).length,1);
  a.type('guide-deck-front','What did you notice?\nName one thing you saw.');a.type('guide-deck-back','Cue: say what you saw.');
  await click(a,'guide-deck-add');
  assert.equal(a.err(),'');
  assert.equal(cardTitles(a).length,2);
  assert.equal(cardTitles(a)[1],'What did you notice? (has a back)');
  a.type('guide-deck-front','What would you try next?');a.type('guide-deck-back','');
  await click(a,'guide-deck-add');
  assert.equal(cardTitles(a).length,3);
  await click(a,'guide-deck-up');
  assert.deepEqual(cardTitles(a).map(t=>t.replace(/ \(has a back\)/,'')),['Write the first card here.','What would you try next?','What did you notice?']);
  a.type('guide-deck-front','What would you do differently?');
  await click(a,'guide-deck-update');
  assert.match(cardTitles(a)[1],/^What would you do differently\?/);
  await click(a,'guide-deck-remove');
  assert.equal(cardTitles(a).length,2);
  assert.equal((await recorded(a)).decks[0].cards.length,2);
  // errors say what is wrong and change nothing
  a.type('guide-deck-front','');
  await click(a,'guide-deck-add');
  assert.match(a.err(),/Write the front of the card first/);
  assert.equal(cardTitles(a).length,2);
  await click(a,'guide-deck-down');await click(a,'guide-deck-down');
  assert.match(a.err(),/This card is the last one/);
});

test('a card that is too full for its size is reported in words, and the size can be changed',async()=>{
  const a=await withTool();
  await startDeck(a,{title:'Dense',type:'activity'});
  a.type('guide-deck-front','x'.repeat(250)+' words');a.type('guide-deck-back','');
  await click(a,'guide-deck-add');
  a.type('guide-deck-perpage','8');await a.n('guide-deck-perpage').onchange();
  const text=a.n('guide-deck-fit').children.map(li=>li.textContent).join(' ');
  assert.match(text,/Card 2 \(front\) has \d+ characters, which is more than fits when 8 cards go on a page/);
  a.type('guide-deck-perpage','2');await a.n('guide-deck-perpage').onchange();
  assert.equal(a.n('guide-deck-fit').children[0].textContent,'Every card fits at this size.');
  assert.equal((await recorded(a)).decks[0].perPage,2);
});

test('the print sheet is made as a Word file on the paper chosen, the deck is listed as an output, and the text view shows every card',async()=>{
  const a=await withTool();
  await startDeck(a,{title:'Scenario cards',activity:'scenario rehearsal#1'});
  a.type('guide-deck-paper','A4');await a.n('guide-deck-paper').onchange();a.n('guide-deck-marks').checked=true;await a.n('guide-deck-marks').onchange();
  await click(a,'guide-deck-print');
  assert.equal(a.err(),'');
  const d=a.adapter.openedDocuments()[0];
  assert.ok(d);
  const zip=await JSZip.loadAsync(Buffer.from(d.base64,'base64'));
  const doc=await zip.file('word/document.xml').async('string');
  assert.match(doc,/w:pgSz w:w="11906"/);assert.match(doc,/Take the roles of performer/);assert.match(doc,/Crop mark/);
  assert.match(a.n('guide-status').textContent,/The print sheet is made on A4 paper, 4 to a page\./);
  const rec=await recorded(a);
  assert.equal(rec.outputs[0].kind,'cards');assert.equal(rec.outputs[0].name,'Card deck: Scenario cards');
  await click(a,'guide-deck-text');
  assert.equal(a.n('guide-deck-textview').hidden,false);
  assert.match(a.n('guide-deck-textview').textContent,/Card 1\nFront: Take the roles/);
});

test('the deck is named in the materials of the activity it came from, once, and the activity still reads as parts',async()=>{
  const a=await withTool();
  await startDeck(a,{title:'Scenario cards',activity:'scenario rehearsal#1'});
  await click(a,'guide-deck-link');
  assert.equal(a.err(),'');
  const find=async()=>{const paras=await a.adapter.paragraphs();const at=paras.findIndex(p=>/Scenario rehearsal \(18 min\)/.test(p.text));return readBlock(paras,at);};
  const b=await find();
  assert.ok(b.parts,'still labelled parts');
  assert.match(b.parts.materials,/Observation checklist\nCard deck: Scenario cards$/);
  assert.equal(b.parts.phase,'Practice');assert.match(b.parts.steps,/3 min: Assign performer/);
  await click(a,'guide-deck-link');
  assert.match(a.n('guide-status').textContent,/already names this deck/);
  assert.equal(((await find()).parts.materials.match(/Card deck/g)||[]).length,1);
  // a blank deck has no activity to name it in
  await startDeck(a,{title:'Blank',type:'activity'});
  await click(a,'guide-deck-link');
  assert.match(a.err(),/was not started from an activity/);
});

test('a deck can be deleted, and the sheet already made is left alone',async()=>{
  const a=await withTool();
  await startDeck(a,{title:'Short-lived',type:'toolkit'});
  assert.match(a.n('guide-deck-editor-title').textContent,/Short-lived \(6 cards\)/);
  await click(a,'guide-deck-print');
  assert.equal(a.err(),'');
  await click(a,'guide-deck-forget');
  assert.equal(decks(a).length,0);assert.equal(a.n('guide-deck-editor').hidden,true);
  assert.equal((await recorded(a)).decks.length,0);
  assert.equal(a.adapter.openedDocuments().length,1);
  assert.equal((await recorded(a)).outputs.length,1);
});

test('the slides: the order is listed, nothing is too crowded, and the PowerPoint file is offered with notes, and listed as an output',async()=>{
  const a=await withTool();
  assert.match(a.n('guide-slide-note').textContent,/^\d+ slides, in this order\./);
  const titles=a.n('guide-slide-list').children.map(li=>li.textContent);
  assert.match(titles[0],/^1\. Skill day$/);assert.match(titles.at(-1),/\. Close$/);
  assert.ok(titles.some(t=>/Scenario rehearsal \(09:22, 18 min\)/.test(t)));
  assert.equal(a.n('guide-slide-fit').children[0].textContent,'Every slide is short enough to read.');
  assert.equal(a.n('guide-slide-make').disabled,false);
  await click(a,'guide-slide-make');
  assert.equal(a.err(),'');
  const file=a.adapter.downloads?.()[0]||null;
  assert.ok(file,'a PowerPoint file was offered');
  assert.equal(file.fileName,'Skill day.pptx');
  assert.equal(file.type,'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  const zip=await JSZip.loadAsync(Buffer.from(file.base64,'base64'));
  assert.ok(zip.file('ppt/notesSlides/notesSlide1.xml'));
  assert.match(await zip.file('ppt/slides/slide3.xml').async('string'),/name="scenario rehearsal#1"|name="model and notice#1"|name="[^"]+#1"/);
  assert.equal((await recorded(a)).outputs[0].kind,'slides');
  a.type('guide-slide-start','late');a.n('guide-slide-start').onchange();
  assert.match(a.n('guide-slide-note').textContent,/clock time such as 9:00/);
  assert.equal(a.n('guide-slide-make').disabled,true);
});

test('with no activities there are no slides to make, and the page says so',async()=>{
  const a=await started('Empty');
  await addModule(a,'M');await addLesson(a,'L','5');
  a.p.useTool('sessionDecks');await a.p.openTool('sessionDecks');
  assert.match(a.n('guide-slide-note').textContent,/no activities to make slides from yet/);
  assert.equal(a.n('guide-slide-make').disabled,true);
});
