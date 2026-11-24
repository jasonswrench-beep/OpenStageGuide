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

// ---- bringing changes back from a presentation
import {DOMParser} from '@xmldom/xmldom';
import {slidesModel} from '../src/session-decks.mjs';
import {pptxBytes} from '../src/pptx-export.mjs';
import {readPptx} from '../src/pptx.mjs';
globalThis.DOMParser=DOMParser;
const editedDeck=async(a,edit)=>{
  const model=slidesModel(await a.adapter.paragraphs(),{start:'09:00'});
  edit(model);
  const read=await readPptx(await pptxBytes(model,JSZip),'Edited.pptx',JSZip);
  return {name:'Edited.pptx',slides:read.slides};
};
const chooseDeck=(a,deck)=>a.adapter.setPicker(async opts=>{assert.equal(opts.kind,'slides');return {deck,files:[]};});
const changeBoxes=a=>a.n('guide-slide-changes').children.flatMap(g=>g.children[1].children).map(li=>li.children[0].children[0]);
const changeTexts=a=>a.n('guide-slide-changes').children.flatMap(g=>g.children[1].children).map(li=>li.children[0].children[1].textContent);
const rehearsal=async a=>{const paras=await a.adapter.paragraphs();return readBlock(paras,paras.findIndex(p=>/Scenario rehearsal \(18 min\)/.test(p.text)));};

test('a presentation edited outside is read back: its changes are listed against the activities, nothing is written until some are ticked, and only those are',async()=>{
  const a=await withTool();
  const deck=await editedDeck(a,m=>{const s=m.slides.find(x=>/^Scenario rehearsal/.test(x.title));s.lines[0]='Take your roles: performer, counterpart and observer.';s.lines=s.lines.map(l=>l==='Grouping: Triads'?'Grouping: Groups of three':l);});
  chooseDeck(a,deck);
  const before=a.rows();
  await click(a,'guide-slide-import');
  assert.equal(a.err(),'');
  assert.match(a.n('guide-slide-import-note').textContent,/^Edited\.pptx: 2 changes in 1 activity\./);
  assert.deepEqual(changeTexts(a).map(t=>t.replace(/".*$/,'').trim()),['Instructions:','Grouping:']);
  assert.match(changeTexts(a)[1],/Grouping: "Triads" becomes "Groups of three"/);
  assert.deepEqual(a.rows(),before,'listing the changes writes nothing');
  assert.deepEqual(changeBoxes(a).map(b=>b.checked),[true,true]);
  assert.match(a.n('guide-slide-apply').textContent,/Apply the ticked changes \(2\)/);
  assert.equal(a.n('guide-slide-unmatched').children.length,2,'the opening and the close are not an activity\'s');
  // untick the first, apply: only the grouping changes
  changeBoxes(a)[0].checked=false;changeBoxes(a)[0].onchange();
  assert.match(a.n('guide-slide-apply').textContent,/Apply the ticked changes \(1\)/);
  await click(a,'guide-slide-apply');
  assert.equal(a.err(),'');
  const b=await rehearsal(a);
  assert.equal(b.parts.grouping,'Groups of three');
  assert.match(b.parts.instructions,/^Take the roles of performer, counterpart and observer\./,'the unticked change was not made');
  assert.equal(b.parts.phase,'Practice');
  assert.match(a.n('guide-status').textContent,/One activity was updated from Edited\.pptx\./);
  assert.equal(a.n('guide-slide-changes').children.length,0,'the list is cleared once applied');
});

test('a change to something the guide also changed since the slides were made is marked and not ticked; ticking it is the author\'s choice',async()=>{
  const a=await withTool();
  await click(a,'guide-slide-make');                         // the slides are made, and the guide's words are remembered
  assert.equal(a.err(),'');
  // the presentation is edited outside, from the slides as they were made: it says something else about the grouping and adds a debrief prompt
  const deck=await editedDeck(a,m=>{const s=m.slides.find(x=>/^Scenario rehearsal/.test(x.title));s.lines=s.lines.map(l=>l==='Grouping: Triads'?'Grouping: Pairs':l);s.notes.push('Debrief: A new prompt from the slides.');});
  // the guide changes the grouping afterward
  const paras=await a.adapter.paragraphs();
  a.adapter.place(paras.findIndex(p=>/Scenario rehearsal \(18 min\)/.test(p.text)));await a.p.quietRefresh();
  const {activityRewrite}=await import('../src/session-slide-import.mjs');
  const {cueBlock}=await import('../src/guide-ooxml.mjs');
  const rw=activityRewrite(paras,'scenario rehearsal#1',{grouping:'Fours'});
  await a.adapter.insert(ctx=>cueBlock(rw.args,ctx),{where:'after',after:rw.span.end,select:'last',remove:{start:rw.span.start,end:rw.span.end,guard:rw.guard}});
  chooseDeck(a,deck);
  await click(a,'guide-slide-import');
  assert.match(a.n('guide-slide-import-note').textContent,/1 of them also changed in the guide since the slides were made/);
  const texts=changeTexts(a),boxes=changeBoxes(a);
  const g=texts.findIndex(t=>/^Grouping:/.test(t)),d=texts.findIndex(t=>/^Debrief prompts:/.test(t));
  assert.match(texts[g],/The guide has also changed this since the slides were made\./);
  assert.equal(boxes[g].checked,false);assert.equal(boxes[d].checked,true);
});

test('no presentation chosen, or a file that cannot be read, leaves the guide and the list as they were',async()=>{
  const a=await withTool();
  const before=a.rows();
  a.adapter.setPicker(async()=>({cancelled:true}));
  await click(a,'guide-slide-import');
  assert.match(a.n('guide-status').textContent,/No presentation was chosen\. The guide is as it was\./);
  a.adapter.setPicker(async()=>{throw new Error('That file is not a presentation this program can read.');});
  await click(a,'guide-slide-import');
  assert.match(a.err(),/not a presentation this program can read/);
  assert.deepEqual(a.rows(),before);
  assert.equal(a.n('guide-slide-apply').disabled,true);
  assert.equal(a.n('guide-slide-changes').children.length,0);
});

test('an activity that is gone by the time the changes are applied is reported, and the others are still applied',async()=>{
  const a=await withTool();
  const deck=await editedDeck(a,m=>{
    const s=m.slides.find(x=>/^Scenario rehearsal/.test(x.title));s.lines=s.lines.map(l=>l==='Grouping: Triads'?'Grouping: Pairs':l);
    const t=m.slides.find(x=>/^Exit check/.test(x.title));t.lines=t.lines.map(l=>/^Make:/.test(l)?'Make: A short written exit response':l);
  });
  chooseDeck(a,deck);
  await click(a,'guide-slide-import');
  assert.equal(a.n('guide-slide-changes').children.length,2);
  // the exit check is deleted from the guide before the author applies
  const paras=await a.adapter.paragraphs();
  const at=paras.findIndex(p=>/Exit check \(5 min\)/.test(p.text));
  const b=readBlock(paras,at);
  await a.adapter.deleteRange({start:b.span.start,end:b.span.end,guard:paras[b.span.start].text});
  await click(a,'guide-slide-apply');
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/One activity was updated from Edited\.pptx\. 1 could not be: That activity is not in the guide any more/);
  assert.equal((await rehearsal(a)).parts.grouping,'Pairs');
});