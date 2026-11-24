// Card decks and slides made from the guide's activities (src/session-decks.mjs), through the writers: the print sheet (src/card-print.mjs) and the presentation file (src/pptx-export.mjs).
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {DECK_TYPES,deckFromGuide,printModel,deckLine,slidesModel,deckTypeLabel} from '../src/session-decks.mjs';
import {activityBlocks} from '../src/session-record.mjs';
import {cardFit,cardSheetFiles,checkDeck as checkCards} from '../src/card-print.mjs';
import {slideFit,pptxBytes,checkDeck as checkSlides} from '../src/pptx-export.mjs';
import {readPptx} from '../src/pptx.mjs';
import {started,addModule,addLesson,addBlock,submit} from './helpers/guide-harness.mjs';

globalThis.DOMParser=DOMParser;   // the program's own reader of a presentation (src/pptx.mjs) parses with the page's parser

async function guide(){
  const a=await started('Skill day');
  await addModule(a,'Practice');await addLesson(a,'Warm up','5');
  await addBlock(a,'objective','Give feedback that names an example.');
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  a.n('guide-tpl-list').children[0].children[2].children[0].onclick();
  a.type('guide-tpl-title','Skill practice');
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  return a;
}

test('three kinds of deck are offered, each with what its front and back hold',()=>{
  assert.deepEqual(DECK_TYPES.map(t=>t.id),['activity','toolkit','facilitator']);
  assert.ok(DECK_TYPES.every(t=>t.label&&t.help));
  assert.equal(deckTypeLabel('toolkit'),'Participant toolkit: techniques to try at work');
});

test('cards for one activity: a card for each step, with a follow-up on the back; a blank deck starts with one card',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const rehearsal=activityBlocks(paras).find(x=>x.title==='Scenario rehearsal');
  const deck=deckFromGuide(paras,{type:'activity',activityKey:rehearsal.key});
  assert.equal(deck.title,'Scenario rehearsal: cards');
  assert.equal(deck.activity,'Scenario rehearsal');
  assert.equal(deck.cards.length,3);
  assert.deepEqual(deck.cards[0].front,['Take the roles of performer, counterpart and observer.']);
  assert.deepEqual(deck.cards[0].back,['Follow up: What changed when you repeated the moment?']);
  const blank=deckFromGuide(paras,{type:'activity'});
  assert.equal(blank.cards.length,1);assert.equal(blank.activity,'');
  assert.throws(()=>deckFromGuide(paras,{type:'activity',activityKey:'gone#1'}),/not in the guide any more/);
  assert.throws(()=>deckFromGuide(paras,{type:'nope'}),/Choose what kind of deck/);
});

test('a toolkit has a card for each activity, naming the technique and when to try it, with short steps and a check; facilitator cards show time, set-up and a cue',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const kit=deckFromGuide(paras,{type:'toolkit'});
  assert.equal(kit.cards.length,6);
  const c=kit.cards.find(x=>x.front[0]==='Scenario rehearsal');
  assert.equal(c.front[1],'Try it when you want a rehearsed response and one revision.');
  assert.equal(c.back.length,4,'three short steps and the check');
  assert.equal(c.back[0],'Assign performer, counterpart, and observer');
  assert.match(c.back.at(-1),/^Check: What changed when you repeated the moment\?/);
  const pocket=deckFromGuide(paras,{type:'facilitator',title:'Pocket run sheet'});
  assert.equal(pocket.title,'Pocket run sheet');
  const p=pocket.cards.find(x=>x.front[0]==='Scenario rehearsal');
  assert.equal(p.front[1],'18 min · Triads');
  assert.match(p.back.join('\n'),/Set up: /);assert.match(p.back.join('\n'),/Steps: Assign performer/);assert.match(p.back.join('\n'),/Cue: What changed/);assert.match(p.back.join('\n'),/Makes: A rehearsed response and one revision/);
  const empty=await started('Nothing');
  await addModule(empty,'M');await addLesson(empty,'L','5');
  assert.throws(()=>deckFromGuide(await_(empty),{type:'toolkit'}),/no activities/);
});
const await_=a=>a.adapter.snapshot().map(r=>({text:r.text,style:r.style}));

test('every deck made from the guide goes through the print sheet writer, fits its cards, and keeps its words',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  for(const type of ['toolkit','facilitator']){
    const deck=deckFromGuide(paras,{type});
    const stored={...deck,paper:'A4',perPage:4,cropMarks:true,flip:'long',sequence:'alternate'};
    const model=printModel(stored);
    assert.deepEqual(cardFit(model),[],`${type} cards fit six to a page`);
    const files=cardSheetFiles(model);
    const doc=files.find(f=>f.name==='word/document.xml').text;
    assert.match(doc,/Scenario rehearsal/);
    assert.ok(checkCards(model).cards.length===deck.cards.length);
  }
  assert.equal(deckLine({title:'  Scenario   cards '}),'Card deck: Scenario cards');
  assert.deepEqual(printModel({title:'T',type:'activity',paper:'Letter',perPage:4,cropMarks:false,flip:'long',sequence:'alternate',cards:[{front:['a'],back:[]},{front:['b'],back:['c']}]}).cards,[{front:['a'],back:null},{front:['b'],back:['c']}]);
});

test('the slides: an opening, a slide for each activity with its task and time, the steps in the notes, a break, and a close; every title is different',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const m=slidesModel(paras,{start:'09:00'});
  const kinds=m.slides.map(s=>s.kind);
  assert.equal(kinds[0],'opening');assert.equal(kinds.at(-1),'close');
  assert.equal(kinds.filter(k=>k==='activity').length,6);
  assert.equal(new Set(m.slides.map(s=>s.title)).size,m.slides.length,'no two slides share a title');
  assert.equal(m.slides[0].title,'Skill day');
  assert.match(m.slides[0].lines.join('\n'),/Agenda: Outcome and agenda briefing, Model and notice, Scenario rehearsal/);
  const s=m.slides.find(x=>/^Scenario rehearsal/.test(x.title));
  assert.equal(s.title,'Scenario rehearsal (09:22, 18 min)');
  assert.equal(s.key,'scenario rehearsal#1');
  assert.deepEqual(s.lines.slice(0,3),['Take the roles of performer, counterpart and observer.','Run the scenario while the observer records evidence.','Give and hear feedback, then repeat one moment.']);
  assert.ok(s.lines.includes('Grouping: Triads'));assert.ok(s.lines.includes('Time: 18 min, 09:22 to 09:40'));assert.ok(s.lines.includes('Make: A rehearsed response and one revision'));
  assert.equal(s.notes[0],'09:22 3 min: Assign performer, counterpart, and observer');
  assert.ok(s.notes.some(n=>/^Supplies: Scenario, Cue card, Observation checklist\./.test(n)));
  assert.ok(s.notes.some(n=>/^Debrief: What changed/.test(n)));
  assert.ok(s.notes.some(n=>/^Other ways to take part: Offer a written simulation\./.test(n)));
  assert.equal(slideFit(m).length,0,'nothing too crowded to read');
  assert.doesNotThrow(()=>checkSlides(m));
});

test('a break slide says when to come back and what comes next; the slides are one deck that a presentation program opens and the program reads back with the same names',async()=>{
  const a=await started('Long day');
  await addModule(a,'M');await addLesson(a,'L','5');
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  const tdx=a.n('guide-tpl-list').children.findIndex(li=>/Training Transfer Planning/.test(li.children[0].textContent));
  a.n('guide-tpl-list').children[tdx].children[2].children[0].onclick();
  await submit(a,'guide-tpl-apply-form');
  const paras=await a.adapter.paragraphs();
  const m=slidesModel(paras,{start:'09:00'});
  const brk=m.slides.find(x=>x.kind==='break');
  assert.ok(brk);
  assert.match(brk.title,/^Break \(\d\d:\d\d, 10 min\)$/);
  assert.match(brk.lines[0],/^Back at \d\d:\d\d$/);
  assert.ok(brk.notes.some(n=>/^Next: /.test(n)));
  const bytes=await pptxBytes(m,JSZip);
  assert.ok(bytes.length>3000);
  const back=await readPptx(bytes,'deck.pptx',JSZip);
  assert.equal(back.slides.length,m.slides.length);
  assert.deepEqual(back.slides.map(s=>s.title),m.slides.map(s=>s.title));
});
