// Bringing changes back from a presentation to the activities (src/session-slide-import.mjs): slides the program made are read back, matched to their activities by name, and
// only the differences are proposed; a conflict is flagged and a slide that cannot be matched is never applied.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {slideName,readSlide,planSlideImport,planSummary,activityRewrite} from '../src/session-slide-import.mjs';
import {slidesModel} from '../src/session-decks.mjs';
import {pptxBytes} from '../src/pptx-export.mjs';
import {readPptx} from '../src/pptx.mjs';
import {readBlock} from '../src/guide-model.mjs';
import {fingerprints} from '../src/session-record.mjs';
import {makeContext,cueBlock} from '../src/guide-ooxml.mjs';
import {started,addModule,addLesson,submit} from './helpers/guide-harness.mjs';

globalThis.DOMParser=DOMParser;

async function guide(){
  const a=await started('Skill day');
  await addModule(a,'Practice');await addLesson(a,'Warm up','5');
  a.p.useTool('sessionTemplates');await a.p.openTool('sessionTemplates');
  a.n('guide-tpl-list').children[0].children[2].children[0].onclick();
  a.type('guide-tpl-title','Skill practice');
  await submit(a,'guide-tpl-apply-form');
  assert.equal(a.err(),'');
  return a;
}
// The slides of the guide, written to a file, optionally edited on the way (as the author would edit them in a presentation program), and read back as the program reads a deck.
async function deckOf(paras,edit=m=>m){
  const model=edit(structuredClone(slidesModel(paras,{start:'09:00'})));
  const read=await readPptx(await pptxBytes(model,JSZip),'Skill day.pptx',JSZip);
  return read.slides;
}
const slideOf=(m,start)=>m.slides.find(s=>s.title.startsWith(start));

test('a slide title gives the activity\'s name without its time, and a slide is read into the parts it came from',()=>{
  assert.equal(slideName('Scenario rehearsal (09:22, 18 min)'),'Scenario rehearsal');
  assert.equal(slideName('Scenario rehearsal (Day 2 · 09:22, 18 min)'),'Scenario rehearsal');
  assert.equal(slideName('Close'),'Close');
  const r=readSlide({title:'Scenario rehearsal (09:22, 18 min)',text:'Scenario rehearsal (09:22, 18 min)\nTake the roles.\nRun it.\nGrouping: Triads\nTime: 18 min, 09:22 to 09:40\nMake: A rehearsed response and one revision',
    notes:'09:22 3 min: Assign roles\n09:25 6 min: Run it\nSupplies: Scenario, Cue card.\nSet up: Print the cards.\nDebrief: What changed?\nIf something goes wrong: Pause it.\nOther ways to take part: Write it.'});
  assert.deepEqual(r,{name:'Scenario rehearsal',title:'Scenario rehearsal (09:22, 18 min)',task:['Take the roles.','Run it.'],grouping:'Triads',output:'A rehearsed response and one revision',
    steps:['3 min: Assign roles','6 min: Run it'],materials:'Scenario, Cue card',preparation:'Print the cards.',debrief:['What changed?'],challenges:['Pause it.'],access:['Write it.']});
});

test('slides that have not been edited propose nothing: every slide matches its activity, and the plan says so',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const plan=planSlideImport(paras,await deckOf(paras));
  assert.deepEqual(plan.changes,[]);
  assert.equal(plan.unchanged.length,6);
  assert.equal(plan.unmatched.length,2,'the opening slide and the close are not an activity\'s');
  assert.deepEqual(plan.unmatched.map(u=>u.title),['Skill day','Close']);
  assert.equal(planSummary(plan),'Nothing in the presentation differs from the guide. 2 slides were not matched to an activity.');
});

test('an edit to a slide\'s task, its grouping and its notes is proposed, part by part, against what the guide says now',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const slides=await deckOf(paras,m=>{
    const s=slideOf(m,'Scenario rehearsal');
    s.lines[0]='Take your roles: performer, counterpart and observer.';
    s.lines=s.lines.map(l=>l==='Grouping: Triads'?'Grouping: Groups of three':l);
    s.notes=s.notes.map(n=>/^09:25 6 min/.test(n)?'09:25 6 min: Run the scenario and write down what you see':n);
    s.notes.push('Debrief: What would you do differently?');
    return m;
  });
  const plan=planSlideImport(paras,slides,{previous:null});
  assert.equal(plan.matched.length,1);
  const m=plan.matched[0];
  assert.equal(m.activity.title,'Scenario rehearsal');assert.equal(m.slide.title,'Scenario rehearsal (09:22, 18 min)');
  assert.deepEqual(m.changes.map(c=>c.part),['instructions','grouping','steps','debrief']);
  const byPart=Object.fromEntries(m.changes.map(c=>[c.part,c]));
  assert.equal(byPart.instructions.to.split('\n')[0],'Take your roles: performer, counterpart and observer.');
  assert.equal(byPart.instructions.to.split('\n').length,3,'the other two lines are kept');
  assert.equal(byPart.grouping.from,'Triads');assert.equal(byPart.grouping.to,'Groups of three');
  assert.match(byPart.steps.to,/2 min: Run the scenario and write down|6 min: Run the scenario and write down what you see/);
  assert.equal(byPart.debrief.to,'What changed when you repeated the moment?\nWhat would you do differently?'.replace('What changed when you repeated the moment?','What changed when you repeated the moment?'));
  assert.ok(m.changes.every(c=>!c.conflict),'no record of the slides, so no conflict is claimed');
  assert.equal(planSummary(plan),'4 changes in 1 activity. 2 slides were not matched to an activity.');
});

test('a part the guide also changed since the slides were made is a conflict, not ticked by default, and a part only the slides changed is not',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const made=fingerprints(paras,'slides');                // what the slides were made from
  const slides=await deckOf(paras,m=>{const s=slideOf(m,'Scenario rehearsal');s.lines[0]='Slide wording for the first step.';s.notes.push('Debrief: A new prompt from the slides.');return m;});
  // the guide changes the first instruction after the slides were made
  const edited=(()=>{
    const at=paras.findIndex(p=>/Scenario rehearsal \(18 min\)/.test(p.text));
    const b=readBlock(paras,at);
    const next=structuredClone(paras);
    const rewrite=activityRewrite(paras,'scenario rehearsal#1',{instructions:'Guide wording for the first step.\nRun the scenario while the observer records evidence.\nGive and hear feedback, then repeat one moment.'});
    const ctx=makeContext();cueBlock(rewrite.args,ctx);
    return {rewrite,b,next};
  })();
  assert.equal(edited.rewrite.args.parts.instructions.split('\n')[0],'Guide wording for the first step.');
  // Rebuild the paragraphs as the guide would hold them after that edit, through the adapter, so the fingerprints are the guide's own.
  const span=edited.rewrite.span;
  await a.adapter.insert(ctx=>cueBlock(edited.rewrite.args,ctx),{where:'after',after:span.end,select:'last',remove:{start:span.start,end:span.end,guard:edited.rewrite.guard}});
  const now=await a.adapter.paragraphs();
  const plan=planSlideImport(now,slides,{previous:made});
  const m=plan.matched[0];
  const byPart=Object.fromEntries(m.changes.map(c=>[c.part,c]));
  assert.equal(byPart.instructions.conflict,true,'both the guide and the slide changed the instructions');
  assert.equal(byPart.debrief.conflict,false,'only the slide changed the debrief prompts');
  assert.match(planSummary(plan),/1 of them also changed in the guide since the slides were made/);
});

test('a slide with a new title, or one that is not an activity\'s, is listed as not matched and nothing is proposed from it',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const slides=await deckOf(paras,m=>{const s=slideOf(m,'Scenario rehearsal');s.title='Rehearsal, renamed on the slide (09:22, 18 min)';s.lines[0]='Changed.';return m;});
  const plan=planSlideImport(paras,slides);
  assert.ok(plan.unmatched.some(u=>/renamed on the slide/.test(u.title)));
  assert.equal(plan.matched.length,0);
  assert.deepEqual(plan.changes,[]);
});

test('accepted parts are written into the activity and nothing else is lost; an activity that is gone or not written as parts is refused in words',async()=>{
  const a=await guide();
  const paras=await a.adapter.paragraphs();
  const rewrite=activityRewrite(paras,'scenario rehearsal#1',{grouping:'Groups of three',debrief:'What changed?\nWhat next?'});
  assert.equal(rewrite.args.parts.grouping,'Groups of three');
  assert.equal(rewrite.args.parts.steps.split('\n').length,4,'the steps are as they were');
  assert.equal(rewrite.args.parts.source,'Scenario rehearsal (A13, version 1)');
  await a.adapter.insert(ctx=>cueBlock(rewrite.args,ctx),{where:'after',after:rewrite.span.end,select:'last',remove:{start:rewrite.span.start,end:rewrite.span.end,guard:rewrite.guard}});
  const now=await a.adapter.paragraphs();
  const b=readBlock(now,now.findIndex(p=>/Scenario rehearsal \(18 min\)/.test(p.text)));
  assert.equal(b.parts.grouping,'Groups of three');assert.equal(b.parts.debrief,'What changed?\nWhat next?');assert.equal(b.parts.phase,'Practice');
  assert.throws(()=>activityRewrite(now,'gone#1',{}),/not in the guide any more/);
});
