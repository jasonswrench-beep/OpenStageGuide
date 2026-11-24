// Card decks and slide decks made from the guide's activities (milestone 3 of the product spec of 2026-11-22). Pure functions over the guide's paragraphs: the cards of a deck
// started from an activity, and the slides of a deck made from the agenda, each as the plain model the writers take (src/card-print.mjs and src/pptx-export.mjs).
//
// A deck stays attached to the activity that supplies its content: a card deck is started from an activity (or all of them) and remembers which, and a slide carries the name of the
// activity it came from, so changes can be matched to it later. Nothing here changes the guide.
import {guideAgenda} from './guide-agenda.mjs';
import {readBlock} from './guide-model.mjs';
import {readSteps} from './session-text.mjs';
import {activityBlocks} from './session-record.mjs';

export const DECK_TYPES=Object.freeze([
  {id:'activity',label:'Prompt or scenario cards for one activity',help:'The front is a prompt, a role, a scenario or an item to sort. The back may hold a follow-up question.'},
  {id:'toolkit',label:'Participant toolkit: techniques to try at work',help:'The front names the technique and when to try it. The back gives short steps and a way to check it worked.'},
  {id:'facilitator',label:'Facilitator cards: a pocket run sheet',help:'The front shows the activity, its time and how people work. The back shows the set-up, the steps, a cue and what it makes.'}
]);
export const deckTypeLabel=id=>DECK_TYPES.find(t=>t.id===id)?.label||id;

const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();
const lines=t=>String(t??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
const clip=(t,n=300)=>{const s=oneLine(t);return s.length>n?`${s.slice(0,n-1).trimEnd()}…`:s;};
const sentence=t=>oneLine(t).replace(/[.:;,\s]+$/,'');
const lower=t=>{const s=sentence(t);return s?s[0].toLowerCase()+s.slice(1):'';};

// ---- card decks
/**
 * The cards of a new deck. `activityKey` is a key of activityBlocks (src/session-record.mjs); with none, the deck is every activity (toolkit and facilitator) or one blank card.
 * Returns {title, type, activity, cards:[{front,back}]}, where `activity` is the title of the activity it came from, for the link back.
 */
export function deckFromGuide(paras,{type='activity',activityKey='',title=''}={}){
  if(!DECK_TYPES.some(t=>t.id===type))throw new Error('Choose what kind of deck to make.');
  const all=activityBlocks(paras),one=activityKey?all.find(a=>a.key===activityKey):null;
  if(activityKey&&!one)throw new Error('That activity is not in the guide any more. Choose another, or start a blank deck.');
  const chosen=one?[one]:all;
  let cards=[];
  if(type==='activity'){
    if(!one)cards=[{front:['Write the first card here.'],back:[]}];
    else{
      const p=one.parts,fronts=lines(p.caseStudy).length?lines(p.caseStudy):lines(p.instructions),follow=lines(p.debrief);
      cards=fronts.slice(0,120).map((f,i)=>({front:[clip(f)],back:follow.length?[`Follow up: ${clip(follow[i%follow.length],280)}`]:[]}));
    }
  }else if(type==='toolkit'){
    cards=chosen.map(a=>{
      const p=a.parts,steps=readSteps(p.steps||'').steps.map(s=>s.text),how=(steps.length?steps:lines(p.instructions)).slice(0,3),out=sentence(p.output);
      return {front:[clip(a.title),out?clip(`Try it when you want ${lower(out)}.`):'Try it when it fits.'],
        back:[...how.map(s=>clip(s)),clip(`Check: ${lines(p.debrief)[0]||(out?`Did you make ${lower(out)}?`:'Did it work?')}`)]};
    });
  }else{
    cards=chosen.map(a=>{
      const p=a.parts,steps=readSteps(p.steps||'').steps.map(s=>s.text),setup=lines(p.preparation)[0]||lines(p.materials).join(', ');
      return {front:[clip(a.title),clip([a.minutes?`${a.minutes} min`:'',sentence(p.grouping)].filter(Boolean).join(' · ')||'Time and grouping to add')],
        // Pocket cards are small: each line is kept short, and a longer one is ended with an ellipsis (the print check says if a card is still too full for the size chosen).
        back:[setup?clip(`Set up: ${setup}`,60):'',steps.length?clip(`Steps: ${steps.slice(0,2).join('; ')}`,60):'',lines(p.debrief)[0]?clip(`Cue: ${lines(p.debrief)[0]}`,60):'',p.output?clip(`Makes: ${sentence(p.output)}`,60):''].filter(Boolean)};
    });
  }
  if(!cards.length)throw new Error('The guide has no activities to make cards from yet. Start a blank deck, or add an activity first.');
  const name=oneLine(title)||(one?`${one.title}: cards`:type==='toolkit'?'Toolkit cards':type==='facilitator'?'Facilitator cards':'Cards');
  return {title:name.slice(0,200),type,activity:one?one.title:'',cards:cards.slice(0,120)};
}
/** The model the print sheet takes from a stored deck (a card with no back words has no back). */
export const printModel=deck=>({title:deck.title,type:deck.type,paper:deck.paper,perPage:deck.perPage,cropMarks:deck.cropMarks,flip:deck.flip,sequence:deck.sequence,
  cards:deck.cards.map(c=>({front:c.front,back:c.back&&c.back.length?c.back:null}))});
/** The line an activity's Materials part gets when the author names the deck there. */
export const deckLine=deck=>`Card deck: ${oneLine(deck.title)}`;

// ---- slides from the agenda
const SOURCE_NAME=(key,n)=>String(key||`slide-${n}`).slice(0,80);
/**
 * The slides of the guide as the agenda has it: an opening slide, a slide for each activity and each break, and a closing slide. Each activity's slide shows its task, grouping, time and
 * what it makes; its speaker notes hold the clock time, the steps with their times, the supplies, the debrief prompts, what often goes wrong and other ways to take part.
 */
export function slidesModel(paras,{start='09:00',title=''}={}){
  const agenda=guideAgenda(paras,start);
  const keyOf=new Map(activityBlocks(paras).map(a=>[a.index,a.key]));
  const rows=agenda.rows.filter(r=>r.kind==='block'&&(r.cue==='activity'||r.cue==='break'));
  const slides=[];
  let breaks=0;
  const detail=r=>{const b=r.cue==='activity'?readBlock(paras,r.index):null;return {b,p:b?.parts||{}};};
  const activities=rows.filter(r=>r.cue==='activity'&&detail(r).p.phase!=='Break');
  const names=activities.map(r=>r.text.replace(/^[^:]+:\s*/,''));
  const objective=paras.map((p,i)=>i).find(i=>/^[^A-Za-z]*Objective\b/.test(String(paras[i].text))&&paras[i].style&&/cue/i.test(paras[i].style));
  const goal=objective===undefined?'':oneLine(String(readBlock(paras,objective)?.text||'')).split('\n')[0];
  const access=[...new Set(rows.flatMap(r=>lines(detail(r).p.access)))].slice(0,3);
  slides.push({kind:'opening',title:oneLine(title)||agenda.title||'Session',key:'opening',
    lines:[goal?`Leave with: ${sentence(goal.replace(/^Participants leave with:\s*/i,''))}`:'',names.length?`Agenda: ${names.slice(0,6).join(', ')}${names.length>6?` and ${names.length-6} more`:''}`:''].filter(Boolean).map(l=>clip(l)),
    notes:['Welcome people, say what the session is for and how they can take part.',...access.map(a=>clip(`Ways to take part: ${sentence(a)}.`,580))]});
  rows.forEach((r,i)=>{
    const {p,b}=detail(r),name=r.text.replace(/^[^:]+:\s*/,''),isBreak=r.cue==='break'||p.phase==='Break';
    if(isBreak){
      breaks++;
      const next=rows.slice(i+1).find(x=>x.cue==='activity');
      const nextPrep=next?lines(detail(next).p.preparation)[0]:'';
      slides.push({kind:'break',title:clip(`Break (${r.start}, ${r.minutes} min)`,190),key:`break#${breaks}`,lines:[`Back at ${r.end}`],
        notes:[next?`Next: ${next.text.replace(/^[^:]+:\s*/,'')}.`:'',nextPrep?clip(`Set up: ${sentence(nextPrep)}.`,580):'','Say when to come back in text and aloud.'].filter(Boolean)});
      return;
    }
    const steps=readSteps(p.steps||'');
    let at=(()=>{const m=/(\d{1,2}):(\d{2})/.exec(r.start);return m?Number(m[1])*60+Number(m[2]):0;})();
    const stepNotes=steps.steps.map(s=>{const hh=`${String(Math.floor(at/60)%24).padStart(2,'0')}:${String(at%60).padStart(2,'0')}`;at+=s.min;return clip(`${hh}  ${s.min} min: ${s.text}`,580);});
    slides.push({kind:'activity',title:clip(`${name} (${r.start}, ${r.minutes} min)`,190),key:SOURCE_NAME(keyOf.get(r.index),i),
      lines:[...lines(p.instructions).slice(0,3).map(l=>clip(l,140)),p.grouping?clip(`Grouping: ${sentence(p.grouping)}`,140):'',`Time: ${r.minutes} min, ${r.start} to ${r.end}`,p.output?clip(`Make: ${sentence(p.output)}`,140):''].filter(Boolean),
      notes:[...stepNotes,...steps.without.map(t=>clip(t,580)),
        lines(p.materials).length?clip(`Supplies: ${lines(p.materials).join(', ')}.`,580):'',lines(p.preparation).length?clip(`Set up: ${lines(p.preparation).join(' ')}`,580):'',
        ...lines(p.debrief).map(d=>clip(`Debrief: ${d}`,580)),...lines(p.challenges).map(c=>clip(`If something goes wrong: ${c}`,580)),...lines(p.access).map(a=>clip(`Other ways to take part: ${a}`,580)),
        ...(b?.notes?[]:[])].filter(Boolean).slice(0,40)});
  });
  const last=[...activities].reverse().find(r=>detail(r).p.phase==='Close'),closing=last?lines(detail(last).p.instructions).slice(0,3):[];
  slides.push({kind:'close',title:'Close',key:'close',lines:(closing.length?closing:['Write one thing you will do next','Write one question you still have']).map(l=>clip(l,140)),
    notes:['Say what happens next and who will follow up, and by when.',last&&detail(last).p.output?clip(`Evidence of completion: ${sentence(detail(last).p.output)}.`,580):'Evidence of completion: the written next steps.']});
  return {title:oneLine(title)||agenda.title||'Session',language:'en-US',slides};
}
