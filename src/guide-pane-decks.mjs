// The Decks and slides tool of the guide pane (milestone 3 of the product spec of 2026-11-22): card decks that are edited, previewed on both sides and printed as an editable Word sheet,
// and the slides of the session as a PowerPoint file. The pure work is in session-decks.mjs, card-print.mjs and pptx-export.mjs; the pane's own helpers are handed in.
import {readBlock,blockSpan} from './guide-model.mjs';
import {cueBlock} from './guide-ooxml.mjs';
import {DECK_TYPES,deckFromGuide,printModel,deckLine,slidesModel,deckTypeLabel} from './session-decks.mjs';
import {activityBlocks,readSessions,sessionsText,emptySessions,saveDeck,removeDeck} from './session-record.mjs';
import {cardFit,cardSheetFiles,deckText,deckFileName,PAPERS,PER_PAGE} from './card-print.mjs';
import {slideFit,pptxBytes,pptxName} from './pptx-export.mjs';
import {planSlideImport,planSummary,activityRewrite} from './session-slide-import.mjs';
import {zipWorkbook} from './guide-workbook.mjs';
import {toBase64} from './image-prepare.mjs';

const PPTX_TYPE='application/vnd.openxmlformats-officedocument.presentationml.presentation';
const oneLine=s=>String(s??'').replace(/\s+/g,' ').trim();
const lines=t=>String(t??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
const norm=t=>oneLine(t).toLowerCase();
const cards=n=>`${n} ${n===1?'card':'cards'}`;

/**
 * @param {{doc:any,$:(id:string)=>any,adapter:any,act:(fn:()=>Promise<void>)=>Promise<void>,status:(t:string)=>void,state:()=>{paras:any[],at:number},
 *   Zip:any,record:(o:object)=>Promise<void>,keepOpen:()=>Promise<void>,load:()=>Promise<void>,focus:(id:string)=>void}} deps
 */
export function mountDecksTool({doc,$,adapter,act,status,state,Zip,record,keepOpen,load,focus}){
  let sessions=emptySessions(),current=null,chosen=0;
  async function readRecord(){try{sessions=readSessions(await adapter.sessions?.());}catch{sessions=emptySessions();}}
  async function keep(next){
    sessions=readSessions(sessionsText(next));
    try{await adapter.setSessions?.(sessionsText(sessions));}catch{throw new Error('The program could not keep the deck. Your changes are still on this page.');}
  }
  const deckById=id=>sessions.decks.find(d=>d.id===id)||null;
  // The record is read again before every change: the materials tool writes to it too (an output made from a deck), and a stale copy would write its own over that.
  async function store(deck){await readRecord();const {sessions:next,id}=saveDeck(sessions,deck);await keep(next);current=id;return deckById(id);}

  const option=(value,label)=>{const o=doc.createElement('option');o.value=value;o.textContent=label;return o;};
  function fillSelects(){
    $('guide-deck-type').replaceChildren(...DECK_TYPES.map(t=>option(t.id,t.label)));
    $('guide-deck-paper').replaceChildren(...PAPERS.map(p=>option(p,p)));
    $('guide-deck-perpage').replaceChildren(...PER_PAGE.map(n=>option(String(n),`${n}`)));
    $('guide-deck-flip').replaceChildren(option('long','long edge (most printers)'),option('short','short edge'));
    $('guide-deck-sequence').replaceChildren(option('alternate','each front, then its back (a printer that prints both sides)'),option('separate','all the fronts, then all the backs (two passes)'));
  }
  function drawActivities(){
    const {paras}=state();
    let list=[];
    try{list=activityBlocks(paras);}catch{list=[];}
    $('guide-deck-activity').replaceChildren(option('','A blank deck (or every activity, for a toolkit or facilitator cards)'),...list.map(a=>option(a.key,`${a.title} (${a.lesson})`)));
  }

  // ---- the list of decks and the deck being edited
  function deckItem(d){
    const li=doc.createElement('li'),title=doc.createElement('h3'),note=doc.createElement('p'),buttons=doc.createElement('div');
    li.className='guide-lib-item';title.id=`guide-deck-${d.id}-title`;title.textContent=d.title;
    note.className='hint';note.textContent=`${deckTypeLabel(d.type)}. ${cards(d.cards.length)}, ${d.paper}, ${d.perPage} to a page.${d.activity?` From ${d.activity}.`:''}`;
    buttons.className='guide-actions';
    const open=doc.createElement('button');
    open.type='button';open.className='primary';open.textContent='Open';open.setAttribute('aria-label',`Open the deck ${d.title}`);open.setAttribute('aria-describedby',title.id);
    open.onclick=()=>act(async()=>openDeck(d.id));
    buttons.append(open);
    li.append(title,note,buttons);
    return li;
  }
  function drawEditor(){
    const d=current&&deckById(current);
    $('guide-deck-editor').hidden=!d;
    if(!d)return;
    $('guide-deck-editor-title').textContent=`${d.title} (${cards(d.cards.length)})`;
    $('guide-deck-paper').value=d.paper;$('guide-deck-perpage').value=String(d.perPage);$('guide-deck-flip').value=d.flip;$('guide-deck-sequence').value=d.sequence;
    $('guide-deck-marks').checked=!!d.cropMarks;
    chosen=Math.max(0,Math.min(chosen,d.cards.length-1));
    $('guide-deck-cards').replaceChildren(...d.cards.map((c,i)=>{
      const li=doc.createElement('li'),b=doc.createElement('button');
      b.type='button';b.className='guide-link';b.textContent=`${c.front[0]}${c.back.length?' (has a back)':''}`;
      b.setAttribute('aria-label',`Card ${i+1} of ${d.cards.length}: ${c.front[0]}`);if(i===chosen)b.setAttribute('aria-current','true');
      b.onclick=()=>{chosen=i;$('guide-deck-front').dirty=false;drawEditor();return Promise.resolve();};
      li.append(b);return li;
    }));
    const card=d.cards[chosen];
    if(card&&!$('guide-deck-front').dirty){$('guide-deck-front').value=card.front.join('\n');$('guide-deck-back').value=card.back.join('\n');}
    drawPreview(card);
    const fit=cardFit(printModel(d));
    $('guide-deck-fit').replaceChildren(...(fit.length?fit.map(f=>f.message):['Every card fits at this size.']).map(t=>{const li=doc.createElement('li');li.textContent=t;return li;}));
  }
  function drawPreview(card){
    const side=(label,text)=>{const p=doc.createElement('p'),b=doc.createElement('strong'),t=doc.createElement('span');b.textContent=`${label}: `;t.textContent=text;p.append(b,t);return p;};
    $('guide-deck-preview').replaceChildren(side('Front',card?card.front.join(' / '):''),side('Back',card&&card.back.length?card.back.join(' / '):'blank'));
  }
  async function openDeck(id){
    await readRecord();
    current=id;chosen=0;$('guide-deck-textview').hidden=true;
    drawEditor();focus('guide-deck-editor-title');
  }

  async function newDeck(){
    await readRecord();
    const title=oneLine($('guide-deck-title').value),{paras}=state();
    const made=deckFromGuide(paras,{type:$('guide-deck-type').value,activityKey:$('guide-deck-activity').value,title});
    const saved=await store({...made,paper:'Letter',perPage:4,cropMarks:false,flip:'long',sequence:'alternate'});
    $('guide-deck-title').value='';
    chosen=0;
    await draw();
    status(`${saved.title} is started with ${cards(saved.cards.length)}. Edit them below, then make the print sheet.`);
    focus('guide-deck-editor-title');
  }
  const read=()=>{const front=lines($('guide-deck-front').value);if(!front.length)throw new Error('Write the front of the card first.');return {front,back:lines($('guide-deck-back').value)};};
  async function change(fn,message){
    await readRecord();
    const d=current&&deckById(current);
    if(!d)throw new Error('Open a deck first.');
    const next=structuredClone(d);fn(next);
    await store(next);
    $('guide-deck-front').dirty=false;
    await draw();
    if(message)status(message(next));
  }
  const addCard=()=>change(d=>{if(d.cards.length>=120)throw new Error('A deck holds up to 120 cards.');d.cards.push(read());chosen=d.cards.length-1;},d=>`Card ${d.cards.length} was added.`);
  const updateCard=()=>change(d=>{d.cards[chosen]=read();},()=>`Card ${chosen+1} was changed.`);
  const removeCard=()=>change(d=>{if(d.cards.length<2)throw new Error('A deck needs at least one card. Delete the deck instead.');d.cards.splice(chosen,1);},d=>`The card was removed. ${cards(d.cards.length)} left.`);
  const move=by=>change(d=>{const to=chosen+by;if(to<0||to>=d.cards.length)throw new Error(by<0?'This card is the first one.':'This card is the last one.');[d.cards[chosen],d.cards[to]]=[d.cards[to],d.cards[chosen]];chosen=to;},()=>`The card is now number ${chosen+1}.`);
  const setting=()=>change(d=>{d.paper=$('guide-deck-paper').value;d.perPage=Number($('guide-deck-perpage').value);d.flip=$('guide-deck-flip').value;d.sequence=$('guide-deck-sequence').value;d.cropMarks=!!$('guide-deck-marks').checked;});

  async function printSheet(){
    const d=current&&deckById(current);
    if(!d)throw new Error('Open a deck first.');
    const model=printModel(d),over=cardFit(model);
    const files=cardSheetFiles(model);
    const Z=Zip||globalThis.JSZip;
    if(!Z)throw new Error('The pane could not load what it needs to make a Word file. Close it, open it again, then try again.');
    const base64=await zipWorkbook(files,Z),name=deckFileName(d.title,d.paper);
    let how;
    try{await adapter.openDocument(base64,name);how=`It is open in a new Word window. Save it as "${name}", or any name you like.`;}
    catch{adapter.offerDownload(base64,name);how=`Word could not open it for you, so it was downloaded as "${name}". Open it from your Downloads folder.`;}
    await record({kind:'cards',name:`Card deck: ${d.title}`});
    status(`The print sheet is made on ${d.paper} paper, ${d.perPage} to a page.${over.length?` ${over.length===1?'One side is':`${over.length} sides are`} too full for that size: see the list under the card.`:''} ${how}`);
  }
  async function showText(){
    const d=current&&deckById(current);
    if(!d)throw new Error('Open a deck first.');
    $('guide-deck-textview').textContent=deckText(printModel(d));
    $('guide-deck-textview').hidden=false;
    status('The deck is shown as plain text under the buttons.');
  }
  // The deck is named in the Materials of the activity it came from, so the run sheet and the preparation checklist list it.
  async function linkDeck(){
    const d=current&&deckById(current);
    if(!d)throw new Error('Open a deck first.');
    if(!d.activity)throw new Error('This deck was not started from an activity, so there is no activity to name it in.');
    const list=await adapter.paragraphs();
    const a=activityBlocks(list).find(x=>norm(x.title)===norm(d.activity));
    if(!a)throw new Error(`${d.activity} is not in the guide any more.`);
    const span=blockSpan(list,a.index),b=span&&readBlock(list,span.start);
    if(!b||!b.parts)throw new Error(`${d.activity} is not written as labelled parts, so a line cannot be added to its materials. Edit it first.`);
    const line=deckLine(d),materials=lines(b.parts.materials);
    if(materials.includes(line)){status(`${d.activity} already names this deck in its materials.`);return;}
    const parts={...b.parts,materials:[...materials,line].join('\n')};
    await adapter.insert(ctx=>cueBlock({cue:'activity',text:b.text.split('\n')[0],minutes:b.minutes||undefined,audience:b.audience,parts,notes:b.notes},ctx),{where:'after',after:span.end,select:'last',remove:{start:span.start,end:span.end,guard:b.firstText}});
    await keepOpen();await load();
    status(`${d.activity} now lists "${line}" in its materials. To take it back, choose Edit, then Undo.`);
  }
  async function forgetDeck(){
    await readRecord();
    const d=current&&deckById(current);
    if(!d)throw new Error('Open a deck first.');
    await keep(removeDeck(sessions,d.id));
    current=null;await draw();
    status(`${d.title} was deleted. A print sheet already made from it is not touched.`);
  }

  // ---- the slides
  function slides(){
    const start=oneLine($('guide-slide-start').value)||'09:00';
    if(!/^([01]?\d|2[0-3]):[0-5]\d$/.test(start))throw new Error('Give the start time as a clock time such as 9:00.');
    return slidesModel(state().paras,{start});
  }
  function drawSlides(){
    let model=null,problem='';
    try{model=slides();}catch(e){problem=e.message;}
    const li=t=>{const x=doc.createElement('li');x.textContent=t;return x;};
    if(!model||model.slides.length<3){
      $('guide-slide-list').replaceChildren();$('guide-slide-fit').replaceChildren();
      $('guide-slide-note').textContent=problem||'There are no activities to make slides from yet. Add a lesson with activities first.';
      $('guide-slide-make').disabled=true;
      return;
    }
    $('guide-slide-make').disabled=false;
    $('guide-slide-note').textContent=`${model.slides.length} slides, in this order. The time is on each slide as words, and the notes hold the steps and the debrief.`;
    $('guide-slide-list').replaceChildren(...model.slides.map((s,i)=>li(`${i+1}. ${s.title}`)));
    const fit=slideFit(model);
    $('guide-slide-fit').replaceChildren(...(fit.length?fit.map(f=>li(f.message)):[li('Every slide is short enough to read.')]));
  }
  async function makeSlides(){
    const model=slides(),Z=Zip||globalThis.JSZip;
    if(!Z)throw new Error('The pane could not load what it needs to make a PowerPoint file. Close it, open it again, then try again.');
    const bytes=await pptxBytes(model,Z),name=pptxName(model.title);
    adapter.offerDownload(toBase64(bytes),name,PPTX_TYPE);
    await record({kind:'slides',name:`Slides: ${model.title}`});
    const over=slideFit(model);
    status(`The slides were downloaded as "${name}": ${model.slides.length} slides with speaker notes.${over.length?` ${over.length===1?'One slide is':`${over.length} slides are`} crowded: see the list.`:''} Open it from your Downloads folder.`);
  }

  // ---- bringing changes back from a presentation (session-slide-import.mjs): the author ticks what to take, and only then is anything written
  let plan=null,ticked=new Set(),importName='';
  const shorten=t=>{const s=oneLine(t);return s.length>70?`${s.slice(0,69).trimEnd()}…`:s;};
  function drawImport(){
    const box=$('guide-slide-changes'),li=t=>{const x=doc.createElement('li');x.textContent=t;return x;};
    if(!plan){
      box.replaceChildren();$('guide-slide-unmatched').replaceChildren();
      $('guide-slide-apply').disabled=true;$('guide-slide-all').disabled=true;$('guide-slide-none').disabled=true;
      return;
    }
    $('guide-slide-import-note').textContent=`${importName}: ${planSummary(plan)}`;
    let n=0;
    box.replaceChildren(...plan.matched.map(m=>{
      const group=doc.createElement('div'),head=doc.createElement('h4'),list=doc.createElement('ul');
      group.className='guide-upd-item';head.textContent=`${m.activity.title} (${m.activity.lesson})`;list.className='guide-issues';
      list.replaceChildren(...m.changes.map(c=>{
        const item=doc.createElement('li'),label=doc.createElement('label'),input=doc.createElement('input'),text=doc.createElement('span');
        label.className='guide-check';input.type='checkbox';input.id=`guide-slide-change-${++n}`;input.checked=ticked.has(c.id);
        input.onchange=()=>{if(input.checked)ticked.add(c.id);else ticked.delete(c.id);syncApply();};
        text.textContent=`${c.label}: "${shorten(c.from)||'empty'}" becomes "${shorten(c.to)}"${c.conflict?'. The guide has also changed this since the slides were made.':''}`;
        label.append(input,text);item.append(label);return item;
      }));
      group.append(head,list);return group;
    }));
    $('guide-slide-unmatched').replaceChildren(...plan.unmatched.map(u=>li(`Slide ${u.order}, "${shorten(u.title)}", is not an activity's slide that this program can match, so nothing is taken from it.`)));
    $('guide-slide-all').disabled=!plan.changes.length;$('guide-slide-none').disabled=!plan.changes.length;
    syncApply();
  }
  function syncApply(){
    const n=ticked.size;
    $('guide-slide-apply').disabled=!n;
    $('guide-slide-apply').textContent=n?`Apply the ticked changes (${n})`:'Apply the ticked changes';
  }
  async function chooseImport(){
    const res=await adapter.pickFiles({kind:'slides'});
    if(!res||res.cancelled||!res.deck){status('No presentation was chosen. The guide is as it was.');return;}
    await readRecord();
    const made=sessions.outputs.find(o=>o.kind==='slides'&&!o.excluded);
    plan=planSlideImport(await adapter.paragraphs(),res.deck.slides,{previous:made?made.sources:null});
    importName=res.deck.name||'The presentation';
    ticked=new Set(plan.changes.filter(c=>!c.conflict).map(c=>c.id));
    drawImport();
    status(`${importName}: ${planSummary(plan)} Nothing in the guide has changed yet.`);
  }
  const tickAll=on=>{if(!plan)return Promise.resolve();ticked=new Set(on?plan.changes.map(c=>c.id):[]);drawImport();return Promise.resolve();};
  async function applyImport(){
    if(!plan||!ticked.size)throw new Error('Tick at least one change first.');
    const groups=new Map();
    for(const c of plan.changes)if(ticked.has(c.id)){const key=c.id.slice(0,c.id.lastIndexOf(':'));groups.set(key,{...(groups.get(key)||{}),[c.part]:c.to});}
    let done=0;const failed=[];
    for(const [key,over] of groups){
      try{
        const list=await adapter.paragraphs(),rw=activityRewrite(list,key,over);
        await adapter.insert(ctx=>cueBlock(rw.args,ctx),{where:'after',after:rw.span.end,select:'last',remove:{start:rw.span.start,end:rw.span.end,guard:rw.guard}});
        done++;
      }catch(e){failed.push(e.message);}
    }
    await keepOpen();await load();
    if(!done)throw new Error(failed[0]||'Nothing was changed.');
    plan=null;ticked=new Set();drawImport();
    status(`${done===1?'One activity was':`${done} activities were`} updated from ${importName}.${failed.length?` ${failed.length} could not be: ${failed.join(' ')}`:''} To take it back, choose Edit, then Undo, once for each activity.`);
  }

  async function draw(){
    await readRecord();
    drawActivities();
    $('guide-deck-list').replaceChildren(...sessions.decks.map(deckItem));
    $('guide-deck-list-heading').hidden=false;
    if(current&&!deckById(current))current=null;
    drawEditor();
    drawSlides();
    drawImport();
  }

  fillSelects();
  $('guide-deck-editor').hidden=true;$('guide-deck-textview').hidden=true;
  $('guide-deck-new-form').onsubmit=e=>{e?.preventDefault?.();return act(newDeck);};
  $('guide-deck-add').onclick=()=>act(addCard);
  $('guide-deck-update').onclick=()=>act(updateCard);
  $('guide-deck-remove').onclick=()=>act(removeCard);
  $('guide-deck-up').onclick=()=>act(()=>move(-1));
  $('guide-deck-down').onclick=()=>act(()=>move(1));
  for(const id of ['guide-deck-paper','guide-deck-perpage','guide-deck-flip','guide-deck-sequence','guide-deck-marks'])$(id).onchange=()=>act(setting);
  $('guide-deck-print').onclick=()=>act(printSheet);
  $('guide-deck-text').onclick=()=>act(showText);
  $('guide-deck-link').onclick=()=>act(linkDeck);
  $('guide-deck-forget').onclick=()=>act(forgetDeck);
  $('guide-deck-front').oninput=()=>{$('guide-deck-front').dirty=true;};
  $('guide-deck-back').oninput=()=>{$('guide-deck-front').dirty=true;};
  $('guide-slide-start').onchange=()=>drawSlides();
  $('guide-slide-make').onclick=()=>act(makeSlides);
  $('guide-slide-import').onclick=()=>act(chooseImport,{slow:true});
  $('guide-slide-all').onclick=()=>tickAll(true);
  $('guide-slide-none').onclick=()=>tickAll(false);
  $('guide-slide-apply').onclick=()=>act(applyImport);
  drawImport();
  return {draw,open:openDeck};
}
