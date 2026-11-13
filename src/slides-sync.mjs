// Update from Slides: compares a changed presentation with the slides already in the guide and says what differs, so the author
// can choose what to bring in. Planned with the author on 2026-11-06 (docs/RIBBON-PLAN.md, "Keeping slides up to date").
//
// What the guide remembers. A slide's number, title, animation clicks, speaker notes and (if it had one) picture are in the document as
// a block (guide-slides.mjs). What a block cannot say is which slide of which presentation it came from, or whether its notes are still
// the presentation's words, so each import also leaves a small record in the document's settings: for each presentation, the permanent
// number PowerPoint keeps inside every slide (it stays when a slide moves or is renamed), the line the slide's block starts with, a
// fingerprint of the notes as imported and of the picture. A guide made before this record existed still works: slides are then matched
// by title, and by number when the title changed, and nothing is offered as new without asking first.
//
// Nothing here touches the document; the pane turns the plan into Word calls (guide-pane.mjs). Fingerprints are only for noticing a
// change, not for security.
import {kindOf,readBlock,cueFromText} from './guide-model.mjs';
import {slideTextLines} from './guide-slides.mjs';

export const MEMORY_KEY='osg.slides.v1';
const MAX_DECKS=8;

// A short fingerprint of a string or of bytes: the same input gives the same fingerprint, a different one almost surely a different one.
export function fingerprint(input){
  let h1=0xdeadbeef,h2=0x41c6ce57;
  const feed=c=>{h1=Math.imul(h1^c,2654435761);h2=Math.imul(h2^c,1597334677);};
  if(typeof input==='string')for(let i=0;i<input.length;i++)feed(input.charCodeAt(i));
  else for(let i=0;i<input.length;i++)feed(input[i]);
  h1=Math.imul(h1^(h1>>>16),2246822507);h1^=Math.imul(h2^(h2>>>13),3266489909);
  h2=Math.imul(h2^(h2>>>16),2246822507);h2^=Math.imul(h1^(h1>>>13),3266489909);
  return (4294967296*(2097151&h2)+(h1>>>0)).toString(16);
}
// Notes compare line by line, ignoring blank lines and runs of spaces, which is how a block holds them.
export const normNotes=s=>String(s??'').split(/\r?\n/).map(l=>l.replace(/\s+/g,' ').trim()).filter(Boolean).join('\n');
const normTitle=s=>String(s??'').replace(/\s+/g,' ').trim().toLowerCase();
const oneLine=s=>String(s??'').replace(/\s+/g,' ').trim();
export const notesPrint=s=>fingerprint(normNotes(s));

// ---- the record

// A presentation is known by its file name without ".pptx" and without a copy number, so "AI deck (2).pptx" is "AI deck".
export const deckKey=name=>String(name??'').toLowerCase().replace(/\.pptx$/,'').replace(/[\s_-]*\(\d+\)$/,'').replace(/\s+/g,' ').trim();

export function readMemory(text){
  try{
    const m=JSON.parse(text);
    if(m&&m.v===1&&Array.isArray(m.decks)){
      return {v:1,decks:m.decks.filter(d=>d&&typeof d.key==='string'&&Array.isArray(d.slides)).map(d=>({
        name:String(d.name??d.key),key:d.key,
        seen:(Array.isArray(d.seen)?d.seen:[]).filter(Number.isInteger),
        slides:d.slides.filter(s=>s&&Number.isInteger(s.id)&&typeof s.label==='string').map(s=>({id:s.id,label:s.label,ns:String(s.ns??''),pic:String(s.pic??''),tx:String(s.tx??'')}))
      }))};
    }
  }catch{/* nothing remembered yet, or not readable: start again */}
  return {v:1,decks:[]};
}
export const writeMemory=memory=>JSON.stringify(memory);
export const recordFor=(memory,name)=>memory.decks.find(d=>d.key===deckKey(name))||null;

// Adds what an import or an update left in the guide to the record of its presentation: `seen` is every slide the presentation had at
// that moment (so a slide the author chose not to bring in is not offered as new later), `entries` the slides that are in the guide.
export function remember(memory,{name,seen,entries}){
  const key=deckKey(name),old=recordFor(memory,name),byId=new Map((old?.slides||[]).map(s=>[s.id,s]));
  for(const e of entries)byId.set(e.id,{id:e.id,label:e.label,ns:String(e.ns??''),pic:String(e.pic??''),tx:String(e.tx??'')});
  const record={name:String(name),key,seen:[...new Set([...(old?.seen||[]),...seen])].sort((a,b)=>a-b),slides:[...byId.values()].sort((a,b)=>a.id-b.id)};
  return {v:1,decks:[record,...memory.decks.filter(d=>d.key!==key)].slice(0,MAX_DECKS)};
}

// ---- the slides in the guide

const LABEL=/^Slide\s+(\d+)\s*(?::\s*(.*))?$/s,CLICKS=/^Animation:\s*(\d+)\s+clicks?\.?$/i;
// The slide blocks of the guide: the line each starts with ("Slide 3: Title"), the animation clicks and speaker notes it holds, who
// it is shown to, and whether it carries a picture. A block the author has retyped as something else is not recognised, and not touched.
export function slideBlocksIn(paras){
  const out=[];
  for(let i=0;i<paras.length;i++){
    if(kindOf(paras[i])!=='cue')continue;
    const cue=cueFromText(paras[i].text);
    if(!cue||cue.id!=='slide')continue;
    const b=readBlock(paras,i);
    if(!b||b.span.start!==i)continue;
    const lines=b.text.split('\n'),m=LABEL.exec(lines[0].trim());
    if(!m)continue;
    const clicks=lines.slice(1).map(l=>CLICKS.exec(l.trim())).find(Boolean);
    out.push({index:i,span:b.span,firstText:b.firstText,label:lines[0].trim(),order:Number(m[1]),title:oneLine(m[2]),clicks:clicks?Number(clicks[1]):0,notes:b.notes,slideText:b.slideText||'',audience:b.audience,hasPicture:b.hasPicture});
  }
  return out;
}
// The line a slide's block starts with, as guide-slides.mjs writes it.
export const labelOf=({order,title})=>`Slide ${order}${oneLine(title)?`: ${oneLine(title)}`:''}`;

// ---- the plan

const plural=(n,w)=>`${n} ${w}${n===1?'':'s'}`;
const clicksText=n=>n?plural(n,'click'):'no clicks';
const quote=s=>s?`"${s}"`:'no title';

// What differs between the presentation as it is now and the slides in the guide.
//   blocks    the guide's slide blocks (slideBlocksIn)
//   deck      {name, slides:[{order, sourceId, title, clicks, notes}]} the changed presentation
//   pictures  Map(slide number -> fingerprint of its new picture) when the author also chose the PDF, else null
//   record    what the guide remembers of this presentation (recordFor), or null
// Returns {items, pairs, unchanged, leftOut, hasRecord}. Each item is one slide with the changes on offer; every change has `on`, whether
// it starts ticked. An item with `blocked` cannot be applied, and says why.
export function planSlideUpdate({blocks,deck,pictures=null,record=null}){
  const slides=deck.slides,claimedB=new Set(),claimedS=new Set(),pairs=[];
  const pair=(block,slide,how)=>{claimedB.add(block);claimedS.add(slide);pairs.push({block,slide,how,entry:record?.slides.find(s=>s.label===block.label)||null});};
  const removedByMemory=[];
  // 1. by what the guide remembers: the permanent number of the slide, and the line its block started with (a presentation read
  // without those numbers is matched by title and place alone)
  if(record&&slides.some(s=>Number.isInteger(s.sourceId))){
    for(const entry of record.slides){
      const block=blocks.find(b=>!claimedB.has(b)&&b.label===entry.label);
      if(!block)continue;
      const slide=slides.find(s=>!claimedS.has(s)&&s.sourceId===entry.id);
      if(slide)pair(block,slide,'memory');else removedByMemory.push(block);
    }
  }
  // 2. by title, then 3. by number (the same place in the presentation, with another title)
  for(const block of blocks){
    if(claimedB.has(block)||removedByMemory.includes(block))continue;
    const t=normTitle(block.title);
    const slide=t&&slides.find(s=>!claimedS.has(s)&&normTitle(s.title)===t);
    if(slide)pair(block,slide,'title');
  }
  for(const block of blocks){
    if(claimedB.has(block)||removedByMemory.includes(block))continue;
    const slide=slides.find(s=>!claimedS.has(s)&&s.order===block.order);
    if(slide)pair(block,slide,'number');
  }
  const items=[];let unchanged=0;
  for(const {block,slide,how,entry} of pairs){
    const changes=[];
    if(normTitle(block.title)!==normTitle(slide.title))changes.push({id:'title',text:`Title: ${quote(block.title)} to ${quote(oneLine(slide.title))}`,on:true});
    if(block.order!==slide.order)changes.push({id:'number',text:`Moved: slide ${block.order} is now slide ${slide.order}`,on:true});
    if((block.clicks||0)!==(slide.clicks||0))changes.push({id:'clicks',text:`Animation: ${clicksText(block.clicks||0)} to ${clicksText(slide.clicks||0)}`,on:true});
    if(normNotes(block.notes)!==normNotes(slide.notes)){
      // Notes the author has rewritten in Word are kept unless they tick them; notes still as the presentation had them follow it.
      const untouched=!!entry&&entry.ns!==''&&entry.ns===notesPrint(block.notes);
      changes.push({id:'notes',text:untouched?'Speaker notes changed in the presentation':'Speaker notes differ from the presentation\'s. You rewrote them in Word, so yours are kept unless you tick this',on:untouched});
    }
    // The slide's own text, under its picture. A guide made before slide text was kept has none and is left alone, unless the guide
    // remembers having had it.
    const deckText=normNotes(slideTextLines(slide)),blockText=normNotes(block.slideText);
    if(deckText!==blockText&&(blockText!==''||(entry&&entry.tx))){
      const untouched=!!entry&&entry.tx!==''&&entry.tx===notesPrint(block.slideText);
      changes.push({id:'text',text:untouched?'Slide text changed in the presentation':`Slide text differs from the presentation's. You changed it in Word, so yours is kept unless you tick this`,on:untouched});
    }
    let blocked='';
    const print=pictures?.get(slide.order);
    if(pictures){
      if(print){
        if(entry&&entry.pic){if(entry.pic!==print)changes.push({id:'picture',text:'The slide\'s picture changed',on:true});}
        else if(block.hasPicture)changes.push({id:'picture',text:'The slide\'s picture (cannot tell whether it changed)',on:false,unsure:true});
        else changes.push({id:'picture',text:'Add the slide\'s picture',on:true});
      }else if(block.hasPicture&&changes.length)blocked='The PDF has no picture for this slide, and the block has one that cannot be kept while its words change.';
    }else if(block.hasPicture&&changes.length)blocked='This slide has a picture. Choose the PDF of the presentation too, so the slide can be rebuilt with its picture.';
    if(!changes.length){unchanged++;continue;}
    items.push({key:`s${slide.order}`,kind:'changed',label:labelOf(block),block,slide,how,changes,...(blocked?{blocked}:{})});
  }
  // Slides in the presentation with no block: new ones are offered; ones the author left out of the guide before are not.
  const known=new Set([...(record?.seen||[]),...(record?.slides.map(s=>s.id)||[])]);
  let leftOut=0;
  for(const slide of slides){
    if(claimedS.has(slide))continue;
    if(record&&known.has(slide.sourceId)){leftOut++;continue;}
    items.push({key:`n${slide.order}`,kind:'added',label:labelOf(slide),slide,changes:[{id:'add',text:record?'New in the presentation: put it in the guide':'Not in the guide (it may be new, or one you left out): put it in the guide',on:!!record}]});
  }
  // Blocks that have gone from the presentation: those it is known they came from, and, when nothing is remembered, those matched to nothing.
  for(const block of blocks){
    if(claimedB.has(block))continue;
    if(record&&!removedByMemory.includes(block))continue;
    items.push({key:`r${block.index}`,kind:'removed',label:block.label,block,changes:[{id:'remove',text:'No longer in the presentation: take it out of the guide',on:false}]});
  }
  items.sort((a,b)=>(a.slide?.order??a.block.order)-(b.slide?.order??b.block.order)||(a.kind==='changed'?-1:1));
  return {items,pairs,unchanged,leftOut,hasRecord:!!record};
}

// The one-sentence summary of a plan.
export function planSummary(plan){
  const n=k=>plan.items.filter(i=>i.kind===k).length;
  const parts=[];
  if(n('changed'))parts.push(`${plural(n('changed'),'slide')} changed`);
  if(n('added'))parts.push(`${plural(n('added'),'slide')} not in the guide`);
  if(n('removed'))parts.push(`${plural(n('removed'),'slide')} no longer in the presentation`);
  const left=plan.unchanged?` ${plural(plan.unchanged,'slide')} unchanged, left alone.`:'';
  const skipped=plan.leftOut?` ${plural(plan.leftOut,'slide')} you left out of the guide earlier ${plan.leftOut===1?'was':'were'} not offered again.`:'';
  return `${parts.length?parts.join(', ')+'.':'Nothing differs.'}${left}${skipped}`;
}

// ---- from the plan to what is done

// The ticked changes as operations on the document, in the order they must be done: from the bottom of the guide up, so the position of
// everything not done yet is still right. Where an insert and a replace share a place, the insert is done first. A replace rebuilds the
// slide's whole block (words, clicks, picture, notes) with the ticked changes and the old values for the rest.
export function operations(plan){
  const ops=[],at={};
  const on=(it,id)=>it.changes.some(c=>c.id===id&&c.on);
  for(const it of plan.items){
    if(it.kind==='changed'&&!it.blocked&&it.changes.some(c=>c.on)){
      const {block:b,slide:d}=it,slide={order:on(it,'number')?d.order:b.order,title:on(it,'title')?d.title:b.title,clicks:on(it,'clicks')?d.clicks:b.clicks,notes:on(it,'notes')?d.notes:b.notes,text:on(it,'text')?slideTextLines(d):normNotes(b.slideText)};
      ops.push({type:'replace',key:it.key,at:b.span.end,span:b.span,guard:b.firstText,slide,share:b.audience==='everyone',deckOrder:d.order,sourceId:d.sourceId,wantsPicture:b.hasPicture||on(it,'picture'),notesFollowDeck:on(it,'notes')});
    }else if(it.kind==='removed'&&it.changes[0].on)ops.push({type:'delete',key:it.key,at:it.block.span.end,span:it.block.span,guard:it.block.firstText});
  }
  // New slides go after the slide before them in the presentation that is in the guide, or before the one after it.
  const adds=plan.items.filter(i=>i.kind==='added'&&i.changes[0].on).map(i=>i.slide).sort((a,b)=>a.order-b.order);
  const sorted=[...plan.pairs].sort((a,b)=>a.slide.order-b.slide.order);
  for(const s of adds){
    const before=[...sorted].reverse().find(p=>p.slide.order<s.order),after=sorted.find(p=>p.slide.order>s.order);
    const anchor=before?before.block.span.end:after?after.block.span.start-1:null;
    const op=ops.find(o=>o.type==='insert'&&o.after===anchor);
    if(op)op.slides.push(s);else ops.push({type:'insert',key:`n${s.order}`,at:anchor??Infinity,after:anchor,slides:[s]});
  }
  return ops.sort((a,b)=>(b.at-a.at)||((a.type==='insert'?0:1)-(b.type==='insert'?0:1)));
}

// What the guide should remember once the operations are done: for every slide of the presentation that is in the guide, the line its block
// starts with now, a fingerprint of the notes as the presentation has them (kept as it was when the author kept their own notes), and
// of the picture. `pictures` is the same Map the plan was made with, or null.
export function entriesAfter(plan,ops,{deck,pictures=null}){
  const replaced=new Map(ops.filter(o=>o.type==='replace').map(o=>[o.key,o]));
  const itemOf=new Map(plan.items.filter(i=>i.kind==='changed').map(i=>[i.slide,i]));
  const entries=[];
  for(const {block,slide,entry} of plan.pairs){
    const item=itemOf.get(slide),op=item&&replaced.get(item.key);
    const notesNow=op?op.slide.notes:block.notes,textNow=op?op.slide.text:block.slideText;
    entries.push({
      id:slide.sourceId,
      label:op?labelOf(op.slide):block.label,
      ns:normNotes(notesNow)===normNotes(slide.notes)?notesPrint(slide.notes):(entry?.ns??''),
      pic:op&&op.wantsPicture?(pictures?.get(slide.order)??''):(entry?.pic??''),
      tx:normNotes(textNow)===normNotes(slideTextLines(slide))?notesPrint(slideTextLines(slide)):(entry?.tx??'')
    });
  }
  for(const o of ops)if(o.type==='insert')for(const s of o.slides)entries.push({id:s.sourceId,label:labelOf(s),ns:notesPrint(s.notes),pic:pictures?.get(s.order)??'',tx:notesPrint(slideTextLines(s))});
  return {seen:deck.slides.map(s=>s.sourceId).filter(Number.isInteger),entries:entries.filter(e=>Number.isInteger(e.id))};
}
