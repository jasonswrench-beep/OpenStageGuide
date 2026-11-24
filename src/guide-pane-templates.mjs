// The Session templates tool of the guide pane (milestone 1 of the product spec of 2026-11-22): browse the ready-made sessions and the author's own, preview one without
// changing anything, apply a copy to the guide, and save a lesson as a template. The pure work is in session-templates.mjs; this file only draws it and puts the result in the
// document, with the pane's own helpers handed in (so it works with the same plain DOM calls the pane uses and the test DOM supports).
import {FAMILIES} from './session-catalog.mjs';
import {TEMPLATES_KEY} from './session-text.mjs';
import {lessonPlacement,modulePlacement,readGuide,lessonMinutes} from './guide-model.mjs';
import {makeContext} from './guide-ooxml.mjs';
import {allTemplates,templateById,matchTemplates,templatePreview,sessionXml,moduleXml,templateFromLesson,cleanMine,addMine,removeMine,serializeMine,groupText,TIME_BANDS,PACK_LABELS} from './session-templates.mjs';

const MAX_SAVED_CHARS=900000;
const plural=(n,word)=>`${n} ${word}${n===1?'':'s'}`;
const activities=n=>`${n} ${n===1?'activity':'activities'}`;
const oneLine=s=>String(s??'').replace(/\s+/g,' ').trim();

/**
 * @param {{doc:any,$:(id:string)=>any,adapter:any,act:(fn:()=>Promise<void>)=>Promise<void>,status:(t:string)=>void,state:()=>{paras:any[],at:number},
 *   keepOpen:()=>Promise<void>,load:()=>Promise<void>,tidy:()=>string,undo:string,storage:any,focus:(id:string)=>void}} deps
 */
export function mountTemplatesTool({doc,$,adapter,act,status,state,keepOpen,load,tidy,undo,storage,focus}){
  let mine=[],chosen=null,titleWas='';
  const loadMine=()=>{try{mine=cleanMine(JSON.parse(storage?.getItem(TEMPLATES_KEY)||'null'));}catch{mine=[];}};
  function keepMine(){
    try{
      if(!storage)return false;
      const text=serializeMine(mine);
      if(text.length>MAX_SAVED_CHARS)return false;
      storage.setItem(TEMPLATES_KEY,text);return true;
    }catch{return false;}
  }

  const option=(value,label)=>{const o=doc.createElement('option');o.value=value;o.textContent=label;return o;};
  function fillFilters(){
    $('guide-tpl-family').replaceChildren(option('','Any purpose'),...FAMILIES.map(f=>option(f,f)));
    $('guide-tpl-band').replaceChildren(option('','Any length'),...TIME_BANDS.map(b=>option(b.id,b.label)));
    $('guide-tpl-pack').replaceChildren(option('','All sets'),...Object.entries(PACK_LABELS).map(([id,label])=>option(id,label)));
  }
  const wanted=()=>({text:$('guide-tpl-find').value,family:$('guide-tpl-family').value,band:$('guide-tpl-band').value,pack:$('guide-tpl-pack').value});

  function card(t){
    const li=doc.createElement('li'),title=doc.createElement('h3'),kind=doc.createElement('span'),summary=doc.createElement('p'),buttons=doc.createElement('div');
    li.className='guide-lib-item';
    title.id=`guide-tpl-${t.id}-title`;
    title.textContent=`${t.title} (${t.minutes} min)`;
    kind.className='guide-lib-kind';kind.textContent=t.mine?'Yours':t.family;title.append(kind);
    summary.className='hint';
    summary.textContent=[[t.audience,groupText(t.group)].filter(Boolean).join(', '),t.purpose,t.output?`Makes: ${t.output}`:''].filter(Boolean).map(s=>s.replace(/\.$/,'')).join('. ')+'.';
    buttons.className='guide-actions';
    const look=doc.createElement('button');
    look.type='button';look.className='primary';look.textContent='Preview';look.setAttribute('aria-label',`Preview ${t.title}`);look.setAttribute('aria-describedby',title.id);
    look.onclick=()=>{preview(t.id);return Promise.resolve();};
    buttons.append(look);
    if(t.mine){
      const remove=doc.createElement('button');
      remove.type='button';remove.textContent='Remove';remove.setAttribute('aria-label',`Remove ${t.title} from your templates`);
      remove.onclick=()=>act(()=>removeSaved(t));
      buttons.append(remove);
    }
    li.append(title,summary,buttons);
    return li;
  }

  function placementNote(){
    const {paras,at}=state(),here=!$('guide-tpl-where-new').checked;
    try{
      if(here){const p=lessonPlacement(paras,Math.max(at,0));return `It goes after ${p.after}, as Lesson ${p.m}.${p.l}.`;}
      const p=modulePlacement(paras,Math.max(at,0));
      return `It goes after the module your cursor is in, as Module ${p.n}, with its own Transition block.`;
    }catch(e){return e.message;}
  }

  function draw(){
    const found=matchTemplates(allTemplates(mine),wanted());
    $('guide-tpl-list').replaceChildren(...found.map(card));
    const all=allTemplates(mine).length;
    $('guide-tpl-count').textContent=found.length===all?`${plural(all,'template')}${mine.length?`, ${mine.length} of them yours`:''}.`:`${found.length} of ${plural(all,'template')} found.`;
    $('guide-tpl-mine').replaceChildren(...mine.map(card));
    $('guide-tpl-mine-heading').hidden=false;
    $('guide-tpl-mine-note').textContent=mine.length?'Your templates are kept on this computer and are here in every guide you open. To change one, apply it, edit the lesson, and save it again with the same name.'
      :'Put your cursor in a lesson, give the template a name and choose Save. It keeps the lesson\'s activities, steps and prompts, and is kept on this computer for every guide you open.';
    if(chosen)drawPreview();
  }

  const para=(label,text)=>{
    const p=doc.createElement('p'),b=doc.createElement('strong'),t=doc.createElement('span');
    b.textContent=`${label}: `;t.textContent=text;p.append(b,t);return p;
  };
  const cell=(text,tag='td')=>{const c=doc.createElement(tag);c.textContent=text;return c;};
  function drawPreview(){
    const t=chosen;
    let p,startProblem='';
    try{p=templatePreview(t,{start:$('guide-tpl-start').value});}
    catch(e){p=templatePreview(t);startProblem=`${e.message} `;}
    $('guide-tpl-preview').hidden=false;
    $('guide-tpl-preview-title').textContent=`${p.title} (${p.minutes} min)`;
    $('guide-tpl-preview-meta').textContent=[PACK_LABELS[p.pack]||p.pack,p.family==='Yours'?'':p.family,p.audience,p.group].filter(Boolean).join('. ')+'. '+p.note;
    const facts=[];
    if(p.purpose)facts.push(para('What it is for',p.purpose));
    if(p.outcomes.length)facts.push(para('What the group leaves with',p.outcomes.join(' ')));
    if(p.prerequisites)facts.push(para('Before you start',p.prerequisites));
    if(p.rationale)facts.push(para('Why the activities come in this order',p.rationale));
    if(p.materials.length)facts.push(para('Materials',p.materials.join(', ')));
    if(p.adjustments.length)facts.push(para('Ways to adapt it',p.adjustments.join(' ')));
    $('guide-tpl-preview-facts').replaceChildren(...facts);
    $('guide-tpl-preview-body').replaceChildren(...p.agenda.map(r=>{
      const tr=doc.createElement('tr');
      tr.append(cell(`${r.from} to ${r.to}`),cell(r.title),cell(r.phase),cell(String(r.minutes)),cell(r.output));
      return tr;
    }));
    $('guide-tpl-where-note').textContent=`${startProblem}${placementNote()}`;
  }
  function preview(id){
    const t=templateById(id,mine);
    if(!t)return;
    chosen=t;
    // The title starts as the template's own, and follows the choice until the author types one.
    const typed=oneLine($('guide-tpl-title').value);
    if(!typed||typed===titleWas){$('guide-tpl-title').value=t.title;titleWas=t.title;}
    drawPreview();
    status(`${t.title}: ${t.minutes} minutes, ${activities(t.rows.length)}. Nothing in the guide has changed.`);
    focus('guide-tpl-preview-title');
  }
  function closePreview(){
    chosen=null;
    $('guide-tpl-preview').hidden=true;
    status('The preview is closed. Nothing in the guide was changed.');
  }

  const minutesBefore=(paras,anchor)=>readGuide(paras).modules.flatMap(m=>m.lessons).filter(l=>l.index<=anchor).reduce((s,l)=>s+lessonMinutes(l),0);
  async function apply(){
    if(!chosen)throw new Error('Choose a template and Preview it first.');
    const t=chosen,title=oneLine($('guide-tpl-title').value),audience=oneLine($('guide-tpl-audience').value),start=$('guide-tpl-start').value;
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),inModule=!$('guide-tpl-where-new').checked;
    let anchor,build,label;
    if(inModule){
      const p=lessonPlacement(list,here);
      anchor=p.anchor;label=`Lesson ${p.m}.${p.l}`;
      build=ctx=>sessionXml(t,{title,audience,start,m:p.m,l:p.l,before:p.before},ctx);
    }else{
      const p=modulePlacement(list,here);
      anchor=p.anchor;label=`Module ${p.n}`;
      build=ctx=>moduleXml(t,{title,moduleTitle:title,audience,start,m:p.n,before:minutesBefore(list,anchor)},ctx);
    }
    build(makeContext());   // the words are tried first: a title or a start time that cannot be used says so before anything is written, and what was typed stays
    await adapter.insert(build,{where:'after',after:anchor,select:'first'});
    await keepOpen();
    await load();
    status(`A copy of "${t.title}" is in the guide as ${label}: ${title}. The template itself is unchanged.${tidy()} ${undo}`);
  }

  async function saveTemplate(){
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex();
    const {template,skipped}=templateFromLesson(list,here,{name:$('guide-tpl-name').value,keepExamples:!!$('guide-tpl-keep-examples').checked,keepNotes:!!$('guide-tpl-keep-notes').checked});
    const {list:next,replaced}=addMine(mine,template);
    const before=mine;mine=next;
    if(!keepMine()){mine=before;throw new Error('This computer would not keep your template (its storage is off or full). Nothing was saved.');}
    $('guide-tpl-name').value='';
    draw();
    status(`${template.title} was ${replaced?'updated in':'saved to'} your templates, with ${activities(template.rows.length)}.${skipped?` ${activities(skipped)} with no instructions could not be kept.`:''}`);
  }
  async function removeSaved(t){
    const before=mine;mine=removeMine(mine,t.id);
    if(!keepMine()){mine=before;throw new Error('This computer would not change your saved templates. Nothing was removed.');}
    if(chosen&&chosen.id===t.id){chosen=null;$('guide-tpl-preview').hidden=true;}
    draw();
    status(`${t.title} was removed from your templates. Copies already in a guide are not touched.`);
  }

  fillFilters();
  loadMine();
  $('guide-tpl-preview').hidden=true;
  $('guide-tpl-find').oninput=draw;
  for(const id of ['guide-tpl-family','guide-tpl-band','guide-tpl-pack'])$(id).onchange=draw;
  for(const id of ['guide-tpl-where-here','guide-tpl-where-new'])$(id).onchange=()=>{$('guide-tpl-where-note').textContent=placementNote();};
  $('guide-tpl-start').onchange=()=>{if(chosen)drawPreview();};
  $('guide-tpl-apply-form').onsubmit=e=>{e?.preventDefault?.();return act(apply);};
  $('guide-tpl-close').onclick=()=>{closePreview();return Promise.resolve();};
  $('guide-tpl-save').onclick=()=>act(saveTemplate);
  draw();
  return {draw,loadMine,preview,mine:()=>mine};
}
