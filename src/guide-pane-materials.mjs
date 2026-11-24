// The Session materials tool of the guide pane (milestone 2 of the product spec of 2026-11-22): make the run sheet and the preparation checklist, send the author to the participant
// workbook, and list what has been made with its status (Suggested, Ready, Review needed or Excluded), what changed since, and a way to make just that one again. The pure work is
// in session-materials.mjs and session-record.mjs; the pane's own helpers are handed in.
import {readGuide} from './guide-model.mjs';
import {runSheetModel,runSheetFiles,runSheetFileName,prepModel,prepFiles,prepFileName} from './session-materials.mjs';
import {OUTPUT_KINDS,readSessions,sessionsText,emptySessions,recordOutput,outputStatus,describeDiff,setExcluded,removeOutput} from './session-record.mjs';
import {zipWorkbook} from './guide-workbook.mjs';

const STATUS_WORDS={Suggested:'Not made yet',Ready:'Ready: the guide has not changed since it was made','Review needed':'Review needed: the guide has changed since it was made',Excluded:'Excluded: left out on purpose'};
const oneLine=s=>String(s??'').replace(/\s+/g,' ').trim();
const when=iso=>{const d=new Date(iso);return Number.isFinite(d.getTime())?d.toISOString().slice(0,16).replace('T',' at '):'';};

/**
 * @param {{doc:any,$:(id:string)=>any,adapter:any,act:(fn:()=>Promise<void>)=>Promise<void>,status:(t:string)=>void,state:()=>{paras:any[],at:number},
 *   Zip:any,group:()=>{people?:number|null,teams?:number|null,tables?:number|null},openTool:(id:string)=>Promise<void>,focus:(id:string)=>void}} deps
 */
export function mountMaterialsTool({doc,$,adapter,act,status,state,Zip,group,openTool,focus}){
  let sessions=emptySessions();
  const ZipClass=()=>Zip||globalThis.JSZip;
  async function readRecord(){try{sessions=readSessions(await adapter.sessions?.());}catch{sessions=emptySessions();}}
  async function keep(next){
    sessions=readSessions(sessionsText(next));
    try{await adapter.setSessions?.(sessionsText(sessions));}catch{throw new Error('The program could not keep the record of what was made. The file was made all the same.');}
  }
  /** Remembers that an output was made now from the guide as it is. The workbook and deck tools call this too. */
  async function record({kind,name,scope='The whole guide',lessonTitle=null}){
    await readRecord();
    await keep(recordOutput(sessions,{kind,name,scope,lessonTitle,paras:state().paras}));
  }
  // The file is opened in a new Word window, or offered as a download when Word cannot be asked.
  async function deliver(files,name){
    const Z=ZipClass();
    if(!Z)throw new Error('The pane could not load what it needs to make a Word file. Close it, open it again, then try again.');
    const base64=await zipWorkbook(files,Z);
    try{await adapter.openDocument(base64,name);return `It is open in a new Word window. Save it as "${name}", or any name you like.`;}
    catch{adapter.offerDownload(base64,name);return `Word could not open it for you, so it was downloaded as "${name}". Open it from your Downloads folder.`;}
  }

  function scope(){
    const {paras,at}=state();
    if(!$('guide-mat-scope-lesson').checked)return {lessonIndex:null,lessonTitle:null,label:'The whole guide'};
    const lesson=readGuide(paras).modules.flatMap(m=>m.lessons).filter(l=>l.index<=Math.max(at,0)).at(-1);
    if(!lesson||at>lesson.end)throw new Error('Put your cursor inside a lesson first, or choose the whole guide.');
    return {lessonIndex:lesson.index,lessonTitle:lesson.text,label:lesson.text};
  }
  const startTime=()=>{
    const v=oneLine($('guide-mat-start').value)||'09:00';
    if(!/^([01]?\d|2[0-3]):[0-5]\d$/.test(v))throw new Error('Give the start time as a clock time such as 9:00.');
    return v;
  };
  const guideTitle=paras=>readGuide(paras).title||'Guide';

  async function makeRunSheet(options={}){
    const {paras}=state(),s=options.scope||scope(),start=options.start||startTime(),notes=options.notes??!!$('guide-mat-notes').checked;
    const model=runSheetModel(paras,{start,lessonIndex:s.lessonIndex,includeNotes:notes});
    if(!model.lessons.length)throw new Error('There are no lessons to put in a run sheet yet.');
    const how=await deliver(runSheetFiles(model),runSheetFileName(model.title));
    await record({kind:'runsheet',name:notes?'Run sheet with private notes':'Run sheet',scope:s.label,lessonTitle:s.lessonTitle});
    status(`The run sheet is made${s.lessonTitle?` for ${s.label}`:''}. ${how}`);
    await draw();
  }
  async function makeChecklist(options={}){
    const {paras}=state(),s=options.scope||scope();
    const model=prepModel(paras,{lessonIndex:s.lessonIndex,group:group()||{},paper:$('guide-mat-paper').value||'Letter'});
    const how=await deliver(prepFiles(model,{title:guideTitle(paras)}),prepFileName(guideTitle(paras)));
    await record({kind:'checklist',name:'Preparation checklist',scope:s.label,lessonTitle:s.lessonTitle});
    status(`The preparation checklist is made${model.rows.length?` with ${model.rows.length} supplies`:''}. ${how}`);
    await draw();
  }
  // Making one output again changes nothing else: the others keep the status they have, and the earlier copy of this one, which may have been edited by hand, is not touched.
  async function again(o){
    const s={lessonIndex:null,lessonTitle:o.lessonTitle,label:o.scope};
    if(o.lessonTitle){
      const lesson=readGuide(state().paras).modules.flatMap(m=>m.lessons).find(l=>l.text===o.lessonTitle);
      if(!lesson)throw new Error(`${o.lessonTitle} is not in the guide any more, so this cannot be made again. Forget it, or make a new one for the whole guide.`);
      s.lessonIndex=lesson.index;
    }
    if(o.kind==='runsheet')return makeRunSheet({scope:s,start:startTime(),notes:/private notes/i.test(o.name)});
    if(o.kind==='checklist')return makeChecklist({scope:s});
    await openTool(o.kind==='workbook'?'participantWorkbook':'sessionDecks');
  }

  function changesOf(o){
    const {status:st,diff}=outputStatus(o,state().paras);
    $('guide-mat-changes-title').textContent=`${o.name}: ${st}`;
    $('guide-mat-changes').replaceChildren(...(diff?describeDiff(diff):[STATUS_WORDS[st]]).map(t=>{const li=doc.createElement('li');li.textContent=t;return li;}));
    $('guide-mat-changes-box').hidden=false;
    status(`${o.name}: ${st}.`);
    focus('guide-mat-changes-title');
  }
  async function toggle(o){await readRecord();await keep(setExcluded(sessions,o.id,!o.excluded));await draw();status(`${o.name} is ${o.excluded?'included again':'left out on purpose'}.`);}
  async function forget(o){await readRecord();await keep(removeOutput(sessions,o.id));$('guide-mat-changes-box').hidden=true;await draw();status(`${o.name} was forgotten. The file you made from it is not touched.`);}

  function item(o){
    const li=doc.createElement('li'),title=doc.createElement('h3'),note=doc.createElement('p'),buttons=doc.createElement('div');
    li.className='guide-lib-item';
    const {status:st}=outputStatus(o,state().paras);
    title.id=`guide-mat-${o.id}-title`;title.textContent=`${OUTPUT_KINDS[o.kind]}: ${o.name}`;
    note.className='hint';note.textContent=`${o.scope}. ${o.builtAt?`Made ${when(o.builtAt)}. `:''}${STATUS_WORDS[st]}.`;
    buttons.className='guide-actions';
    const button=(label,fn,aria)=>{const b=doc.createElement('button');b.type='button';b.textContent=label;b.setAttribute('aria-label',`${aria} ${o.name}`);b.setAttribute('aria-describedby',title.id);b.onclick=()=>act(fn);return b;};
    const direct=['runsheet','checklist'].includes(o.kind);
    buttons.append(button('What changed',async()=>changesOf(o),'What changed in'),button(direct?'Make this one again':'Open its tool',()=>again(o),direct?'Make again':'Open the tool for'),
      button(o.excluded?'Include it':'Leave it out',()=>toggle(o),o.excluded?'Include':'Leave out'),button('Forget it',()=>forget(o),'Forget'));
    li.append(title,note,buttons);
    return li;
  }

  async function draw(){
    await readRecord();
    const {paras,at}=state();
    const here=readGuide(paras).modules.flatMap(m=>m.lessons).filter(l=>l.index<=Math.max(at,0)).at(-1);
    $('guide-mat-scope-note').textContent=$('guide-mat-scope-lesson').checked?(here&&at<=here.end?`It will be made for ${here.text}.`:'Put your cursor inside a lesson first.'):'It will be made for every lesson in the guide.';
    const g=group()||{};
    $('guide-mat-counts').textContent=g.people||g.teams?`The checklist counts ${[g.people?`${g.people} people`:'',g.teams?`${g.teams} teams`:''].filter(Boolean).join(' and ')}, as the Forms tool has them.`
      :'The guide does not say how many people and teams there will be, so the checklist leaves the quantities to fill in. The Forms tool takes those numbers.';
    $('guide-mat-outputs').replaceChildren(...sessions.outputs.map(item));
    $('guide-mat-outputs-note').textContent=sessions.outputs.length?`${sessions.outputs.length} made. Each is a copy. Make one again after the guide changes, and the rest stay as they are.`
      :'Nothing has been made yet. A run sheet or a checklist made here, a participant workbook, a deck or slides will be listed with whether the guide has changed since.';
  }

  const option=(value,label)=>{const o=doc.createElement('option');o.value=value;o.textContent=label;return o;};
  $('guide-mat-paper').replaceChildren(option('Letter','Letter'),option('A4','A4'));
  $('guide-mat-paper').value='Letter';
  $('guide-mat-changes-box').hidden=true;
  for(const id of ['guide-mat-scope-guide','guide-mat-scope-lesson'])$(id).onchange=()=>draw();
  $('guide-mat-runsheet').onclick=()=>act(()=>makeRunSheet());
  $('guide-mat-checklist').onclick=()=>act(()=>makeChecklist());
  $('guide-mat-workbook').onclick=()=>act(()=>openTool('participantWorkbook'));
  return {draw,record};
}
