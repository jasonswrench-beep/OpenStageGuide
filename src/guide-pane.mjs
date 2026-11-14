// The guide pane: Start a guide, Build, Review. It reads the document, shows what is there, and inserts
// ordinary Word content. It keeps no record of its own, and it only uses plain DOM properties (no classList,
// no event listeners) so the existing minimal test DOM can drive it.
import {allCues,CUE_GROUPS,AUDIENCES,ACTIVITY_PARTS,cueById,iconUrl,setCustomCues,customCueNames,checkCustomCueNames,CUSTOM_CUE_SLOTS} from './guide-cues.mjs';
import {moduleHeading,lessonHeading,timeBlock,cueBlock,startGuide as startGuideBody,sampleGuide,acknowledgementsBody,perspectiveBody,appendixBody} from './guide-ooxml.mjs';
import {cueFromText,updatePlan,guideIssues,outline,lessonMinutes,describeAt,blockAnchor,blockSpan,kindOf,transitionAnchor,lessonPlacement,modulePlacement,contentsState,readGuide,standardPageState,missingStandardPages,appendixPlacement,tidyTarget,tidyMove,tidyDeletion} from './guide-model.mjs';
import {glanceBody,glanceState,glancePlacement,glanceRows,cuesBody,cuesState,cuesPlacement,cuesUsed,moduleAt,overviewFacts,overviewBody,overviewState,overviewPlacement,overviewsSummary} from './guide-pages.mjs';
import {checkParts} from './guide-check.mjs';
import {readSettings as readFormSettings} from './guide-forms.mjs';
import {itemPlacement,sectionPlacement,itemSpan,itemTarget} from './guide-model.mjs';
import {ITEM_TYPES,itemTypeById,LEVELS,levelById,parsePairs,parseLines,parseLabels,parseCriteria,parseRubricRows,DEFAULT_RUBRIC_LEVELS} from './guide-items.mjs';
import {itemBlock,sectionBlock} from './guide-assess-ooxml.mjs';
import {readPastedQuestions,MAX_PASTE} from './guide-paste.mjs';
import {readBankFile,MAX_BANK_ADD} from './guide-bank.mjs';
import {summarize,readSource,buildWorkbook,zipWorkbook,pdfFileName,WORKBOOK_LOOKS} from './guide-workbook.mjs';
import {workbookStructure,printCheck,checkPdfTags,WORKBOOK_HUMAN_CHECKS} from './workbook-checks.mjs';
import {leakCheck} from './workbook-leak.mjs';
import {toBase64,pictureSize,PICTURE_WIDTHS,PICTURE_HELP} from './image-prepare.mjs';
import {matchPictures,slideBlock,slideTextLines,MAX_SLIDES_AT_ONCE,MAX_TOTAL_PICTURE_BYTES} from './guide-slides.mjs';
import {GUIDE_LOOKS,cleanLook,parseLook,resolveGuideLook,lookSpecs,headerPackage,footerPackage} from './guide-looks.mjs';
import {screensOf,narrationStats,elearningIssues,screenBlocks,scriptText,scriptFiles,scriptFileName,WORDS_PER_MINUTE} from './guide-elearning.mjs';
import {authorProject,authorZipName} from './guide-author.mjs';
import {toolFor,GUIDE_SECTIONS} from './guide-tools.mjs';
import {backupFileName,inspectBackup,copyProblem,backupFacts,sizeText} from './guide-backup.mjs';
import {readCheckpoint} from './checkpoint.mjs';
import {guideAgenda,agendaSummary,agendaFiles,agendaFileName,START_TIME} from './guide-agenda.mjs';
import {alignment,alignmentSummary} from './guide-alignment.mjs';
import {readVersions,writeVersions,lessonList,lessonKey,newVersion,toggleLesson,moveLesson,versionAgenda,versionSummary,parseTarget} from './guide-versions.mjs';
import {STARTER_ACTIVITIES,STORE_KEY as LIBRARY_KEY,blockFromEntry,entryFromBlock,cleanMine,addMine,removeMine,matches as libraryMatches} from './activity-library.mjs';
import {fingerprint,notesPrint,readMemory,writeMemory,recordFor,remember,slideBlocksIn,labelOf,planSlideUpdate,planSummary,operations,entriesAfter} from './slides-sync.mjs';

const BUSY_IDS=['guide-leave-block','guide-bk-save','guide-bk-file','guide-bk-open','guide-upd-choose','guide-upd-which','guide-upd-all','guide-upd-none','guide-upd-go','guide-help-sample','guide-tab-guide','guide-tab-assess','guide-tab-workbook','guide-workbook-go','guide-sec-level-one','guide-sec-level-two','guide-sec-level-three','guide-item-go','guide-paste-go','guide-bank-go','guide-item-cancel','guide-item-start-two','guide-assess-edit','guide-retry','guide-start-go','guide-sample','guide-add-module','guide-add-lesson','guide-lib-save','guide-agenda-refresh','guide-agenda-download','guide-align-refresh','guide-ver-new','guide-ver-make','guide-block-go','guide-block-cancel','guide-update','guide-build-glance','guide-build-cues','guide-build-overview','guide-open-review','guide-refresh','guide-review-run','guide-review-back','guide-review-pdf','guide-edit-selected','guide-edit-save','guide-edit-cancel','guide-move-up','guide-move-down','guide-delete','guide-delete-yes','guide-delete-no','guide-add-standard','guide-add-appendix','guide-own-save','guide-pic-choose','guide-slides-choose','guide-slides-all','guide-slides-none','guide-slides-go','guide-look-apply','guide-el-add','guide-el-script','guide-el-text','guide-el-author'];
const CHECK_YOURSELF=[
  'Does each cue read sensibly when spoken aloud?',
  'Run Word\'s own check: Review > Check Accessibility.',
  'Read one page with a screen reader before you share the guide.'
];
// A file window waits for the author, who may take minutes to find a file; its own limit (file-pick.mjs) is ten.
const SLOW_TIMEOUT_MS=15*60*1000;
const UNDO='To take it back, click in your document and press Ctrl+Z.';
// A move or a delete is done in several steps in Word (the change, the renumbering, the contents table), so one Ctrl+Z is not enough: five were needed for a move.
const UNDO_MANY='To take it back, click in your document and press Ctrl+Z several times (a move or a delete is done in several steps) until it is as it was.';
const SAMPLE_TITLE='Sample guide: Handling workplace conflict';
const parseMinutes=(raw,label='Minutes')=>{
  const t=String(raw??'').trim();
  if(!t)return 0;
  if(!/^\d{1,3}$/.test(t))throw new Error(`${label} must be a whole number, like 5.`);
  return Number(t);
};
const plural=(n,word)=>`${n} ${word}${n===1?'':'s'}`;
const countText=n=>n===1?'1 number or time':`${n} numbers or times`;
const clean=t=>String(t??'').replace(/[\r\n\u0007]+$/,'').trim();
const snippet=t=>{const s=clean(t).replace(/^[^A-Za-z0-9]+/,'');return s.length>40?`${s.slice(0,40)}...`:s;};

// soon(fn) runs fn shortly after the cursor moves; tests pass an immediate version.
// timeoutMs is how long one action may wait for Word before the buttons are given back.
// Zip is the JSZip class and xml holds the browser's DOMParser and XMLSerializer; tests pass their own. storage remembers the
// look chosen for the workbook (the browser's local storage, which may be missing or refused).
const browserStorage=()=>{try{return globalThis.localStorage||null;}catch{return null;}};
// onTool(id) is told whenever the pane itself moves to another tool (see "the open tool" below), so the strip above it can say so.
export function mountGuidePane(doc,{adapter,soon=fn=>setTimeout(fn,600),timeoutMs=30000,Zip,xml={},storage=browserStorage(),onTool}){
  const $=id=>doc.getElementById(id);
  let workbookReady=false;
  // The open tool (src/guide-tools.mjs), or null: with none open every part shows, tab strips and all, as on the first pane.
  // The page sets one at start (api.useTool), so in use there is always a tool open and the pane shows that tool alone.
  let tool=null;
  let queuedTool=null,busyIsSlow=false;
  let formSettings=readFormSettings(null),busy=false,plan=null,contents={present:false,stale:false},paras=[],at=-1,mode='start',selectedCue='say',connected=false,editing=null,pinned=false,contentsSig='',waiting=false,pendingDelete=null,lastOwn='',pictureReady=null;

  const status=text=>{$('guide-error').hidden=true;$('guide-error').textContent='';$('guide-status').textContent=text;};
  const problem=e=>{
    $('guide-status').textContent='';$('guide-error').textContent=(e&&e.message)||String(e);$('guide-error').hidden=false;
    // If the pane has never managed to read the document, there is nothing else on screen to press, so offer to try again.
    $('guide-retry-box').hidden=connected;
  };
  const focus=id=>$(id)?.focus?.();
  const tidy=()=>plan.edits.length||contents.stale?' Choose Update guide to bring the numbers, times and contents table up to date.':'';
  function syncButtons(){
    for(const id of BUSY_IDS)if($(id))$(id).disabled=busy;
    // Nothing is made for learners until the guide has something that is for them.
    $('guide-workbook-go').disabled=busy||!workbookReady;
    if(plan){
      // Update guide always works: it also refreshes the contents table's page numbers, which cannot be checked from here.
      const parts=[];
      if(plan.edits.length)parts.push(`${countText(plan.edits.length)} ${plan.edits.length===1?'is':'are'} out of date.`);
      if(contents.stale)parts.push(`The contents table is out of date.${contents.why?` ${contents.why}`:''}`);
      $('guide-update-note').textContent=parts.length?parts.join(' '):'Numbers and times are right. Update guide also refreshes the contents table and its page numbers.';
    }
  }
  // `slow` is for the one thing that waits for the author: a file window, which can stay open for minutes.
  async function act(fn,{slow=false}={}){
    if(busy)return;
    busy=true;busyIsSlow=slow;syncButtons();
    let timer;
    try{
      await Promise.race([fn(),new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('Word did not answer in time. Close the pane, open it again from the ribbon, then try again.')),slow?SLOW_TIMEOUT_MS:timeoutMs);})]);
    }catch(e){problem(e);}
    finally{clearTimeout(timer);busy=false;busyIsSlow=false;syncButtons();}
    // A tool asked for while this was still running (found in a real Word, 2026-11-07: a click on "More tools" a moment after Review
    // Guide opened did nothing, and said nothing) opens now. If several were asked for, the last one wins.
    if(queuedTool){const id=queuedTool;queuedTool=null;await act(()=>openToolScreen(id));}
  }
  function show(which,{moveFocus=false}={}){
    mode=which;
    // With a tool open, a screen the pane moves to by itself (Review from "Keep it tidy", the sample guide, the end of Start) is
    // that screen's own tool, so the strip and the screen never say different things.
    if(tool){
      const own={assess:'assessments',workbook:'participantWorkbook',review:'reviewGuide',help:'help'}[which];
      if(own&&tool!==own)setTool(own);
      else if(which==='build'&&toolFor(tool).kind!=='guide')setTool('openDesigner');
    }
    $('guide-start').hidden=which!=='start';
    $('guide-build').hidden=which!=='build';
    $('guide-assess').hidden=which!=='assess';
    $('guide-workbook').hidden=which!=='workbook';
    $('guide-review').hidden=which!=='review';
    $('guide-help').hidden=which!=='help';
    // The switch between the Guide, Assessments and Workbook screens shows on those three screens only, and not at all while a tool
    // is open: the ribbon (or "More tools") is how the author moves between them then.
    $('guide-switch').hidden=!!tool||(which!=='build'&&which!=='assess'&&which!=='workbook');
    $('guide-tab-guide').setAttribute('aria-pressed',String(which==='build'));
    $('guide-tab-assess').setAttribute('aria-pressed',String(which==='assess'));
    $('guide-tab-workbook').setAttribute('aria-pressed',String(which==='workbook'));
    paintTool();
    if(moveFocus)focus(headingFor(which));
  }
  // ---- the open tool: the Guide screen is one page of five parts and eleven pieces; a tool shows just its own
  // announce: the strip above the pane is told even when the pane thinks it is already on this tool. The older panels are opened by the
  // page, not by the pane, so the strip can have moved on to one of them while the pane still holds the tool before it (found in a real
  // Word, 2026-11-07: Review Guide, then Timing & Agenda, then Review Guide again left the strip saying Timing & Agenda).
  function setTool(id,{announce=false}={}){
    if(tool===id&&!announce)return;
    tool=id;
    onTool?.(id);
  }
  // The tab strips give way to the tool, and each piece of the Guide screen shows only for the tools that name it.
  function paintTool(){
    const t=tool?toolFor(tool):null,guide=!!t&&t.kind==='guide';
    $('guide-parts').hidden=!!t;
    for(const id of GUIDE_SECTIONS){const el=$(id);if(el)el.hidden=guide&&!t.sections.includes(id);}
    $('guide-sec-titlebar').hidden=guide&&!t.titlebar;
    $('guide-sec-selected').hidden=guide&&!t.selected;
    // A piece whose own heading says what the strip above already says (Add a module or lesson, E-Learning screens, Update from slides)
    // leaves the saying to the strip, so a screen reader does not read the same heading twice.
    for(const id of ['guide-module-heading','guide-el-heading','guide-upd-heading','guide-lib-heading','guide-agenda-heading','guide-align-heading','guide-ver-heading','guide-bk-heading','guide-rs-heading']){
      const heading=$(id);
      if(heading)heading.hidden=!!t&&clean(heading.textContent).toLowerCase()===t.title.toLowerCase();
    }
    if(guide)showPart(t.part);
  }
  // Where the focus goes when a screen opens: its own heading, or the tool's name when the tool has no heading of its own.
  function headingFor(which){
    if(tool&&which==='build'){const t=toolFor(tool);return t.titlebar?'guide-build-heading':'tool-title';}
    return {start:'guide-start-heading',build:'guide-build-heading',assess:'guide-assess-heading',workbook:'guide-workbook-heading',review:'guide-review-heading',help:'guide-help-heading'}[which];
  }
  // The block form serves two tools and the module form a third: whatever the pane is asked to focus, the tool that has it is open.
  function ensureTool(ids,fallback){
    if(!tool||ids.includes(tool))return;
    setTool(fallback);
    paintTool();
  }
  // The two writing tools open on the block form with different starting points: Add Activity on Activity written as parts,
  // Instructional Cues on an ordinary cue. A block being edited keeps what it has.
  function applyPreset(t){
    if(editing||!t.preset)return;
    if(t.preset.cue&&selectedCue!==t.preset.cue)chooseCue(t.preset.cue);
    if(t.preset.parts===true&&selectedCue==='activity'&&!$('guide-activity-parts').checked){$('guide-activity-parts').checked=true;syncParts();}
    if(t.preset.parts===false&&selectedCue==='activity')chooseCue('say');
  }
  // Opens one tool: reads the document, then shows the screen that tool lives on. Without a guide in the document, a tool that works
  // on the guide shows Start instead (Help, Review and the older panels do not need one).
  async function openToolScreen(id){
    const t=toolFor(id);
    if(!t)return;
    setTool(id,{announce:true});
    status('');
    if(t.kind==='screen'&&t.screen==='help'){show('help',{moveFocus:true});return;}
    await load();
    if(t.kind==='screen'&&t.screen==='review'){await review();return;}
    if(!plan.guide.isGuide){show('start',{moveFocus:true});return;}
    if(t.kind==='guide'){applyPreset(t);show('build',{moveFocus:true});return;}
    if(t.screen==='workbook')hideWorkbookMessages();
    show(t.screen,{moveFocus:true});
  }

  // ---- the Guide screen is five short parts, one at a time, so the pane is not one long page to scroll
  // One tab list: its tabs, their panels, the one shown, and the arrow keys. Returns show(which).
  function tabSet(tabPrefix,panelPrefix,ids){
    const show=(which,{moveFocus=false}={})=>{
      if(!ids.includes(which))return;
      for(const id of ids){
        $(`${panelPrefix}${id}`).hidden=id!==which;
        const tab=$(`${tabPrefix}${id}`);
        tab.setAttribute('aria-selected',String(id===which));
        tab.tabIndex=id===which?0:-1;
      }
      if(moveFocus)focus(`${tabPrefix}${which}`);
    };
    ids.forEach((id,i)=>{
      const tab=$(`${tabPrefix}${id}`);
      tab.onclick=()=>show(id);
      // Arrow keys, Home and End move along the tabs, as in any tab list.
      tab.onkeydown=e=>{
        const next={ArrowRight:i+1,ArrowLeft:i-1,Home:0,End:ids.length-1}[e?.key];
        if(next===undefined)return;
        e.preventDefault?.();
        show(ids[(next+ids.length)%ids.length],{moveFocus:true});
      };
    });
    show(ids[0]); // the page's own markup says the same; this makes it so wherever the page is read
    return show;
  }
  const PART_IDS=['add','slides','elearning','look','organize'];
  const showPart=tabSet('guide-part-tab-','guide-part-',PART_IDS);
  const showAssessPart=tabSet('guide-apart-tab-','guide-apart-',['items','paste','bank']);
  // The block form lives in the Add part, and the item form in the Items part, so anything that sends the author to one shows it first.
  const focusBlockText=()=>{ensureTool(['addActivity','instructionalCues'],'instructionalCues');showPart('add');focus('guide-block-text');};
  const focusItemStem=()=>{showAssessPart('items');focus('guide-item-stem');};

  // ---- reading the document
  async function load(){
    const list=await adapter.paragraphs();
    // Not being able to tell where the cursor is must never stop the pane showing the guide.
    let here=-1;
    try{here=await adapter.cursorIndex();}catch{/* no cursor: the "In your document" line stays hidden */}
    // The author's own cues live in the document. Reading them is never allowed to stop the pane: the built-in cues still work.
    let own=[];
    try{own=(await adapter.customCues?.())||[];}catch{/* unreadable: only the built-in cues are offered */}
    const ownKey=JSON.stringify(Array.from({length:CUSTOM_CUE_SLOTS},(_,i)=>String(own[i]??'')));
    if(ownKey!==lastOwn){lastOwn=ownKey;setCustomCues(own);buildCues();drawOwnCues();}
    await readStoredLook();
    // How forms are numbered. A document that cannot say (Word, through the add-in) has no forms and the automatic way.
    try{formSettings=readFormSettings(await adapter.formSettings?.());}catch{formSettings=readFormSettings(null);}
    paras=list;at=here;
    plan=updatePlan(list,{forms:formSettings});
    contents=contentsState(list);
    connected=true;
    $('guide-retry-box').hidden=true;
    // A guide carrying the reopen request an earlier build left behind has it removed, once, so Word stops complaining.
    if(plan.guide.isGuide)await keepOpen();
    if(plan.guide.isGuide){drawBuild();if(mode==='start')show('build');}    else if(mode!=='review'&&mode!=='help')show('start');
    syncButtons();
    await paintPlanning();
  }
  function drawBuild(){
    $('guide-build-heading').textContent=plan.guide.title||'Your guide';
    $('guide-total').textContent=`Total teaching time: ${plan.totalMinutes} min`;
    drawSelected();
    drawTidy();
    drawAssessSelected();
    drawStandard();
    drawPages();
    drawContents();
    drawWorkbook();
    drawElearning();
  }
  // Which standard pages are in the document, and which a button could still add.
  function drawStandard(){
    const s=standardPageState(paras),{pages,cannot}=missingStandardPages(paras);
    const parts=[];
    if(pages.length)parts.push(`Missing: ${[...pages].reverse().map(p=>p.name).join(' and ')}.`);
    if(cannot.length)parts.push(`${cannot.join(' and ')} cannot be placed, because this guide has none of the pages it would sit beside.`);
    if(!pages.length&&!cannot.length)parts.push('Acknowledgements and The Program in Perspective are in your guide.');
    parts.push(s.appendix>=0?'It has an Appendix.':'It has no Appendix yet. That is optional.');
    $('guide-standard-note').textContent=parts.join(' ');
  }
  // Move and delete work on the whole lesson the cursor is in, or the module when it is not in a lesson.
  function drawTidy(){
    const target=editing?null:tidyTarget(paras,at);
    if(pendingDelete&&(!target||target.span.start!==pendingDelete.start||target.text!==pendingDelete.text))cancelDelete();
    $('guide-tidy').hidden=!target;
    if(!target)return;
    $('guide-tidy-text').textContent=`Your cursor is in ${target.text}. These buttons work on the whole ${target.kind}, with everything in it.`;
    $('guide-move-up').textContent=`Move ${target.kind} up`;
    $('guide-move-down').textContent=`Move ${target.kind} down`;
    $('guide-delete').textContent=`Delete ${target.kind}`;
  }
  function cancelDelete(){
    pendingDelete=null;
    $('guide-delete-confirm').hidden=true;
    $('guide-delete-text').textContent='';
  }
  // Whether each generated Getting Started page is in the document and still says what the guide says now. Nothing rebuilds
  // them by itself, so this is the only sign that one has fallen out of date.
  function drawPages(){
    // A page that cannot be built yet says what it is waiting for, here beside its button, so nothing depends on a message elsewhere.
    const note=(name,s,waiting='')=>s.built?(s.stale?`${name} is out of date. Choose its button to build it again.`:`${name} matches your guide.`):`${name} has not been built yet.${waiting}`;
    const lessons=plan.guide.modules.reduce((n,m)=>n+m.lessons.length,0);
    $('guide-glance-note').textContent=note('Training at a Glance',glanceState(paras),lessons?'':' It needs a module and a lesson: add them first.');
    $('guide-cues-note').textContent=note('The Graphic Cues page',cuesState(paras),cuesUsed(paras).length?'':' It needs some blocks: add them first.');
    // The overview belongs to one module, the one the cursor is in; the count says how the rest stand.
    const mod=moduleAt(paras,at),all=overviewsSummary(paras);
    const here=mod?note(`The overview of ${mod.text||'this module'}`,overviewState(paras,mod)):'Put your cursor in a module to build its overview.';
    $('guide-overview-note').textContent=all.total?`${here} ${all.built} of ${plural(all.total,'module')} ${all.built===1?'has':'have'} an overview${all.stale?`, ${all.stale} out of date`:''}.`:here;
  }
  // The contents list is only rebuilt when it changed, so a quiet refresh never takes the focus off a button.
  function drawContents(){
    const rows=outline(plan),sig=JSON.stringify(rows),list=$('guide-contents');
    if(sig===contentsSig)return;
    contentsSig=sig;
    if(!rows.length){
      const li=doc.createElement('li');li.textContent='No modules or lessons yet. Add a module to begin.';
      list.replaceChildren(li);
    }else list.replaceChildren(...rows.map(r=>{
      const li=doc.createElement('li'),b=doc.createElement('button');
      li.className=`guide-${r.kind}`;b.type='button';b.className='guide-link';
      b.textContent=r.kind==='lesson'?`${r.text} (${r.minutes} min)`:r.text;
      b.onclick=()=>act(async()=>{await adapter.go(r.index,r.text);await load();status(`Moved to ${r.text}.`);});
      li.append(b);return li;
    }));
  }
  // "In your document": what the cursor is in, with a way to edit it.
  function drawSelected(){
    const found=editing?null:describeAt(paras,at),here=found&&found.type!=='item'?found:null; // items are edited on the Assessments screen
    $('guide-selected').hidden=!here;
    if(!here)return;
    const noun={block:'block',lesson:'lesson',module:'module'}[here.type];
    $('guide-selected-text').textContent=`Your cursor is in ${here.type==='block'?`the ${here.label}`:here.label}.`;
    $('guide-edit-selected').textContent=`Edit this ${noun}`;
    // In a block's shaded notes Enter only starts another line of notes, so a block has its own way out to plain text below it.
    $('guide-leave-block').hidden=here.type!=='block';
    $('guide-selected-hint').hidden=here.type!=='block';
    $('guide-selected-hint').textContent=here.type==='block'?'In the notes, Enter starts another line of notes. To write below this block instead, choose Write below this block.':'';
  }
  // Puts a plain line under the block the cursor is in and moves the cursor into it, so the author can write below the block (and below
  // slides, whose notes Enter would only extend). Text typed there is ordinary text: it belongs to no block.
  async function writeBelow(){
    await load();
    const here=describeAt(paras,at);
    if(!here||here.type!=='block')throw new Error('Click inside a block first (a block is a cue and its notes, such as a Say or a Slide). Then choose Write below this block.');
    const {created}=await adapter.plainLineAfter(here.block.span.end);
    await load();
    status(`${created?'A plain line was added below':'Your cursor is on the plain line below'} the ${here.label}${created?', and your cursor is in it':''}. What you type here is ordinary text and belongs to no block. To add another block here, choose Add a block. To go back to the notes, click in them.${created?` ${UNDO}`:''}`);
  }

  // ---- the cue picker: a plain radio group, so arrow keys and a screen reader work without any widget code
  function chooseCue(id,{audience=true}={}){
    selectedCue=id;
    const c=cueById(id);
    $('guide-cue-help').textContent=c.help;
    $('guide-cue-current').textContent=c.label;$('guide-cue-current-icon').src=iconUrl(c.icon);
    for(const label of $('guide-cues').children)label.children[0].checked=label.children[0].id===`guide-cue-${id}`;
    if(audience)for(const a of AUDIENCES)$(`guide-audience-${a.id}`).checked=a.id===c.audience;
    // Only an Activity can be written as labelled parts; choosing another cue puts the form back to plain text.
    $('guide-activity-box').hidden=id!=='activity';
    if(id!=='activity')$('guide-activity-parts').checked=false;
    // Any block can carry a picture; a Picture block must.
    $('guide-pic-box').open=id==='picture'||!!pictureReady;
    $('guide-pic-summary').textContent=id==='picture'?'Picture (needed for this block)':'Add a picture to this block (optional)';
    syncParts();
  }
  // ---- a picture in a block: choose the file, describe it, and the block carries it
  // The file is chosen in a window of its own (a file chooser cannot be opened from inside the pane); the picture comes back already read and, if large, shrunk.
  async function choosePicture(){
    const res=await adapter.pickFiles({kind:'picture'});
    if(!res||res.cancelled||!res.files?.length){status('No picture was chosen.');return;}
    const p=res.files[0];
    pictureReady=p;
    $('guide-pic-chosen').textContent=`${p.name}: ${p.width} by ${p.height} pixels${p.shrunk?', made smaller to keep the document small':''}. Now describe it below.`;
    $('guide-pic-box').open=true;
    focus('guide-pic-alt');
  }
  function clearPicture(){
    pictureReady=null;
    $('guide-pic-alt').value='';$('guide-pic-decorative').checked=false;
    $('guide-pic-chosen').textContent='';
    $('guide-pic-width').value=PICTURE_WIDTHS[0].id;
    $('guide-pic-box').open=selectedCue==='picture';
  }
  function readPictureForm(){
    if(!pictureReady){
      if(selectedCue==='picture')throw new Error(`Choose a picture file first. ${PICTURE_HELP}`);
      return undefined;
    }
    const decorative=!!$('guide-pic-decorative').checked,alt=$('guide-pic-alt').value;
    if(!decorative&&!alt.trim())throw new Error('Describe the picture in a few words for people who cannot see it, or tick "This picture is only decoration".');
    const p=pictureReady;
    return {base64:toBase64(p.bytes),type:p.type,mime:p.mime,size:pictureSize(p,$('guide-pic-width').value),alt,decorative,name:p.name};
  }
  // The labelled-parts fields show only while an Activity is chosen and the box is ticked; the Text box is then the title.
  const partsMode=()=>selectedCue==='activity'&&!!$('guide-activity-parts').checked;
  function syncParts(){
    const on=partsMode();
    $('guide-activity-fields').hidden=!on;
    $('guide-block-text-label').textContent=on?'Title of the activity':'Text';
    $('guide-block-text').rows=on?2:4;
  }
  $('guide-activity-parts').onchange=syncParts;
  function buildGroups(){
    $('guide-cue-group').replaceChildren(...[{id:'',label:'All cues'},...CUE_GROUPS].map(g=>{const o=doc.createElement('option');o.value=g.id;o.textContent=g.label;return o;}));
    $('guide-cue-group').value='';
  }
  // The list shows the cues that match what is typed in Find and the kind chosen beside it.
  function filterCues(){
    const q=$('guide-cue-find').value.trim().toLowerCase(),g=$('guide-cue-group').value;
    let shown=0;
    for(const label of $('guide-cues').children){const hide=(!!q&&!label.dataset.label.includes(q))||(!!g&&label.dataset.group!==g);label.hidden=hide;if(!hide)shown++;}
    $('guide-cue-count').textContent=q||g?`${plural(shown,'cue')} match.`:'';
  }
  function buildCues(){
    if(!cueById(selectedCue))selectedCue='say'; // an own cue that has gone is not left chosen
    $('guide-cues').replaceChildren(...allCues().map(c=>{
      const label=doc.createElement('label'),input=doc.createElement('input'),img=doc.createElement('img'),name=doc.createElement('span');
      label.className='guide-cue';label.dataset.label=c.label.toLowerCase();label.dataset.group=c.group||'';
      input.type='radio';input.name='guide-cue';input.id=`guide-cue-${c.id}`;input.value=c.id;input.checked=c.id===selectedCue;
      input.onchange=()=>chooseCue(c.id);
      // Choosing with the mouse or Enter puts the list away and moves on to the text; arrow keys through the radios leave it open.
      input.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault?.();setCuePanel(false);focus('guide-block-text');}};
      label.onclick=e=>{if(e&&e.detail>0)globalThis.setTimeout?.(()=>{setCuePanel(false);focus('guide-block-text');},0);};
      img.src=iconUrl(c.icon);img.alt='';img.width=28;img.height=28;
      name.textContent=c.label;
      label.append(input,img,name);
      return label;
    }));
    chooseCue(selectedCue);
    filterCues();
  }
  $('guide-cue-find').oninput=filterCues;
  $('guide-cue-group').onchange=$('guide-cue-group').oninput=filterCues;
  // The long list of cues stays put away until the author asks for it: the chosen cue and a Change button are all that show.
  function setCuePanel(open){
    // Opened with nothing typed or chosen, the list starts on the chosen cue's own kind (a dozen or so cues, not all sixty);
    // "All cues" is still the first choice in the kind box.
    if(open&&!$('guide-cue-find').value&&!$('guide-cue-group').value){$('guide-cue-group').value=cueById(selectedCue)?.group||'';filterCues();}
    $('guide-cue-panel').hidden=!open;
    $('guide-cue-toggle').setAttribute('aria-expanded',String(open));
    $('guide-cue-toggle-action').textContent=open?'Close':'Change';
  }
  $('guide-cue-toggle').onclick=()=>{const open=$('guide-cue-panel').hidden;setCuePanel(open);if(open)focus('guide-cue-group');};
  setCuePanel(false);

  // ---- the author's own cues: up to three names, kept in the document so they travel with it
  function drawOwnCues(){
    const names=customCueNames();
    for(let i=1;i<=CUSTOM_CUE_SLOTS;i++)$(`guide-own-${i}`).value=names[i-1];
  }
  async function saveOwnCues(){
    if(!adapter.setCustomCues)throw new Error('This document cannot keep your own cues.');
    $('guide-own-note').textContent=''; // an old "Saved." must not stay beside a refusal (seen in real Word, 2026-11-06)
    const names=Array.from({length:CUSTOM_CUE_SLOTS},(_,i)=>$(`guide-own-${i+1}`).value);
    const check=checkCustomCueNames(names);
    if(!check.ok)throw new Error(check.problems.join(' '));
    // A name that blocks already use cannot change or go: those blocks would stop being recognised as that cue.
    const before=customCueNames();
    for(let i=0;i<CUSTOM_CUE_SLOTS;i++){
      if(before[i]&&before[i]!==check.names[i]&&paras.some(p=>kindOf(p)==='cue'&&cueFromText(p.text)?.id===`custom-${i+1}`)){
        throw new Error(`Blocks in your guide use the cue "${before[i]}". Change or delete those blocks first, then change the name.`);
      }
    }
    await adapter.setCustomCues(check.names);
    lastOwn=JSON.stringify(check.names);setCustomCues(check.names);buildCues();drawOwnCues();
    const n=check.names.filter(Boolean).length;
    $('guide-own-note').textContent=n?`Saved. Your cue${n===1?'':'s'} ${n===1?'is':'are'} in the list under "For you".`:'Saved. You have no cues of your own.';
    status('Your own cues are saved in this document.');
  }
  $('guide-own-save').onclick=()=>act(saveOwnCues);

  // After the first change the pane asks Word to open it again with this document.
  const keepOpen=async()=>{if(!pinned){pinned=true;await adapter.keepPaneOpen?.();}};

  // ---- actions
  async function startGuide(){
    const title=$('guide-title').value.trim();
    const hadText=(await adapter.paragraphs()).some(p=>p.text.trim());
    // A look chosen here is checked first, and styles the pages bring with them are made from it.
    const wanted=$('guide-start-look').value?cleanLook({id:$('guide-start-look').value}):null;
    if(wanted){const r=resolveGuideLook(wanted);if(!r.ok)throw new Error(r.issues.join(' '));adapter.setTheme?.(r.theme);}
    await adapter.insert(ctx=>startGuideBody({title,standard:$('guide-standard').checked},ctx),{where:'start',select:'last'});
    await adapter.setTitle(title);
    // The guide exists from here on. If the look cannot be applied, the guide is still there and the author is told.
    let lookNote='';
    if(wanted){
      try{lookNote=` ${lookReport(await applyLookTo(wanted,title))}`;}
      catch(e){lookNote=` The look could not be applied (${(e&&e.message)||e}). Choose Change the look to try again.`;}
    }
    await adapter.refreshFields();
    await keepOpen();
    await load();
    show('build',{moveFocus:true});
    status(`Guide started.${hadText?' Your existing text is below the new pages.':''}${lookNote} Add a module to begin. ${UNDO}`);
  }
  async function trySample(){
    await adapter.insert(sampleGuide,{where:'start',select:'last'});
    // A document title is part of what makes a file accessible, so the sample has one unless the document already does.
    if(!String(await adapter.title()).trim())await adapter.setTitle(SAMPLE_TITLE);
    await adapter.refreshFields();
    await keepOpen();
    await load();
    show('build',{moveFocus:true});
    status(`The sample guide is at the top of this document. ${UNDO}`);
  }
  async function addModule(){
    const title=$('guide-new-title').value.trim();
    if(!title)throw new Error('Type a title for the module first.');
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex();
    const {anchor,n}=modulePlacement(list,here);
    await adapter.insert(ctx=>moduleHeading(n,title)+cueBlock({cue:'transition',text:'Say how this module leads into the next one.'},ctx),{where:'after',after:anchor,select:'first'});
    $('guide-new-title').value='';
    await keepOpen();
    await load();
    status(`Module ${n} added, with a Transition block at its end. Choose Add lesson for its first lesson.${tidy()} ${UNDO}`);
  }
  async function addLesson(){
    const title=$('guide-new-title').value.trim();
    if(!title)throw new Error('Type a title for the lesson first.');
    const minutes=parseMinutes($('guide-new-minutes').value,'Lesson minutes');
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex();
    const {anchor,m,l,before,after}=lessonPlacement(list,here);
    await adapter.insert(ctx=>lessonHeading(m,l,title)+timeBlock({minutes,cumulative:before},ctx),{where:'after',after:anchor,select:'last'});
    $('guide-new-title').value='';
    await keepOpen();
    await load();
    status(`Lesson ${m}.${l} added after ${after}. Add its blocks below.${tidy()} ${UNDO}`);
    focusBlockText();
  }
  function readBlockForm(){
    const cue=cueById(selectedCue),text=$('guide-block-text').value;
    const picture=editing?undefined:readPictureForm();
    // A block that carries a picture may have no words of its own: the picture is its content.
    if(!text.trim()&&!picture)throw new Error(partsMode()?'Type a title for the activity first.':'Type the text for this block first.');
    const minutes=parseMinutes($('guide-block-minutes').value),notes=$('guide-block-notes').value;
    const audience=(AUDIENCES.find(a=>$(`guide-audience-${a.id}`).checked)||{}).id||cue.audience;
    const parts=partsMode()?Object.fromEntries(ACTIVITY_PARTS.map(p=>[p.id,$(`guide-act-${p.id}`).value])):undefined;
    return {cue,text,notes,minutes,audience,parts,picture};
  }
  function clearBlockForm(){
    clearPicture();
    for(const id of ['guide-block-text','guide-block-notes','guide-block-minutes',...ACTIVITY_PARTS.map(p=>`guide-act-${p.id}`)])$(id).value='';
    editing=null;
    $('guide-block-heading').textContent='Add a block';
    $('guide-block-go').textContent='Insert block';
    $('guide-block-cancel').hidden=true;
    $('guide-block-note').textContent='';
  }
  async function addBlock(){
    if(editing?.type==='block')return updateBlock();
    const block=readBlockForm();
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex();
    // A Transition leads out of its module, so it always goes at the module's end; any other block goes after the block the cursor is in.
    const anchor=block.cue.id==='transition'?transitionAnchor(list,here).anchor:blockAnchor(list,here);
    await adapter.insert(ctx=>cueBlock(block,ctx),{where:'after',after:anchor,select:'last'});
    clearBlockForm();
    await keepOpen();
    await load();
    status(`${block.cue.label} block added.${tidy()} ${UNDO}`);
    focusBlockText();
  }

  // ---- the activity library: ready-made activities and the author's own, each put into the guide as an Activity written as labelled parts
  let mine=[];
  function loadMine(){try{mine=cleanMine(JSON.parse(storage?.getItem(LIBRARY_KEY)||'null'));}catch{mine=[];}}
  // Not being able to keep them (storage missing or refused) is said, not hidden: the activity is still in the guide.
  function keepMine(){try{if(!storage)return false;storage.setItem(LIBRARY_KEY,JSON.stringify(mine));return true;}catch{return false;}}
  function libraryItem(e,own){
    const li=doc.createElement('li'),title=doc.createElement('h3'),kind=doc.createElement('span'),summary=doc.createElement('p');
    li.className='guide-lib-item';
    title.id=`guide-lib-${e.id}-title`;
    title.textContent=`${e.title}${e.minutes?` (${e.minutes} min)`:''}`;
    kind.className='guide-lib-kind';kind.textContent=e.kind;title.append(kind);
    summary.className='hint';summary.textContent=e.summary;
    const details=doc.createElement('details'),sum=doc.createElement('summary'),inside=doc.createElement('div');
    sum.textContent='What is in it';
    const part=(label,text,{list=false,facilitator=false}={})=>{
      if(!String(text||'').trim())return;
      const p=doc.createElement('p');p.className='guide-lib-part';
      const b=doc.createElement('strong');b.textContent=`${label}${facilitator?' (facilitator only)':''}: `;
      p.append(b);
      const rows=String(text).split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
      if(!list){const t=doc.createElement('span');t.textContent=rows.join(' ');p.append(t);inside.append(p);return;}
      inside.append(p);
      const bullets=doc.createElement('ul');
      bullets.replaceChildren(...rows.map(x=>{const step=doc.createElement('li');step.textContent=x;return step;}));
      inside.append(bullets);
    };
    part('Objective',e.objective);part('Materials',e.materials,{facilitator:true});part('Preparation',e.preparation,{facilitator:true});
    part('Instructions',e.instructions,{list:true});part('Case study',e.caseStudy);part('Facilitator notes',e.notes,{facilitator:false});
    details.append(sum,inside);
    const buttons=doc.createElement('div'),insert=doc.createElement('button');
    buttons.className='guide-actions';
    insert.type='button';insert.className='primary';insert.textContent='Insert';insert.setAttribute('aria-label',`Insert ${e.title} into the guide`);
    insert.setAttribute('aria-describedby',title.id);
    insert.onclick=()=>act(()=>insertActivity(e));
    buttons.append(insert);
    if(own){
      const remove=doc.createElement('button');
      remove.type='button';remove.textContent='Remove';remove.setAttribute('aria-label',`Remove ${e.title} from your activities`);
      remove.onclick=()=>act(()=>removeSaved(e));
      buttons.append(remove);
    }
    li.append(title,summary,details,buttons);
    return li;
  }
  function drawLibrary(){
    const q=$('guide-lib-find').value,ready=STARTER_ACTIVITIES.filter(e=>libraryMatches(e,q)),own=mine.filter(e=>libraryMatches(e,q));
    $('guide-lib-list').replaceChildren(...ready.map(e=>libraryItem(e,false)));
    $('guide-lib-mine').replaceChildren(...own.map(e=>libraryItem(e,true)));
    const activities=n=>`${n} ${n===1?'activity':'activities'}`;
    $('guide-lib-count').textContent=q.trim()?`${activities(ready.length+own.length)} found.`:`${activities(STARTER_ACTIVITIES.length)} ready to use${mine.length?`, and ${mine.length} of your own`:''}.`;
    $('guide-lib-mine-heading').hidden=false;
    $('guide-lib-mine-note').textContent=mine.length?'Your own activities are kept on this computer and are here in every guide you open. To change one, insert it, edit it in the guide, and save it again with the same title.'
      :'You have not saved any yet. Put your cursor in an Activity you wrote, then choose Save. It is kept on this computer and is here in every guide you open.';
  }
  async function insertActivity(e){
    const block=blockFromEntry(e);
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),anchor=blockAnchor(list,here);
    await adapter.insert(ctx=>cueBlock(block,ctx),{where:'after',after:anchor,select:'last'});
    await keepOpen();
    await load();
    status(`${e.title} was added after the block your cursor was in. Change its words in Word, or put your cursor in it and choose Edit.${tidy()} ${UNDO}`);
  }
  async function saveToLibrary(){
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),d=describeAt(list,here);
    const saved=entryFromBlock(d&&d.type==='block'?d.block:null);
    const {list:next,replaced}=addMine(mine,saved);
    const before=mine;mine=next;
    if(!keepMine()){mine=before;throw new Error('This computer would not keep your activity (storage for the add-in is off or full). Nothing was saved.');}
    drawLibrary();
    status(`${saved.title} was ${replaced?'updated in':'saved to'} your activities.`);
  }
  async function removeSaved(e){
    const before=mine;mine=removeMine(mine,e.id);
    if(!keepMine()){mine=before;throw new Error('This computer would not change your saved activities. Nothing was removed.');}
    drawLibrary();
    status(`${e.title} was removed from your activities. The copy already in a guide is not touched.`);
  }

  // ---- the planning tools, each read from the guide as it is (nothing here needs the older course record): timing and agenda,
  // learning alignment and delivery versions
  const AGENDA_START_KEY='osg.agendaStart.v1';
  let agendaShown='',versions=[],chosenVersion='',confirmingDelete=false;
  const cellOf=(tag,text,scope)=>{const c=doc.createElement(tag);c.textContent=text;if(scope)c.setAttribute('scope',scope);return c;};
  const listItem=text=>{const li=doc.createElement('li');li.textContent=text;return li;};
  function agendaRowsInto(body,rows){
    body.replaceChildren(...rows.map(r=>{
      const tr=doc.createElement('tr');
      if(r.kind==='module'){tr.className='guide-agenda-module';const th=cellOf('th',r.text,'rowgroup');th.setAttribute('colspan','3');tr.append(th);return tr;}
      tr.className=r.kind==='lesson'?'guide-agenda-lesson':'guide-agenda-block';
      tr.append(cellOf('td',`${r.start} to ${r.end}`),cellOf('td',r.text),cellOf('td',String(r.minutes)));
      return tr;
    }));
  }
  // The Word file of an agenda opens in a window of its own (or is downloaded when Word cannot), as the narration script does.
  async function openAgendaFile(a){
    const ZipClass=Zip||globalThis.JSZip;
    if(!ZipClass)throw new Error('The pane could not load what it needs to make the agenda. Close it, open it again from the ribbon, then try again.');
    const base64=await zipWorkbook(agendaFiles(a),ZipClass),name=agendaFileName(a.title);
    try{await adapter.openDocument(base64,name);status(`The agenda is open in a new Word window. Save it as "${name}", or any name you like.`);}
    catch{adapter.offerDownload(base64,name);status(`Word could not open it for you, so it was downloaded as "${name}". Open it from your Downloads folder.`);}
  }
  // -- timing and agenda
  function drawAgenda(){
    const start=$('guide-agenda-start').value||'09:00';
    let a;
    try{a=guideAgenda(paras,start);}
    catch(e){agendaShown='';$('guide-agenda-table').hidden=true;$('guide-agenda-body').replaceChildren();$('guide-agenda-notes').replaceChildren();$('guide-agenda-summary').textContent=e.message;return;}
    agendaShown=JSON.stringify([start,a.rows,a.notes]);
    $('guide-agenda-summary').textContent=agendaSummary(a);
    $('guide-agenda-table').hidden=!a.lessons;
    agendaRowsInto($('guide-agenda-body'),a.rows);
    $('guide-agenda-notes').replaceChildren(...a.notes.map(n=>listItem(n.text)));
    try{storage?.setItem(AGENDA_START_KEY,start);}catch{/* the start time is just not remembered */}
  }
  async function makeAgenda(){
    const list=await adapter.paragraphs(),start=$('guide-agenda-start').value||'09:00',a=guideAgenda(list,start);
    if(!a.lessons)throw new Error('There are no lessons to put in an agenda yet. Add a module and a lesson, then give them minutes.');
    if(JSON.stringify([start,a.rows,a.notes])!==agendaShown){paras=list;drawAgenda();throw new Error('Your guide or the start time changed since the agenda was shown. It is shown again above. Check it, then choose Make the agenda in Word.');}
    await openAgendaFile({...a,title:String(await adapter.title()).trim()||a.title});
  }
  // -- learning alignment
  function drawAlignment(){
    const a=alignment(paras);
    $('guide-align-summary').textContent=alignmentSummary(a);
    const names=list=>list.map(x=>`${x.title||x.label}${x.similar?' (similar words)':''}`).join('; ');
    const goTo=(index,label)=>{
      const b=doc.createElement('button');
      b.type='button';b.textContent='Go there';b.setAttribute('aria-label',`Go to ${label}`);
      b.onclick=()=>act(async()=>{await adapter.go(index,paras[index]?paras[index].text:'');await load();status(`Moved to ${label}.`);});
      return b;
    };
    $('guide-align-list').replaceChildren(...a.objectives.map(o=>{
      const li=doc.createElement('li'),h=doc.createElement('h3'),p1=doc.createElement('p'),p2=doc.createElement('p');
      li.className='guide-align-item';
      h.textContent=`${o.lesson}: ${o.text}`;
      p1.textContent=o.activities.length?`Activity: ${names(o.activities)}`:'No activity yet.';
      p2.textContent=o.items.length?`Assessment: ${names(o.items)}`:'No assessment item yet.';
      if(!o.activities.length)p1.className='guide-align-gap';
      if(!o.items.length)p2.className='guide-align-gap';
      li.append(h,p1,p2,goTo(o.index,`the objective in ${o.lesson}`));
      return li;
    }));
    $('guide-align-loose-heading').hidden=!a.loose.length;
    $('guide-align-loose').replaceChildren(...a.loose.map(l=>{const li=listItem(`${l.text}: ${l.why} `);li.append(goTo(l.index,l.text));return li;}));
  }
  // -- delivery versions
  const versionNow=()=>versions.find(v=>v.id===chosenVersion)||null;
  async function readVersionsNow(){
    let text=null;
    try{text=await adapter.versionMemory?.();}catch{text=null;}
    versions=readVersions(text);
    if(!versions.some(v=>v.id===chosenVersion))chosenVersion=versions[0]?versions[0].id:'';
  }
  // A change is kept in the document; if Word will not keep it, the pane goes back to what it had.
  async function changeVersions(next,pick){
    confirmingDelete=false;
    const before=versions,was=chosenVersion;
    versions=next;if(pick!==undefined)chosenVersion=pick;
    if(!versions.some(v=>v.id===chosenVersion))chosenVersion=versions[0]?versions[0].id:'';
    try{await adapter.setVersionMemory(writeVersions(versions));}
    catch(e){versions=before;chosenVersion=was;drawVersions();throw e;}
    status(''); // a problem shown for an earlier try goes away once this one is kept
    drawVersions();
  }
  function drawVersions(){
    const v=versionNow(),all=lessonList(paras);
    $('guide-ver-select').replaceChildren(...versions.map(x=>{const o=doc.createElement('option');o.value=x.id;o.textContent=x.name;return o;}));
    $('guide-ver-select').value=chosenVersion;
    $('guide-ver-delete').disabled=!v;
    if(!confirmingDelete)$('guide-ver-delete').textContent='Delete this version';
    $('guide-ver-edit').hidden=!v;
    if(!v)return;
    $('guide-ver-name').value=v.name;$('guide-ver-target').value=v.target?String(v.target):'';$('guide-ver-start').value=v.start;
    const inVersion=v.lessons.map(t=>({title:t,lesson:all.find(l=>l.key===lessonKey(t))||null}));
    const out=all.filter(l=>!v.lessons.some(t=>lessonKey(t)===l.key));
    const row=(title,lesson,included,position)=>{
      const li=doc.createElement('li'),label=doc.createElement('label'),box=doc.createElement('input'),span=doc.createElement('span');
      li.className='guide-lib-item';
      label.className='guide-check';box.type='checkbox';box.checked=included;
      box.onchange=()=>act(()=>changeVersions(versions.map(x=>x.id===v.id?toggleLesson(v,all,title):x)));
      span.textContent=lesson?` ${lesson.text} (${lesson.minutes} min)`:` ${title} (no longer in the guide)`;
      label.append(box,span);li.append(label);
      if(included){
        const bar=doc.createElement('div');bar.className='guide-actions';
        const move=(text,delta,disabled)=>{const b=doc.createElement('button');b.type='button';b.textContent=text;b.disabled=disabled;b.setAttribute('aria-label',`${text}: ${title}`);b.onclick=()=>act(()=>changeVersions(versions.map(x=>x.id===v.id?moveLesson(v,title,delta):x)));return b;};
        bar.append(move('Earlier',-1,position===0),move('Later',1,position===v.lessons.length-1));
        li.append(bar);
      }
      return li;
    };
    $('guide-ver-lessons').replaceChildren(...inVersion.map((x,i)=>row(x.title,x.lesson,true,i)),...out.map(l=>row(l.title,l,false,-1)));
    let a;
    try{a=versionAgenda(paras,v,{title:String(plan?.guide?.title||'Guide')});}
    catch(e){$('guide-ver-summary').textContent=e.message;$('guide-ver-table').hidden=true;return;}
    $('guide-ver-summary').textContent=versionSummary(a)+(a.missing.length?` ${a.missing.length===1?'One lesson is':`${a.missing.length} lessons are`} no longer in the guide.`:'');
    $('guide-ver-table').hidden=!a.lessons;
    agendaRowsInto($('guide-ver-body'),a.rows);
  }
  async function newVersionNow(){
    await readVersionsNow();
    const v=newVersion(lessonList(await adapter.paragraphs()),versions);
    await changeVersions([...versions,v],v.id);
    status(`${v.name} was made with every lesson in it. Take out the lessons you will not teach, and set a target.`);
  }
  async function deleteVersionNow(){
    const v=versionNow();
    if(!v)return;
    if(!confirmingDelete){confirmingDelete=true;$('guide-ver-delete').textContent='Yes, delete it';status(`Choose "Yes, delete it" to delete ${v.name}. Anything else keeps it.`);return;}
    confirmingDelete=false;
    await changeVersions(versions.filter(x=>x.id!==v.id),'');
    status(`${v.name} was deleted.`);
  }
  async function renameVersion(){
    const v=versionNow(),name=oneLine($('guide-ver-name').value).slice(0,80);
    if(!v)return;
    if(!name){$('guide-ver-name').value=v.name;throw new Error('A version needs a name.');}
    if(versions.some(x=>x.id!==v.id&&lessonKey(x.name)===lessonKey(name))){$('guide-ver-name').value=v.name;throw new Error('Another version already has that name.');}
    await changeVersions(versions.map(x=>x.id===v.id?{...x,name}:x));
  }
  async function retargetVersion(){
    const v=versionNow();
    if(!v)return;
    let target;
    try{target=parseTarget($('guide-ver-target').value);}catch(e){$('guide-ver-target').value=v.target?String(v.target):'';throw e;}
    await changeVersions(versions.map(x=>x.id===v.id?{...x,target}:x));
  }
  async function restartVersion(){
    const v=versionNow(),start=$('guide-ver-start').value||'09:00';
    if(!v)return;
    if(!START_TIME.test(start)){$('guide-ver-start').value=v.start;throw new Error('Choose a valid start time.');}
    await changeVersions(versions.map(x=>x.id===v.id?{...x,start}:x));
  }
  async function makeVersionAgenda(){
    const v=versionNow();
    if(!v)throw new Error('Choose or make a version first.');
    const list=await adapter.paragraphs(),a=versionAgenda(list,v,{title:String(await adapter.title()).trim()||plan?.guide?.title||'Guide'});
    if(!a.lessons)throw new Error('No lessons are in this version yet. Tick the lessons to teach.');
    await openAgendaFile(a);
  }
  // Whichever of the three is the open tool is drawn from the document each time it is read. A problem here never stops the pane.
  async function paintPlanning(){
    try{
      if(tool==='timingAndAgenda')drawAgenda();
      else if(tool==='learningAlignment')drawAlignment();
      else if(tool==='deliveryVersions'){await readVersionsNow();drawVersions();}
    }catch{/* the panel stays as it was */}
  }

  // ---- slides from a presentation: a deck read for its titles, notes and clicks, and the author's own slide pictures
  let deck=null,slidePics=[];
  const oneLine=s=>String(s??'').replace(/\s+/g,' ').trim();
  async function chooseSlides(){
    const res=await adapter.pickFiles({kind:'slides'});
    if(!res||res.cancelled||!res.deck){status('No presentation was chosen.');return;}
    deck={name:res.deck.name,slides:res.deck.slides.map(s=>({...s})),chosen:new Set(res.deck.slides.filter(s=>!s.hidden).map(s=>s.order))};
    slidePics=res.files||[];
    drawDeck();
    status(`Read ${plural(deck.slides.length,'slide')} from ${deck.name}${slidePics.length?` and ${plural(slidePics.length,'picture')}`:''}. Choose the slides to insert below.`);
  }
  function drawDeck(){
    $('guide-slides-fieldset').hidden=!deck;
    $('guide-slides-list').replaceChildren(...(deck?deck.slides.map(s=>{
      const label=doc.createElement('label'),input=doc.createElement('input'),span=doc.createElement('span');
      label.className='guide-check';input.type='checkbox';input.id=`guide-slide-${s.order}`;input.checked=deck.chosen.has(s.order);
      input.onchange=()=>{if(input.checked)deck.chosen.add(s.order);else deck.chosen.delete(s.order);drawSlidesNote();};
      span.textContent=` Slide ${s.order}${oneLine(s.title)?`: ${oneLine(s.title)}`:''}${s.hidden?' (hidden)':''}${s.clicks?`, ${plural(s.clicks,'click')}`:''}`;
      label.append(input,span);return label;
    }):[]));
    drawSlidesNote();
  }
  function drawSlidesNote(){
    $('guide-slides-summary').textContent=deck?`${deck.name}: ${plural(deck.slides.length,'slide')}, ${deck.chosen.size} chosen.${deck.slides.some(s=>s.titleGuessed)?' This presentation has no slide titles, so each slide is named by the largest text on it.':''}`:'';
    if(!slidePics.length){
      // Said plainly, so a slide that arrives without its picture is never a surprise (found in a recording, 2026-11-06).
      $('guide-slides-pics-note').textContent=deck?'No pictures of the slides were chosen, so each slide goes in as text: its number, title, clicks and notes. To add the pictures, choose the presentation again and also give a PDF of it.':'';
      return;
    }
    const {byOrder,problems}=deck?matchPictures(deck.slides,slidePics):{byOrder:new Map(),problems:[]};
    const lack=deck?[...deck.chosen].filter(o=>!byOrder.has(o)).length:0;
    $('guide-slides-pics-note').textContent=problems.length?problems.join(' '):`${plural(slidePics.length,'picture')} chosen${deck?`; ${lack?`${plural(lack,'chosen slide')} will have no picture`:'every chosen slide has one'}`:''}.`;
  }
  async function insertSlides(){
    if(!deck)throw new Error('Choose a presentation first.');
    const slides=deck.slides.filter(s=>deck.chosen.has(s.order));
    if(!slides.length)throw new Error('Choose at least one slide to insert.');
    const {byOrder,problems}=matchPictures(deck.slides,slidePics);
    if(problems.length)throw new Error(problems.join(' '));
    const share=!!$('guide-slides-share').checked;
    // Any number of slides can be chosen. They go into Word in runs (at most MAX_SLIDES_AT_ONCE slides and MAX_TOTAL_PICTURE_BYTES of
    // pictures to a run), one run after another and in order, so a long presentation is still one press of the button.
    const runs=[];let run=[],runBytes=0;
    for(const s of slides){
      const file=byOrder.get(s.order),size=file?file.bytes.length:0; // already read and shrunk by the file window
      if(size>MAX_TOTAL_PICTURE_BYTES)throw new Error(`The picture for slide ${s.order} is larger than ${MAX_TOTAL_PICTURE_BYTES/1048576} MB. Choose a smaller picture or a PDF.`);
      if(run.length&&(run.length>=MAX_SLIDES_AT_ONCE||runBytes+size>MAX_TOTAL_PICTURE_BYTES)){runs.push(run);run=[];runBytes=0;}
      run.push({s,file});runBytes+=size;
    }
    if(run.length)runs.push(run);
    const inserted=[];let failure=null;
    try{
      for(const part of runs){
        if(runs.length>1)status(`Inserting slides ${inserted.length+1} to ${inserted.length+part.length} of ${slides.length}...`);
        const blocks=part.map(({s,file})=>slideBlock(s,{share,picture:file?{base64:toBase64(file.bytes),type:file.type,mime:file.mime,size:pictureSize(file,'large'),name:file.name}:null}));
        // The cursor is left at the end of what was just inserted, so each run follows the one before it.
        const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),anchor=blockAnchor(list,here);
        await adapter.insert(ctx=>blocks.map(b=>cueBlock(b,ctx)).join(''),{where:'after',after:anchor,select:'last'});
        inserted.push(...part);
      }
    }catch(e){failure=e;}
    const slidesIn=inserted.map(x=>x.s),withPicture=inserted.filter(x=>x.file).length;
    if(slidesIn.length){
      // What the guide now remembers of these slides is what Update from slides compares with later. Not being able to keep it is no reason to stop.
      try{
        const entries=slidesIn.filter(s=>Number.isInteger(s.sourceId)).map(s=>({id:s.sourceId,label:labelOf(s),ns:notesPrint(s.notes),tx:notesPrint(slideTextLines(s)),pic:byOrder.get(s.order)?fingerprint(byOrder.get(s.order).bytes):''}));
        const memory=readMemory(await adapter.slideMemory?.());
        await adapter.setSlideMemory?.(writeMemory(remember(memory,{name:deck.name,seen:deck.slides.map(s=>s.sourceId).filter(Number.isInteger),entries})));
      }catch{/* the next update then matches these slides by their titles */}
      for(const s of slidesIn)deck.chosen.delete(s.order);
      drawDeck();
      await keepOpen();
      await load();
    }
    if(failure){
      // Said plainly: what went in stays in, and what did not is still ticked, so pressing the button again carries on from there.
      const msg=(failure&&failure.message)||String(failure);
      throw new Error(slidesIn.length?`${plural(slidesIn.length,'slide')} of ${slides.length} went in, then it stopped: ${msg} The others are still ticked; choose Insert again to carry on. ${UNDO}`:msg);
    }
    const without=slides.length-withPicture;
    const why=!slidePics.length?' (no pictures of the slides were chosen; they went in as text)':' (no matching picture file was chosen)';
    status(`${plural(slides.length,'slide')} inserted${without?`, ${without===slides.length&&!slidePics.length?'with no picture':`${without} of them without a picture`}${why}`:''}.${tidy()} Your cursor is in the last slide. To type below the slides, choose Write below this block. ${UNDO}`);
  }

  // ---- backups: a complete copy of the document saved as an ordinary .docx, and a saved backup opened again (guide-backup.mjs)
  let chosenBackup=null;
  const backupZip=()=>{
    const ZipClass=Zip||globalThis.JSZip;
    if(!ZipClass)throw new Error('The pane could not load what it needs to work with a backup. Close it, open it again from the ribbon, then try again.');
    return ZipClass;
  };
  const parserOf=()=>xml.DOMParserClass||globalThis.DOMParser;
  async function saveBackup(){
    status('Making a copy of the document...');
    await load();
    const modules=plan.guide.modules.filter(m=>m.index>=0);
    const live={isGuide:!!plan.guide.isGuide,modules:modules.length,lessons:modules.reduce((n,m)=>n+m.lessons.length,0)};
    const bytes=await adapter.captureFile();
    const info=await inspectBackup(bytes,backupZip(),{DOMParserClass:parserOf()});
    const bad=copyProblem(info,live);
    if(bad)throw new Error(bad);
    const name=backupFileName(plan.guide.title||info.title);
    adapter.offerDownload(toBase64(bytes),name);
    status(`Backup made: "${name}" (${sizeText(info.bytes)}). It was checked: it opens as a Word file and holds ${plural(info.modules,'module')}, ${plural(info.lessons,'lesson')}, ${plural(info.slides,'slide')} and ${plural(info.paragraphs,'paragraph')}, the same guide as the document. Look in your Downloads folder, or where Word asked you to save it. Keep another copy somewhere other than this computer's drive. To use it later, choose Restore Backup.`);
  }
  function forgetBackup(){
    chosenBackup=null;
    $('guide-bk-found').hidden=true;
    $('guide-bk-facts').replaceChildren();
  }
  // A .docx is read as it is. A .zip is the older backup: its package is checked as it always was and the Word copy is taken out of it.
  async function chooseBackup(){
    forgetBackup();
    const file=$('guide-bk-file').files?.[0];
    if(!file)return;
    status(`Reading ${file.name}...`);
    let bytes=new Uint8Array(await file.arrayBuffer()),from='',name=file.name;
    if(/\.zip$/i.test(file.name)){
      bytes=(await readCheckpoint(bytes,backupZip())).master;
      from=' (an older backup package: the Word copy was taken out of it and checked)';
      name=`${file.name.replace(/\.zip$/i,'')} - Restored.docx`;
    }
    const info=await inspectBackup(bytes,backupZip(),{DOMParserClass:parserOf()});
    chosenBackup={bytes,name,info};
    $('guide-bk-facts').replaceChildren(...backupFacts(info).map(t=>{const li=doc.createElement('li');li.textContent=t;return li;}));
    $('guide-bk-found').hidden=false;
    status(`${file.name} is a readable backup${from}. Check that this is the one you want, then choose Open this backup in a new window.`);
  }
  async function openBackup(){
    if(!chosenBackup)throw new Error('Choose a backup file first.');
    const {bytes,name}=chosenBackup,base64=toBase64(bytes);
    try{await adapter.openDocument(base64,name);status('The backup is open in a new Word window. The document you were working in has not changed. To go back to the backup, save that copy over your guide (File, Save As), or keep working in it.');}
    catch{adapter.offerDownload(base64,name);status(`Word could not open it for you, so it was downloaded as "${name}". Open it from your Downloads folder.`);}
  }

  // ---- update from slides: the changed presentation compared with the slides already in the guide (slides-sync.mjs)
  let upd=null;
  async function chooseUpdate(){
    const res=await adapter.pickFiles({kind:'slides'});
    if(!res||res.cancelled||!res.deck){status('No presentation was chosen.');return;}
    const d={name:res.deck.name,slides:res.deck.slides.map(s=>({...s}))},files=res.files||[];
    let byOrder=new Map(),pictures=null;
    if(files.length){
      const m=matchPictures(d.slides,files);
      if(m.problems.length)throw new Error(m.problems.join(' '));
      byOrder=m.byOrder;pictures=new Map([...byOrder].map(([order,f])=>[order,fingerprint(f.bytes)]));
    }
    let memory;
    try{memory=readMemory(await adapter.slideMemory?.());}catch{memory=readMemory('');}
    // Only a presentation with the same name is taken to be an earlier version of this one; the author can say otherwise below.
    const named=recordFor(memory,d.name);
    upd={deck:d,byOrder,pictures,memory,which:named?named.key:'',plan:null,record:null};
    await planUpdate();
  }
  async function planUpdate(){
    const u=upd;
    paras=await adapter.paragraphs();
    const blocks=slideBlocksIn(paras);
    if(!blocks.length){upd=null;drawUpdate();throw new Error('There are no slides from a presentation in this guide yet. Use Import slides first.');}
    u.record=u.which?(u.memory.decks.find(x=>x.key===u.which)||null):null;
    u.plan=planSlideUpdate({blocks,deck:u.deck,pictures:u.pictures,record:u.record});
    drawUpdate();
    const picNote=u.pictures?'':' No PDF was chosen, so a slide that has a picture cannot be rebuilt.';
    status(`${u.deck.name}: ${planSummary(u.plan)}${u.plan.items.length?' Tick what to bring in, then choose Update the guide.':''}${picNote}`);
  }
  function drawUpdate(){
    const u=upd;
    $('guide-upd-fieldset').hidden=!u||!u.plan||!u.plan.items.length;
    $('guide-upd-which-box').hidden=!u||!u.memory.decks.length;
    if(!u){$('guide-upd-summary').textContent='';$('guide-upd-list').replaceChildren();drawUpdateNote();return;}
    const none=doc.createElement('option');none.value='';none.textContent='A different presentation (slides are matched by their titles)';
    $('guide-upd-which').replaceChildren(none,...u.memory.decks.map(x=>{const o=doc.createElement('option');o.value=x.key;o.textContent=x.name;return o;}));
    $('guide-upd-which').value=u.which;
    $('guide-upd-summary').textContent=`${u.deck.name}: ${plural(u.deck.slides.length,'slide')}. ${planSummary(u.plan)}`;
    $('guide-upd-list').replaceChildren(...u.plan.items.map(it=>{
      const group=doc.createElement('div'),head=doc.createElement('p');
      group.className='guide-upd-item';head.className='guide-upd-head';
      head.textContent=`${it.label}: ${it.kind==='changed'?'changed':it.kind==='added'?'not in the guide':'no longer in the presentation'}`;
      group.append(head);
      if(it.blocked){const why=doc.createElement('p');why.className='hint';why.textContent=it.blocked;group.append(why);}
      for(const c of it.changes){
        const label=doc.createElement('label'),input=doc.createElement('input'),span=doc.createElement('span');
        label.className='guide-check';input.type='checkbox';input.id=`guide-upd-${it.key}-${c.id}`;input.checked=c.on&&!it.blocked;input.disabled=!!it.blocked;
        input.onchange=()=>{c.on=input.checked;drawUpdateNote();};
        span.textContent=` ${c.text}`;label.append(input,span);group.append(label);
      }
      return group;
    }));
    drawUpdateNote();
  }
  function drawUpdateNote(){
    const ticked=upd?.plan?operations(upd.plan).length:0;
    $('guide-upd-go').textContent=ticked?`Update the guide (${plural(ticked,'change')})`:'Update the guide';
  }
  const tickAll=on=>{if(!upd?.plan)return;for(const it of upd.plan.items)if(!it.blocked)for(const c of it.changes)c.on=on;drawUpdate();};
  async function applyUpdate(){
    const u=upd;
    if(!u?.plan)throw new Error('Choose the changed presentation first.');
    // The guide may have changed since the list was made: the plan is made again from what is there now, with the author's ticks carried
    // over, and if that gives different work the list is stale and nothing is done.
    const ticks=new Map(u.plan.items.flatMap(i=>i.changes.map(c=>[`${i.key}/${c.id}`,c.on])));
    const before=operations(u.plan).map(o=>`${o.type}:${o.key}`).join(',');
    paras=await adapter.paragraphs();
    const fresh=planSlideUpdate({blocks:slideBlocksIn(paras),deck:u.deck,pictures:u.pictures,record:u.record});
    for(const i of fresh.items)for(const c of i.changes)if(ticks.has(`${i.key}/${c.id}`))c.on=ticks.get(`${i.key}/${c.id}`);
    const ops=operations(fresh);
    if(ops.map(o=>`${o.type}:${o.key}`).join(',')!==before)throw new Error('Your guide changed since the list was made. Choose the presentation again to list what differs now.');
    if(!ops.length)throw new Error('Nothing is ticked, so nothing was changed.');
    // Every block is built before any change is made, so a refusal (too many pictures) leaves the guide as it was.
    let bytes=0;
    const picture=order=>{
      const f=u.byOrder.get(order);
      if(!f)return null;
      bytes+=f.bytes.length;
      if(bytes>MAX_TOTAL_PICTURE_BYTES)throw new Error(`These pictures add up to more than ${MAX_TOTAL_PICTURE_BYTES/1048576} MB. Update fewer slides at a time.`);
      return {base64:toBase64(f.bytes),type:f.type,mime:f.mime,size:pictureSize(f,'large'),name:f.name};
    };
    const share=!!$('guide-upd-share').checked;
    const work=ops.map(op=>{
      if(op.type==='replace'){
        const pic=op.wantsPicture?picture(op.deckOrder):null;
        if(op.wantsPicture&&!pic)throw new Error(`The PDF has no picture for slide ${op.deckOrder}.`);
        return {op,block:slideBlock(op.slide,{picture:pic,share:op.share})};
      }
      if(op.type==='insert')return {op,blocks:op.slides.map(s=>slideBlock(s,{picture:picture(s.order),share}))};
      return {op};
    });
    let done=0;
    try{
      for(const {op,block,blocks} of work){
        if(op.type==='replace')await adapter.insert(ctx=>cueBlock(block,ctx),{where:'after',after:op.span.end,select:'last',remove:{start:op.span.start,end:op.span.end,guard:op.guard}});
        else if(op.type==='delete')await adapter.deleteRange({start:op.span.start,end:op.span.end,guard:op.guard});
        else{
          const after=op.after===null?blockAnchor(await adapter.paragraphs(),await adapter.cursorIndex()):op.after;
          await adapter.insert(ctx=>blocks.map(b=>cueBlock(b,ctx)).join(''),{where:after<0?'start':'after',after,select:'last'});
        }
        done++;
      }
    }catch(e){
      upd=null;drawUpdate();
      await load().catch(()=>{});
      throw new Error(`${done} of ${plural(ops.length,'change')} were made before this stopped: ${(e&&e.message)||e} To take back what was done, click in your document and press Ctrl+Z several times.`);
    }
    let kept='';
    try{
      const {seen,entries}=entriesAfter(fresh,ops,{deck:u.deck,pictures:u.pictures});
      await adapter.setSlideMemory?.(writeMemory(remember(u.memory,{name:u.deck.name,seen,entries})));
    }catch{kept=' The guide could not keep what it remembers about these slides, so the next update will match them by their titles.';}
    upd=null;drawUpdate();
    await keepOpen();
    await load();
    status(`${plural(done,'change')} made.${kept}${tidy()} ${UNDO_MANY}`);
  }

  // ---- E-Learning: screens (a run of blocks), their narration counted and timed, and the script for the voice talent
  function drawElearning(){
    const screens=screensOf(paras);
    if(!screens.length){$('guide-el-summary').textContent='No screens yet. Add one above, or use the E-Learning cues in the list of cues.';return;}
    const s=narrationStats(screens);
    $('guide-el-summary').textContent=`${plural(s.screens,'screen')}. Narration: ${plural(s.words,'word')}, about ${s.minutes} minute${s.minutes===1?'':'s'} at ${WORDS_PER_MINUTE} words a minute.${s.lessons.length>1?` By lesson: ${s.lessons.map(l=>`${l.lesson||'no lesson'} ${l.minutes} min`).join('; ')}.`:''}`;
  }
  async function addScreen(){
    const title=$('guide-el-title').value.trim(),minutes=parseMinutes($('guide-el-minutes').value);
    const interaction=!!$('guide-el-interaction').checked;
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),anchor=blockAnchor(list,here);
    await adapter.insert(ctx=>screenBlocks({title,minutes,interaction},ctx),{where:'after',after:anchor,select:'last'});
    $('guide-el-title').value='';$('guide-el-minutes').value='';$('guide-el-interaction').checked=false;
    await keepOpen();
    await load();
    status(`Screen "${title}" added. Type over the sample sentence under each part.${tidy()} ${UNDO}`);
    focus('guide-el-title');
  }
  async function makeScript(kind){
    const list=await adapter.paragraphs(),screens=screensOf(list);
    if(!screens.length)throw new Error('There are no screens yet. Add a screen first.');
    if(!screens.some(s=>s.narration))throw new Error('No screen has narration yet. Type over the sample sentence under Narration, then make the script.');
    const title=String(await adapter.title()).trim()||plan?.guide?.title||'Guide';
    if(kind==='text'){
      const name=scriptFileName(title,'txt');
      adapter.offerDownload(toBase64(new TextEncoder().encode(scriptText(screens,title))),name,'text/plain');
      status(`The script was saved as "${name}" in your Downloads folder.`);
      return;
    }
    const ZipClass=Zip||globalThis.JSZip;
    if(!ZipClass)throw new Error('The pane could not load what it needs to make the script. Close it, open it again from the ribbon, then try again.');
    const base64=await zipWorkbook(scriptFiles(screens,title),ZipClass),name=scriptFileName(title,'docx');
    try{await adapter.openDocument(base64,name);status(`The narration script is open in a new Word window. Save it as "${name}", or any name you like.`);}
    catch{adapter.offerDownload(base64,name);status(`Word could not open it for you, so it was downloaded as "${name}". Open it from your Downloads folder.`);}
  }
  // The storyboard as an OpenStage Author project, in a zip: nothing is changed in the document.
  async function sendToAuthor(){
    const list=await adapter.paragraphs(),screens=screensOf(list);
    if(!screens.length)throw new Error('There are no screens yet. Add a screen first.');
    const title=String(await adapter.title()).trim()||plan?.guide?.title||'Course';
    const ZipClass=Zip||globalThis.JSZip;
    if(!ZipClass)throw new Error('The pane could not load what it needs to make the project. Close it, open it again from the ribbon, then try again.');
    const base64=await zipWorkbook(authorProject(screens,{title}),ZipClass),name=authorZipName(title);
    adapter.offerDownload(base64,name,'application/zip');
    status(`${plural(screens.length,'screen')} sent: the project was saved as "${name}" in your Downloads folder. Unzip it, then open the folder that ends in .osproj in OpenStage Author. The notes file beside it lists everything else the storyboard says.`);
  }
  // The Start screen's choice of what is being made: a storyboard starts with its own look and without the facilitator's pages.
  function buildKinds(){
    $('guide-start-kind').replaceChildren(...[{id:'',name:'A facilitator guide'},{id:'storyboard',name:'An E-Learning storyboard'}].map(k=>{const o=doc.createElement('option');o.value=k.id;o.textContent=k.name;return o;}));
    $('guide-start-kind').value='';
    $('guide-start-kind').onchange=()=>{
      const story=$('guide-start-kind').value==='storyboard';
      $('guide-standard').checked=!story;$('guide-start-look').value=story?'storyboard':START_LOOK;
    };
  }

  // ---- the look: the guide's colors and fonts, kept as real styles and remembered in the document
  // Until the new way of changing styles has been proven in the author's Word, a new guide starts with the standard styles.
  const START_LOOK='';
  let look=null,lookRaw=null;
  const guideLookInputs=new Map();
  const lookName=id=>GUIDE_LOOKS.find(l=>l.id===id)?.name||id;
  function buildGuideLooks(){
    $('guide-looks').replaceChildren(...GUIDE_LOOKS.map(l=>{
      const label=doc.createElement('label'),input=doc.createElement('input'),name=doc.createElement('span');
      label.className='guide-cue';
      input.type='radio';input.name='guide-look';input.id=`guide-look-${l.id}`;input.value=l.id;input.checked=l.id==='clearline';guideLookInputs.set(l.id,input);
      name.textContent=`${l.name}: ${l.note}`;
      label.append(input,name);
      return label;
    }));
    $('guide-start-look').replaceChildren(...[{id:'',name:'Standard styles (no look yet)',note:''},...GUIDE_LOOKS].map(l=>{
      const o=doc.createElement('option');o.value=l.id;o.textContent=l.id?`${l.name}: ${l.note}`:l.name;return o;
    }));
    $('guide-start-look').value=START_LOOK;
  }
  const chosenGuideLook=()=>GUIDE_LOOKS.find(l=>guideLookInputs.get(l.id)?.checked)?.id||'clearline';
  const readLookForm=()=>({id:chosenGuideLook(),accent:$('guide-look-accent').value,organization:$('guide-look-org').value,header:!!$('guide-look-header').checked,copyright:$('guide-look-copyright').value,revised:$('guide-look-revised').value});
  function drawLook(){
    $('guide-look-now').textContent=look
      ?`This guide uses the ${lookName(look.id)} look${look.accent?`, accent color ${look.accent}`:''}${look.header?', with a running header and footer':''}.`
      :'This guide has no look yet, so it uses the standard styles. Open "Change the look" to choose one.';
    const s=look||cleanLook(null);
    for(const l of GUIDE_LOOKS)guideLookInputs.get(l.id).checked=l.id===s.id;
    $('guide-look-accent').value=s.accent;$('guide-look-org').value=s.organization;$('guide-look-header').checked=s.header;
    $('guide-look-copyright').value=s.copyright;$('guide-look-revised').value=s.revised;
  }
  // The workbook starts from the guide's look unless the author has already chosen one for it.
  function shareLookWithWorkbook(){
    if(!look)return;
    let saved=null;try{saved=storage?.getItem(LOOK_KEY)??null;}catch{saved=null;}
    if(saved!==null)return;
    const family=GUIDE_LOOKS.find(l=>l.id===look.id)?.family||look.id; // a look of the guide's own uses its family's workbook look
    for(const [id,input] of lookInputs)input.checked=id===family;
    $('guide-wb-org').value=look.organization;$('guide-wb-accent').value=look.accent;
  }
  // Reads the look stored in the document and tells the adapter, so styles added later are made in the same look.
  async function readStoredLook(){
    let raw=null;
    try{raw=(await adapter.getLook?.())??null;}catch{/* an unreadable look is the same as none */}
    if(raw===lookRaw)return;
    lookRaw=raw;look=raw?parseLook(raw):null;
    const r=look?resolveGuideLook(look):null;
    adapter.setTheme?.(r?.ok?r.theme:undefined);
    drawLook();shareLookWithWorkbook();offerStoryboardCues();
  }
  // A storyboard offers the E-Learning cues first; the author can still choose any other kind.
  function offerStoryboardCues(){
    if(look?.id==='storyboard'&&!$('guide-cue-group').value){$('guide-cue-group').value='elearning';filterCues();}
  }
  // Changes the styles that exist, the header and footer and the stored look. Returns what Word did and did not accept.
  async function applyLookTo(settings,titleHint=''){
    const r=resolveGuideLook(settings);
    if(!r.ok)throw new Error(r.issues.join(' '));
    const s=r.settings,parser=xml.DOMParserClass||globalThis.DOMParser;
    adapter.setTheme?.(r.theme);
    const styles=adapter.applyLook?await adapter.applyLook(lookSpecs(r.theme,parser)):{applied:[],skipped:[],refused:[]};
    const title=String(await adapter.title()).trim()||titleHint||plan?.guide?.title||'';
    const had=!!look?.header;
    // undefined leaves the header or footer alone; null takes the one the look added away.
    const header=s.header?headerPackage({title}):(had?null:undefined);
    const footer=s.header?footerPackage({organization:s.organization,copyright:s.copyright,revised:s.revised}):(had?null:undefined);
    const hf=adapter.setHeaderFooter&&(header!==undefined||footer!==undefined)?await adapter.setHeaderFooter({header,footer}):{done:[],failed:[]};
    const text=JSON.stringify(s);
    await adapter.setLook?.(text);
    look=s;lookRaw=text;
    // Drawn here too: when a guide is started with a look, the next load finds the stored look equal to this one and draws nothing
    // (seen in real Word, 2026-11-06: the pane said "no look yet" and showed Clearline for a guide started in the Storyboard look).
    drawLook();shareLookWithWorkbook();
    offerStoryboardCues();
    return {styles,hf,name:lookName(s.id)};
  }
  function lookReport({styles,hf,name}){
    const parts=[`The ${name} look is applied.`];
    if(styles.refused.length){
      const byStyle=new Map();
      for(const f of styles.refused)byStyle.set(f.style,[...(byStyle.get(f.style)||[]),f.what]);
      parts.push(`Word would not change ${[...byStyle].map(([style,what])=>`${style} (${what.join(', ')})`).join('; ')}, and the rest was changed. To finish by hand, open the Styles pane in Word (Home, the small arrow beside Styles) and change those styles' settings.`);
    }
    // Word's own styles (body text, title, headings) are looked up by their English names; if one is not found, the look is incomplete.
    const missing=styles.skipped.filter(n=>['Normal','Title','Heading 1','Heading 2'].includes(n));
    if(missing.length)parts.push(`Word did not find its own ${missing.join(', ')} style${missing.length===1?'':'s'}, so ${missing.length===1?'it was':'they were'} not changed. Change ${missing.length===1?'it':'them'} by hand in the Styles pane.`);
    for(const f of hf.failed)parts.push(`The ${f.what} could not be set. Add it in Word: Insert > ${f.what==='header'?'Header':'Footer'}.`);
    return parts.join(' ');
  }
  async function applyLookNow(){
    $('guide-look-note').textContent=''; // an old "... look is applied." must not stay beside a refusal (seen in real Word, 2026-11-06)
    const out=await applyLookTo(readLookForm());
    shareLookWithWorkbook();
    await load();
    drawLook();
    $('guide-look-note').textContent=lookReport(out);
    status(`${lookReport(out)} To go back, choose the earlier look and apply it again.`);
  }

  // ---- standard pages
  async function addStandardPages(){
    const list=await adapter.paragraphs(),{pages,cannot}=missingStandardPages(list);
    if(!pages.length){
      status(cannot.length?`${cannot.join(' and ')} could not be added: this guide has none of the pages it would sit beside.`:'All the standard pages are already in your guide.');
      return;
    }
    // The page that comes last in the document goes in first, so the place of the next one does not move.
    for(const page of pages)
      await adapter.insert(()=>page.page==='acknowledgements'?acknowledgementsBody():perspectiveBody(),{where:page.after<0?'start':'after',after:page.after,select:'first'});
    await adapter.refreshFields();
    await keepOpen();
    await load();
    const names=[...pages].reverse().map(p=>p.name).join(' and ');
    status(`Added ${names}. Type over the sentence under each heading.${cannot.length?` ${cannot.join(' and ')} could not be placed.`:''}${tidy()} ${UNDO}`);
  }
  async function addAppendix(){
    const list=await adapter.paragraphs(),place=appendixPlacement(list);
    if(place.exists){
      await adapter.go(place.at,list[place.at].text);
      await load();
      status('This guide already has an Appendix. Moved to it.');
      return;
    }
    await adapter.insert(()=>appendixBody(),{where:'after',after:place.anchor,select:'first'});
    await adapter.refreshFields();
    await keepOpen();
    await load();
    status(`Appendix added at the end of the guide. Type over the sentence under its heading.${tidy()} ${UNDO}`);
  }

  // ---- moving and deleting a whole lesson or module
  // After either, the document is read again, headings are renumbered, lesson times follow the new order and the contents table is refreshed.
  async function tidyUp(){
    await load();
    if(plan.edits.length){await adapter.applyEdits(plan.edits);await load();}
    await adapter.refreshFields();
    await load();
  }
  async function moveUnit(dir){
    cancelDelete();
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),target=tidyTarget(list,here);
    if(!target)throw new Error('Put your cursor inside a module or lesson first.');
    const move=tidyMove(list,target,dir);
    if(move.message){status(move.message);return;}
    await adapter.moveRange({start:move.span.start,end:move.span.end,guard:list[move.span.start].text,after:move.after});
    await tidyUp();
    status(`${clean(paras[at]?.text)||target.text} moved ${dir<0?'up':'down'}. The numbers and times are up to date, and the contents table is refreshed. ${UNDO_MANY}`);
  }
  // Delete asks first: it names what will go, and nothing is deleted until the author says yes.
  async function askDelete(){
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),target=tidyTarget(list,here);
    if(!target)throw new Error('Put your cursor inside a module or lesson first.');
    const d=tidyDeletion(list,target);
    pendingDelete={start:d.span.start,text:target.text};
    const inside=target.kind==='module'?`the module with its ${plural(d.lessons,'lesson')}, its Transition block and anything you typed in it`:'the lesson with its time block, its blocks and anything you typed in it';
    $('guide-delete-text').textContent=`Delete ${target.text}? This removes ${inside}: ${plural(d.paragraphs,'paragraph')} in all. You can take it back: click in your document and press Ctrl+Z several times.`;
    $('guide-delete-confirm').hidden=false;
  }
  async function confirmDelete(){
    const asked=pendingDelete;
    if(!asked)return;
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),target=tidyTarget(list,here);
    cancelDelete();
    if(!target||target.span.start!==asked.start||target.text!==asked.text)throw new Error('Your cursor has moved since you chose Delete, so nothing was deleted. Put your cursor in the one you want to delete, then choose Delete again.');
    await adapter.deleteRange({start:target.span.start,end:target.span.end,guard:list[target.span.start].text});
    await tidyUp();
    status(`${target.text} deleted. The numbers and times are up to date, and the contents table is refreshed. ${UNDO_MANY}`);
  }

  // ---- editing what is already there
  function startEdit(){
    const here=describeAt(paras,at);
    if(!here)return;
    if(here.type==='block'){
      const b=here.block;
      if(b.hasPicture)throw new Error('This block has a picture, and the pane cannot rewrite a block that holds one. To change its words, type in Word. To change the picture, select the block\'s lines in Word, press Delete, and add the block again. (Move or delete works on a whole lesson, not on one block.)');
      editing={type:'block',block:b};
      chooseCue(b.cue,{audience:false});
      $('guide-block-text').value=b.text;$('guide-block-notes').value=b.notes;$('guide-block-minutes').value=b.minutes;
      // An activity written as labelled parts opens with its parts filled in; any other activity opens as plain text.
      if(b.cue==='activity'){
        $('guide-activity-parts').checked=!!b.parts;
        for(const p of ACTIVITY_PARTS)$(`guide-act-${p.id}`).value=b.parts?.[p.id]??'';
        syncParts();
      }
      for(const a of AUDIENCES)$(`guide-audience-${a.id}`).checked=a.id===b.audience;
      $('guide-block-heading').textContent='Edit a block';
      $('guide-block-go').textContent='Update block';
      $('guide-block-cancel').hidden=false;
      $('guide-block-note').textContent=`You are editing the ${here.label} that starts "${snippet(b.text)}". Change it, then choose Update block.`;
      drawSelected();
      focusBlockText();
      return;
    }
    const lesson=here.type==='lesson'?here.lesson:null,heading=lesson||here.module;
    editing={type:here.type,index:heading.index,text:clean(heading.text),lesson};
    $('guide-edit-heading').textContent=lesson?'Edit this lesson':'Edit this module';
    $('guide-edit-title').value=heading.title;
    $('guide-edit-minutes').hidden=$('guide-edit-minutes-label').hidden=!lesson;
    $('guide-edit-minutes').disabled=false;
    $('guide-edit-note').textContent='';
    if(lesson){
      $('guide-edit-minutes').value=String(lessonMinutes(lesson));
      if(lesson.blockMinutes>0){
        $('guide-edit-minutes').disabled=true;
        $('guide-edit-note').textContent=`This lesson's time is the total of its blocks' minutes (${lesson.blockMinutes}). To change it, edit the minutes on those blocks.`;
      }
    }
    $('guide-edit-form').hidden=false;
    drawSelected();
    ensureTool(['addModule'],'addModule');
    showPart('add');
    focus('guide-edit-heading');
  }
  function endEdit(){
    const wasBlock=editing?.type==='block';
    editing=null;
    $('guide-edit-form').hidden=true;
    if(wasBlock)clearBlockForm();
    drawSelected();
  }
  async function updateBlock(){
    const b=editing.block,block=readBlockForm();
    // A slide's own text is not part of the block form, so editing the block's words or notes keeps it.
    if(b.slideText)block.slideText=b.slideText;
    // Find the block again in case the document moved on since it was chosen.
    const list=await adapter.paragraphs();
    const hits=list.map((p,i)=>kindOf(p)==='cue'&&clean(p.text)===b.firstText?i:-1).filter(i=>i>=0).sort((x,y)=>Math.abs(x-b.span.start)-Math.abs(y-b.span.start));
    const span=hits.length?blockSpan(list,hits[0]):null;
    if(!span)throw new Error('That block has changed since you chose it. Select it in your document again, then choose Edit this block.');
    await adapter.insert(ctx=>cueBlock(block,ctx),{where:'after',after:span.end,select:'last',remove:{start:span.start,end:span.end,guard:b.firstText}});
    endEdit();
    await load();
    status(`${block.cue.label} block updated.${tidy()} ${UNDO}`);
    focusBlockText();
  }
  async function saveEdit(){
    const e=editing;
    if(!e||e.type==='block')return;
    const title=$('guide-edit-title').value.trim();
    if(!title)throw new Error('Type a title first.');
    const minutes=e.lesson&&!$('guide-edit-minutes').disabled?parseMinutes($('guide-edit-minutes').value,'Lesson minutes'):null;
    const list=await adapter.paragraphs(),now=updatePlan(list,{forms:formSettings});
    if(clean(list[e.index]?.text)!==e.text)throw new Error(`${e.lesson?'That lesson':'That module'} has changed since you chose it. Select it in your document again, then choose Edit.`);
    const prefix=new RegExp(`^${e.lesson?'Lesson\\s+\\d+\\.\\d+':'Module\\s+\\d+'}\\s*:\\s*`).exec(e.text);
    const edits=[{kind:'heading',index:e.index,text:`${prefix?prefix[0]:''}${title}`}];
    const lesson=e.lesson&&now.guide.modules.flatMap(m=>m.lessons).find(l=>l.index===e.index);
    if(lesson&&minutes!==null&&minutes!==lessonMinutes(lesson)&&lesson.timeIndex>=0)edits.push({kind:'time',index:lesson.timeIndex,minutes,cumulative:lesson.timeCumulative??0});
    await adapter.applyEdits(edits);
    endEdit();
    await load();
    if(plan.edits.length){await adapter.applyEdits(plan.edits);await load();}
    status(`${e.lesson?'Lesson':'Module'} updated, and the numbers and times are up to date. ${UNDO}`);
  }
  // Brings the heading numbers, the lesson times and the contents table up to date, in that order.
  async function updateGuide(){
    const count=plan.edits.length;
    if(count)await adapter.applyEdits(plan.edits);
    const refreshed=await adapter.refreshFields();
    await load();
    const changed=count?`Updated ${countText(count)}${refreshed?' and the contents table':''}.`:`Numbers and times were already right.${refreshed?' The contents table is refreshed.':''}`;
    status(`${changed}${refreshed?'':' Word could not refresh the contents table here: right-click it and choose Update Field.'} ${UNDO}`);
  }
  // The Getting Started pages built from the guide. Each replaces only what an earlier build wrote (or the untouched
  // placeholder), so the author's own text on the page is never removed; nothing else ever rebuilds them.
  const pageInsert=place=>({where:place.after<0?'start':'after',after:place.after,select:'first',remove:place.remove});
  const pageHow=(place,name)=>place.withHeading?`There was no ${name} heading, so one was added.`:place.replaces==='page'?'The earlier one was replaced.':place.replaces==='placeholder'?'It took the place of the placeholder text.':'It is under the heading, above your own text.';
  async function buildGlancePage(){
    const list=await adapter.paragraphs(),guide=readGuide(list),place=glancePlacement(list);
    await adapter.insert(()=>glanceBody(guide,{withHeading:place.withHeading}),pageInsert(place));
    await adapter.refreshFields();
    await keepOpen();
    await load();
    const lessons=glanceRows(guide).rows.filter(r=>r.kind==='lesson').length;
    status(`Training at a Glance built from ${plural(lessons,'lesson')}. ${pageHow(place,'Training at a Glance')}${tidy()} ${UNDO}`);
  }
  async function buildCuesPage(){
    const list=await adapter.paragraphs(),entries=cuesUsed(list),place=cuesPlacement(list);
    await adapter.insert(ctx=>cuesBody(entries,ctx,{withHeading:place.withHeading}),pageInsert(place));
    await adapter.refreshFields();
    await keepOpen();
    await load();
    status(`The Graphic Cues page built with ${plural(entries.length,'cue')}. ${pageHow(place,'Graphic Cues')}${tidy()} ${UNDO}`);
  }
  async function buildOverviewPage(){
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex(),mod=moduleAt(list,here);
    if(!mod)throw new Error('Put your cursor inside a module first. The overview is built for the module the cursor is in.');
    const facts=overviewFacts(list,mod),place=overviewPlacement(list,mod);
    await adapter.insert(()=>overviewBody(facts,place),pageInsert(place));
    await keepOpen();
    await load();
    const tail=place.withAuthored?' Type over the sentence after Goal, Overview, Materials needed and References and resources.':place.replaces?' Your own lines were not touched.':'';
    status(`The overview of ${mod.text||'the module'} is built from its ${plural(facts.lessons.length,'lesson')}.${tail}${tidy()} ${UNDO}`);
  }
  async function review(){
    paras=await adapter.paragraphs();
    plan=updatePlan(paras,{forms:formSettings});
    const parts=await adapter.reviewParts();
    // Question numbers the guide holds, so a Knowledge check can be checked against them.
    const itemNumbers=new Set(plan.assess.sections.flatMap(s=>s.items.map(i=>i.number)).filter(n=>n!==null&&n!==undefined));
    const found=[...guideIssues(plan),...elearningIssues(paras,itemNumbers),...checkParts(parts).findings.filter(f=>f.level==='fail'||f.level==='warn')];
    const fix=found.filter(f=>f.level==='fail'),look=found.filter(f=>f.level==='warn');
    const fill=(id,items,empty)=>{
      const list=$(id);
      if(!items.length){const li=doc.createElement('li');li.textContent=empty;list.replaceChildren(li);return;}
      list.replaceChildren(...items.slice(0,40).map(item=>{
        const li=doc.createElement('li'),text=doc.createElement('span');
        text.textContent=item.message;li.append(text);
        if(item.paragraph!==undefined){
          // The words at that place are remembered, so the link finds them again if the document has moved on since.
          const there=paras[item.paragraph]?.text;
          const b=doc.createElement('button');
          b.type='button';b.className='guide-link';b.textContent='Go there';b.setAttribute('aria-label',`Go to: ${item.message}`);
          b.onclick=()=>act(async()=>{await adapter.go(item.paragraph,there);await load();status('Moved to the place in the document.');});
          li.append(b);
        }
        return li;
      }));
    };
    fill('guide-review-fix',fix,'Nothing to fix.');
    fill('guide-review-look',look,'Nothing to look at.');
    fill('guide-review-check',CHECK_YOURSELF.map(message=>({message})),'');
    $('guide-review-summary').textContent=`${fix.length} to fix, ${look.length} to look at, ${CHECK_YOURSELF.length} to check yourself. This is not Word's Accessibility Checker and not a screen-reader test.`;
    status('');
    hidePdfMessages();
    show('review',{moveFocus:true});
    return {fix,look};
  }
  // A reviewed PDF: the review runs first, and with anything to fix no PDF is made. Otherwise Word makes its own PDF, which is
  // checked for the markers of a tagged one before it is saved.
  function hidePdfMessages(){$('guide-review-pdf-problem').hidden=true;$('guide-review-pdf-ok').hidden=true;}
  function pdfMessage(kind,text){
    $('guide-review-pdf-problem').hidden=kind!=='problem';$('guide-review-pdf-ok').hidden=kind!=='ok';
    $(kind==='problem'?'guide-review-pdf-problem':'guide-review-pdf-ok').textContent=text;
    $(kind==='problem'?'guide-review-pdf-problem':'guide-review-pdf-ok').scrollIntoView?.({block:'nearest'});
  }
  async function makePdf(){
    const {fix}=await review();
    if(fix.length){pdfMessage('problem',`No PDF was made: ${plural(fix.length,'thing')} to fix first, listed above. Fix them, then choose Save as PDF again.`);return;}
    let base64;
    try{base64=await adapter.exportPdf();}
    catch(e){
      const said=(e&&e.message)||'';
      // The program (no Word behind it) says its own whole message, with how to get a PDF.
      if(/^The program does not make PDFs yet\./.test(said)){pdfMessage('problem',said);return;}
      pdfMessage('problem',`Word could not make the PDF. ${said} You can make one yourself: File > Export > Create PDF/XPS, with "Document structure tags for accessibility" ticked.`.trim());return;
    }
    const tags=checkPdfTags(Uint8Array.from(atob(base64),ch=>ch.charCodeAt(0)));
    const name=pdfFileName(plan.guide.title);
    adapter.offerDownload(base64,name,'application/pdf');
    pdfMessage(tags.tagged==='yes'?'ok':'problem',`The PDF was saved as "${name}" in your downloads. ${tags.tagged==='yes'?'':'It was saved, but: '}${tags.message}`);
  }

  // ---- assessments: sections and the items inside them
  let selectedType='multiple-choice',editingItem=null;
  const OPTION_NUMBERS=[1,2,3,4,5,6];
  function drawAssessSelected(){
    const found=editingItem?null:describeAt(paras,at),here=found&&found.type==='item'?found:null;
    $('guide-assess-selected').hidden=!here;
    if(here)$('guide-assess-selected-text').textContent=`Your cursor is in ${here.label}.`;
    // Before anything is typed, say where an item would go, so a missing section is never a surprise.
    const target=editingItem?null:itemTarget(paras,at),where=$('guide-item-where');
    where.hidden=!!editingItem;
    where.setAttribute('data-state',target?'ready':'none');
    where.textContent=target?`Items go in: ${target.title}${target.after?`, after ${target.after}`:''}.`:'No assessment section here yet. Start one above, or click inside one in your document, then choose Insert item.';
    if(target)hideItemProblem();
  }
  // Messages about the item form sit just above Insert item, where the author is looking, not at the top of the pane.
  const syncItemMessages=()=>{$('guide-item-msg-box').hidden=$('guide-item-problem').hidden&&$('guide-item-ok').hidden&&$('guide-item-offer').hidden;};
  function hideItemProblem(){$('guide-item-problem').hidden=true;$('guide-item-offer').hidden=true;syncItemMessages();}
  function hideItemMessages(){hideItemProblem();$('guide-item-ok').hidden=true;syncItemMessages();}
  function itemProblem(e,{offer=false}={}){
    $('guide-status').textContent='';$('guide-error').hidden=true;$('guide-error').textContent='';
    $('guide-item-ok').hidden=true;
    $('guide-item-msg-box').hidden=false;
    $('guide-item-problem').hidden=false;
    $('guide-item-problem').textContent=typeof e==='string'?e:(e&&e.message)||String(e);
    $('guide-item-offer').hidden=!offer;
    $('guide-item-msg-box').scrollIntoView?.({block:'nearest'});
  }
  function itemOk(text){
    $('guide-item-problem').hidden=true;$('guide-item-offer').hidden=true;
    $('guide-item-msg-box').hidden=false;
    $('guide-item-ok').hidden=false;
    $('guide-item-ok').textContent=text;
    $('guide-item-msg-box').scrollIntoView?.({block:'nearest'});
  }
  function chooseType(id,{keepLines=false}={}){
    selectedType=id;
    const t=itemTypeById(id);
    $('guide-item-help').textContent=t.help;
    $('guide-type-current').textContent=t.label;
    for(const label of $('guide-item-types').children)label.children[0].checked=label.children[0].id===`guide-item-type-${id}`;
    $('guide-item-options').hidden=!t.options;
    $('guide-item-options-legend').textContent=t.pick==='many'?'Options (tick every correct one)':'Options (tick the correct one)';
    $('guide-item-truefalse').hidden=!t.truefalse;
    $('guide-item-answer-box').hidden=!t.answer;
    if(t.answer)$('guide-item-answer-label').textContent=t.answer;
    $('guide-item-numeric').hidden=!t.numeric;
    $('guide-item-rubric').hidden=!t.rubric;
    // A new rubric starts with three levels, which the author can change.
    if(t.rubric&&!$('guide-item-rubric-levels').value.trim())$('guide-item-rubric-levels').value=DEFAULT_RUBRIC_LEVELS;
    $('guide-item-matching').hidden=!t.matching;
    $('guide-item-rating').hidden=!t.rating;
    $('guide-item-ordering').hidden=!t.ordering;
    $('guide-item-performance').hidden=!t.performance;
    $('guide-item-semantic').hidden=!t.semantic;
    $('guide-item-action').hidden=!t.action;
    $('guide-item-lines-box').hidden=!t.lines;
    if(t.lines&&!keepLines)$('guide-item-lines').value=String(t.lines);
    $('guide-item-key-box').hidden=!!t.noKey;
    $('guide-item-stem-label').textContent=t.group==='survey'?'Prompt':'Question';
    $('guide-item-stem-hint').textContent=t.stemHint||'';
  }
  // Like the cue list: the type of item shows as its name and a Change button, and the list stays put away until asked for.
  function setTypePanel(open){
    $('guide-type-panel').hidden=!open;
    $('guide-type-toggle').setAttribute('aria-expanded',String(open));
    $('guide-type-toggle-action').textContent=open?'Close':'Change';
  }
  $('guide-type-toggle').onclick=()=>{const open=$('guide-type-panel').hidden;setTypePanel(open);if(open)focus(`guide-item-type-${selectedType}`);};
  setTypePanel(false);
  function buildTypes(){
    $('guide-item-types').replaceChildren(...ITEM_TYPES.map(t=>{
      const label=doc.createElement('label'),input=doc.createElement('input'),name=doc.createElement('span');
      label.className='guide-cue';
      input.type='radio';input.name='guide-item-type';input.id=`guide-item-type-${t.id}`;input.value=t.id;input.checked=t.id===selectedType;
      input.onchange=()=>{chooseType(t.id);hideItemMessages();};
      input.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault?.();setTypePanel(false);focus('guide-item-stem');}};
      label.onclick=e=>{if(e&&e.detail>0)globalThis.setTimeout?.(()=>{setTypePanel(false);focus('guide-item-stem');},0);};
      name.textContent=t.label;
      label.append(input,name);
      return label;
    }));
    chooseType(selectedType);
  }
  const parsePoints=raw=>{
    const t=String(raw??'').trim();
    if(t&&!/^\d{1,3}(\.\d{1,2})?$/.test(t))throw new Error('Points must be a number, like 1 or 2.5.');
    return t;
  };
  function readItemForm(){
    const t=itemTypeById(selectedType);
    const linesText=String($('guide-item-lines').value??'').trim();
    if(t.lines&&linesText&&!/^\d{1,2}$/.test(linesText))throw new Error('Lines to write on must be a whole number, like 3.');
    return {
      type:t.id,stem:$('guide-item-stem').value,
      options:OPTION_NUMBERS.map(n=>$(`guide-item-opt-${n}`).value),
      correct:OPTION_NUMBERS.filter(n=>$(`guide-item-correct-${n}`).checked).map(n=>n-1),
      truth:$('guide-item-true').checked?true:$('guide-item-false').checked?false:undefined,
      answer:$('guide-item-answer').value,tolerance:$('guide-item-tolerance').value,units:$('guide-item-units').value,
      why:$('guide-item-why').value,points:parsePoints($('guide-item-points').value),
      objective:$('guide-item-objective').value,lines:t.lines&&linesText?Number(linesText):undefined,
      pairs:parsePairs($('guide-item-pairs').value),extras:parseLines($('guide-item-extras').value),
      statements:parseLines($('guide-item-statements').value),
      scale:$('guide-item-scale-custom').checked?'custom':$('guide-item-scale-success').checked?'success':'agree',
      labels:parseLabels($('guide-item-labels').value),
      orderItems:parseLines($('guide-item-order-items').value),criteria:parseCriteria($('guide-item-criteria').value),
      levels:parseCriteria($('guide-item-rubric-levels').value),rubricRows:parseRubricRows($('guide-item-rubric-rows').value),
      learnerRubric:t.rubric&&!!$('guide-item-rubric-learners').checked,
      sdPairs:parsePairs($('guide-item-sd-pairs').value),sdPoints:$('guide-item-sd-7').checked?7:5,
      actions:String($('guide-item-actions').value??'').trim()===''?3:Number(String($('guide-item-actions').value).trim())
    };
  }
  const setScale=id=>{
    $('guide-item-scale-agree').checked=id!=='success'&&id!=='custom';
    $('guide-item-scale-success').checked=id==='success';
    $('guide-item-scale-custom').checked=id==='custom';
  };
  function clearItemForm(){
    for(const id of ['guide-item-stem','guide-item-answer','guide-item-tolerance','guide-item-units','guide-item-why','guide-item-points','guide-item-objective','guide-item-pairs','guide-item-extras','guide-item-statements','guide-item-labels','guide-item-order-items','guide-item-criteria','guide-item-rubric-levels','guide-item-rubric-rows','guide-item-sd-pairs',...OPTION_NUMBERS.map(n=>`guide-item-opt-${n}`)])$(id).value='';
    setScale('agree');
    $('guide-item-sd-5').checked=true;$('guide-item-sd-7').checked=false;$('guide-item-actions').value='3';
    for(const n of OPTION_NUMBERS)$(`guide-item-correct-${n}`).checked=false;
    $('guide-item-rubric-learners').checked=false;
    $('guide-item-true').checked=false;$('guide-item-false').checked=false;
    editingItem=null;
    $('guide-item-heading').textContent='Add an item';
    $('guide-item-go').textContent='Insert item';
    $('guide-item-cancel').hidden=true;
    $('guide-item-note').textContent='';
    chooseType(selectedType);
    if(plan)drawAssessSelected();
  }
  async function startSection(levelId,{fromItem=false}={}){
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex();
    const {anchor}=sectionPlacement(list,here);
    await adapter.insert(()=>sectionBlock(levelId),{where:'after',after:anchor,select:'first'});
    await keepOpen();
    await load();
    status(`${levelById(levelId).title} added. Add its items below.${tidy()} ${UNDO}`);
    if(fromItem){
      // Started from the message above Insert item: stay by the button, with what was typed still in the form.
      itemOk(`${levelById(levelId).title} added. What you typed is still in the form. Choose Insert item to put it in.`);
    }else focusItemStem();
  }
  async function addItem(){
    hideItemMessages();
    try{
      if(editingItem)return await updateItem();
      const item=readItemForm(),t=itemTypeById(item.type);
      const list=await adapter.paragraphs(),here=await adapter.cursorIndex();
      if(!itemTarget(list,here)){
        // Nothing has been changed and nothing typed is lost; say what to do right where the author clicked.
        itemProblem('Items go inside an assessment section, and your cursor is not in one. What you typed is still here. Start a section, or click inside one in your document, then choose Insert item again.',{offer:true});
        return;
      }
      const {anchor,number}=itemPlacement(list,here);
      await adapter.insert(ctx=>itemBlock({...item,number},ctx),{where:'after',after:anchor,select:'last'});
      clearItemForm();
      await keepOpen();
      await load();
      status(`${t.label} item added.${tidy()} ${UNDO}`);
      focusItemStem();
    }catch(e){itemProblem(e);}
  }
  // ---- pasting many questions at once, in the plain-text layout test banks use. Nothing is inserted unless every question reads.
  const hidePasteMessages=()=>{
    $('guide-paste-problem').hidden=true;$('guide-paste-problem').textContent='';
    $('guide-paste-ok').hidden=true;$('guide-paste-ok').textContent='';
    $('guide-paste-list').hidden=true;$('guide-paste-list').replaceChildren();
  };
  const pasteProblem=text=>{$('guide-paste-problem').textContent=text;$('guide-paste-problem').hidden=false;};
  async function pasteQuestions(){
    hidePasteMessages();
    const text=$('guide-paste-text').value;
    if(!text.trim()){pasteProblem('Paste the questions first.');return;}
    const r=readPastedQuestions(text);
    if(!r.found&&!r.problems.length){pasteProblem('No questions were found. Each question starts with its number, like 1) or 1. , and its options are lettered, like a. or a).');return;}
    if(r.found>MAX_PASTE){pasteProblem(`${r.found} questions were found and up to ${MAX_PASTE} can be added at a time. Paste fewer, then paste the rest.`);return;}
    if(r.problems.length){
      pasteProblem(`Nothing was added, because ${plural(r.problems.length,'problem')} must be fixed first. What you pasted is still here.`);
      const shown=r.problems.slice(0,20);
      $('guide-paste-list').replaceChildren(...shown.map(p=>{
        const li=doc.createElement('li');
        li.textContent=`${p.number===null?'':`Question ${p.number}, `}line ${p.line}: ${p.message}`;
        return li;
      }),...(r.problems.length>shown.length?(()=>{const li=doc.createElement('li');li.textContent=`and ${r.problems.length-shown.length} more.`;return [li];})():[]));
      $('guide-paste-list').hidden=false;
      return;
    }
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex();
    if(!itemTarget(list,here)){pasteProblem('Items go inside an assessment section, and your cursor is not in one. What you pasted is still here. Start a section, or click inside one in your document, then choose Add the pasted questions again.');return;}
    const {anchor,number}=itemPlacement(list,here);
    await adapter.insert(ctx=>r.items.map((it,i)=>itemBlock({...it,number:number+i},ctx)).join(''),{where:'after',after:anchor,select:'last'});
    $('guide-paste-text').value='';
    await keepOpen();
    await load();
    const kinds=andList([...r.counts].map(([id,n])=>`${n} ${itemTypeById(id).label.toLowerCase()}`));
    const extra=[r.notes.length?r.notes.join(' '):'',r.ignored?`${plural(r.ignored,'line')} (titles, feedback and publisher notes) ${r.ignored===1?'was':'were'} not used.`:''].filter(Boolean).join(' ');
    const msg=`${plural(r.items.length,'question')} added, numbered ${number} to ${number+r.items.length-1}: ${kinds}.${extra?` ${extra}`:''}${tidy()} ${UNDO}`;
    $('guide-paste-ok').textContent=msg;$('guide-paste-ok').hidden=false;
    status(msg);
  }
  // ---- importing questions from an OpenAssess bank file. The questions that read are added (up to 50 at a time, from where the
  // author says); the ones that cannot be carried across are listed with the reason, since they cannot be fixed here.
  let bank=null;
  const hideBankMessages=()=>{$('guide-bank-problem').hidden=true;$('guide-bank-problem').textContent='';$('guide-bank-ok').hidden=true;$('guide-bank-ok').textContent='';};
  const bankProblem=text=>{$('guide-bank-problem').textContent=text;$('guide-bank-problem').hidden=false;};
  function bankSelection(){
    if(!bank)return {readable:[],skipped:[]};
    const folder=bank.folders.find(f=>f.id===$('guide-bank-folder').value);
    const list=folder?folder.ids.map(id=>bank.all.find(q=>q.id===id)).filter(Boolean):bank.all;
    return {readable:list.filter(q=>q.item),skipped:list.filter(q=>q.problem)};
  }
  function drawBankWindow(){
    const {readable}=bankSelection(),raw=$('guide-bank-start').value.trim()||'1',from=/^\d{1,5}$/.test(raw)?Number(raw):0;
    $('guide-bank-window').textContent=!readable.length?'':from<1||from>readable.length?`Start at a number from 1 to ${readable.length}.`:`Next to add: questions ${from} to ${Math.min(from+MAX_BANK_ADD-1,readable.length)} of ${readable.length}.`;
  }
  function drawBank(){
    $('guide-bank-box').hidden=!bank;
    if(!bank)return;
    const {readable,skipped}=bankSelection();
    const kinds=new Map();
    for(const q of readable)kinds.set(q.item.type,(kinds.get(q.item.type)||0)+1);
    $('guide-bank-summary').textContent=`${bank.title}: ${plural(readable.length+skipped.length,'question')} here. ${plural(readable.length,'question')} can be added${kinds.size?` (${andList([...kinds].map(([id,n])=>`${n} ${itemTypeById(id).label.toLowerCase()}`))})`:''}${skipped.length?`; ${skipped.length} cannot, and ${skipped.length===1?'is':'are'} listed below`:''}.`;
    $('guide-bank-skipped-box').hidden=!skipped.length;
    const shown=skipped.slice(0,30).map(q=>{const li=doc.createElement('li');li.textContent=`${q.title}: ${q.problem.message}`;return li;});
    if(skipped.length>shown.length){const li=doc.createElement('li');li.textContent=`and ${skipped.length-shown.length} more.`;shown.push(li);}
    $('guide-bank-skipped').replaceChildren(...shown);
    drawBankWindow();
  }
  async function chooseBank(){
    hideBankMessages();bank=null;
    const file=$('guide-bank-file').files?.[0];
    if(!file){drawBank();return;}
    try{
      const ZipClass=Zip||globalThis.JSZip;
      const read=await readBankFile(await file.arrayBuffer(),{Zip:ZipClass});
      bank=read;
      const options=[{value:'',text:`All questions (${read.all.length})`},...read.folders.map(f=>({value:f.id,text:`${'  '.repeat(f.depth)}${f.name} (${f.ids.length})`}))];
      $('guide-bank-folder').replaceChildren(...options.map(o=>{const op=doc.createElement('option');op.value=o.value;op.textContent=o.text;return op;}));
      $('guide-bank-folder').value='';$('guide-bank-start').value='1';
    }catch(e){bankProblem((e&&e.message)||String(e));}
    drawBank();
  }
  async function addFromBank(){
    hideBankMessages();
    if(!bank){bankProblem('Choose a bank file first.');return;}
    const {readable,skipped}=bankSelection();
    if(!readable.length){bankProblem('There are no questions here that can be added.');return;}
    const raw=$('guide-bank-start').value.trim()||'1';
    if(!/^\d{1,5}$/.test(raw)||Number(raw)<1||Number(raw)>readable.length){bankProblem(`Start at a number from 1 to ${readable.length}.`);return;}
    const from=Number(raw)-1,chosen=readable.slice(from,from+MAX_BANK_ADD);
    const list=await adapter.paragraphs(),here=await adapter.cursorIndex();
    if(!itemTarget(list,here)){bankProblem('Items go inside an assessment section, and your cursor is not in one. Start a section, or click inside one in your document, then choose Add questions from the bank again.');return;}
    const {anchor,number}=itemPlacement(list,here);
    await adapter.insert(ctx=>chosen.map((q,i)=>itemBlock({...q.item,number:number+i},ctx)).join(''),{where:'after',after:anchor,select:'last'});
    const next=from+chosen.length+1;
    $('guide-bank-start').value=String(Math.min(next,readable.length));
    await keepOpen();
    await load();
    drawBank();
    const left=readable.length-from-chosen.length;
    const notes=chosen.flatMap(q=>q.notes||[]);
    const msg=`${plural(chosen.length,'question')} added from the bank, numbered ${number} to ${number+chosen.length-1} (questions ${from+1} to ${from+chosen.length} of ${readable.length} that can be added).${left?` ${left} more can still be added: choose the button again.`:' That is all of them.'}${notes.length?` ${notes.join(' ')}`:''}${skipped.length?` ${plural(skipped.length,'question')} could not be added and ${skipped.length===1?'is':'are'} listed above.`:''}${tidy()} ${UNDO}`;
    $('guide-bank-ok').textContent=msg;$('guide-bank-ok').hidden=false;
    status(msg);
  }
  function startEditItem(){
    const here=describeAt(paras,at);
    if(!here||here.type!=='item')return;
    const it=here.item;
    editingItem=it;
    selectedType=it.type;
    chooseType(it.type,{keepLines:true});
    $('guide-item-stem').value=it.stem;
    OPTION_NUMBERS.forEach((n,i)=>{$(`guide-item-opt-${n}`).value=it.options[i]??'';$(`guide-item-correct-${n}`).checked=it.correct.includes(i);});
    $('guide-item-true').checked=it.truth===true;$('guide-item-false').checked=it.truth===false;
    $('guide-item-answer').value=it.answer;$('guide-item-lines').value=it.lines?String(it.lines):'';
    $('guide-item-tolerance').value=it.tolerance??'';$('guide-item-units').value=it.units??'';
    $('guide-item-why').value=it.why;$('guide-item-points').value=it.points;$('guide-item-objective').value=it.objective;
    $('guide-item-pairs').value=(it.pairs||[]).map(p=>`${p.item} | ${p.match}`).join('\n');
    $('guide-item-extras').value=(it.extras||[]).join('\n');
    $('guide-item-statements').value=(it.statements||[]).join('\n');
    setScale(it.scale||'agree');
    $('guide-item-labels').value=it.scale==='custom'?(it.labels||[]).join(', '):'';
    $('guide-item-order-items').value=(it.orderItems||[]).join('\n');
    $('guide-item-criteria').value=(it.criteria||[]).map(c=>`${c.text}${c.points?` | ${c.points}`:''}`).join('\n');
    $('guide-item-rubric-levels').value=(it.levels||[]).map(l=>`${l.text}${l.points?` | ${l.points}`:''}`).join('\n');
    $('guide-item-rubric-rows').value=(it.rubricRows||[]).map(r=>r.join(' | ')).join('\n');
    $('guide-item-rubric-learners').checked=!!it.learnerRubric;
    $('guide-item-sd-pairs').value=(it.sdPairs||[]).map(p=>`${p.item} | ${p.match}`).join('\n');
    $('guide-item-sd-7').checked=it.sdPoints===7;$('guide-item-sd-5').checked=it.sdPoints!==7;
    $('guide-item-actions').value=it.actions?String(it.actions):'3';
    $('guide-item-heading').textContent='Edit an item';
    $('guide-item-go').textContent='Update item';
    $('guide-item-cancel').hidden=false;
    $('guide-item-note').textContent=`You are editing ${it.label}. Change it, then choose Update item.`;
    drawAssessSelected();
    focusItemStem();
  }
  async function updateItem(){
    const it=editingItem,item=readItemForm();
    // Find the item again in case the document moved on since it was chosen.
    const list=await adapter.paragraphs();
    const hits=list.map((p,i)=>kindOf(p)==='question'&&clean(p.text)===it.firstText?i:-1).filter(i=>i>=0).sort((x,y)=>Math.abs(x-it.span.start)-Math.abs(y-it.span.start));
    const span=hits.length?itemSpan(list,hits[0]):null;
    if(!span)throw new Error('That item has changed since you chose it. Select it in your document again, then choose Edit this item.');
    await adapter.insert(ctx=>itemBlock({...item,number:it.number??1},ctx),{where:'after',after:span.end,select:'last',remove:{start:span.start,end:span.end,guard:it.firstText}});
    const label=itemTypeById(item.type).label;
    clearItemForm();
    await load();
    status(`${label} item updated.${tidy()} ${UNDO}`);
    focusItemStem();
  }

  // ---- the learner workbook: a separate file made from what is for learners
  const andList=items=>items.length<2?items.join(''):`${items.slice(0,-1).join(', ')} and ${items.at(-1)}`;
  function leftText(s){
    const l=s.left,parts=[];
    if(l.facilitatorBlocks)parts.push(plural(l.facilitatorBlocks,'facilitator-only block'));
    if(l.notes)parts.push(`${plural(l.notes,'set')} of facilitator notes`);
    if(l.keys)parts.push(plural(l.keys,'answer key'));
    if(l.timeBlocks)parts.push(plural(l.timeBlocks,'time block'));
    if(l.front)parts.push('the contents and Getting Started pages');
    if(l.overview)parts.push(plural(l.overview,'module overview'));
    if(l.own)parts.push(`${plural(l.own,'paragraph')} of your own text`);
    return parts.length?`Left out for learners: ${andList(parts)}.`:'Nothing is left out: every part of this guide is for learners.';
  }
  function drawWorkbook(){
    const s=summarize(paras);
    workbookReady=s.kept>0;
    $('guide-workbook-summary').textContent=s.kept?`${andList([plural(s.modules,'module'),plural(s.lessons,'lesson'),`${plural(s.blocks,'block')} for learners`,plural(s.items,'assessment item')])}.`
      :'Nothing in this guide is for learners yet. Add a block for Everyone, or an assessment, then come back.';
    $('guide-workbook-left').textContent=leftText(s);
    $('guide-workbook-added').textContent=`The workbook adds a title page with a Name line, a contents table${s.objectives?` and a Learning objectives page from ${plural(s.objectives,'Objective block')}`:'. There are no Objective blocks, so no Learning objectives page'}.`;
    const own=s.leftOut.slice(0,40);
    $('guide-workbook-leftout-box').hidden=!own.length;
    $('guide-workbook-leftout').replaceChildren(...own.map(o=>{
      const li=doc.createElement('li'),b=doc.createElement('button');
      li.textContent=`${snippet(o.text)} `;b.type='button';b.className='guide-link';b.textContent='Go there';
      b.onclick=()=>act(async()=>{await adapter.go(o.index,o.text);await load();});
      li.append(b);return li;
    }));
    if(s.leftOut.length>own.length){const li=doc.createElement('li');li.textContent=`and ${s.leftOut.length-own.length} more.`;$('guide-workbook-leftout').append(li);}
    syncButtons();
  }
  const syncWorkbookMessages=()=>{$('guide-workbook-msg-box').hidden=$('guide-workbook-problem').hidden&&$('guide-workbook-ok').hidden;};
  function hideWorkbookMessages(){$('guide-workbook-problem').hidden=true;$('guide-workbook-ok').hidden=true;syncWorkbookMessages();}
  function workbookMessage(kind,text){
    $('guide-status').textContent='';$('guide-error').hidden=true;
    $('guide-workbook-msg-box').hidden=false;
    $('guide-workbook-problem').hidden=kind!=='problem';$('guide-workbook-ok').hidden=kind!=='ok';
    $(kind==='problem'?'guide-workbook-problem':'guide-workbook-ok').textContent=text;
    $('guide-workbook-msg-box').scrollIntoView?.({block:'nearest'});
  }
  function showWorkbookChecks(structure,print){
    const list=(id,items,empty)=>{
      const ul=$(id);
      if(!items.length){const li=doc.createElement('li');li.textContent=empty;ul.replaceChildren(li);return;}
      ul.replaceChildren(...items.slice(0,40).map(m=>{const li=doc.createElement('li');li.textContent=m;return li;}));
    };
    list('guide-workbook-fix',structure.fix.map(f=>f.message),'Nothing to fix.');
    list('guide-workbook-look-at',structure.look.map(f=>f.message),'Nothing to look at.');
    list('guide-workbook-print',print.map(f=>`${f.level==='warn'?'Look at this: ':''}${f.message}`),'');
    list('guide-workbook-human',WORKBOOK_HUMAN_CHECKS,'');
    $('guide-workbook-checks-box').hidden=false;
  }
  // The look and the brand: what the author chose is read from the screen, and remembered in this browser (a convenience only;
  // the pane keeps no record of the guide).
  const LOOK_KEY='openstage-guide.workbook-look';
  const lookInputs=new Map();
  const chosenLook=()=>WORKBOOK_LOOKS.find(l=>lookInputs.get(l.id)?.checked)?.id||'clearline';
  function buildLooks(){
    $('guide-wb-looks').replaceChildren(...WORKBOOK_LOOKS.map(l=>{
      const label=doc.createElement('label'),input=doc.createElement('input'),name=doc.createElement('span');
      label.className='guide-cue';
      input.type='radio';input.name='guide-wb-look';input.id=`guide-wb-look-${l.id}`;input.value=l.id;input.checked=l.id==='clearline';
      input.onchange=saveLook;lookInputs.set(l.id,input);
      name.textContent=`${l.name}: ${l.note}`;
      label.append(input,name);
      return label;
    }));
  }
  function workbookChoices(){
    const brand={},org=String($('guide-wb-org').value??'').trim();
    if(org)brand.organizationName=org;
    let accent=String($('guide-wb-accent').value??'').trim();
    if(accent){
      if(!accent.startsWith('#'))accent=`#${accent}`;
      if(!/^#[0-9a-fA-F]{6}$/.test(accent))throw new Error('Type the accent color as a # and six letters or digits, like #164B6B, or leave it empty.');
      brand.accent=accent;
    }
    if($('guide-wb-print-gray').checked)brand.printMode='grayscale';
    if($('guide-wb-page-a4').checked)brand.pageSize='A4';
    const raw=String($('guide-wb-lines').value??'').trim();
    if(raw&&!/^\d$/.test(raw))throw new Error('Lines to write on must be a whole number from 0 to 6.');
    return {familyId:chosenLook(),brand,lines:raw===''?3:Number(raw)};
  }
  function saveLook(){
    try{
      storage?.setItem(LOOK_KEY,JSON.stringify({look:chosenLook(),org:$('guide-wb-org').value,accent:$('guide-wb-accent').value,
        gray:!!$('guide-wb-print-gray').checked,a4:!!$('guide-wb-page-a4').checked,lines:$('guide-wb-lines').value}));
    }catch{/* nothing remembered is no harm */}
  }
  function restoreLook(){
    let saved=null;
    try{saved=JSON.parse(storage?.getItem(LOOK_KEY)||'null');}catch{saved=null;}
    if(!saved||typeof saved!=='object')return;
    const known=WORKBOOK_LOOKS.some(l=>l.id===saved.look);
    for(const l of WORKBOOK_LOOKS)lookInputs.get(l.id).checked=known?l.id===saved.look:l.id==='clearline';
    $('guide-wb-org').value=String(saved.org??'');$('guide-wb-accent').value=String(saved.accent??'');
    $('guide-wb-print-gray').checked=!!saved.gray;$('guide-wb-print-color').checked=!saved.gray;
    $('guide-wb-page-a4').checked=!!saved.a4;$('guide-wb-page-letter').checked=!saved.a4;
    $('guide-wb-lines').value=String(saved.lines??'3');
  }
  async function makeWorkbook(){
    hideWorkbookMessages();
    try{
      const choices=workbookChoices();
      saveLook();
      const parser=xml.DOMParserClass||globalThis.DOMParser,serializer=xml.SerializerClass||globalThis.XMLSerializer,ZipClass=Zip||globalThis.JSZip;
      if(!ZipClass)throw new Error('The pane could not load what it needs to make a workbook. Close it, open it again from the ribbon, then try again.');
      const source=readSource(await adapter.sourcePackage(),{DOMParserClass:parser});
      const built=buildWorkbook(source,choices,{DOMParserClass:parser,SerializerClass:serializer});
      // The file is checked before it is offered. If anything meant for the facilitator is in it, it is not offered.
      const leak=leakCheck({files:built.files,facilitatorTexts:built.facilitatorTexts,allowedTexts:built.allowedTexts,expectedTitle:built.workbookTitle},{DOMParserClass:parser});
      if(!leak.passed){workbookMessage('problem',`The workbook was not made, because its check found a problem. ${leak.summary}`);return;}
      const base64=await zipWorkbook(built.files,ZipClass);
      // The structure rules and the print check do not stop the workbook: what they find can be fixed in the copy.
      const structure=workbookStructure(built.files,{DOMParserClass:parser}),print=printCheck(built.effective,{lines:built.lines});
      showWorkbookChecks(structure,print);
      const checked=`Leak check passed: nothing meant for the facilitator is in the file. ${leftText(built.summary)} Structure check: ${structure.fix.length} to fix, ${structure.look.length} to look at.`;
      try{
        await adapter.openDocument(base64,built.fileName);
        workbookMessage('ok',`The workbook is open in a new Word window. Save it as "${built.fileName}", or any name you like. ${checked}`);
      }catch{
        adapter.offerDownload(base64,built.fileName);
        workbookMessage('ok',`Word could not open it for you, so it was downloaded as "${built.fileName}". Open it from your Downloads folder. ${checked}`);
      }
    }catch(e){workbookMessage('problem',(e&&e.message)||String(e));}
  }

  // ---- wiring
  $('guide-start-form').onsubmit=e=>{e?.preventDefault?.();return act(startGuide);};
  $('guide-sample').onclick=()=>act(trySample);
  $('guide-help-sample').onclick=()=>act(trySample);
  $('guide-add-module').onclick=()=>act(addModule);
  $('guide-add-lesson').onclick=()=>act(addLesson);
  $('guide-agenda-start').onchange=()=>drawAgenda();
  $('guide-agenda-refresh').onclick=()=>act(async()=>{await load();drawAgenda();});
  $('guide-agenda-download').onclick=()=>act(makeAgenda);
  $('guide-bk-save').onclick=()=>act(saveBackup);
  $('guide-bk-file').onchange=()=>act(chooseBackup);
  $('guide-bk-open').onclick=()=>act(openBackup);
  $('guide-align-refresh').onclick=()=>act(async()=>{await load();drawAlignment();});
  $('guide-ver-select').onchange=()=>{confirmingDelete=false;chosenVersion=$('guide-ver-select').value;drawVersions();};
  $('guide-ver-new').onclick=()=>act(newVersionNow);
  $('guide-ver-delete').onclick=()=>act(deleteVersionNow);
  $('guide-ver-name').onchange=()=>act(renameVersion);
  $('guide-ver-target').onchange=()=>act(retargetVersion);
  $('guide-ver-start').onchange=()=>act(restartVersion);
  $('guide-ver-make').onclick=()=>act(makeVersionAgenda);
  try{const saved=storage?.getItem(AGENDA_START_KEY);if(START_TIME.test(saved||''))$('guide-agenda-start').value=saved;}catch{/* the usual 09:00 */}
  $('guide-lib-find').oninput=drawLibrary;
  $('guide-lib-save').onclick=()=>act(saveToLibrary);
  loadMine();drawLibrary();
  // The buttons are disabled while an action runs, so focus returns to the one pressed only once it has finished.
  $('guide-add-standard').onclick=async()=>{await act(addStandardPages);focus('guide-add-standard');};
  $('guide-add-appendix').onclick=async()=>{await act(addAppendix);focus('guide-add-appendix');};
  $('guide-move-up').onclick=async()=>{await act(()=>moveUnit(-1));focus('guide-move-up');};
  $('guide-move-down').onclick=async()=>{await act(()=>moveUnit(1));focus('guide-move-down');};
  $('guide-delete').onclick=async()=>{await act(askDelete);focus(pendingDelete?'guide-delete-no':'guide-delete');};
  $('guide-delete-yes').onclick=async()=>{await act(confirmDelete);focus('guide-build-heading');};
  $('guide-delete-no').onclick=async()=>{cancelDelete();status('Nothing was deleted.');focus('guide-delete');};
  $('guide-block-form').onsubmit=e=>{e?.preventDefault?.();return act(addBlock);};
  $('guide-block-text').onkeydown=$('guide-block-notes').onkeydown=e=>{if(e.ctrlKey&&e.key==='Enter'){e.preventDefault?.();return act(addBlock);}};
  for(const p of ACTIVITY_PARTS)$(`guide-act-${p.id}`).onkeydown=e=>{if(e.ctrlKey&&e.key==='Enter'){e.preventDefault?.();return act(addBlock);}};
  $('guide-block-cancel').onclick=()=>act(async()=>{endEdit();status('Edit cancelled. Nothing was changed.');});
  $('guide-edit-selected').onclick=()=>act(async()=>startEdit());
  $('guide-leave-block').onclick=()=>act(writeBelow);
  $('guide-edit-form').onsubmit=e=>{e?.preventDefault?.();return act(saveEdit);};
  $('guide-edit-cancel').onclick=()=>act(async()=>{endEdit();status('Edit cancelled. Nothing was changed.');});
  $('guide-retry').onclick=()=>act(async()=>{await load();status('');});
  $('guide-update').onclick=()=>act(updateGuide);
  // The buttons are disabled while an action runs, so focus returns to the one pressed only once it has finished.
  $('guide-build-glance').onclick=async()=>{await act(buildGlancePage);focus('guide-build-glance');};
  $('guide-build-cues').onclick=async()=>{await act(buildCuesPage);focus('guide-build-cues');};
  $('guide-build-overview').onclick=async()=>{await act(buildOverviewPage);focus('guide-build-overview');};
  $('guide-open-review').onclick=()=>act(review);
  $('guide-review-run').onclick=()=>act(review);
  $('guide-review-back').onclick=()=>act(async()=>{show('build',{moveFocus:true});await load();});
  $('guide-review-pdf').onclick=()=>act(makePdf);
  hidePdfMessages();
  $('guide-refresh').onclick=()=>act(async()=>{await load();status('Read the document again.');});
  $('guide-pic-width').replaceChildren(...PICTURE_WIDTHS.map(w=>{const o=doc.createElement('option');o.value=w.id;o.textContent=w.label;return o;}));
  $('guide-pic-width').value=PICTURE_WIDTHS[0].id;
  $('guide-pic-choose').onclick=()=>act(choosePicture,{slow:true});
  $('guide-slides-choose').onclick=()=>act(chooseSlides,{slow:true});
  $('guide-slides-all').onclick=()=>{if(deck){deck.chosen=new Set(deck.slides.map(s=>s.order));drawDeck();}};
  $('guide-slides-none').onclick=()=>{if(deck){deck.chosen=new Set();drawDeck();}};
  $('guide-slides-go').onclick=()=>act(insertSlides);
  $('guide-upd-choose').onclick=()=>act(chooseUpdate,{slow:true});
  $('guide-upd-which').onchange=()=>act(async()=>{if(!upd)return;upd.which=$('guide-upd-which').value;await planUpdate();});
  $('guide-upd-all').onclick=()=>tickAll(true);
  $('guide-upd-none').onclick=()=>tickAll(false);
  $('guide-upd-go').onclick=()=>act(applyUpdate,{slow:true});
  buildGuideLooks();drawLook();buildKinds();
  $('guide-look-apply').onclick=()=>act(applyLookNow);
  $('guide-el-add').onclick=async()=>{await act(addScreen);};
  $('guide-el-script').onclick=()=>act(()=>makeScript('docx'));
  $('guide-el-text').onclick=()=>act(()=>makeScript('text'));
  $('guide-el-author').onclick=()=>act(sendToAuthor);
  buildGroups();
  buildCues();
  buildTypes();
  // Switching screens drops the last screen's message, so the new screen opens on its own content.
  $('guide-tab-guide').onclick=()=>act(async()=>{status('');show('build',{moveFocus:true});await load();});
  $('guide-tab-assess').onclick=()=>act(async()=>{status('');show('assess',{moveFocus:true});await load();});
  $('guide-tab-workbook').onclick=()=>act(async()=>{status('');hideWorkbookMessages();show('workbook',{moveFocus:true});await load();});
  $('guide-workbook-go').onclick=()=>act(makeWorkbook);
  buildLooks();restoreLook();
  for(const id of ['guide-wb-org','guide-wb-accent','guide-wb-lines','guide-wb-print-color','guide-wb-print-gray','guide-wb-page-letter','guide-wb-page-a4'])$(id).onchange=saveLook;
  hideWorkbookMessages();
  for(const level of LEVELS)$(`guide-sec-${level.id}`).onclick=()=>act(()=>startSection(level.id));
  // The buttons are disabled while an action runs, so focus goes to Insert item only once it has finished.
  $('guide-item-start-two').onclick=async()=>{await act(()=>startSection('level-two',{fromItem:true}));focus('guide-item-go');};
  hideItemMessages();
  $('guide-item-form').onsubmit=e=>{e?.preventDefault?.();return act(addItem);};
  $('guide-item-stem').onkeydown=$('guide-item-answer').onkeydown=$('guide-item-why').onkeydown=e=>{if(e.ctrlKey&&e.key==='Enter'){e.preventDefault?.();return act(addItem);}};
  $('guide-item-cancel').onclick=()=>act(async()=>{clearItemForm();status('Edit cancelled. Nothing was changed.');});
  $('guide-assess-edit').onclick=()=>act(async()=>startEditItem());
  hidePasteMessages();
  $('guide-paste-form').onsubmit=e=>{e?.preventDefault?.();return act(pasteQuestions);};
  hideBankMessages();
  $('guide-bank-file').onchange=()=>act(chooseBank);
  $('guide-bank-folder').onchange=()=>{$('guide-bank-start').value='1';drawBank();};
  $('guide-bank-start').oninput=drawBankWindow;
  $('guide-bank-form').onsubmit=e=>{e?.preventDefault?.();return act(addFromBank);};

  // For Open Designer and its keyboard shortcut. Reading the document alone leaves Review, Assessments or Workbook on the screen,
  // so this also goes back to the Guide screen (or to Start, when the document has no guide yet).
  async function showGuide(){
    await load();
    if(mode==='review'||mode==='assess'||mode==='workbook'){status('');show(plan.guide.isGuide?'build':'start',{moveFocus:true});}
  }

  const api={
    refresh:()=>act(load),
    showGuide:()=>act(showGuide),
    // The ribbon buttons and "More tools": open one tool of src/guide-tools.mjs. useTool only names the tool the pane starts on,
    // before the document has been read.
    // While the pane is busy (a few seconds on a long guide) the tool asked for waits its turn instead of being dropped. A file window
    // that may stay open for minutes is the exception: a tool chosen while it is open is not carried out long after.
    openTool:id=>{
      if(busy){if(!busyIsSlow)queuedTool=id;return Promise.resolve();}
      return act(()=>openToolScreen(id));
    },
    useTool:id=>{if(toolFor(id)){setTool(id);paintTool();}},
    // For when the cursor moves or the pane regains focus: read the document without disabling anything under the author's mouse.
    quietRefresh:async()=>{if(busy||!connected)return;try{await load();}catch{/* the next action reports a real problem */}},
    showReview:()=>act(review),
    // For the keyboard shortcuts that move about the guide: a line in the status area, whether or not the pane is showing.
    say:text=>status(text),
    // Word tells the pane when the cursor moves or the text changes (typing, Undo, Redo), so what it shows keeps up.
    // Called once Office is ready, because Word's events cannot be hooked before that.
    watch:()=>adapter.onSelectionChanged?.(()=>{if(waiting)return;waiting=true;return soon(()=>{waiting=false;return api.quietRefresh();});}),
    state:()=>({mode,plan,busy,editing,tool})
  };
  return api;
}
