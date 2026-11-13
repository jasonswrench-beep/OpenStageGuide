import {mountWorkbookDesignFiles} from './workbook-design-files-panel.mjs';
import {explainWordError} from './word-errors.mjs';
import {validateWorkbookDesignDraft,createWorkbookDesignFile,serializeWorkbookDesignFile} from './workbook-design-file.mjs';
import {WORKBOOK_COVER_FIELDS,WORKBOOK_COPYRIGHT_FIELDS} from './workbook-public-frontmatter.mjs';
import {WORKBOOK_RESEARCH_FIELDS} from './workbook-research.mjs';
import {workbookLogoRequest} from './workbook-logo.mjs';
import {mountWorkbookLogoPanel} from './workbook-logo-panel.mjs';
import {workbookResearchRequest} from './workbook-research.mjs';
import {mountWorkbookResearchPanel} from './workbook-research-panel.mjs';
import {workbookContentsRequest} from './workbook-contents.mjs';
import {workbookFrontMatterRequest} from './workbook-public-frontmatter.mjs';
import {mountWorkbookFrontMatterPanel} from './workbook-frontmatter-panel.mjs';
import {materialsPreview,materialsParagraphs,materialsBlob,materialsName} from './materials-output.mjs';
import {facilitatorPreview,facilitatorParagraphs,facilitatorBlob,facilitatorName} from './facilitator-output.mjs';
import {planFacilitatorPreview,planFacilitatorBlob,planFacilitatorName} from './selected-plan-facilitator.mjs';
import {ACTIVITY_DRAFT_FIELDS,rawControl,rawInvalid,sameRaw,wouldReplaceRaw} from './activity-draft.mjs';
import {ASSESSMENT_STARTERS,assessmentStarter} from './assessment-starters.mjs';
import {mountAssessmentPanel} from './assessment-panel.mjs';
import {COVERAGE_ROLES,objectiveCoverage,coverageReport,editObjectiveCoverage} from './objective-coverage.mjs';
import { course,wrenchCourse,addModule,addActivity,editActivitySettings,courseDetails,editCourseDetails,COURSE_DETAIL_FIELDS,ICONS,total,participantView } from './model.mjs';
import {objectiveText,objectiveLabel,unsupportedObjective} from './objective-text.mjs';
import {worksheetBlob,worksheetName,workbookPreview,WORKBOOK_SCOPE} from './worksheet.mjs';
import {agendaPreview,agendaView,agendaBlob,agendaName,AGENDA_SCOPE} from './agenda.mjs';
import {deliveryPlans,editDeliveryPlan,selectedPlanObjectiveReport} from './delivery-plans.mjs';
import {planAgendaPreview,planAgendaBlob,planAgendaName} from './plan-agenda.mjs';
import {planWorkbookPreview,planWorkbookBlob,planWorkbookName,planMaterialsPreview,planMaterialsBlob,planMaterialsName} from './selected-plan-outputs.mjs';
import {WORKBOOK_FAMILIES,resolveWorkbookBrand} from './workbook-presentation.mjs';
import {WORKBOOK_DEFAULT_BASE_LANGUAGE,validateWorkbookFamilyBrand} from './workbook-family-output.mjs';
import {FAMILY_BRAND_FIELDS,familyBrandRequest,prepareFamilyWorkbook,familyReviewText,familyWorkbookOutputBlob,familyWorkbookName} from './workbook-family-flow.mjs';
import {readBrandPresetCatalog,writeBrandPresetCatalog,nextBrandPresetCatalog} from './workbook-brand-presets.mjs';
import {deliveryPackPreview,buildDeliveryPack,deliveryPackName,PACK_PRIVATE} from './delivery-pack.mjs';
import * as word from './word.mjs';
import {registerRibbonActions,registerNavigationActions} from './ribbon.mjs';
import {mountGuidePane} from './guide-pane.mjs';
import {toolFor} from './guide-tools.mjs';
import {mountToolStrip} from './tool-strip.mjs';
import {createWordAdapter,createDemoAdapter} from './guide-word.mjs';
import {createProjectAdapter} from './guide-project-adapter.mjs';
import {webProjectStore} from './guide-project-web.mjs';
import {installFileInputBridge} from './file-input-bridge.mjs';
import {buildAgenda,reviewCourse} from './planning.mjs';
import {createCheckpoint,readCheckpoint,checkpointName,MAX_PACKAGE_BYTES} from './checkpoint.mjs';
import {emptyLibrary,saveLibraryCopy,mergeLibrary,readLibrary,serializeLibrary,libraryFields,libraryPlaceholders,libraryName,renameLibraryCopy,removeLibraryCopy,undoLibraryRemoval,editLibraryContent,undoLibraryContent,MAX_LIBRARY_BYTES} from './library.mjs';
import {readPptx,pptxDraft,editSlideDraft,MAX_PPTX_BYTES} from './pptx.mjs';
import {restoreReview,serializeReview,MAX_REVIEW_BYTES} from './pptx-review.mjs';
import {searchOutline} from './outline.mjs';
import {comparePptx} from './pptx-compare.mjs';
import {moduleAsset,moduleIdentity,inspectBuild} from './diagnostics.mjs';
import {renderBuildStatus} from './build-status.mjs';
import {activityMaterials,editActivityMaterials} from './model.mjs';
import {normalizeMaterials} from './materials.mjs';
import {planMaterials} from './planning.mjs';
import {serializeMaterialsDraft,restoreMaterialsDraft,MAX_MATERIAL_DRAFT_BYTES} from './materials-draft.mjs';
const $=id=>document.getElementById(id);
// ?project=1 is the program-of-its-own spike: the same pane and the same in-memory document as the preview, but the guide is read from
// and saved to one project file by the program's own server (scripts/project-server.mjs). It is a preview in every other way.
const projectMode=new URLSearchParams(location.search).get('project')==='1';
const demo=projectMode||new URLSearchParams(location.search).get('demo')==='1';
// &program=1: the pane is the tools of the program (program/), shown in a frame of its page. The program's own look replaces the add-in's
// header and footer, and the program is told when the tools are ready and how the guide's saving stands (it draws the guide from the project).
const programMode=projectMode&&new URLSearchParams(location.search).get('program')==='1';
if(programMode){
  document.documentElement.classList.add('in-program');
  const programLook=document.createElement('link');programLook.rel='stylesheet';programLook.href='/program/pane.css';document.head.append(programLook);
}
const toProgram=message=>{if(programMode&&globalThis.parent&&globalThis.parent!==globalThis)globalThis.parent.postMessage(message,location.origin);};
let project=null,busy=false,ready=false,blocked=false,previewRevision=null;
let assessmentPanel=null;
let verifiedBackup=null,backupSelection=0;
let activityDraft=null;
let courseDraft=null;
let agendaDraft=null;
let planDraft=null,planGroups=[],planGeneration=0;
let planReview=null;
let workbookDraft=null;
let familyDraft=null,familyAlternative=null,familyGeneration=0,familyConsentEpoch=0,familyFrontMatterPanel=null,familyResearchPanel=null,familyLogoPanel=null,familyDesignFilesPanel=null;
let presetStore=null,presetState=null,presetPending=null,presetNeedsReload=false;
let facilitatorDraft=null;
let packDraft=null;
let library=emptyLibrary(),libraryOrigin=null,librarySelection=0;
let libraryEdit=null,libraryRemoved=null,libraryPreferred=null;
let libraryContentDraft=null,libraryContentUndo=null;
let libraryMaterialRows=null,libraryMaterialEditing=null,libraryMaterialDirty=false;
const LIBRARY_MATERIAL_FIELDS=['item','quantity','unit','use','notes'];
const LIBRARY_CONTENT_FIELDS=['title','icon','minutes','facilitator','participant','answers','lines'];
let pptSource=null,pptReview=null,pptSelection=0,pptCurrent='',pptRestoreSelection=0,pptAlias='';
let pptComparison=null,pptUpdateSelection=0;
const loadedBuild=moduleIdentity(import.meta.url);
let diagnosticReport=null;
let materialsDraft=null,materialsExportDraft=null;
let copiedMaterials=null;
let starterOrigin=null,starterReview=null,activityGeneration=0,activityReplacement=null;
const copiedMaterialNodes=new Map();
const ACTIVITY_DEFAULTS={title:'',icon:'Learning activity',minutes:'10',facilitator:'',participant:'',answers:'',lines:'5'};
let coverageDraft=null,coverageLinks=[],coverageEditing='',coverageReview=null,coverageGeneration=0;
let gapDraft=null,gapGeneration=0,gapTargets='';
let materialsRestoreSelection=0;
function message(text){$('status').textContent=text;}
function invalidateMaterialsChecklist(){materialsExportDraft=null;$('materials-checklist-reviewed').checked=false;$('materials-checklist-output').textContent='Refresh and review the checklist from saved materials.';}
function invalidateCoverageReport(){$('coverage-review-summary').textContent='Refresh Review Guide to inspect the current authored objective mapping.';$('coverage-review-rows').replaceChildren();}
function clearAgendaCoverage(){$('agenda-coverage').replaceChildren();$('agenda-coverage').hidden=true;$('agenda-coverage-heading').hidden=true;}
function invalidateOutputs(){
  invalidateCoverageReport();
  invalidateFacilitator();
  invalidatePack();
  invalidateMaterialsChecklist();
  agendaDraft=null;
  $('agenda-reviewed').checked=false;
  workbookDraft=null;
  $('workbook-plan-reviewed').checked=false;
  previewRevision=null;$('workbook-preview').textContent='Refresh the preview after the course is checked.';
  invalidateFamily();
  $('agenda-body').replaceChildren();$('agenda-table').hidden=true;clearAgendaCoverage();$('agenda-summary').textContent='Refresh after the course is checked.';
  $('review-list').replaceChildren();$('review-summary').textContent='Refresh after the course is checked.';
}
function error(e){
  invalidateOutputs();
  if(!demo&&(e.code==='OSG_INTEGRITY'||word.isBlocked()))blocked=true;
  const shown=explainWordError(e); // Word's own codes ("InvalidArgument") become a sentence that says what to do
  $('error').textContent=shown.message||String(shown);$('error').hidden=false;$('error').scrollIntoView({block:'nearest'});
  render(); // Ribbon errors can arrive outside run(), so update action gating here too.
}
function clearError(){$('error').hidden=true;$('error').textContent='';}
function view(name){if(!['guide','course','module','activity','edit-activity','output','outline','agenda','pack','review','checkpoints','library','slides','materials','assessments'].includes(name))name='guide';{const summary=document.querySelector('.course-summary');if(summary)summary.hidden=name==='guide';}document.querySelectorAll('main>section').forEach(x=>x.hidden=x.id!=='view-'+name);document.querySelectorAll('[data-view]').forEach(b=>b.setAttribute('aria-selected',String(b.dataset.view===name)));document.querySelector('main').scrollTop=0;assessmentPanel?.renderBusy();familyDesignFilesPanel?.renderBusy();}
// SWITCHED OFF 2026-11-08 at the author's word ("Legacy tools" made it sound like a revised product, and he does not need them): the section
// at the bottom of the pane is hidden in taskpane.html, so none of the older panels can be reached. Their code and tests stay until they are
// deleted on purpose. Three panels under "Legacy tools" still worked from the older course record (made with the older Course tool, which lives under
// "Legacy tools" at the bottom of the pane). The ribbon buttons of the same names were rebuilt on 2026-11-07 and read the guide
// itself, so a guide written with the new tools has no use for these. Said plainly, so the author is not sent to a tab that is not there.
const NEEDS_OLDER_COURSE=Object.freeze({'edit-activity':'Learning Alignment',agenda:'Timing & Agenda',pack:'Delivery Versions'});
function needsCourseText(name){
  const tool=NEEDS_OLDER_COURSE[name];
  return tool?`${tool} on the ribbon now works from your guide. This older panel works only from a course made with the older Course tool, which a guide written here does not have.`
    :'There is no course here yet. Open Legacy tools at the bottom of the pane, choose Course, and create it.';
}
async function navigate(name){
  if(busy){message('Please wait for the current Word operation to finish.');return;}
  view(name);
  if(name==='guide'||name==='guide-review'){if(name==='guide-review')await guidePane.showReview();else await guidePane.showGuide();return;}
  if(!project&&!['course','checkpoints','library','slides','materials','assessments'].includes(name)){message(needsCourseText(name));return;}
  if(name==='activity'&&!project.modules.length)message('Add a module on the Module tab first.');
  if(project&&!blocked&&['agenda','review'].includes(name))await run(async()=>{await refresh();if(name==='agenda')showAgenda();else showReview();});
}
// Opens one tool of src/guide-tools.mjs: from a ribbon button, a keyboard shortcut or "More tools". The guide's own tools are the
// guide pane's work; a tool that is not redesigned yet opens the older panel it always had (the plan of 2026-11-06, point 4).
const LEGACY_HEADINGS={library:'library-heading','edit-activity':'edit-activity-heading',agenda:'agenda-heading',pack:'pack-heading',checkpoints:'checkpoints-heading'};
async function openTool(id){
  const tool=toolFor(id);
  if(!tool){await navigate('guide');return;}
  if(busy){message('Please wait for the current Word operation to finish.');return;}
  message(''); // a line left by the tool before (such as "now works from your guide") does not stay over this one
  if(tool.kind==='legacy'){
    toolStrip.show(id);
    await navigate(tool.view);
    const heading=document.getElementById(LEGACY_HEADINGS[tool.view]);
    if(heading){heading.tabIndex=-1;heading.focus?.();}
    return;
  }
  view('guide');
  await guidePane.openTool(id);
}
function render(){
  document.querySelectorAll('[data-acceptance-action]').forEach(b=>b.disabled=busy||b.dataset.acceptanceLocked==='1');
  document.querySelector('main').setAttribute('aria-busy',String(busy));
  document.querySelectorAll('[data-view],form input,form textarea,form select').forEach(x=>x.disabled=busy);
  $('refresh').disabled=busy||!ready||blocked;
  $('recover').disabled=busy||!ready;
  $('backup-save').disabled=busy||!ready||!project||blocked||demo;
  $('backup-file').disabled=busy;
  $('backup-inspect').disabled=busy;
  $('backup-restore').disabled=busy||!verifiedBackup;
  $('show-outline').disabled=busy;
  $('agenda-refresh').disabled=$('review-refresh').disabled=busy||!project||!ready||blocked;
  $('agenda-download').disabled=busy||blocked||!ready||!agendaDraft||!project||agendaDraft.revision!==project.revision||agendaDraft.startInput!==$('agenda-start').value||(agendaDraft.kind==='plan'&&!$('agenda-reviewed').checked);
  $('course-name').textContent=project?.title||'Your next course starts here';$('course-time').textContent=project?`${total(project)} min`:'0 min';
  $('course-form').hidden=!!project;$('existing-course').hidden=!project;
  $('course-details').hidden=!project;
  $('course-details-load').disabled=busy||blocked||!ready||!project;
  $('course-details-form').querySelector('button').disabled=busy||blocked||!ready||!courseDraft||!project;
  for(const key of COURSE_DETAIL_FIELDS)$('course-detail-'+key).disabled=busy||blocked||!ready||!courseDraft||!project;
  $('module-form').querySelector('button').disabled=!project||busy||!ready||blocked;
  $('activity-insert').disabled=!project?.modules.length||busy||!ready||blocked;
  const existing=$('edit-activity-id'),chosen=existing.value;
  existing.replaceChildren();
  project?.modules.forEach((m,i)=>m.activities.forEach((a,j)=>{const o=document.createElement('option');o.value=a.id;o.textContent=`${i+1}.${j+1} ${a.title}`;existing.append(o);}));
  existing.value=project?.modules.some(m=>m.activities.some(a=>a.id===chosen))?chosen:project?.modules.flatMap(m=>m.activities)[0]?.id||'';
  existing.disabled=busy||blocked||!ready||!existing.value;
  $('edit-activity-load').disabled=busy||blocked||!ready||!existing.value;
  $('edit-activity-form').querySelector('button').disabled=busy||blocked||!ready||!activityDraft||activityDraft.id!==existing.value;
  for(const id of ['edit-activity-minutes','edit-activity-lines'])$(id).disabled=busy||blocked||!ready||!activityDraft||activityDraft.id!==existing.value;
  $('course-form').querySelector('button').disabled=busy||!ready||blocked;
  const target=$('activity-module'),selected=target.value;
  target.replaceChildren();
  project?.modules.forEach((m,i)=>{const o=document.createElement('option');o.value=m.id;o.textContent=`${i+1}. ${m.title}`;target.append(o);});
  target.value=project?.modules.some(m=>m.id===selected)?selected:project?.modules.at(-1)?.id||'';
  target.disabled=busy||blocked||!ready||!project?.modules.length;
  const destination=project?.modules.find(m=>m.id===target.value);
  $('activity-target').textContent=destination?`Adding to ${destination.title}. New activities go at the end of this module.`:'Add a module before creating an activity.';
  $('download').disabled=busy||blocked||!ready||!project||!workbookDraft||previewRevision!==project.revision||workbookDraft.kind==='plan'&&!$('workbook-plan-reviewed').checked||workbookDraft.kind!=='plan'&&!!$('workbook-plan').value;
  $('preview').disabled=busy||blocked||!project||!ready;
  $('workbook-plan-review-label').hidden=workbookDraft?.kind!=='plan';
  $('workbook-plan-reviewed').disabled=busy||blocked||!ready||workbookDraft?.kind!=='plan';
  renderFacilitator();
  renderPack();
  renderLibrary();
  renderSlides();
  renderOutline();
  renderBuildStatus($('build-loaded'),$('build-check'),loadedBuild?.buildId,diagnosticReport?.status,busy);
  $('build-download').disabled=busy||!diagnosticReport;
  renderMaterials();
  renderCopiedMaterials();
  renderActivityDraft();
  renderCoverageEditor();
  renderPlanEditor();
  renderOutputPlanSelectors();
  renderFamily();
  assessmentPanel?.renderBusy();familyDesignFilesPanel?.renderBusy();
}
const facilitatorChoice=()=>({answers:!!$('facilitator-answers').checked,participant:!!$('facilitator-participant').checked});
const facilitatorDisplayStructure=()=>JSON.stringify(Array.from($('facilitator-output').children,node=>({tag:node.tagName||'',className:node.className||'',hidden:!!node.hidden,style:node.style?.cssText||'',text:node.textContent})));
function invalidateFacilitator(){facilitatorDraft=null;$('facilitator-reviewed').checked=false;$('facilitator-output').textContent='Refresh and review the private copy before downloading.';}
function renderFacilitator(){
  if(facilitatorDraft&&(!project||facilitatorDraft.snapshot!==JSON.stringify(project)||facilitatorDraft.options!==JSON.stringify(facilitatorChoice())||facilitatorDraft.planId!==$('facilitator-plan').value||facilitatorDraft.text!==$('facilitator-output').textContent||facilitatorDraft.displayStructure!==facilitatorDisplayStructure()))invalidateFacilitator();
  for(const id of ['facilitator-preview','facilitator-answers','facilitator-participant'])$(id).disabled=busy||blocked||!ready||!project;
  $('facilitator-reviewed').disabled=busy||blocked||!ready||!facilitatorDraft;
  $('facilitator-download').disabled=busy||blocked||!ready||!project||!facilitatorDraft||!$('facilitator-reviewed').checked;
}
const packChoice=()=>$('pack-route').value===PACK_PRIVATE?{route:PACK_PRIVATE,start:$('pack-start').value,answers:!!$('pack-answers').checked,participant:!!$('pack-participant').checked}:{route:$('pack-route').value,start:$('pack-start').value};
const packReviewText=preview=>{
 const id=preview.identity,lines=[preview.route===PACK_PRIVATE?'PRIVATE FACILITATOR PACK':'SHAREABLE LEARNER PACK',`Plan: ${id.planName} · Plan ID: ${id.planId}`,`Project ID: ${id.projectId} · Revision: ${id.revision} · Generated: ${id.generation}`,`Selected time: ${id.minutes} minutes · ${id.start}–${id.end}`,`Exact selected scope: ${id.scope.map(group=>`${group.moduleId}: ${group.activityIds.join(', ')}`).join(' | ')}`,`Files: ${preview.members.join(', ')}`];
 if(preview.options)lines.push(`PRIVATE guide options: answers/debrief ${preview.options.answers?'included':'excluded'}; participant instructions in PRIVATE guide ${preview.options.participant?'included':'excluded'}. The participant workbook is always included.`);
 lines.push('Limits and assumptions:',...preview.assumptions.map(text=>`• ${text}`));
 for(const member of preview.contentPreviews)lines.push(`\n===== ${member.name} · complete structured content =====`,member.text);
 return lines.join('\n');
};
const packDisplayStructure=()=>{const root=$('pack-output');return JSON.stringify({root:{tag:root.tagName||'',className:root.className||'',hidden:!!root.hidden,style:root.style?.cssText||'',text:root.textContent},children:Array.from(root.children,node=>({tag:node.tagName||'',className:node.className||'',hidden:!!node.hidden,style:node.style?.cssText||'',text:node.textContent}))});};
function invalidatePack(){packDraft=null;$('pack-reviewed').checked=false;$('pack-private-consent').checked=false;$('pack-output').textContent='Choose a saved plan and refresh the pack review.';}
function packReviewed(draft){
 if(blocked||!ready||!project||packDraft!==draft||!$('pack-reviewed').checked||draft.planId!==$('pack-plan').value||draft.snapshot!==JSON.stringify(project)||draft.choice!==JSON.stringify(packChoice())||draft.text!==$('pack-output').textContent||draft.structure!==packDisplayStructure())return false;
 if(draft.route===PACK_PRIVATE&&!$('pack-private-consent').checked)return false;
 try{return draft.preview===JSON.stringify(deliveryPackPreview(project,draft.planId,{...packChoice(),stamp:draft.stamp}));}catch{return false;}
}
function renderPack(){
 const privateRoute=$('pack-route').value===PACK_PRIVATE;
 $('pack-private-options').hidden=!privateRoute;$('pack-private-consent-label').hidden=!privateRoute;
 if(packDraft&&(!project||packDraft.snapshot!==JSON.stringify(project)||packDraft.planId!==$('pack-plan').value||packDraft.choice!==JSON.stringify(packChoice())||packDraft.text!==$('pack-output').textContent||packDraft.structure!==packDisplayStructure()))invalidatePack();
 for(const id of ['pack-start','pack-route','pack-preview','pack-answers','pack-participant'])$(id).disabled=busy||blocked||!ready||!project;
 $('pack-reviewed').disabled=busy||blocked||!ready||!packDraft;
 $('pack-private-consent').disabled=busy||blocked||!ready||!packDraft||!privateRoute;
 $('pack-download').disabled=busy||blocked||!ready||!packDraft||!packReviewed(packDraft);
}
function renderCopiedMaterials(){
  if(libraryOrigin)$('activity-library-note').textContent=`Independent copy of ${libraryOrigin.entryName}. Review all text and replace unfinished prompts before insertion. ${copiedMaterials?.length||0} material entries follow this copy; review and edit them below. Earlier copies stay unchanged.`;
  const root=$('activity-copied-materials-list');$('activity-copied-materials').hidden=!copiedMaterials?.length;
  for(const [row,view] of copiedMaterialNodes)if(!copiedMaterials?.includes(row)){view.group.remove();copiedMaterialNodes.delete(row);}
  for(const [index,row] of (copiedMaterials||[]).entries()){
    let entry=copiedMaterialNodes.get(row);
    if(!entry){
      const group=document.createElement('fieldset'),legend=document.createElement('legend'),inputs={};group.append(legend);entry={group,legend,inputs};
      for(const [key,title] of [['item','Name'],['quantity','Activity total quantity'],['unit','Unit (case matters)'],['use','Supply type'],['notes','Private preparation notes']]){
        const id=`copied-material-${crypto.randomUUID()}-${key}`,label=document.createElement('label'),input=document.createElement(key==='use'?'select':key==='notes'?'textarea':'input');label.textContent=title;label.setAttribute('for',id);input.id=id;if(key==='notes')input.rows=3;
        if(key==='use')for(const [value,text] of [['consumable','Consumed or retained'],['reusable','Reusable shared pool']]){const option=document.createElement('option');option.value=value;option.textContent=text;input.append(option);}
        else if(key==='quantity'){input.type='number';input.min='0.001';input.max='100000';input.step='0.001';}else input.maxLength=key==='item'?200:key==='unit'?80:2000;
        input.value=String(row[key]);input.oninput=input.onchange=()=>{row[key]=input.value;activityChanged();};inputs[key]=input;group.append(label,input);
      }
      const remove=document.createElement('button');remove.type='button';remove.textContent='Remove copied material';remove.onclick=()=>run(async()=>{if(!copiedMaterials?.includes(row))throw new Error('This copied material changed. Use its current controls.');copiedMaterials=copiedMaterials.filter(item=>item!==row);activityChanged();});group.append(remove);entry.remove=remove;copiedMaterialNodes.set(row,entry);root.append(group);
    }
    // Keep live input nodes/values/validity intact, including native badInput.
    // A refusal or review render must not reconstruct an unfinished number.
    entry.legend.textContent=`Copied material ${index+1}`;
    for(const input of Object.values(entry.inputs))input.disabled=busy||blocked||!ready;
    entry.remove.disabled=busy||blocked||!ready;
  }
}
function renderMaterials(){
  if(materialsExportDraft&&(!project||materialsExportDraft.snapshot!==JSON.stringify(project)))invalidateMaterialsChecklist();
  if(materialsExportDraft&&materialsExportDraft.planId!==$('materials-checklist-plan').value)invalidateMaterialsChecklist();
  $('materials-checklist-preview').disabled=busy||blocked||!ready||!project;
  $('materials-checklist-reviewed').disabled=busy||blocked||!ready||!materialsExportDraft;
  $('materials-checklist-download').disabled=busy||blocked||!ready||!project||!materialsExportDraft||!$('materials-checklist-reviewed').checked||materialsExportDraft.kind==='plan'&&materialsExportDraft.displayed!==$('materials-checklist-output').textContent;
  const select=$('materials-activity'),chosen=select.value;select.replaceChildren();
  for(const m of project?.modules||[])for(const a of m.activities){const o=document.createElement('option');o.value=a.id;o.textContent=`${m.title} · ${a.title}`;select.append(o);}
  select.value=project?.modules.some(m=>m.activities.some(a=>a.id===chosen))?chosen:project?.modules.flatMap(m=>m.activities)[0]?.id||'';
  const editable=materialsDraft&&materialsDraft.id===select.value;
  select.disabled=$('materials-load').disabled=busy||blocked||!ready||!select.value;
  for(const id of ['material-item','material-quantity','material-unit','material-use','material-notes','materials-save'])$(id).disabled=busy||blocked||!ready||!editable;
  $('materials-add-form').querySelector('button').disabled=busy||blocked||!ready||!editable;
  $('materials-library-form').querySelector('button').disabled=busy||blocked||!ready||!editable;
  $('materials-refresh').disabled=busy||blocked||!ready||!project;
  const draft=$('materials-draft-list');draft.replaceChildren();
  materialsDraft?.items.forEach(m=>{const li=document.createElement('li'),text=document.createElement('p'),remove=document.createElement('button');text.textContent=`${m.quantity} ${m.unit} · ${m.item} · ${m.use}${m.notes?' · '+m.notes:''}`;remove.textContent='Remove draft item';remove.type='button';remove.disabled=busy||blocked||!editable;remove.onclick=()=>run(async()=>{if(!materialsDraft||materialsDraft.id!==$('materials-activity').value||!materialsDraft.items.includes(m))throw new Error('This draft item changed. Use its current remove button.');materialsDraft.items=materialsDraft.items.filter(item=>item!==m);$('materials-review-approved').checked=false;});li.append(text,remove);draft.append(li);});
  $('materials-draft-download').disabled=busy||!materialsDraft;
  $('materials-discard').disabled=busy||blocked||!materialsDraft;
  $('materials-restore-file').disabled=busy||blocked||!ready||!project;
  $('materials-restore').disabled=busy||blocked||!ready||!select.value;
  const restored=!!materialsDraft?.restored&&editable;$('materials-restore-review').hidden=!restored;$('materials-review-approved').disabled=busy||blocked||!restored;
  const current=$('materials-current-list');current.replaceChildren();
  if(restored){$('materials-restore-origin').textContent=`Restored recorded revision ${materialsDraft.restored.recordedRevision} as an unsaved draft. Current checked revision: ${materialsDraft.project.revision}. No requirements have been replaced.`;for(const m of materialsDraft.restored.savedItems){const li=document.createElement('li');li.textContent=`${m.quantity} ${m.unit} · ${m.item} · ${m.use}${m.notes?' · '+m.notes:''}`;current.append(li);}if(!materialsDraft.restored.savedItems.length){const li=document.createElement('li');li.textContent='No saved material entries for this activity.';current.append(li);}if(!$('materials-review-approved').checked)$('materials-save').disabled=true;}
  const plan=$('materials-plan-list');plan.replaceChildren();if(!project||blocked){$('materials-plan-summary').textContent=blocked?'Course operations paused. Saved materials list is unavailable until recovery.':'Create a course to plan materials.';$('materials-plan-assumption').textContent='';return;}
  const p=planMaterials(project);$('materials-plan-summary').textContent=`Revision ${p.revision} · ${p.rows.length} material groups · ${p.configuredActivities} of ${p.totalActivities} activities have entries. Current saved panel view; refresh to include newer Word edits.`;$('materials-plan-assumption').textContent=p.assumption;
  for(const row of p.rows){const li=document.createElement('li');li.textContent=`${row.quantity} ${row.unit} · ${row.item} · ${row.use} · ${row.activities.length} activities${row.notes.length?' · Preparation: '+row.notes.join('; '):''}`;plan.append(li);}
}
function renderOutline(){
  for(const id of ['outline-query','outline-icon','outline-status','outline-clear'])$(id).disabled=busy||blocked||!ready||!project;
  const root=$('outline');root.replaceChildren();
  if(blocked){$('outline-results').textContent='Course operations are paused. Use Check and recover before searching or navigating.';return;}
  if(!project){$('outline-results').textContent='Create a course to see its outline.';return;}
  const result=searchOutline(project,{query:$('outline-query').value,icon:$('outline-icon').value,status:$('outline-status').value});
  $('outline-results').textContent=`Showing ${result.moduleCount} of ${result.totalModules} modules and ${result.activityCount} of ${result.totalActivities} activities · Revision ${project.revision}. Module headings include context for matching activities.${$('outline-status').value==='no-recorded-materials'?' No saved material entries does not mean no supplies are required.':''}`;
  if(!result.rows.length){const empty=document.createElement('p');empty.className='muted';empty.textContent=result.filtered?'No items match these filters. Clear filters or refresh from Word.':'Add a module to build your course outline.';root.append(empty);}
  for(const row of result.rows){
    const button=document.createElement('button');button.type='button';button.disabled=busy||!ready;button.dataset.outlineId=row.id;
    button.textContent=`${row.moduleIndex+1}${row.type==='activity'?'.'+(row.activityIndex+1):''}. ${row.title}`;
    if(row.type==='activity')button.className='outline-activity';
    const detail=document.createElement('small');detail.textContent=row.type==='module'?`${row.activityCount} activities · ${row.minutes} min`:`${row.icon} · ${row.minutes} min`;button.append(detail);
    button.onclick=()=>run(()=>demo?Promise.resolve(message('Item selected in the preview.')):word.jump(row.id,row.type==='activity'?'activity':undefined));root.append(button);
  }
}
function pptFields(){return Object.fromEntries(['title','text','notes','moduleId','lesson'].map(key=>[key,$(key==='moduleId'?'ppt-module':'ppt-'+key).value]));}
function clearComparison(){pptComparison=null;renderComparison();}
function renderComparison(){
  if(pptComparison&&(pptComparison.source!==pptSource||pptComparison.review!==pptReview||pptComparison.editor!==JSON.stringify({id:pptCurrent,...pptFields()})||pptComparison.selection!==pptUpdateSelection||!$('ppt-update-lineage').checked))pptComparison=null;
  for(const id of ['ppt-update-file','ppt-update-lineage','ppt-compare'])$(id).disabled=busy||blocked||!ready||!pptSource;
  const select=$('ppt-change'),chosen=select.value;select.replaceChildren();select.disabled=busy||blocked||!ready||!pptComparison;
  if(!pptComparison){$('ppt-compare-summary').textContent='Choose a revised PPTX to compare. Refresh this comparison after review edits.';$('ppt-compare-warning').textContent='';$('ppt-change-preview').textContent='';return;}
  const result=pptComparison.result,c=result.counts;
  $('ppt-compare-summary').textContent=`Proposals: ${c.added} added · ${c.missing} missing · ${c.changed} changed · ${c.moved} reordered · ${c.unchanged} unchanged text/order · ${c.conflicts} with local conflicts · ${c.ambiguous} with alternate content matches.`;
  $('ppt-compare-warning').textContent=`${result.warning}\nOriginal snapshot: ${result.originalHash}\nIncoming snapshot: ${result.incomingHash}`;
  for(const row of result.rows){const option=document.createElement('option');option.value=row.key;const flags=[row.kind==='added-candidate'?'Added':row.kind==='missing-candidate'?'Missing':row.changed?'Changed':'Matched text',row.moved?'Reordered':'',row.conflict?'Local conflict':''].filter(Boolean).join(' · ');option.textContent=`${flags}: ${(row.incoming?.title||row.original?.title||row.id).slice(0,120)}`;select.append(option);}
  select.value=result.rows.some(r=>r.key===chosen)?chosen:result.rows[0]?.key||'';
  const row=result.rows.find(r=>r.key===select.value);$('ppt-change-preview').textContent=row?`PRIVATE AUTHOR COMPARISON — no changes applied\n${JSON.stringify(row,null,2)}`:'';
}
function loadSlideReview(){
  const slide=pptReview?.slides.find(s=>s.id===pptCurrent),baseline=pptSource?.slides.find(s=>s.id===pptCurrent);
  if(!slide)return;
  const module=$('ppt-module');module.replaceChildren();const empty=document.createElement('option');empty.value='';empty.textContent='Unassigned';module.append(empty);
  for(const m of project?.modules||[]){const option=document.createElement('option');option.value=m.id;option.textContent=m.title;module.append(option);}
  if(slide.moduleId&&!project?.modules.some(m=>m.id===slide.moduleId)){const option=document.createElement('option');option.value=slide.moduleId;option.textContent='Module no longer available — reassign before saving';module.append(option);}
  for(const key of ['title','text','notes','moduleId','lesson'])$(key==='moduleId'?'ppt-module':'ppt-'+key).value=slide[key];
  $('ppt-baseline').textContent=`SOURCE BASELINE — PRIVATE SPEAKER NOTES\n${JSON.stringify(baseline,null,2)}`;
  $('ppt-warnings').textContent=[...pptSource.warnings,...baseline.warnings].join('\n');
}
function renderSlides(){
  const select=$('ppt-slide');select.replaceChildren();for(const s of pptSource?.slides||[]){const option=document.createElement('option');option.value=s.id;option.textContent=`${s.order}. ${s.title||'Untitled slide'} (${s.sourceId})`;select.append(option);}select.value=pptCurrent;
  $('ppt-summary').textContent=pptSource?`${pptSource.source.name}${pptAlias&&pptAlias!==pptSource.source.name?` · Selected file: ${pptAlias}`:''} · ${pptSource.slides.length} slides · SHA-256 ${pptSource.source.sha256} · Review only; no Word insertion.`:'Choose a PPTX to inspect its editable text and speaker notes.';
  for(const id of ['ppt-file','ppt-read','ppt-slide','ppt-title','ppt-text','ppt-notes','ppt-module','ppt-lesson','ppt-review-file','ppt-restore'])$(id).disabled=busy||blocked||!ready||(!['ppt-file','ppt-read','ppt-review-file','ppt-restore'].includes(id)&&!pptReview);
  $('ppt-review-form').querySelector('button').disabled=busy||blocked||!ready||!pptReview;
  $('ppt-export').disabled=busy||!ready||!pptReview;
  $('ppt-reset').disabled=busy||blocked||!ready||!pptReview;
  renderComparison();
}
function renderLibrary(){
  const select=$('library-item'),chosen=libraryPreferred??select.value,filter=$('library-filter').value.trim().toLowerCase();libraryPreferred=null;select.replaceChildren();
  const entries=library.entries.filter(e=>[e.name,...e.tags].join(' ').toLowerCase().includes(filter));
  for(const entry of entries){const option=document.createElement('option');option.value=entry.id;option.textContent=entry.name+(entry.tags.length?' · '+entry.tags.join(', '):'');select.append(option);}
  select.value=entries.some(e=>e.id===chosen)?chosen:entries[0]?.id||'';
  const selected=entries.find(e=>e.id===select.value);
  $('library-preview').textContent=selected?`AUTHORING COPY — includes private teaching notes\n${JSON.stringify(selected.fields,null,2)}`:'';
  $('library-summary').textContent=`${library.entries.length} session copies · ${entries.length} matching`+(selected?` · Unfinished prompts: ${libraryPlaceholders(selected.fields).join(', ')||'none detected'}`:'');
  $('library-save-form').querySelector('button').disabled=busy||blocked||!ready;
  for(const id of ['library-item','library-use','library-export','library-import','library-file','library-filter'])$(id).disabled=busy||blocked||!ready;
  $('library-use').disabled=busy||blocked||!ready||!selected||!project?.modules.length;
  $('library-export').disabled=busy||!ready;
  for(const id of ['library-item','library-filter'])$(id).disabled=busy||!ready;
  for(const id of ['library-manage-load','library-remove'])$(id).disabled=busy||blocked||!ready||!selected;
  const canEdit=!!libraryEdit&&library.entries.some(e=>e.id===libraryEdit.id)&&select.value===libraryEdit.id;
  for(const id of ['library-edit-name','library-edit-tags'])$(id).disabled=busy||blocked||!ready||!canEdit;
  $('library-manage-form').querySelector('button').disabled=busy||blocked||!ready||!canEdit;
  $('library-undo').disabled=busy||!ready||!libraryRemoved;
  $('library-manage-summary').textContent=(libraryEdit?`Loaded name/tags for ${libraryEdit.name}. `:'Load the selected copy to edit its name and tags. ')+(libraryRemoved?`Undo available for removed copy: ${libraryRemoved.entry.name}.`:'');
  $('library-content-load').disabled=busy||!ready||!selected;
  $('library-content-discard').disabled=busy||!ready||!libraryContentDraft;
  for(const key of LIBRARY_CONTENT_FIELDS)$('library-content-'+key).disabled=busy||!ready||!libraryContentDraft;
  $('library-content-reviewed').disabled=busy||!ready||!libraryContentDraft||libraryMaterialDirty;
  $('library-content-save').disabled=busy||!ready||!libraryContentDraft||!selected||selected.id!==libraryContentDraft.id||!$('library-content-reviewed').checked||libraryMaterialDirty;
  $('library-content-undo').disabled=busy||!ready||!libraryContentUndo;
  $('library-content-summary').textContent=libraryContentDraft?`Editing stored content for ${libraryContentDraft.name}. ${library.entries.some(e=>e.id===libraryContentDraft.id)?'Review before saving; newer changes are checked on save.':'This copy was removed. Draft retained; saving cannot recreate it.'}`:'Load selected content to start a separate draft. Download the library after saving to keep it.';
  renderLibraryMaterials();
}
function clearLibraryMaterialEntry(){for(const key of LIBRARY_MATERIAL_FIELDS)$('library-material-'+key).value=key==='quantity'?'1':key==='use'?'consumable':'';libraryMaterialEditing=null;libraryMaterialDirty=false;}
function renderLibraryMaterials(){
  for(const key of LIBRARY_MATERIAL_FIELDS)$('library-material-'+key).disabled=busy||!ready||!libraryContentDraft;
  for(const id of ['library-material-apply','library-material-clear'])$(id).disabled=busy||!ready||!libraryContentDraft;
  $('library-material-summary').textContent=libraryMaterialDirty?'Unfinished material entry: add/apply or clear it before reviewing and saving.':'Material rows are staged only; review and save the stored library content to keep them.';
  const root=$('library-content-materials');root.replaceChildren();const draft=libraryContentDraft;
  for(const row of libraryMaterialRows||[]){
    const li=document.createElement('li'),text=document.createElement('p'),edit=document.createElement('button'),remove=document.createElement('button');text.textContent=`${row.item} · ${row.quantity} ${row.unit} · ${row.use}${row.notes?' · '+row.notes:''}`;
    for(const button of [edit,remove]){button.type='button';button.disabled=busy||!ready||!draft;}
    edit.textContent='Edit stored material';remove.textContent='Remove stored material';
    const check=()=>{if(!draft||libraryContentDraft!==draft||!libraryMaterialRows?.includes(row))throw new Error('This stored material changed. Use its current controls.');if(libraryMaterialDirty)throw new Error('Add/apply or clear the unfinished material entry first.');};
    edit.onclick=()=>run(async()=>{check();libraryMaterialEditing=row;for(const key of LIBRARY_MATERIAL_FIELDS)$('library-material-'+key).value=String(row[key]);libraryMaterialDirty=true;$('library-content-reviewed').checked=false;},{recovery:true});
    remove.onclick=()=>run(async()=>{check();libraryMaterialRows=libraryMaterialRows.filter(item=>item!==row);$('library-content-reviewed').checked=false;},{recovery:true});
    li.append(text,edit,remove);root.append(li);
  }
}
async function run(action,{recovery=false}={}){if(busy)return;if(blocked&&!recovery){message('Use Check and recover before continuing.');return;}busy=true;clearError();render();try{await action();}catch(e){error(e);}finally{busy=false;render();}}
async function refresh(){const before=JSON.stringify(project);if(!demo){const fresh=await word.readProject(project);project=await word.persist(fresh);}if(before!==JSON.stringify(project)){invalidateCoverageReport();agendaDraft=null;$('agenda-reviewed').checked=false;$('agenda-body').replaceChildren();$('agenda-table').hidden=true;clearAgendaCoverage();$('agenda-summary').textContent='The course changed. Refresh and review the agenda again.';}render();}
async function commit(next,kind,item){project=demo?next:await word.insert(project,next,kind,item);invalidateOutputs();render();}
function inputFields(){return {...Object.fromEntries(ACTIVITY_DRAFT_FIELDS.map(key=>[key,$('activity-'+key).value])),...(copiedMaterials===null?{}:{materials:copiedMaterials.map(row=>Object.fromEntries(LIBRARY_MATERIAL_FIELDS.map(key=>[key,copiedMaterialNodes.get(row)?.inputs[key].value??row[key]])))})};}
function activityRaw(){return {generation:activityGeneration,destination:$('activity-module').value,fields:Object.fromEntries(ACTIVITY_DRAFT_FIELDS.map(key=>[key,rawControl($('activity-'+key))])),materials:copiedMaterials===null?null:copiedMaterials.map(row=>Object.fromEntries(LIBRARY_MATERIAL_FIELDS.map(key=>[key,rawControl(copiedMaterialNodes.get(row)?.inputs[key]||{value:String(row[key])})]))),libraryOrigin:structuredClone(libraryOrigin),starterOrigin:structuredClone(starterOrigin)};}
function candidateRaw(fields){return {fields:Object.fromEntries(ACTIVITY_DRAFT_FIELDS.map(key=>[key,rawControl({value:String(fields[key])})])),materials:fields.materials===undefined?null:fields.materials.map(row=>Object.fromEntries(LIBRARY_MATERIAL_FIELDS.map(key=>[key,rawControl({value:String(row[key])})])))};}
function replacementState(){const {destination,...raw}=activityRaw();return {raw,review:starterReview,reviewed:!!$('assessment-reviewed').checked};}
const activityOrigin=()=>libraryOrigin?{kind:'library',...libraryOrigin}:starterOrigin?{kind:'starter',...starterOrigin}:null;
function prepareActivityCandidate(kind){
  if(!project?.modules.length)throw new Error('Create a module before loading an Activity copy.');
  if(kind==='library'){
    const entry=library.entries.find(e=>e.id===$('library-item').value);if(!entry)throw new Error('Select a saved activity copy.');
    const fields=libraryFields(entry.fields);return {kind,id:entry.id,fingerprint:JSON.stringify(entry),fields,origin:{entryId:entry.id,entryName:entry.name,formatVersion:1}};
  }
  const entry=assessmentStarter($('assessment-starter').value);return {kind:'starter',id:entry.origin.id,fingerprint:JSON.stringify(entry),fields:entry.fields,origin:entry.origin};
}
function applyActivityCandidate(candidate){
  for(const key of ACTIVITY_DRAFT_FIELDS)$('activity-'+key).value=String(candidate.fields[key]);
  copiedMaterials=candidate.fields.materials===undefined?null:structuredClone(candidate.fields.materials);
  libraryOrigin=candidate.kind==='library'?structuredClone(candidate.origin):null;
  starterOrigin=candidate.kind==='starter'?structuredClone(candidate.origin):null;
  starterReview=null;$('assessment-reviewed').checked=false;activityGeneration++;activityReplacement=null;
  $('activity-library-note').textContent='';cue();renderCopiedMaterials();view('activity');
  message(candidate.kind==='starter'?'Original assessment starter loaded. Replace every prompt, justify suitability and review this exact draft before insertion.':'Activity copy loaded into the form. Review and edit it, choose the destination module, then insert.');
}
function requestActivityCandidate(kind){
  const candidate=prepareActivityCandidate(kind),defaults=candidateRaw(ACTIVITY_DEFAULTS);
  for(const key of ['title','participant'])defaults.fields[key].invalid.valueMissing=true;
  if(wouldReplaceRaw(activityRaw(),defaults,candidateRaw(candidate.fields),activityOrigin(),{kind:candidate.kind,...candidate.origin})){
    activityReplacement={candidate,state:structuredClone(replacementState())};view('activity');message('Your current Activity draft is retained. Keep it or explicitly discard it before loading the selected copy.');return;
  }
  applyActivityCandidate(candidate);
}
function activityChanged(){activityGeneration++;starterReview=null;$('assessment-reviewed').checked=false;renderActivityDraft();}
function renderActivityDraft(){
  if(starterReview&&!sameRaw(starterReview,activityRaw())){starterReview=null;$('assessment-reviewed').checked=false;}
  $('activity-replacement').hidden=!activityReplacement;
  $('activity-replacement-summary').textContent=activityReplacement?`Loading ${activityReplacement.candidate.kind==='starter'?activityReplacement.candidate.origin.name:activityReplacement.candidate.origin.entryName} would replace the current Activity draft. Nothing has been discarded.`:'';
  $('activity-replace-confirm').disabled=busy||blocked||!ready||!activityReplacement;
  $('activity-replace-cancel').disabled=busy||!activityReplacement;
  $('assessment-starter').disabled=$('assessment-load').disabled=busy||blocked||!ready||!project?.modules.length;
  $('assessment-review').hidden=!starterOrigin;
  $('assessment-origin').textContent=starterOrigin?`${starterOrigin.name} · original editable starter v${starterOrigin.version}. This is session authoring guidance, not a validated measurement instrument.`:'';
  const prompts=starterOrigin?libraryPlaceholders(inputFields()):[];let valid=true;if(starterOrigin)try{libraryFields(inputFields());const raw=activityRaw();valid=!Object.values(raw.fields).some(rawInvalid)&&!(raw.materials||[]).some(row=>Object.values(row).some(rawInvalid));}catch{valid=false;}
  $('assessment-unresolved').textContent=starterOrigin?(prompts.length?`Replace unfinished prompts in: ${prompts.join(', ')}.`:valid?'No recognized prompts remain. Review suitability, private administration/key, content and destination.':'Complete supported fields and material quantities before reviewing.') :'';
  $('assessment-reviewed').disabled=busy||blocked||!ready||!starterOrigin||!!prompts.length||!valid;
  if(starterOrigin)$('activity-insert').disabled=busy||blocked||!ready||!project?.modules.length||!!prompts.length||!valid||!$('assessment-reviewed').checked||!starterReview;
}

function coverageEntry(){return {target:$('coverage-target').value,practice:!!$('coverage-practice').checked,assessment:!!$('coverage-assessment').checked,rationale:$('coverage-rationale').value,current:!!$('coverage-current').checked,editing:coverageEditing};}
function coverageEntryDirty(){const e=coverageEntry();return !!(e.target||e.practice||e.assessment||e.rationale||e.current||e.editing);}
function coverageState(){return JSON.stringify({draft:coverageDraft,links:coverageLinks,entry:coverageEntry(),generation:coverageGeneration});}
function coverageChanged(){coverageGeneration++;coverageReview=null;$('coverage-reviewed').checked=false;}
function clearCoverageEntry(){coverageEditing='';$('coverage-target').value='';$('coverage-practice').checked=$('coverage-assessment').checked=$('coverage-current').checked=false;$('coverage-rationale').value='';coverageChanged();}
function coveragePayload(){return coverageLinks.map(({objectiveModuleId,roles,reviewedObjective,rationale,reviewCurrent})=>({objectiveModuleId,roles:[...roles],reviewedObjective,rationale,reviewCurrent}));}
function setCoverageDraft(base,id){
 const parsed=objectiveCoverage(base);if(!parsed.supported)throw new Error(parsed.message);if(!base.modules.some(m=>m.activities.some(a=>a.id===id)))throw new Error('Choose an existing activity for objective links.');
 coverageDraft={id,project:structuredClone(base)};coverageLinks=parsed.data.links.filter(l=>l.activityId===id).map(({objectiveModuleId,roles,reviewedObjective,rationale})=>({objectiveModuleId,roles:[...roles],reviewedObjective,rationale,reviewCurrent:false}));clearCoverageEntry();
 const target=$('coverage-target');target.replaceChildren();const blank=document.createElement('option');blank.value='';blank.textContent='Choose a nonblank module objective';target.append(blank);
 for(const m of base.modules)if(objectiveText(m).trim()||unsupportedObjective(m)||coverageLinks.some(l=>l.objectiveModuleId===m.id)){const option=document.createElement('option');option.value=m.id;option.textContent=m.title+(unsupportedObjective(m)?' — unsupported objective':objectiveText(m).trim()?'':' — blank/stale target');target.append(option);}target.value='';
}
function renderCoverageEditor(){
 const locked=busy||blocked||!ready||!project,id=$('edit-activity-id').value,editable=!!coverageDraft&&coverageDraft.id===id;
 const target=coverageDraft?.project.modules.find(m=>m.id===$('coverage-target').value),targetUnsupported=!!target&&unsupportedObjective(target),unsupportedLinks=coverageLinks.some(l=>unsupportedObjective(coverageDraft?.project.modules.find(m=>m.id===l.objectiveModuleId)));
 $('coverage-load').disabled=locked||!id;$('coverage-discard').disabled=busy||!coverageDraft;
 for(const key of ['target','practice','assessment','rationale','current','apply','entry-clear'])$('coverage-'+key).disabled=locked||!editable||key==='target'&&!!coverageEditing||targetUnsupported&&['current','apply'].includes(key);
 if(coverageReview!==coverageState()){$('coverage-reviewed').checked=false;coverageReview=null;}
 $('coverage-reviewed').disabled=locked||!editable||coverageEntryDirty()||unsupportedLinks;$('coverage-save').disabled=locked||!editable||coverageEntryDirty()||unsupportedLinks||!$('coverage-reviewed').checked||!coverageReview;
 $('coverage-summary').textContent=coverageDraft?`Loaded activity links at revision ${coverageDraft.project.revision}; ${coverageLinks.length} links in this private draft. ${project&&project.revision!==coverageDraft.project.revision?'The saved course is newer; preserve this draft before discarding/reloading.':'Apply or clear the entry, then review and save the complete list.'}`:'Load current links for the selected existing activity. Other Activity form work stays independent.';
 const old=coverageLinks.find(l=>l.objectiveModuleId===coverageEditing);
 $('coverage-objective').textContent=target?`Loaded current objective (revision ${coverageDraft.project.revision}):\n${objectiveLabel(target)}\nPreviously reviewed text:\n${old?.reviewedObjective||'(new link; explicit review required)'}`:'';
 const root=$('coverage-links');root.replaceChildren();
 for(const row of coverageLinks){const m=project?.modules.find(m=>m.id===row.objectiveModuleId)||coverageDraft.project.modules.find(m=>m.id===row.objectiveModuleId),li=document.createElement('li'),p=document.createElement('p'),edit=document.createElement('button'),remove=document.createElement('button');
  p.textContent=`${m?.title||row.objectiveModuleId} · ${row.roles.join(' + ')} · ${m&&unsupportedObjective(m)?'Objective unavailable; not counted':m&&objectiveText(m).trim()&&m.objective===row.reviewedObjective?'Current declared mapping':'Stale / blank; not counted'}\nReviewed: ${row.reviewedObjective}\nPrivate rationale: ${row.rationale}`;
  edit.type=remove.type='button';edit.textContent='Edit objective link';remove.textContent='Remove link from draft';edit.disabled=locked||!editable;remove.disabled=locked||!editable||unsupportedObjective(m);
  edit.onclick=()=>run(async()=>{if(!coverageLinks.includes(row))throw new Error('This objective-link row changed. Use its current controls.');if(coverageEntryDirty())throw new Error('Apply or explicitly clear the current raw objective-link entry first.');coverageEditing=row.objectiveModuleId;$('coverage-target').value=row.objectiveModuleId;$('coverage-practice').checked=row.roles.includes('practice');$('coverage-assessment').checked=row.roles.includes('assessment');$('coverage-rationale').value=row.rationale;$('coverage-current').checked=false;coverageChanged();});
  remove.onclick=()=>run(async()=>{if(!coverageLinks.includes(row))throw new Error('This objective-link row changed.');if(coverageEntryDirty())throw new Error('Apply or explicitly clear the current raw objective-link entry first.');coverageLinks=coverageLinks.filter(l=>l!==row);coverageChanged();});li.append(p,edit,remove);root.append(li);
 }
 const targetKey=project?JSON.stringify(project.modules.map(m=>[m.id,m.title,m.objective])):'';
 if(!gapDraft&&targetKey!==gapTargets){const select=$('coverage-gap-target'),chosen=select.value;select.replaceChildren();for(const m of project?.modules||[]){const option=document.createElement('option');option.value=m.id;option.textContent=m.title+(unsupportedObjective(m)?' — unsupported objective':objectiveText(m).trim()?'':' — blank objective');select.append(option);}select.value=project?.modules.some(m=>m.id===chosen)?chosen:project?.modules[0]?.id||'';gapTargets=targetKey;}
 for(const key of ['target','role','load'])$('coverage-gap-'+key).disabled=locked||!project?.modules.length;
 $('coverage-gap-discard').disabled=busy||!gapDraft;
 const gapEditable=!!gapDraft&&gapDraft.id===$('coverage-gap-target').value&&gapDraft.role===$('coverage-gap-role').value;
 const gapUnsupported=gapEditable&&unsupportedObjective(gapDraft.project.modules.find(m=>m.id===gapDraft.id));
 for(const key of ['note','current','save'])$('coverage-gap-'+key).disabled=locked||!gapEditable||gapUnsupported&&key!=='note';
 $('coverage-gap-remove').disabled=locked||!gapEditable||!gapDraft.existing||gapUnsupported;
 $('coverage-gap-objective').textContent=gapDraft?`Loaded objective (revision ${gapDraft.project.revision}):\n${objectiveLabel(gapDraft.project.modules.find(m=>m.id===gapDraft.id))}\nPreviously reviewed: ${gapDraft.existing?.reviewedObjective||'(none; explicit current review required)'}`:'';
}
function gapState(){return JSON.stringify({draft:gapDraft,target:$('coverage-gap-target').value,role:$('coverage-gap-role').value,note:$('coverage-gap-note').value,current:!!$('coverage-gap-current').checked,generation:gapGeneration});}
function setGapDraft(base,id,role){const parsed=objectiveCoverage(base);if(!parsed.supported)throw new Error(parsed.message);if(!base.modules.some(m=>m.id===id)||!COVERAGE_ROLES.includes(role))throw new Error('Choose a supported module objective and role.');const existing=parsed.data.exceptions.find(e=>e.objectiveModuleId===id&&e.role===role);gapDraft={project:structuredClone(base),id,role,existing:existing?structuredClone(existing):null};$('coverage-gap-note').value=existing?.note||'';$('coverage-gap-current').checked=false;gapGeneration++;}
function showCoverageReport(){
 const report=coverageReport(project),root=$('coverage-review-rows');root.replaceChildren();if(!report.supported){$('coverage-review-summary').textContent=report.message;return;}
 $('coverage-review-summary').textContent=`Revision ${project.revision}. ${report.scope} Current counts use exact reviewed text only; gap notes never supply coverage.`;
 for(const row of report.rows){const section=document.createElement('section'),heading=document.createElement('h3'),objective=document.createElement('p');heading.textContent=row.title;objective.textContent=row.invalid?'Objective has an unsupported value. Current coverage is unavailable; annotations remain preserved.':row.blank?'Objective is blank. Old annotations remain stale and uncounted.':row.objective;section.append(heading,objective);
  for(const role of COVERAGE_ROLES){const state=row.roles[role],line=document.createElement('p');line.textContent=row.invalid?`${role}: Current coverage unavailable; ${state.unavailable.length} saved links preserved and uncounted.`:`${role}: ${state.current.length} current declared links; ${state.stale.length} stale links. ${state.acknowledgedGap?'Acknowledged GAP.':state.gap?'Unmapped GAP (suggestion).':''}${state.conflict?' Declared links and a gap note coexist; review this conflict.':''}`;section.append(line);
   for(const link of [...state.current,...state.stale,...state.unavailable]){const p=document.createElement('p');p.textContent=`${link.activity?.title||link.activityId} · reviewed: ${link.reviewedObjective} · private rationale: ${link.rationale}`;section.append(p);}
   if(state.exception){const p=document.createElement('p');p.textContent=`${row.invalid?'Unverified private gap note (not counted)':state.exceptionCurrent?'Current private gap note (never coverage)':'Stale private gap note (not counted)'}: ${state.exception.note}\nReviewed: ${state.exception.reviewedObjective}`;section.append(p);}
  }root.append(section);
 }
}
$('coverage-load').onclick=()=>run(async()=>{if(coverageDraft)throw new Error('Preserve and discard the existing objective-link draft before loading another.');const id=$('edit-activity-id').value;await refresh();if(id!==$('edit-activity-id').value)throw new Error('The selected activity changed while loading.');setCoverageDraft(project,id);});
$('coverage-discard').onclick=()=>run(async()=>{coverageDraft=null;coverageLinks=[];clearCoverageEntry();$('coverage-target').replaceChildren();message('Objective-link draft discarded. Saved annotations and Activity insertion draft are unchanged.');},{recovery:true});
$('coverage-entry-clear').onclick=()=>run(async()=>clearCoverageEntry());
for(const id of ['coverage-target','coverage-practice','coverage-assessment','coverage-rationale','coverage-current'])$(id).oninput=$(id).onchange=()=>{coverageChanged();renderCoverageEditor();};
$('coverage-entry-form').onsubmit=e=>{e.preventDefault();run(async()=>{
 if(!coverageDraft||coverageDraft.id!==$('edit-activity-id').value)throw new Error('Load this activity’s links before editing.');const entry=coverageEntry(),m=coverageDraft.project.modules.find(m=>m.id===entry.target);if(!m)throw new Error('Choose a current module objective.');if(!coverageEditing&&coverageLinks.some(l=>l.objectiveModuleId===entry.target))throw new Error('Edit the existing link rather than creating a duplicate.');
 const old=coverageLinks.find(l=>l.objectiveModuleId===coverageEditing),payload={objectiveModuleId:m.id,roles:COVERAGE_ROLES.filter(r=>entry[r]),reviewedObjective:entry.current?m.objective:old?.reviewedObjective||'',rationale:entry.rationale,reviewCurrent:entry.current},links=coveragePayload(),at=links.findIndex(l=>l.objectiveModuleId===coverageEditing);if(at>=0)links[at]=payload;else links.push(payload);
 const checked=editObjectiveCoverage(coverageDraft.project,{kind:'links',activityId:coverageDraft.id,links}),saved=checked.objectiveCoverage.links.find(l=>l.activityId===coverageDraft.id&&l.objectiveModuleId===m.id),row={objectiveModuleId:m.id,roles:[...saved.roles],reviewedObjective:saved.reviewedObjective,rationale:saved.rationale,reviewCurrent:payload.reviewCurrent};if(at>=0)coverageLinks[at]=row;else coverageLinks.push(row);clearCoverageEntry();message('Objective link applied to the private draft only. Review and save the complete list.');
 });};
$('coverage-reviewed').onchange=()=>{coverageReview=$('coverage-reviewed').checked&&coverageDraft&&!coverageEntryDirty()?coverageState():null;renderCoverageEditor();};
$('coverage-save').onclick=()=>run(async()=>{
 const draft=coverageDraft,state=coverageState();if(!draft||draft.id!==$('edit-activity-id').value||coverageEntryDirty()||!$('coverage-reviewed').checked||coverageReview!==state)throw new Error('Apply or clear the raw entry, then review this complete objective-link draft before saving.');const change={kind:'links',activityId:draft.id,links:coveragePayload()};
 const next=demo?(JSON.stringify(project)===JSON.stringify(draft.project)?editObjectiveCoverage(project,change):(()=>{throw new Error('The course changed. Preserve the draft and reload.');})()):await word.updateObjectiveCoverage(draft.project,change);
 const unchanged=coverageDraft===draft&&state===coverageState();project=next;invalidateOutputs();if(unchanged)setCoverageDraft(next,draft.id);showCoverageReport();message('Private objective links saved. Any changed/blank objective remains stale until explicitly re-reviewed. Newer draft work is retained.');
});
$('coverage-gap-load').onclick=()=>run(async()=>{if(gapDraft)throw new Error('Preserve and discard the loaded gap-note draft before loading another.');const id=$('coverage-gap-target').value,role=$('coverage-gap-role').value;await refresh();if(id!==$('coverage-gap-target').value||role!==$('coverage-gap-role').value)throw new Error('The selected gap target changed while loading.');setGapDraft(project,id,role);$('coverage-gap-status').textContent='Loaded private gap note; explicit current-objective review is required for a new/reapproved snapshot.';});
$('coverage-gap-discard').onclick=()=>run(async()=>{gapDraft=null;$('coverage-gap-note').value='';$('coverage-gap-current').checked=false;gapGeneration++;$('coverage-gap-status').textContent='Gap-note draft discarded. Saved notes are unchanged.';},{recovery:true});
for(const id of ['coverage-gap-target','coverage-gap-role','coverage-gap-note','coverage-gap-current'])$(id).oninput=$(id).onchange=()=>{gapGeneration++;renderCoverageEditor();};
async function saveGap(remove){
 const draft=gapDraft,state=gapState();if(!draft||draft.id!==$('coverage-gap-target').value||draft.role!==$('coverage-gap-role').value)throw new Error('Load this objective and role before saving its note.');if(remove&&(!$('coverage-gap-note').value||$('coverage-gap-note').value!==draft.existing?.note||$('coverage-gap-current').checked))throw new Error('Preserve/discard raw gap-note changes before removing the saved note.');
 const change={kind:'exception',objectiveModuleId:draft.id,role:draft.role,remove,...(remove?{}:{note:$('coverage-gap-note').value,reviewCurrent:!!$('coverage-gap-current').checked,reviewedObjective:$('coverage-gap-current').checked?draft.project.modules.find(m=>m.id===draft.id).objective:draft.existing?.reviewedObjective||''})};
 const next=demo?(JSON.stringify(project)===JSON.stringify(draft.project)?editObjectiveCoverage(project,change):(()=>{throw new Error('The course changed. Preserve the note and reload.');})()):await word.updateObjectiveCoverage(draft.project,change);
 const unchanged=gapDraft===draft&&state===gapState();project=next;invalidateOutputs();if(unchanged)setGapDraft(next,draft.id,draft.role);showCoverageReport();$('coverage-gap-status').textContent=remove?'Saved gap note removed.':'Private gap note saved. It acknowledges a gap and never counts as coverage.';
}
$('coverage-gap-save').onclick=()=>run(()=>saveGap(false));$('coverage-gap-remove').onclick=()=>run(()=>saveGap(true));

function showPreview(){
  workbookDraft=null;$('workbook-plan-reviewed').checked=false;
  const planId=$('workbook-plan').value,root=$('workbook-preview');root.replaceChildren();
  if(planId){
    const stamp=new Date().toISOString(),v=planWorkbookPreview(project,planId,stamp),add=(tag,value)=>{const x=document.createElement(tag);x.textContent=value;root.append(x);};
    add('h2',v.title);add('p','Participant workbook');add('p',v.scopeLabel);add('p',`Project ID: ${v.projectId} · Generation revision: ${v.revision} · Generated: ${v.stamp}`);
    for(const line of v.identityLines)add('p',line);
    for(const m of v.modules){add('h3',m.title);if(m.objective)add('p',m.objective);for(const a of m.activities){add('h3',a.title);add('p',a.participant);for(let i=0;i<a.lines;i++){const line=document.createElement('div');line.className='writing-line';root.append(line);}}}
    previewRevision=project.revision;workbookDraft={kind:'plan',planId,revision:project.revision,stamp,snapshot:JSON.stringify(project),preview:JSON.stringify(v),displayed:root.textContent,displayStructure:workbookDisplayStructure(root)};return;
  }
  const v=workbookPreview(project);
  const issues=reviewCourse(project),errors=issues.filter(x=>x.severity==='error').length;
  const add=(tag,text)=>{const x=document.createElement(tag);x.textContent=text;root.append(x);};
  add('h2',v.title);add('p','Participant workbook');add('p',WORKBOOK_SCOPE);add('p',`Project ID: ${v.projectId} · Generation revision: ${v.revision}`);
  add('p',`${errors} items need attention · ${issues.length-errors} suggestions. Review Guide lists the unresolved items. Resolve items marked Needs attention before downloading.`);
  for(const m of v.modules){add('h3',m.title);if(m.objective)add('p',m.objective);for(const a of m.activities){add('h3',a.title);add('p',a.participant);for(let i=0;i<a.lines;i++){const line=document.createElement('div');line.className='writing-line';root.append(line);}}}
  previewRevision=project.revision;
  if(!errors)workbookDraft={revision:project.revision,snapshot:JSON.stringify(project)};
}
function workbookDisplayStructure(root){return JSON.stringify(Array.from(root.children,node=>({tag:node.tagName||'',className:node.className||'',hidden:!!node.hidden,style:node.style?.cssText||'',text:node.textContent})));}
function selectedWorkbookReviewed(draft){return !blocked&&ready&&workbookDraft===draft&&draft.kind==='plan'&&$('workbook-plan').value===draft.planId&&$('workbook-plan-reviewed').checked&&draft.snapshot===JSON.stringify(project)&&draft.preview===JSON.stringify(planWorkbookPreview(project,draft.planId,draft.stamp))&&draft.displayed===$('workbook-preview').textContent&&draft.displayStructure===workbookDisplayStructure($('workbook-preview'));}
const familyWorkspaceStyle=node=>node.getAttribute?.('style')||node.style?.cssText||node.attributes?.style||'';
const familyWorkspaceAncestors=()=>{const nodes=[];for(let node=$('view-output');node;node=node.parentElement||node.parent)nodes.push(node);return nodes;};
const familyWorkspaceState=()=>familyWorkspaceAncestors().map(node=>({tag:node.tagName||'',className:node.className||'',hidden:!!node.hidden,style:familyWorkspaceStyle(node)}));
const familyWorkspaceVisible=()=>{
  const view=$('view-output');if(!view||view.isConnected===false)return false;
  for(const node of familyWorkspaceAncestors()){
    const inline=familyWorkspaceStyle(node);
    if(node.hidden||/display\s*:\s*none|visibility\s*:\s*(?:hidden|collapse)|content-visibility\s*:\s*hidden|(?:^|;)\s*opacity\s*:\s*0(?:\s*!important)?\s*(?:;|$)/u.test(inline))return false;
    if(typeof globalThis.getComputedStyle==='function'){let computed;try{computed=globalThis.getComputedStyle(node);}catch{return false;}if(computed&&(computed.display==='none'||['hidden','collapse'].includes(computed.visibility)||computed.opacity!==''&&computed.opacity!==undefined&&Number(computed.opacity)===0))return false;}
  }
  return true;
};
const refuseHiddenFamilyAction=()=>{
  if(familyWorkspaceVisible())return false;
  familyGeneration++;invalidateFamily();$('family-brand-status').textContent='Open the Workbook workspace before reviewing or applying a new copy. Your drafts were kept.';renderFamily();return true;
};
const familyRawBrand=()=>Object.fromEntries(FAMILY_BRAND_FIELDS.map(key=>[key,rawControl($('family-brand-'+key))]));
const familyRawContents=()=>({checked:!!$('family-toc-include').checked,...rawControl($('family-toc-include'))});
const familyChoice=()=>({planId:$('workbook-plan').value,familyId:$('family-template').value,
  baseLanguage:$('family-base-language').value,rawBrand:familyRawBrand(),rawFrontMatter:familyFrontMatterPanel?.raw()??null,rawResearch:familyResearchPanel?.raw()??null,rawToc:familyRawContents(),rawLogo:familyLogoPanel?.raw()??null,logoIdentity:familyLogoPanel?.identity()??null,generation:familyGeneration});
const familyConfig=(choice,stamp)=>({planId:choice.planId,familyId:choice.familyId,baseLanguage:choice.baseLanguage,rawBrand:choice.rawBrand,stamp,...(choice.rawFrontMatter?{frontMatter:workbookFrontMatterRequest(choice.rawFrontMatter)}:{}),toc:workbookContentsRequest(choice.rawToc),...(choice.rawResearch?{research:workbookResearchRequest(choice.rawResearch)}:{}),...(choice.rawLogo?{logo:workbookLogoRequest(choice.rawLogo)}:{})});
const familyDesignUnavailable=()=>!!familyDesignFilesPanel?.isPending()||!!familyDesignFilesPanel?.isLoading();
const familyPending=()=>familyDesignUnavailable()||!!familyFrontMatterPanel?.isPending()||!!familyResearchPanel?.isPending()||!!familyLogoPanel?.isPending()||!!familyLogoPanel?.isLoading();
const familyDisplayStructure=()=>{
  const root=$('family-preview-content'),status=$('family-brand-status'),alternative=$('family-use-safe-alternative');
  const shape=node=>({tag:node.tagName||'',className:node.className||'',hidden:!!node.hidden,style:node.style?.cssText||'',text:node.textContent,
    children:Array.from(node.children||[],shape)});
  return JSON.stringify({workspace:familyWorkspaceState(),section:shape($('workbook-family-section')),brandControls:shape($('family-brand-controls')),frontControls:shape($('family-front-controls')),researchControls:shape($('family-research-controls')),tocControls:shape($('family-toc-controls')),tocChoice:shape($('family-toc-include')),
    logo:familyLogoPanel?.displayState()??null,reviewLabel:shape($('family-review-label')),preview:shape(root),status:shape(status),
    alternative:{hidden:!!alternative.hidden,text:alternative.textContent}});
};
function invalidateFamily(){
  familyDraft=null;familyAlternative=null;$('family-reviewed').checked=false;
  $('family-brand-status').textContent='Choose a design and refresh its preview.';
  $('family-preview-content').textContent='Refresh and review the new workbook design before applying it to a new copy.';
  $('family-use-safe-alternative').hidden=true;$('family-use-safe-alternative').disabled=true;
}
const familyPreparedState=prepared=>JSON.stringify({semantic:prepared.semantic,brand:prepared.brand,
  profileVersion:prepared.output.profileVersion,assetProfileVersion:prepared.output.assetProfileVersion??null,preview:prepared.output.preview,files:prepared.output.files,selectedPlanName:prepared.selectedPlanName});
function familyCurrent(draft,requireConsent=false){
  if(!draft||familyDraft!==draft||!project||blocked||!ready||!familyWorkspaceVisible()||draft.snapshot!==JSON.stringify(project)||
     draft.choice!==JSON.stringify(familyChoice())||draft.display!==familyDisplayStructure()||
     requireConsent&&(!$('family-reviewed').checked||familyPending()))return false;
  try{
    const choice=familyChoice(),prepared=prepareFamilyWorkbook(project,familyConfig(choice,draft.stamp));
    return prepared.eligible&&draft.prepared===familyPreparedState(prepared)&&draft.reviewText===familyReviewText(prepared);
  }catch{return false;}
}
function renderFamily(){
  if(familyDraft&&!familyCurrent(familyDraft))invalidateFamily();
  const unavailable=busy||blocked||!ready||!project||!familyWorkspaceVisible()||familyDesignUnavailable();
  $('family-template').disabled=unavailable;$('family-base-language').disabled=unavailable;$('family-toc-include').disabled=unavailable;$('family-preview').disabled=unavailable||familyPending();
  for(const key of FAMILY_BRAND_FIELDS)$('family-brand-'+key).disabled=unavailable;
  $('family-use-safe-alternative').disabled=unavailable||!familyAlternative;
  $('family-reviewed').disabled=unavailable||!familyDraft||familyPending();
  $('family-apply').disabled=unavailable||!familyCurrent(familyDraft,true);
  familyFrontMatterPanel?.renderBusy();
  familyResearchPanel?.renderBusy();
  familyLogoPanel?.renderBusy();
  renderBrandPresets();
}
const presetStatus=text=>{$('family-preset-status').textContent=text;};
const presetIdentity=()=>JSON.stringify({generation:familyGeneration,family:$('family-template').value,
  raw:familyRawBrand(),selected:$('family-preset-list').value,name:$('family-preset-name').value,
  catalogRaw:presetState?.raw??null,project:JSON.stringify(project)});
const presetReady=()=>{if(familyDesignUnavailable())throw new Error('Finish or Keep the design bundle choice before changing brand presets.');if(busy)throw new Error('Brand presets: finish the current operation before changing presets.');};
function renderBrandPresets(){
  const designBlocked=familyDesignUnavailable();
  const usable=!!presetState&&!busy&&!presetNeedsReload&&!designBlocked;
  for(const id of ['family-preset-save-new','family-preset-overwrite','family-preset-load'])$(id).disabled=!usable;
  $('family-preset-reset').disabled=busy||designBlocked;
  $('family-preset-reload').disabled=busy||designBlocked;
  $('family-preset-list').disabled=busy||!presetState||designBlocked;
  $('family-preset-name').disabled=busy||!presetState||designBlocked;
  $('family-preset-keep').disabled=busy||!presetPending;
  $('family-preset-discard').disabled=busy||!presetPending;
}
function populateBrandPresets(preferred=''){
  const select=$('family-preset-list');select.replaceChildren();
  const blank=document.createElement('option');blank.value='';blank.textContent='Choose a saved preset';select.append(blank);
  for(const row of presetState?.catalog.presets||[]){const option=document.createElement('option');option.value=row.id;option.textContent=row.name;select.append(option);}
  select.value=presetState?.catalog.presets.some(row=>row.id===preferred)?preferred:'';
  renderBrandPresets();
}
function reloadBrandPresets(){
  try{const storage=globalThis.localStorage;const fresh=readBrandPresetCatalog(storage);
    presetStore=storage;presetState=fresh;presetPending=null;presetNeedsReload=false;$('family-preset-confirm').hidden=true;
    populateBrandPresets($('family-preset-list').value);
    presetStatus(`${fresh.catalog.presets.length} local brand preset(s). These belong to this WebView origin and may not follow another Word window or device.`);
  }catch(e){presetPending=null;$('family-preset-confirm').hidden=true;
    presetStatus(e.message||String(e));renderBrandPresets();}
}
function presetRawBrand(brand){
  const current=familyRawBrand();
  return Object.fromEntries(FAMILY_BRAND_FIELDS.map(key=>[key,{value:brand[key],
    invalid:Object.fromEntries(Object.keys(current[key].invalid).map(flag=>[flag,false]))}]));
}
function checkedPresetBrand(brand){
  const choice=familyChoice(),request=familyBrandRequest(presetRawBrand(brand));
  const resolved=resolveWorkbookBrand(choice.familyId,request);
  if(!resolved.eligible)throw new Error(`Brand presets: this brand cannot be used with ${choice.familyId} under current rules. ${resolved.issues.join(' ')} Current branding was kept.`);
  try{validateWorkbookFamilyBrand(resolved);}catch(e){throw new Error(`Brand presets: ${e.message} Current branding was kept.`);}
}
function currentPresetBrand(){
  familyBrandRequest(familyRawBrand());
  return Object.fromEntries(FAMILY_BRAND_FIELDS.map(key=>[key,$('family-brand-'+key).value]));
}
function presetSelected(){return presetState?.catalog.presets.find(row=>row.id===$('family-preset-list').value)||null;}
function presetWrite(action){
  presetReady();
  if(!presetState)throw new Error('Brand presets: reload local storage first.');
  if(presetNeedsReload)throw new Error('Brand presets: the previous write outcome is unconfirmed. Reload saved presets before writing again.');
  const name=$('family-preset-name').value,brand=currentPresetBrand(),selected=presetSelected();
  checkedPresetBrand(brand);
  if(action==='overwrite'&&!selected)throw new Error('Brand presets: choose a saved preset to overwrite.');
  const id=action==='create'?crypto.randomUUID():selected.id;
  const next=nextBrandPresetCatalog(presetState.catalog,action,{id,name,brand,uuid:crypto.randomUUID()});
  let written;
  try{written=writeBrandPresetCatalog(presetStore,presetState.raw,next);}
  catch(e){if(/write outcome is unconfirmed/u.test(e.message||''))presetNeedsReload=true;throw e;}
  presetState=written;populateBrandPresets(id);presetStatus(`${action==='create'?'Saved':'Overwrote'} “${name}” locally. Current workbook review and branding were kept.`);
}
function beginPresetAction(action){
  presetReady();
  if(action!=='reset'&&!presetState)throw new Error('Brand presets: reload local storage first.');
  const selected=presetSelected();
  if(action==='load'&&!selected)throw new Error('Brand presets: choose a saved preset to load.');
  if(action==='overwrite'&&!selected)throw new Error('Brand presets: choose a saved preset to overwrite.');
  if(action==='load')checkedPresetBrand(selected.brand);
  const label=action==='load'?`Load “${selected.name}” into the currently selected design? This replaces your active typed brand fields and preset-name work.`:
    action==='reset'?'Reset the active brand fields to this design’s defaults? Your saved presets stay intact.':
    `Overwrite “${selected.name}” with your current typed brand and name?`;
  presetPending={action,identity:presetIdentity(),presetId:action==='reset'?null:selected?.id||null,catalogRaw:presetState?.raw??null};
  $('family-preset-confirm-text').textContent=label;
  $('family-preset-discard').textContent=action==='load'?'Discard current branding and load preset':action==='reset'?'Discard current branding and reset':'Confirm overwrite';
  $('family-preset-confirm').hidden=false;presetStatus('Choose Keep or the explicit action below. No branding has changed.');renderBrandPresets();
}
function confirmPresetAction(){
  presetReady();
  const pending=presetPending;
  if(!pending||pending.identity!==presetIdentity()||pending.catalogRaw!==(presetState?.raw??null))
    throw new Error('Brand presets: controls changed during confirmation; current branding was kept. Start again.');
  const fresh=pending.action==='reset'?null:readBrandPresetCatalog(presetStore);
  if(fresh&&fresh.raw!==pending.catalogRaw)throw new Error('Brand presets: another session changed the saved catalog; current branding was kept. Reload presets.');
  const selected=presetSelected();
  if(pending.presetId&&(selected?.id!==pending.presetId||JSON.stringify(selected)!==JSON.stringify(fresh.catalog.presets.find(row=>row.id===pending.presetId))))
    throw new Error('Brand presets: the selected preset changed; current branding was kept.');
  if(pending.action==='overwrite'){presetWrite('overwrite');}
  else{
    const brand=pending.action==='load'?selected.brand:Object.fromEntries(FAMILY_BRAND_FIELDS.map(key=>[key,'']));
    if(pending.action==='load')checkedPresetBrand(brand);
    const changed=JSON.stringify(familyRawBrand())!==JSON.stringify(presetRawBrand(brand));
    if(changed)for(const key of FAMILY_BRAND_FIELDS){const control=$('family-brand-'+key);control.value=brand[key];control.setCustomValidity?.('');}
    if(pending.action==='load')$('family-preset-name').value=selected.name;
    if(changed){familyGeneration++;invalidateFamily();}
    presetStatus(pending.action==='load'?`Loaded “${selected.name}” into the current design.${changed?' Refresh and review the new copy.':' Existing workbook review remains current.'}`:
      `Active brand reset to design defaults. Saved presets were kept.${changed?' Refresh and review the new copy.':' Existing workbook review remains current.'}`);
  }
  presetPending=null;$('family-preset-confirm').hidden=true;render();
}
async function showFamilyPreview(){
  invalidateFamily();await refresh();
  const choice=familyChoice(),stamp=new Date().toISOString();let prepared;
  try{prepared=prepareFamilyWorkbook(project,familyConfig(choice,stamp));}
  catch(e){$('family-brand-status').textContent=e.message||String(e);$('family-preview-content').textContent='No new workbook copy is ready. Correct the displayed setting or saved learner content, then refresh this preview.';render();return;}
  if(!prepared.eligible){
    $('family-brand-status').textContent=`This requested design is unavailable: ${prepared.issues.join(' ')}${Object.keys(prepared.alternative).length?` Suggested values (not applied): ${JSON.stringify(prepared.alternative)}.`:''}`;
    $('family-preview-content').textContent='No new workbook copy is ready. Your typed branding values are preserved.';
    if(Object.keys(prepared.alternative).length){familyAlternative={fields:prepared.alternative,choice:JSON.stringify(choice)};$('family-use-safe-alternative').hidden=false;}
    render();return;
  }
  const reviewText=familyReviewText(prepared);$('family-preview-content').textContent=reviewText;
  $('family-brand-status').textContent='Review the complete learner content and actual applied design values below. This is a source preview, not a rendered Word page.';
  familyDraft={choice:JSON.stringify(choice),snapshot:JSON.stringify(project),stamp,prepared:familyPreparedState(prepared),
    output:prepared.output,files:prepared.output.files,title:prepared.output.preview.title,selectedPlanName:prepared.selectedPlanName,familyId:choice.familyId,
    reviewText,display:familyDisplayStructure()};render();
}
document.querySelectorAll('[data-view]').forEach(b=>b.onclick=()=>{if(b.dataset.view==='guide')return openTool('openDesigner');toolStrip.hide();return navigate(b.dataset.view);});
$('outline-status').value='all';
for(const [label] of ICONS){const option=document.createElement('option');option.value=label;option.textContent=label;$('outline-icon').append(option);}
$('outline-query').oninput=()=>{if(!busy&&!blocked)renderOutline();};
for(const id of ['outline-icon','outline-status'])$(id).onchange=()=>{if(!busy&&!blocked)renderOutline();};
$('outline-clear').onclick=()=>{if(busy||blocked)return;$('outline-query').value='';$('outline-icon').value='';$('outline-status').value='all';renderOutline();};
$('show-outline').onclick=()=>view('outline');
for(const [label,file] of ICONS){const o=document.createElement('option');o.value=label;o.textContent=label;$('activity-icon').append(o);}
for(const [label] of ICONS){const o=document.createElement('option');o.value=label;o.textContent=label;$('library-content-icon').append(o);}
$('activity-icon').value='Learning activity';
function cue(){const [label,file]=ICONS.find(([x])=>x===$('activity-icon').value);const img=document.createElement('img');img.src=moduleAsset('assets/icons/'+encodeURIComponent(file),import.meta.url);img.alt='';const text=document.createElement('span');text.textContent=label+' cue';$('cue-preview').replaceChildren(img,text);}
for(const key of ACTIVITY_DRAFT_FIELDS)$('activity-'+key).oninput=$('activity-'+key).onchange=activityChanged;
$('activity-icon').onchange=()=>{activityChanged();cue();};cue();
$('activity-module').onchange=render;
for(const entry of ASSESSMENT_STARTERS){const option=document.createElement('option');option.value=entry.id;option.textContent=entry.name;$('assessment-starter').append(option);}
$('assessment-load').onclick=()=>run(async()=>requestActivityCandidate('starter'));
$('activity-replace-cancel').onclick=()=>run(async()=>{activityReplacement=null;message('Current Activity draft kept. Nothing was replaced.');},{recovery:true});
$('activity-replace-confirm').onclick=()=>run(async()=>{
  const pending=activityReplacement;if(!pending||!sameRaw(pending.state,replacementState()))throw new Error('The Activity draft changed. Keep it and request the replacement again before discarding anything.');
  const current=prepareActivityCandidate(pending.candidate.kind);
  if(current.id!==pending.candidate.id||current.fingerprint!==pending.candidate.fingerprint)throw new Error('The selected source changed. Your Activity draft is unchanged; request its current copy again.');
  applyActivityCandidate(current);
});
$('assessment-reviewed').onchange=()=>{
  starterReview=null;
  if(starterOrigin&&$('assessment-reviewed').checked){try{libraryFields(inputFields());const raw=activityRaw();if(libraryPlaceholders(inputFields()).length||Object.values(raw.fields).some(rawInvalid)||(raw.materials||[]).some(row=>Object.values(row).some(rawInvalid)))throw new Error('Unfinished prompts or invalid entries');starterReview=raw;}catch{$('assessment-reviewed').checked=false;}}
  renderActivityDraft();
};
$('material-quantity').value='1';$('material-use').value='consumable';
$('materials-restore-file').onchange=()=>{materialsRestoreSelection++;};
$('materials-review-approved').onchange=render;
$('materials-discard').onclick=()=>run(async()=>{
  materialsDraft=null;$('materials-review-approved').checked=false;
  for(const id of ['material-item','material-unit','material-notes'])$(id).value='';$('material-quantity').value='1';$('material-use').value='consumable';
  $('materials-draft-summary').textContent='Draft discarded. Load current materials or restore a saved private draft.';
});
$('materials-restore').onclick=()=>run(async()=>{
  if(materialsDraft)throw new Error('Preserve and discard the current materials draft before restoring another.');
  const file=$('materials-restore-file').files?.[0],selection=materialsRestoreSelection,id=$('materials-activity').value,base=project;
  if(!file||file.size>MAX_MATERIAL_DRAFT_BYTES)throw new Error('Choose a private materials draft up to 2 MB.');
  const text=await file.text();
  if(selection!==materialsRestoreSelection||id!==$('materials-activity').value||base!==project)throw new Error('The selected file or destination changed. Check the draft again.');
  const restored=restoreMaterialsDraft(text,base,id),fresh=demo?project:await word.readProject();
  if(selection!==materialsRestoreSelection||id!==$('materials-activity').value||base!==project)throw new Error('The selected file or destination changed. Check the draft again.');
  if(JSON.stringify(fresh)!==JSON.stringify(base))throw new Error('The course changed. Refresh the saved list before restoring this draft.');
  materialsDraft=restored.draft;
  for(const [key,value] of Object.entries(restored.unsavedEntry))$('material-'+key).value=value;
  $('materials-review-approved').checked=false;$('materials-draft-summary').textContent='Private draft restored for review. Compare with the checked saved requirements before saving.';
});
$('materials-activity').onchange=()=>{$('materials-draft-summary').textContent='Load the selected activity. The previous draft stays available if you select that activity again.';render();};
$('materials-load').onclick=()=>run(async()=>{
  if(materialsDraft)throw new Error('Preserve and discard the current materials draft before loading requirements again.');
  const id=$('materials-activity').value;await refresh();const items=activityMaterials(project,id);
  materialsDraft={id,project:structuredClone(project),items};$('materials-activity').value=id;
  $('materials-review-approved').checked=false;
  $('materials-draft-summary').textContent=`Loaded materials at revision ${project.revision}. Edit the draft, then save.`;
});
$('materials-add-form').onsubmit=e=>{e.preventDefault();run(async()=>{
  if(!materialsDraft||materialsDraft.id!==$('materials-activity').value)throw new Error('Load the selected activity materials first.');
  const entry={item:$('material-item').value,quantity:$('material-quantity').value,unit:$('material-unit').value,use:$('material-use').value,notes:$('material-notes').value};
  const items=normalizeMaterials([...materialsDraft.items,entry]);materialsDraft.items=items;
  $('materials-review-approved').checked=false;
  message('Item added to the materials draft. Save activity materials to store it with the course.');
});};
$('materials-save').onclick=()=>run(async()=>{
  if(!materialsDraft||materialsDraft.id!==$('materials-activity').value)throw new Error('Load the selected activity materials first.');
  if(materialsDraft.restored&&!$('materials-review-approved').checked)throw new Error('Review the restored list against the checked saved requirements before saving.');
  const next=editActivityMaterials(materialsDraft.project,materialsDraft.id,materialsDraft.items);
  project=demo?next:await word.updateActivityMaterials(materialsDraft.project,materialsDraft.id,materialsDraft.items);
  materialsDraft={id:materialsDraft.id,project:structuredClone(project),items:activityMaterials(project,materialsDraft.id)};invalidateOutputs();
  $('materials-draft-summary').textContent=`Saved materials at revision ${project.revision}. Save the Word file to keep them.`;
  message('Activity materials saved in the authoring record. Word text and images stay unchanged.');
});
$('materials-checklist-reviewed').onchange=render;
$('materials-checklist-plan').onchange=()=>{if(busy||blocked)return;invalidateMaterialsChecklist();render();};
$('materials-checklist-preview').onclick=()=>run(async()=>{
  invalidateMaterialsChecklist();await refresh();
  const stamp=new Date().toISOString(),planId=$('materials-checklist-plan').value,preview=planId?planMaterialsPreview(project,planId,stamp):materialsPreview(project,stamp);
  $('materials-checklist-output').textContent=materialsParagraphs(preview).map(([,text])=>text).join('\n');
  materialsExportDraft={kind:planId?'plan':'full',planId,snapshot:JSON.stringify(project),preview:JSON.stringify(preview),displayed:$('materials-checklist-output').textContent,stamp};
  message('Review the exact checklist below. It uses saved requirements; private notes and unfinished drafts are excluded. Completeness is not certified.');
});
function selectedMaterialsReviewed(draft){return !blocked&&ready&&materialsExportDraft===draft&&draft.kind==='plan'&&$('materials-checklist-plan').value===draft.planId&&$('materials-checklist-reviewed').checked&&draft.snapshot===JSON.stringify(project)&&draft.preview===JSON.stringify(planMaterialsPreview(project,draft.planId,draft.stamp))&&draft.displayed===$('materials-checklist-output').textContent;}
$('materials-checklist-download').onclick=()=>run(async()=>{
  const draft=materialsExportDraft;
  if(!draft||!$('materials-checklist-reviewed').checked)throw new Error('Refresh and review the materials checklist before downloading.');
  if(draft.kind==='plan'){
    await refresh();if(!selectedMaterialsReviewed(draft))throw new Error('The selected plan, course or checklist review changed. Refresh and review the selected checklist.');
    if(!globalThis.JSZip)throw new Error('The materials checklist exporter has not loaded.');
    const snapshot=structuredClone(project),blob=await planMaterialsBlob(snapshot,draft.planId,draft.stamp,globalThis.JSZip);
    await refresh();if(!selectedMaterialsReviewed(draft))throw new Error('The selected plan, course or checklist review changed. Refresh and review the selected checklist.');
    downloadBlob(blob,planMaterialsName(planMaterialsPreview(snapshot,draft.planId,draft.stamp).planName));
    message('New selected-plan materials checklist prepared from the exact reviewed scope. Recorded requirements do not certify completeness.');return;
  }
  await refresh();
  if($('materials-checklist-plan').value)throw new Error('The checklist scope changed. Refresh and review it before downloading.');
  if(materialsExportDraft!==draft||draft.snapshot!==JSON.stringify(project)||draft.preview!==JSON.stringify(materialsPreview(project,draft.stamp)))throw new Error('The course changed. Refresh and review the materials checklist.');
  if(!globalThis.JSZip)throw new Error('The materials checklist exporter has not loaded.');
  const snapshot=structuredClone(project),blob=await materialsBlob(snapshot,draft.stamp,globalThis.JSZip);
  await refresh();
  if($('materials-checklist-plan').value)throw new Error('The checklist scope changed. Refresh and review it before downloading.');
  if(blocked||!ready||materialsExportDraft!==draft||!$('materials-checklist-reviewed').checked||draft.snapshot!==JSON.stringify(project)||draft.preview!==JSON.stringify(materialsPreview(project,draft.stamp)))throw new Error('The course or checklist review changed. Refresh and review the materials checklist.');
  downloadBlob(blob,materialsName());message('New materials checklist prepared from the reviewed saved snapshot. Completeness is not certified; review layout and accessibility before sharing.');
});
$('materials-refresh').onclick=()=>run(async()=>{await refresh();message('Saved preparation list refreshed from Word. Unsaved materials draft is unchanged.');});
$('materials-library-form').onsubmit=e=>{e.preventDefault();run(async()=>{
  if(!materialsDraft||materialsDraft.id!==$('materials-activity').value)throw new Error('Load the selected activity materials first.');
  const a=materialsDraft.project.modules.flatMap(m=>m.activities).find(a=>a.id===materialsDraft.id);
  const fields=Object.fromEntries(['title','icon','minutes','facilitator','participant','answers','lines'].map(key=>[key,a[key]]));fields.materials=materialsDraft.items;
  const next=saveLibraryCopy(library,$('materials-library-name').value,$('materials-library-tags').value.split(',').map(t=>t.trim()).filter(Boolean),fields);
  library=next;libraryPreferred=next.entries.at(-1).id;$('library-filter').value='';view('library');
  message('Independent library copy saved from the loaded activity text and materials draft. Download the private library to keep it; Word is unchanged.');
});};
$('materials-draft-download').onclick=()=>run(async()=>{
  if(!materialsDraft)throw new Error('Load activity materials before downloading a draft.');
  const text=serializeMaterialsDraft(materialsDraft.project,materialsDraft.id,materialsDraft.items,{item:$('material-item').value,quantity:$('material-quantity').value,unit:$('material-unit').value,use:$('material-use').value,notes:$('material-notes').value});
  downloadBlob(new Blob([text],{type:'application/json'}),`OpenStage Private Materials Draft - ${crypto.randomUUID()}.json`);
  message('Private materials draft prepared without reading Word. Restore it against the matching course/activity for review; this is not a full-master backup.');
},{recovery:true});
$('build-check').onclick=()=>run(async()=>{
  const supported=(name,version)=>{try{return globalThis.Office?.context?.requirements?.isSetSupported(name,version)===true;}catch{return false;}};
  diagnosticReport=await inspectBuild({entryURL:import.meta.url,registry:globalThis[Symbol.for('openstage.guide.runtimeModules')]||[],runtime:{mode:demo?'preview':'word',host:globalThis.Office?.context?.host,wordApi13:supported('WordApi','1.3'),sharedRuntime11:supported('SharedRuntime','1.1')}});
  $('build-diagnostic-summary').textContent=`${diagnosticReport.status}: ${diagnosticReport.reason} Checked ${diagnosticReport.checkedAt}`;
  message('Build information checked. Native acceptance remains pending; no course contents or settings were read or changed.');
},{recovery:true});
$('build-download').onclick=()=>run(async()=>{
  if(!diagnosticReport)throw new Error('Check the loaded build before downloading its diagnostic record.');
  downloadBlob(new Blob([JSON.stringify(diagnosticReport)],{type:'application/json'}),`OpenStage Build Diagnostic - ${crypto.randomUUID()}.json`);
  message('Local diagnostic record prepared. It contains build hashes and capability metadata, without document content or user identity.');
},{recovery:true});
$('ppt-file').onchange=()=>{pptSelection++;};
$('ppt-review-file').onchange=()=>{pptRestoreSelection++;};
$('ppt-update-file').onchange=()=>{pptUpdateSelection++;clearComparison();};
$('ppt-update-lineage').onchange=clearComparison;
$('ppt-change').onchange=()=>{if(!busy&&!blocked)renderComparison();};
for(const id of ['ppt-title','ppt-text','ppt-notes','ppt-lesson'])$(id).oninput=clearComparison;
$('ppt-module').onchange=clearComparison;
$('ppt-compare').onclick=()=>run(async()=>{
  if(!pptSource||!pptReview)throw new Error('Read or restore the original presentation review first.');
  if(!$('ppt-update-lineage').checked)throw new Error('Confirm that you intend to compare a revised version of this source. Numeric IDs alone cannot establish that relationship.');
  const file=$('ppt-update-file').files?.[0],selection=pptUpdateSelection,source=pptSource,review=pptReview,editor={id:pptCurrent,...pptFields()},editorText=JSON.stringify(editor);
  if(!file||file.size>MAX_PPTX_BYTES)throw new Error('Choose a revised standard PPTX up to 16 MB.');
  if(!globalThis.JSZip)throw new Error('The PowerPoint preview reader has not loaded.');
  const incoming=await readPptx(await file.arrayBuffer(),file.name,globalThis.JSZip),incomingDraft=pptxDraft(incoming);
  serializeReview(incoming,incomingDraft,incomingDraft.slides[0]);
  if(selection!==pptUpdateSelection||source!==pptSource||review!==pptReview||editorText!==JSON.stringify({id:pptCurrent,...pptFields()})||!$('ppt-update-lineage').checked)throw new Error('The revised file or author review changed. Compare again from the current review.');
  const result=comparePptx(source,incoming,review,editor);pptComparison={result,source,review,editor:editorText,selection};
  message('Read-only comparison prepared. Source matches and conflicts are proposals only; your source baseline, review text and Word document stay unchanged.');
});
$('ppt-restore').onclick=()=>run(async()=>{
  if(pptReview)throw new Error('Discard the current slide review before restoring another. Download its private review record first to keep your work.');
  const file=$('ppt-review-file').files?.[0],deck=$('ppt-file').files?.[0],selection=pptRestoreSelection,deckSelection=pptSelection;
  if(!file||file.size>MAX_REVIEW_BYTES||!deck||deck.size>MAX_PPTX_BYTES)throw new Error('Choose a private review record up to 32 MB and its matching original PPTX up to 16 MB.');
  if(!globalThis.JSZip)throw new Error('The PowerPoint preview reader has not loaded.');
  const text=await file.text(),data=await deck.arrayBuffer();
  const restored=await restoreReview(text,data,deck.name,globalThis.JSZip);
  if(selection!==pptRestoreSelection||deckSelection!==pptSelection)throw new Error('A selected restore file changed. Choose and check both files again.');
  pptSource=restored.source;pptReview=restored.review;pptCurrent=restored.unsavedEditor.id;pptAlias=restored.selectedName;loadSlideReview();
  for(const key of ['title','text','notes','moduleId','lesson']){
    const element=$(key==='moduleId'?'ppt-module':'ppt-'+key),value=restored.unsavedEditor[key];
    if(key==='moduleId'&&value&&!project?.modules.some(m=>m.id===value)&&value!==pptReview.slides.find(s=>s.id===pptCurrent).moduleId){const option=document.createElement('option');option.value=value;option.textContent='Restored proposal — choose a current module before saving';element.append(option);}
    element.value=value;
  }
  clearComparison();
  message('Private review restored against matching source bytes and extracted baseline. Saved review and unsaved editor are separate; reassess proposed assignments before applying any content. Word stays unchanged.');
});
$('ppt-read').onclick=()=>run(async()=>{
  if(pptReview)throw new Error('Discard the current slide review before reading another deck. Download its private review record first to keep your work.');
  const file=$('ppt-file').files?.[0],selection=pptSelection;
  if(!file||file.size>MAX_PPTX_BYTES)throw new Error('Choose a standard PPTX file up to 16 MB.');
  if(!globalThis.JSZip)throw new Error('The PowerPoint preview reader has not loaded.');
  const source=await readPptx(await file.arrayBuffer(),file.name,globalThis.JSZip);
  if(selection!==pptSelection)throw new Error('The selected PowerPoint file changed. Read the new file.');
  const draft=pptxDraft(source);serializeReview(source,draft,draft.slides[0]);
  pptSource=source;pptReview=draft;pptCurrent=source.slides[0].id;pptAlias='';loadSlideReview();
  clearComparison();
  message('Source read into an author-only review draft. No Word content or source file changed.');
});
$('ppt-reset').onclick=()=>run(async()=>{
  pptSource=null;pptReview=null;pptCurrent='';pptAlias='';
  clearComparison();
  for(const id of ['ppt-title','ppt-text','ppt-notes','ppt-module','ppt-lesson'])$(id).value='';
  $('ppt-baseline').textContent='';$('ppt-warnings').textContent='';
  message('Current slide review discarded. Choose a presentation to start another review.');
});
$('ppt-review-form').onsubmit=e=>{e.preventDefault();run(async()=>{
  if(!pptReview)throw new Error('Read a presentation before reviewing slides.');
  const next=editSlideDraft(pptReview,pptCurrent,pptFields(),(project?.modules||[]).map(m=>m.id));
  serializeReview(pptSource,next,{id:pptCurrent,...pptFields()});pptReview=next;clearComparison();message('Slide edits and proposed assignment saved in this session only.');
});};
$('ppt-slide').onchange=()=>{const target=$('ppt-slide').value;return run(async()=>{
  if(!pptReview?.slides.some(s=>s.id===target))throw new Error('Select a source slide.');
  const saved=pptReview.slides.find(s=>s.id===pptCurrent),fields=pptFields();
  if(Object.keys(fields).some(k=>fields[k]!==saved[k]))throw new Error('Save this slide review before switching slides. Your text is retained.');
  pptCurrent=target;loadSlideReview();
});};
$('ppt-export').onclick=()=>run(async()=>{
  if(!pptReview)throw new Error('Read a presentation before downloading its review.');
  // Preserve both saved review and current unsaved editor, even during quarantine.
  const text=serializeReview(pptSource,pptReview,{id:pptCurrent,...pptFields()});
  downloadBlob(new Blob([text],{type:'application/json'}),`OpenStage Private Slide Review - ${crypto.randomUUID()}.json`);
  message('Private source/review record prepared. Speaker notes are included. Keep the matching original PPTX to restore this review later.');
},{recovery:true});
$('library-filter').oninput=renderLibrary;
$('library-item').onchange=renderLibrary;
for(const key of LIBRARY_CONTENT_FIELDS)$('library-content-'+key).oninput=()=>{$('library-content-reviewed').checked=false;renderLibrary();};
$('library-content-icon').onchange=$('library-content-icon').oninput;
$('library-content-reviewed').onchange=renderLibrary;
for(const key of LIBRARY_MATERIAL_FIELDS)$('library-material-'+key).oninput=()=>{libraryMaterialDirty=true;$('library-content-reviewed').checked=false;renderLibrary();};
$('library-material-use').onchange=$('library-material-use').oninput;
$('library-material-clear').onclick=()=>run(async()=>{if(!libraryContentDraft)throw new Error('Load stored content first.');clearLibraryMaterialEntry();$('library-content-reviewed').checked=false;},{recovery:true});
$('library-material-apply').onclick=()=>run(async()=>{
  if(!libraryContentDraft)throw new Error('Load stored content first.');
  const row=Object.fromEntries(LIBRARY_MATERIAL_FIELDS.map(key=>[key,$('library-material-'+key).value])),current=libraryMaterialRows||[];
  if(libraryMaterialEditing&&!current.includes(libraryMaterialEditing))throw new Error('This material draft changed. Preserve and clear the entry before continuing.');
  const proposed=libraryMaterialEditing?current.map(item=>item===libraryMaterialEditing?row:item):[...current,row];
  const checked=normalizeMaterials(proposed);libraryMaterialRows=checked;clearLibraryMaterialEntry();$('library-content-reviewed').checked=false;message('Material row staged. Review and save the stored library content.');
},{recovery:true});
$('library-content-load').onclick=()=>run(async()=>{
  if(libraryContentDraft)throw new Error('Preserve and discard the current stored-content draft before loading another copy.');
  const entry=library.entries.find(e=>e.id===$('library-item').value);if(!entry)throw new Error('Select a saved library copy.');
  libraryContentDraft=structuredClone(entry);
  for(const key of LIBRARY_CONTENT_FIELDS)$('library-content-'+key).value=String(entry.fields[key]);
  $('library-content-reviewed').checked=false;libraryMaterialRows=entry.fields.materials===undefined?null:structuredClone(entry.fields.materials);clearLibraryMaterialEntry();
  message('Stored content loaded separately. The Activity form and Word are unchanged.');
},{recovery:true});
$('library-content-discard').onclick=()=>run(async()=>{
  libraryContentDraft=null;libraryMaterialRows=null;clearLibraryMaterialEntry();for(const key of LIBRARY_CONTENT_FIELDS)$('library-content-'+key).value='';$('library-content-reviewed').checked=false;message('Stored-content draft discarded. The saved library is unchanged.');
},{recovery:true});
$('library-content-form').onsubmit=e=>{e.preventDefault();run(async()=>{
  if(!libraryContentDraft||libraryContentDraft.id!==$('library-item').value)throw new Error('Select the loaded copy before saving its stored content.');
  if(libraryMaterialDirty)throw new Error('Add/apply or clear the unfinished material entry before saving stored content.');
  if(!$('library-content-reviewed').checked)throw new Error('Review the stored-content draft before saving.');
  const fields=Object.fromEntries(LIBRARY_CONTENT_FIELDS.map(key=>[key,$('library-content-'+key).value]));
  if(libraryMaterialRows!==null)fields.materials=libraryMaterialRows;
  const {next,change}=editLibraryContent(library,libraryContentDraft.id,fields,libraryContentDraft);
  library=next;if(change)libraryContentUndo=change;libraryContentDraft=structuredClone(next.entries.find(e=>e.id===libraryContentDraft.id));
  libraryMaterialRows=libraryContentDraft.fields.materials===undefined?null:structuredClone(libraryContentDraft.fields.materials);
  libraryPreferred=libraryContentDraft.id;$('library-filter').value='';$('library-content-reviewed').checked=false;
  message('Stored library content saved in this session. Download library to keep it. Loaded drafts and Word activities stay unchanged.');
},{recovery:true});};
$('library-content-undo').onclick=()=>run(async()=>{
  if(libraryContentDraft)throw new Error('Preserve and discard the stored-content draft before undoing.');
  if(!libraryContentUndo)throw new Error('No stored-content edit is available to undo.');
  const next=undoLibraryContent(library,libraryContentUndo),id=libraryContentUndo.before.id;library=next;libraryContentUndo=null;libraryPreferred=id;$('library-filter').value='';message('Last stored-content edit undone in this session. Download library to keep it.');
},{recovery:true});
$('library-file').onchange=()=>{librarySelection++;};
$('library-manage-load').onclick=()=>run(async()=>{
  const entry=library.entries.find(e=>e.id===$('library-item').value);if(!entry)throw new Error('Select a library copy to edit.');
  libraryEdit=structuredClone(entry);$('library-edit-name').value=entry.name;$('library-edit-tags').value=entry.tags.join(', ');
  message('Library name and tags loaded. Activity content and earlier insertions are unchanged.');
});
$('library-manage-form').onsubmit=e=>{e.preventDefault();run(async()=>{
  if(!libraryEdit||libraryEdit.id!==$('library-item').value)throw new Error('Load the selected library copy before editing its name and tags.');
  const next=renameLibraryCopy(library,libraryEdit.id,$('library-edit-name').value,$('library-edit-tags').value.split(',').map(t=>t.trim()).filter(Boolean),libraryEdit);
  library=next;libraryEdit=structuredClone(next.entries.find(e=>e.id===libraryEdit.id));$('library-filter').value='';libraryPreferred=libraryEdit.id;
  message('Session library name and tags updated. Download the library to keep these changes. Existing inserted activities stay unchanged.');
});};
$('library-remove').onclick=()=>run(async()=>{
  const {next,removed}=removeLibraryCopy(library,$('library-item').value);library=next;libraryRemoved=removed;
  message(`Removed ${removed.entry.name} from this session library. Undo is available for this removal. Source files and inserted activities stay unchanged.`);
});
$('library-undo').onclick=()=>run(async()=>{
  if(!libraryRemoved)throw new Error('There is no library removal to undo.');
  const next=undoLibraryRemoval(library,libraryRemoved),id=libraryRemoved.entry.id;library=next;libraryRemoved=null;$('library-filter').value='';libraryPreferred=id;
  message('Removed library copy restored in this session. Download the library to keep it.');
},{recovery:true});
$('library-save-form').onsubmit=e=>{e.preventDefault();run(async()=>{
  const next=saveLibraryCopy(library,$('library-name').value,$('library-tags').value.split(',').map(s=>s.trim()).filter(Boolean),inputFields());
  library=next;$('library-filter').value='';libraryPreferred=next.entries.at(-1).id;
  message('Independent activity copy saved in this panel session. Download the library to keep it after reload.');
});};
$('library-use').onclick=()=>run(async()=>requestActivityCandidate('library'));
$('library-export').onclick=()=>run(async()=>{
  const text=serializeLibrary(library);
  readLibrary(text);downloadBlob(new Blob([text],{type:'application/json'}),libraryName());
  message('Authoring library prepared as a new file. It includes facilitator instructions and answers; keep it private.');
},{recovery:true});
$('library-import').onclick=()=>run(async()=>{
  const file=$('library-file').files?.[0],selection=librarySelection;
  if(!file)throw new Error('Choose an OpenStage activity library file.');
  if(file.size>MAX_LIBRARY_BYTES)throw new Error('Choose a library file up to 1 MB.');
  const next=mergeLibrary(library,readLibrary(await file.text()));
  if(selection!==librarySelection)throw new Error('The selected file changed. Check the new file.');
  library=next;message('Library entries checked and added to this session. Download the combined library to keep it.');
});
$('course-details-load').onclick=()=>run(async()=>{
  courseDraft=null;await refresh();
  if(!project)throw new Error('Create a course before adding its details.');
  courseDraft=structuredClone(project);
  for(const [key,text] of Object.entries(courseDetails(project)))$('course-detail-'+key).value=text;
  $('course-details-summary').textContent=`Loaded revision ${project.revision}. Details are stored with this authoring master.`;
  message('Course details loaded. Edit the course title directly in Word.');
});
$('course-details-form').onsubmit=e=>{e.preventDefault();run(async()=>{
  if(!courseDraft)throw new Error('Load current course details before saving.');
  const fields=Object.fromEntries(COURSE_DETAIL_FIELDS.map(key=>[key,$('course-detail-'+key).value]));
  const next=editCourseDetails(courseDraft,fields);
  project=demo?next:await word.updateCourseDetails(courseDraft,fields);
  courseDraft=structuredClone(project);invalidateOutputs();
  $('course-details-summary').textContent=`Saved revision ${project.revision}. Details are stored with this authoring master.`;
  message(demo?'Course details saved in this browser preview.':'Course details saved in the authoring record. Save your Word document to keep them.');
});};
$('edit-activity-id').onchange=()=>{activityDraft=null;$('edit-activity-summary').textContent='Load this activity before editing its settings.';render();};
$('edit-activity-load').onclick=()=>run(async()=>{
  const id=$('edit-activity-id').value;activityDraft=null;
  await refresh();
  const item=project?.modules.flatMap(m=>m.activities).find(a=>a.id===id);
  if(!item)throw new Error('This activity no longer exists. Select another activity.');
  activityDraft={id,project:structuredClone(project)};$('edit-activity-id').value=id;
  $('edit-activity-minutes').value=String(item.minutes);$('edit-activity-lines').value=String(item.lines);
  $('edit-activity-summary').textContent=`${item.title} · ${item.icon} · Revision ${project.revision}. Text stays in Word. Save a backup before changing settings.`;
  message('Activity settings loaded from Word.');
});
$('edit-activity-form').onsubmit=e=>{e.preventDefault();run(async()=>{
  const draft=activityDraft;
  if(!draft||draft.id!==$('edit-activity-id').value)throw new Error('Load this activity before editing its settings.');
  const fields={minutes:$('edit-activity-minutes').value,lines:$('edit-activity-lines').value};
  const {next}=editActivitySettings(draft.project,draft.id,fields);
  project=demo?next:await word.updateActivitySettings(draft.project,draft.id,fields);
  activityDraft={id:draft.id,project:structuredClone(project)};invalidateOutputs();
  $('edit-activity-summary').textContent=`Settings saved · Revision ${project.revision}. Text stays in Word.`;
  message('Activity settings updated. Save your Word document, then refresh the agenda or workbook.');
});};
$('course-starter').onchange=()=>{$('course-frontmatter').hidden=$('course-starter').value!=='wrench';};
$('course-form').onsubmit=e=>{e.preventDefault();run(async()=>{if(project)throw new Error('This document already has a course. Use its course details controls instead.');const details=Object.fromEntries(COURSE_DETAIL_FIELDS.map(key=>[key,$('course-setup-'+key).value]));const starter=$('course-starter').value||'minimal';if(!['minimal','wrench'].includes(starter))throw new Error('Choose a supported guide starter.');const selected=['getting-started','background','team','overview'].filter(kind=>$('course-section-'+kind).checked);const next=starter==='wrench'?wrenchCourse($('course-title').value,details,selected):course($('course-title').value,details);await commit(next,'course');message(demo?'Course created in browser preview.':'Course created in Word. Save this document to keep your course.');view('module');});};
$('module-form').onsubmit=e=>{e.preventDefault();run(async()=>{await refresh();const {next,item}=addModule(project,$('module-title').value,$('module-objective').value);await commit(next,'module',item);$('activity-module').value=item.id;render();$('module-form').reset();message(demo?'Module added to browser preview.':'Module added to the end of your Word document.');view('activity');});};
$('activity-form').onsubmit=e=>{e.preventDefault();const submitted=activityRaw(),moduleId=submitted.destination,origin=structuredClone(libraryOrigin),starter=structuredClone(starterOrigin),review=structuredClone(starterReview),rawFields=inputFields();run(async()=>{
  const fields=origin||starter?libraryFields(rawFields):rawFields;
  const placeholders=origin||starter?libraryPlaceholders(fields):[];
  if(placeholders.length)throw new Error(`Replace unfinished prompts in these copied activity fields before inserting: ${placeholders.join(', ')}.`);
  if(starter&&(!$('assessment-reviewed').checked||!review||!sameRaw(review,submitted)))throw new Error('Review this exact assessment starter draft and destination before inserting.');
  await refresh();const {next,item}=addActivity(project,moduleId,fields);
  if(origin)item.librarySource=origin;
  await commit(next,'activity',item);
  const keptNewer=!sameRaw(submitted,activityRaw());
  if(!keptNewer){
    libraryOrigin=null;starterOrigin=null;starterReview=null;copiedMaterials=null;activityReplacement=null;$('assessment-reviewed').checked=false;$('activity-library-note').textContent='';$('activity-form').reset();
    for(const key of ACTIVITY_DRAFT_FIELDS)$('activity-'+key).value=ACTIVITY_DEFAULTS[key];
    $('activity-module').value=moduleId;activityGeneration++;cue();
  }
  message((demo?'Activity added to the selected module in browser preview.':'Activity and instructional cue inserted at the end of the selected module. Save your Word document.')+(keptNewer?' Newer form work has been retained.':''));
});};
$('refresh').onclick=()=>run(async()=>{await refresh();message('Course refreshed from the document.');});
$('workbook-plan').onchange=()=>{familyGeneration++;invalidateFamily();if(busy||blocked)return;workbookDraft=null;previewRevision=null;$('workbook-plan-reviewed').checked=false;$('workbook-preview').textContent='Refresh and review the chosen workbook scope.';render();};
$('workbook-plan-reviewed').onchange=render;
$('preview').onclick=()=>run(async()=>{await refresh();showPreview();message('Preview updated. Facilitator content is excluded.');});
for(const family of WORKBOOK_FAMILIES){const option=document.createElement('option');option.value=family.id;option.textContent=family.name;$('family-template').append(option);}
$('family-template').value=WORKBOOK_FAMILIES[0].id;
$('family-base-language').value=WORKBOOK_DEFAULT_BASE_LANGUAGE;
$('family-toc-include').checked=false;
const familyChanged=()=>{familyGeneration++;invalidateFamily();render();};
familyFrontMatterPanel=mountWorkbookFrontMatterPanel(document,{isUnavailable:()=>familyDesignUnavailable()||busy||blocked||!ready||!project||!!familyResearchPanel?.isPending()||!!familyLogoPanel?.isPending()||!!familyLogoPanel?.isLoading(),
  onChange:changed=>{if(changed)familyChanged();else render();},context:()=>JSON.stringify({project,family:familyChoice()})});
familyResearchPanel=mountWorkbookResearchPanel(document,{isUnavailable:()=>familyDesignUnavailable()||busy||blocked||!ready||!project||!familyWorkspaceVisible()||!!familyFrontMatterPanel?.isPending()||!!familyLogoPanel?.isPending()||!!familyLogoPanel?.isLoading(),
  onChange:changed=>{if(changed)familyChanged();else render();},context:()=>JSON.stringify({project,family:familyChoice()})});
familyLogoPanel=mountWorkbookLogoPanel(document,{isUnavailable:()=>familyDesignUnavailable()||busy||blocked||!ready||!project||!familyWorkspaceVisible()||!!familyFrontMatterPanel?.isPending()||!!familyResearchPanel?.isPending(),
  onChange:changed=>{if(changed)familyChanged();else render();},context:()=>JSON.stringify({project,family:familyChoice()})});

const designControls=()=>[
 ['family-template','familyId'],['family-base-language','baseLanguage'],
 ...FAMILY_BRAND_FIELDS.map(k=>['family-brand-'+k,'brand.'+k]),
 ['family-front-include-cover','frontMatter.includeCover','checked'],['family-front-include-copyright','frontMatter.includeCopyright','checked'],
 ...WORKBOOK_COVER_FIELDS.map(k=>['family-front-cover-'+k,'frontMatter.cover.'+k]),...WORKBOOK_COPYRIGHT_FIELDS.map(k=>['family-front-copyright-'+k,'frontMatter.copyright.'+k]),
 ['family-toc-include','contents.include','checked'],['family-research-include','research.include','checked'],...WORKBOOK_RESEARCH_FIELDS.map(k=>['family-research-'+k,'research.'+k]),
 ['family-logo-include','logo.include','checked'],['family-logo-role','logo.role'],['family-logo-alt','logo.alt'],['family-logo-width','logo.widthIn']
];
const designValue=(value,path)=>path.split('.').reduce((out,key)=>out[key],value);
function workbookDesignDraft(){return validateWorkbookDesignDraft({familyId:$('family-template').value,baseLanguage:$('family-base-language').value,brand:Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,$('family-brand-'+k).value])),frontMatter:{includeCover:!!$('family-front-include-cover').checked,includeCopyright:!!$('family-front-include-copyright').checked,cover:Object.fromEntries(WORKBOOK_COVER_FIELDS.map(k=>[k,$('family-front-cover-'+k).value])),copyright:Object.fromEntries(WORKBOOK_COPYRIGHT_FIELDS.map(k=>[k,$('family-front-copyright-'+k).value]))},contents:{include:!!$('family-toc-include').checked},research:{include:!!$('family-research-include').checked,...Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(k=>[k,$('family-research-'+k).value]))},logo:{include:!!$('family-logo-include').checked,role:$('family-logo-role').value,alt:$('family-logo-alt').value,widthIn:$('family-logo-width').value}});}
function designFingerprint(){return JSON.stringify({project,ready,blocked,plan:$('workbook-plan').value,family:familyChoice(),consentEpoch:familyConsentEpoch,reviewed:!!$('family-reviewed').checked,display:familyDisplayStructure(),raw:designControls().map(([id,,kind])=>({id,checked:kind==='checked'?!!$(id).checked:null,raw:rawControl($(id))})),presetPending:!!presetPending,frontPending:familyFrontMatterPanel.isPending(),researchPending:familyResearchPanel.isPending(),logoPending:familyLogoPanel.isPending(),logoLoading:familyLogoPanel.isLoading()});}
function checkDesignCandidate(prepared){
 const draft=validateWorkbookDesignDraft(prepared.draft);
 for(const[id,path,kind]of designControls()){
  if(kind==='checked')continue;const source=$(id),value=designValue(draft,path);
  if(source.tagName==='SELECT'){const options=Array.from(source.options||source.children||[]);if(options.length&&!options.some(o=>o.value===value))throw Error('A current design choice cannot represent the imported value. Current work was kept.');continue;}
  const node=document.createElement((source.tagName||'input').toLowerCase());if((node.tagName||'INPUT')==='INPUT')node.type=source.type||'text';node.value=value;if(node.value!==value)throw Error('An imported raw design field cannot be represented exactly by this editor. Current work and the original bundle were kept.');
 }
 let warning='';try{const rawBrand=Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,{value:draft.brand[k],invalid:Object.fromEntries(Object.keys(familyRawBrand()[k].invalid).map(f=>[f,false]))}]));const resolved=resolveWorkbookBrand(draft.familyId,familyBrandRequest(rawBrand));if(!resolved.eligible)warning='The imported brand is a draft that fails current policy. Correct it and refresh the workbook preview before export; no suggested alternative was applied.';}catch{warning='Some imported design values are unfinished drafts. Correct them and refresh the workbook preview before export.';}return warning;
}
function applyDesignFile(prepared,complete){
 checkDesignCandidate(prepared);const draft=validateWorkbookDesignDraft(prepared.draft),logoToken=familyLogoPanel.snapshotDesignFile();
 const before=designControls().map(([id,,kind])=>({node:$(id),kind,value:$(id).value,checked:$(id).checked}));
 const saved={familyDraft,familyAlternative,familyGeneration,familyConsentEpoch,reviewed:$('family-reviewed').checked,status:$('family-brand-status').textContent,preview:$('family-preview-content').textContent,previewChildren:Array.from($('family-preview-content').childNodes||$('family-preview-content').children||[]),previewMockText:Object.hasOwn($('family-preview-content'),'text')?$('family-preview-content').text:null,alternativeHidden:$('family-use-safe-alternative').hidden,alternativeDisabled:$('family-use-safe-alternative').disabled};
 try{
  for(const[id,path,kind]of designControls())if(!id.startsWith('family-logo-')){if(kind==='checked')$(id).checked=designValue(draft,path);else $(id).value=designValue(draft,path);}
  familyLogoPanel.replaceDesignFile(logoToken,draft.logo,prepared.logoAsset);
  const actual=createWorkbookDesignFile(workbookDesignDraft(),familyLogoPanel.designFileAsset()),expected=createWorkbookDesignFile(draft,prepared.logoAsset);
  if(JSON.stringify(actual)!==JSON.stringify(expected))throw Error('The restored design differed from the checked bundle. Current work was restored.');
  familyGeneration++;familyConsentEpoch++;invalidateFamily();render();complete();familyLogoPanel.finishDesignFile(logoToken);
 }catch(e){
  for(const row of before){if(row.kind==='checked')row.node.checked=row.checked;else row.node.value=row.value;}
  familyLogoPanel.rollbackDesignFile(logoToken);({familyDraft,familyAlternative,familyGeneration,familyConsentEpoch}=saved);$('family-reviewed').checked=saved.reviewed;$('family-brand-status').textContent=saved.status;$('family-preview-content').replaceChildren(...saved.previewChildren);if(saved.previewMockText!==null)$('family-preview-content').text=saved.previewMockText;$('family-use-safe-alternative').hidden=saved.alternativeHidden;$('family-use-safe-alternative').disabled=saved.alternativeDisabled;throw e;
 }
}
familyDesignFilesPanel=mountWorkbookDesignFiles(document,{isBusy:()=>busy,hasPending:()=>!!presetPending||familyFrontMatterPanel.isPending()||familyResearchPanel.isPending()||familyLogoPanel.isPending()||familyLogoPanel.isLoading(),getFingerprint:designFingerprint,getDraft:workbookDesignDraft,getAsset:()=>familyLogoPanel.designFileAsset(),checkCandidate:checkDesignCandidate,applyPrepared:applyDesignFile,download:downloadBlob,onStateChange:render});

$('family-template').onchange=familyChanged;
$('family-base-language').onchange=familyChanged;
$('family-toc-include').oninput=$('family-toc-include').onchange=familyChanged;
const contentsBeforeUnload=e=>{const raw=familyRawContents();if(raw.checked||Object.values(raw.invalid).some(Boolean)){e.preventDefault?.();e.returnValue='';return '';}};
globalThis.addEventListener?.('beforeunload',contentsBeforeUnload);
for(const key of FAMILY_BRAND_FIELDS)$('family-brand-'+key).oninput=$('family-brand-'+key).onchange=familyChanged;
$('family-reviewed').onchange=()=>{familyConsentEpoch++;render();};
$('family-preview').onclick=()=>{if(refuseHiddenFamilyAction())return;if(familyLogoPanel?.isPending()||familyLogoPanel?.isLoading()){familyLogoPanel.reportPending();return;}if(familyResearchPanel?.isPending()){familyResearchPanel.reportPending();return;}if(familyFrontMatterPanel?.isPending()){familyFrontMatterPanel.reportPending();return;}return run(showFamilyPreview);};
$('family-use-safe-alternative').onclick=()=>{
  const alternative=familyAlternative;
  if(!alternative||alternative.choice!==JSON.stringify(familyChoice())){invalidateFamily();render();return;}
  for(const [key,value] of Object.entries(alternative.fields))$('family-brand-'+key).value=String(value);
  familyGeneration++;invalidateFamily();$('family-brand-status').textContent='Suggested values were applied only to the named controls. Refresh and review the updated request before creating a copy.';render();
};
$('family-apply').onclick=()=>{
  if(refuseHiddenFamilyAction())return;
  if(familyLogoPanel?.isPending()||familyLogoPanel?.isLoading()){familyLogoPanel.reportPending();return;}
  if(familyResearchPanel?.isPending()){familyResearchPanel.reportPending();return;}
  if(familyFrontMatterPanel?.isPending()){familyFrontMatterPanel.reportPending();return;}
  return run(async()=>{
    const draft=familyDraft,consentEpoch=familyConsentEpoch;
    const current=()=>consentEpoch===familyConsentEpoch&&familyCurrent(draft,true);
    if(!current())throw new Error('Refresh and review the exact new workbook design before applying it.');
    await refresh();if(!current())throw new Error('The course, scope, public front matter, branding or displayed workbook design changed. Refresh and review it.');
    if(!globalThis.JSZip)throw new Error('The workbook exporter has not loaded. Reopen OpenStage Guide.');
    const blob=await familyWorkbookOutputBlob(draft.output,globalThis.JSZip);
    if(!current())throw new Error('The workbook review or acknowledgement changed during packaging. No copy was offered.');
    await refresh();if(!current())throw new Error('The course, scope, public front matter, branding or displayed workbook design changed during packaging. Refresh and review it.');
    downloadBlob(blob,familyWorkbookName(draft.title,draft.familyId,draft.selectedPlanName));
    message('New designed participant workbook prepared from the exact reviewed learner content. Save it as a new copy and inspect its layout and accessibility before sharing.');
  });
};
const presetAttempt=action=>{try{action();}catch(e){presetPending=null;$('family-preset-confirm').hidden=true;presetStatus(e.message||String(e));renderBrandPresets();}};
$('family-preset-list').onchange=()=>{presetPending=null;$('family-preset-confirm').hidden=true;renderBrandPresets();};
$('family-preset-name').oninput=()=>{presetPending=null;$('family-preset-confirm').hidden=true;renderBrandPresets();};
$('family-preset-save-new').onclick=()=>presetAttempt(()=>presetWrite('create'));
for(const [id,action] of [['family-preset-overwrite','overwrite'],['family-preset-load','load'],['family-preset-reset','reset']])
  $(id).onclick=()=>presetAttempt(()=>beginPresetAction(action));
$('family-preset-keep').onclick=()=>{presetPending=null;$('family-preset-confirm').hidden=true;presetStatus('Current branding and review were kept.');renderBrandPresets();};
$('family-preset-discard').onclick=()=>presetAttempt(confirmPresetAction);
$('family-preset-reload').onclick=reloadBrandPresets;
for(const id of ['facilitator-answers','facilitator-participant'])$(id).onchange=()=>{invalidateFacilitator();render();};
$('facilitator-plan').onchange=()=>{if(busy||blocked)return;invalidateFacilitator();render();};
$('facilitator-reviewed').onchange=render;
$('facilitator-preview').onclick=()=>run(async()=>{
  invalidateFacilitator();await refresh();
  const options=facilitatorChoice(),stamp=new Date().toISOString(),planId=$('facilitator-plan').value,preview=planId?planFacilitatorPreview(project,planId,options,stamp):facilitatorPreview(project,options,stamp),text=facilitatorParagraphs(preview).map(([,text])=>text).join('\n');
  $('facilitator-output').textContent=text;
  facilitatorDraft={kind:planId?'plan':'full',planId,snapshot:JSON.stringify(project),options:JSON.stringify(options),stamp,preview:JSON.stringify(preview),text,displayStructure:facilitatorDisplayStructure()};
  message('Review the complete private text, content choices, unresolved suggestions and omissions. Answers/debrief and participant instructions are included only when selected.');
});
function facilitatorReviewed(draft){return !blocked&&ready&&facilitatorDraft===draft&&$('facilitator-reviewed').checked&&$('facilitator-plan').value===draft.planId&&draft.snapshot===JSON.stringify(project)&&draft.options===JSON.stringify(facilitatorChoice())&&draft.preview===JSON.stringify(draft.kind==='plan'?planFacilitatorPreview(project,draft.planId,facilitatorChoice(),draft.stamp):facilitatorPreview(project,facilitatorChoice(),draft.stamp))&&draft.text===$('facilitator-output').textContent&&draft.displayStructure===facilitatorDisplayStructure();}
$('facilitator-download').onclick=()=>run(async()=>{
  const draft=facilitatorDraft;if(!draft||!$('facilitator-reviewed').checked)throw new Error('Refresh and review the private facilitator copy before downloading.');
  await refresh();if(!facilitatorReviewed(draft))throw new Error('The course, choices or private-copy review changed. Refresh and review this private copy.');
  if(!globalThis.JSZip)throw new Error('The private-copy exporter has not loaded.');
  const snapshot=structuredClone(project),options=facilitatorChoice(),blob=draft.kind==='plan'?await planFacilitatorBlob(snapshot,draft.planId,options,draft.stamp,globalThis.JSZip):await facilitatorBlob(snapshot,options,draft.stamp,globalThis.JSZip);
  await refresh();if(!facilitatorReviewed(draft))throw new Error('The course, choices or private-copy review changed. Refresh and review this private copy.');
  downloadBlob(blob,draft.kind==='plan'?planFacilitatorName(planFacilitatorPreview(snapshot,draft.planId,options,draft.stamp).planName):facilitatorName());message('New private structured facilitator copy prepared. Keep your master and backup; edits here do not flow back. Review layout and accessibility before use.');
});
$('pack-plan').onchange=()=>{if(busy||blocked)return;invalidatePack();render();};
$('pack-start').oninput=()=>{if(busy||blocked)return;invalidatePack();render();};
$('pack-route').onchange=()=>{if(busy||blocked)return;invalidatePack();$('pack-answers').checked=false;$('pack-participant').checked=false;render();};
for(const id of ['pack-answers','pack-participant'])$(id).onchange=()=>{if(busy||blocked)return;invalidatePack();render();};
for(const id of ['pack-reviewed','pack-private-consent'])$(id).onchange=render;
$('pack-preview').onclick=()=>run(async()=>{
  const planId=$('pack-plan').value;if(!planId)throw new Error('Choose a saved delivery plan for the pack.');
  await refresh();const choice=packChoice(),stamp=new Date().toISOString(),preview=deliveryPackPreview(project,planId,{...choice,stamp});
  invalidatePack();$('pack-output').textContent=packReviewText(preview);
  packDraft={planId,route:choice.route,choice:JSON.stringify(choice),snapshot:JSON.stringify(project),stamp,preview:JSON.stringify(preview),text:$('pack-output').textContent,structure:packDisplayStructure()};
  message(choice.route===PACK_PRIVATE?'PRIVATE pack review prepared. Confirm the separate facilitator consent before downloading.':'Shareable learner pack review prepared. No facilitator file will be included.');
});
$('pack-download').onclick=()=>run(async()=>{
  const draft=packDraft;if(!draft||!packReviewed(draft))throw new Error('Refresh and review the exact delivery pack before downloading.');
  if(!globalThis.JSZip)throw new Error('The pack exporter has not loaded. Reopen OpenStage Guide.');
  const guard=async()=>{await refresh();if(!packReviewed(draft))throw new Error('The delivery pack source, route, choices or review changed. Refresh and review it again.');};
  await guard();const snapshot=structuredClone(project),choice=packChoice();
  const built=await buildDeliveryPack(snapshot,draft.planId,{...choice,stamp:draft.stamp},globalThis.JSZip,guard);
  await guard();downloadBlob(new Blob([built.data],{type:'application/zip'}),deliveryPackName(draft.route));
  message(draft.route===PACK_PRIVATE?'New PRIVATE facilitator ZIP prepared. Keep it with intended facilitators and inspect every file before use.':'New shareable learner ZIP prepared. Inspect every file before use.');
});
$('download').onclick=()=>run(async()=>{
  const draft=workbookDraft;
  if(!draft)throw new Error('Refresh and review the workbook preview. Resolve items marked Needs attention before downloading.');
  if(draft.kind==='plan'){
    await refresh();if(!selectedWorkbookReviewed(draft))throw new Error('The selected plan, course or workbook review changed. Refresh and review the selected workbook.');
    if(!globalThis.JSZip)throw new Error('The workbook exporter has not loaded. Reopen OpenStage Guide.');
    const snapshot=structuredClone(project),blob=await planWorkbookBlob(snapshot,draft.planId,draft.stamp,globalThis.JSZip);
    await refresh();if(!selectedWorkbookReviewed(draft))throw new Error('The selected plan, course or workbook review changed. Refresh and review the selected workbook.');
    downloadBlob(blob,planWorkbookName(snapshot.title,planWorkbookPreview(snapshot,draft.planId,draft.stamp).planName));
    message('New selected-plan participant workbook prepared from the exact reviewed scope. Save it as a new copy; this selection does not approve objective coverage.');return;
  }
  await refresh();
  if($('workbook-plan').value)throw new Error('The workbook scope changed. Refresh and review it before downloading.');
  if(draft.snapshot!==JSON.stringify(project))throw new Error('The course changed. Refresh and review the workbook before downloading.');
  if(!globalThis.JSZip)throw new Error('The workbook exporter has not loaded. Reopen OpenStage Guide.');
  const snapshot=structuredClone(project),blob=await worksheetBlob(snapshot,globalThis.JSZip);
  await refresh();
  if($('workbook-plan').value)throw new Error('The workbook scope changed. Refresh and review it before downloading.');
  if(workbookDraft!==draft||draft.snapshot!==JSON.stringify(project))throw new Error('The course changed. Refresh and review the workbook before downloading.');
  downloadBlob(blob,worksheetName(snapshot.title));
  message('New participant workbook prepared. Save it as a new copy. Replacing or regenerating an existing output requires a verified backup and is not available here.');
});
$('recover').onclick=()=>run(async()=>{if(demo){message('No interrupted Word insertion in browser preview.');return;}project=await word.recover();blocked=false;invalidateOutputs();message('Course records checked and recovered. Save your Word document.');},{recovery:true});
function planState(){return JSON.stringify({name:$('plan-name').value,target:$('plan-target').value,groups:planGroups,generation:planGeneration});}
function planChanged(){planGeneration++;planReview=null;$('plan-reviewed').checked=false;agendaDraft=null;render();}
function planScope(){return project?deliveryPlans(project):null;}
const invalidPlanOption=row=>`!invalid:${row.index}`;
function renderOutputPlanSelectors(){
 const state=planScope();
 for(const id of ['workbook-plan','materials-checklist-plan','facilitator-plan','pack-plan']){
  const select=$(id),prior=select.value;select.replaceChildren();const full=document.createElement('option');full.value='';full.textContent='Full course (default)';select.append(full);
  if(id==='pack-plan')full.textContent='Choose a saved plan';
  if(state?.supported)for(const row of state.plans){const option=document.createElement('option');option.value=row.id||invalidPlanOption(row);option.textContent=`${row.name||`Plan ${row.index+1}`}${row.reasons.length?' — needs repair':''}`;select.append(option);}
  if(prior&&!state?.plans?.some(row=>row.id===prior||invalidPlanOption(row)===prior)){const stale=document.createElement('option');stale.value=prior;stale.textContent='Previously selected plan is missing or unavailable';select.append(stale);}
  select.value=prior;select.disabled=busy||blocked||!ready||!project;
 }
}
function renderPlanEditor(){
  const state=planScope(),select=$('plan-select'),prior=select.value;
  select.replaceChildren();const full=document.createElement('option');full.value='';full.textContent='Full course (default)';select.append(full);
  if(state?.supported)for(const row of state.plans){const option=document.createElement('option');option.value=row.id||invalidPlanOption(row);option.textContent=`${row.name||`Plan ${row.index+1}`}${row.reasons.length?' — needs repair':''}`;select.append(option);}
  if(prior&&!state?.plans?.some(p=>p.id===prior||invalidPlanOption(p)===prior)){const missing=document.createElement('option');missing.value=prior;missing.textContent='Previously selected plan is missing';select.append(missing);}
  select.value=prior;select.disabled=busy||blocked||!ready||!project||!state?.supported;
  $('plan-state').textContent=!project?'Create a course before planning.':!state?.supported?state.reason:prior?state.plans.find(p=>p.id===prior)?.reasons.join(' ')||'This selection affects only the agenda. Choose scope separately in Workbook and Materials.':'Full course is selected here. Workbook and Materials have separate scope choices.';
  $('plan-new').disabled=busy||blocked||!ready||!project||!state?.supported||!!planDraft;
  $('plan-load').disabled=busy||blocked||!ready||!project||!state?.supported||!prior||!!planDraft;
  $('plan-discard').disabled=busy||!planDraft;
  $('plan-editor').hidden=!planDraft;
  for(const id of ['plan-name','plan-target','plan-module','plan-activity','plan-add','plan-reviewed','plan-save'])$(id).disabled=busy||blocked||!ready||!planDraft;
  $('agenda-reviewed').disabled=busy||blocked||!ready||!agendaDraft||agendaDraft.kind!=='plan';
  const moduleSelect=$('plan-module'),chosen=moduleSelect.value;moduleSelect.replaceChildren();
  for(const module of project?.modules||[]){const option=document.createElement('option');option.value=module.id;option.textContent=module.title;moduleSelect.append(option);}
  moduleSelect.value=project?.modules.some(m=>m.id===chosen)?chosen:project?.modules[0]?.id||'';
  renderPlanActivities();
  const list=$('plan-groups');list.replaceChildren();
  for(const [gi,group] of planGroups.entries()){
    const sourceDraft=planDraft,sourceGeneration=planGeneration;
    const liveGroup=()=>!busy&&planDraft===sourceDraft&&planGeneration===sourceGeneration&&planGroups[gi]===group;
    const staleRow=()=>{message('Plan controls changed. Use the current row buttons; your draft is preserved.');render();};
    const item=document.createElement('li'),module=project?.modules.find(m=>m.id===group.moduleId),heading=document.createElement('strong');heading.textContent=`${module?.title||'Missing module'} (${group.moduleId})`;item.append(heading);
    for(const [label,delta] of [['Move module earlier',-1],['Move module later',1]]){const b=document.createElement('button');b.type='button';b.textContent=label;b.disabled=busy||!planDraft||gi+delta<0||gi+delta>=planGroups.length;b.onclick=()=>{if(!liveGroup()||gi+delta<0||gi+delta>=planGroups.length)return staleRow();[planGroups[gi],planGroups[gi+delta]]=[planGroups[gi+delta],planGroups[gi]];planChanged();};item.append(b);}
    const activities=document.createElement('ol');for(const [ai,id] of group.activityIds.entries()){
      const row=document.createElement('li'),a=module?.activities.find(a=>a.id===id),label=document.createElement('span');label.textContent=`${a?.title||'Missing or moved activity'} (${id})`;row.append(label);
      const liveRow=()=>liveGroup()&&group.activityIds[ai]===id;
      for(const [title,delta] of [['Earlier',-1],['Later',1]]){const b=document.createElement('button');b.type='button';b.textContent=title;b.disabled=busy||!planDraft||ai+delta<0||ai+delta>=group.activityIds.length;b.onclick=()=>{if(!liveRow()||ai+delta<0||ai+delta>=group.activityIds.length)return staleRow();[group.activityIds[ai],group.activityIds[ai+delta]]=[group.activityIds[ai+delta],group.activityIds[ai]];planChanged();};row.append(b);}
      const remove=document.createElement('button');remove.type='button';remove.textContent='Remove';remove.disabled=busy||!planDraft;remove.onclick=()=>{if(!liveRow())return staleRow();group.activityIds.splice(ai,1);if(!group.activityIds.length)planGroups.splice(gi,1);planChanged();};row.append(remove);activities.append(row);
    }item.append(activities);list.append(item);
  }
  $('plan-draft-status').textContent=planDraft?`${planGroups.reduce((n,g)=>n+g.activityIds.length,0)} selected activities in ${planGroups.length} ordered module groups. Unsaved changes stay here until saved or discarded.`:'';
}
function renderPlanActivities(){const selected=$('plan-activity'),old=selected.value,module=project?.modules.find(m=>m.id===$('plan-module').value);selected.replaceChildren();for(const a of module?.activities||[]){const option=document.createElement('option');option.value=a.id;option.textContent=a.title;selected.append(option);}selected.value=module?.activities.some(a=>a.id===old)?old:module?.activities[0]?.id||'';selected.disabled=busy||blocked||!ready||!planDraft||!selected.value;}
function loadPlanDraft(id,record){planDraft={id,project:structuredClone(project)};planGroups=structuredClone(record?.groups||[]);$('plan-name').value=record?.name||'';$('plan-target').value=record?.targetMinutes??'';planGeneration++;planReview=null;$('plan-reviewed').checked=false;render();}
$('plan-name').oninput=$('plan-target').oninput=planChanged;
$('plan-reviewed').onchange=()=>{planReview=$('plan-reviewed').checked?planState():null;render();};
$('plan-module').onchange=()=>{renderPlanActivities();render();};
$('plan-select').onchange=()=>run(async()=>{agendaDraft=null;$('agenda-reviewed').checked=false;await refresh();showAgenda();});
$('plan-new').onclick=()=>run(async()=>{if(planDraft)throw new Error('Discard the current plan draft before starting another.');await refresh();if(!planScope().supported)throw new Error(planScope().reason);loadPlanDraft(null,{name:'',groups:[]});message('New plan draft started. Select and order existing activities.');});
$('plan-load').onclick=()=>run(async()=>{if(planDraft)throw new Error('Discard the current plan draft before loading another.');const id=$('plan-select').value;await refresh();const rows=planScope();if(!rows.supported)throw new Error(rows.reason);const matches=rows.plans.filter(p=>p.id===id);if(matches.length!==1||!id)throw new Error('Select one unique saved plan.');const row=matches[0],groups=row.record?.groups;
  if(row.reasons.some(reason=>reason.includes('plan ID')))throw new Error('This saved plan ID cannot be displayed exactly. Preserve the master and create a separate plan instead of silently changing this ID.');
  if(!Array.isArray(groups)||groups.length>200||groups.reduce((n,g)=>n+(Array.isArray(g?.activityIds)?g.activityIds.length:0),0)>2000||groups.some(g=>!g||typeof g!=='object'||Array.isArray(g)||typeof g.moduleId!=='string'||!Array.isArray(g.activityIds)||g.activityIds.length>2000||g.activityIds.some(a=>typeof a!=='string')))throw new Error('This plan has malformed groups and cannot be edited in this version. Preserve the master; other plans and the full-course agenda remain available.');
  loadPlanDraft(id,row.record);message('Plan loaded for editing. Missing references remain visible until you repair them.');});
$('plan-discard').onclick=()=>run(async()=>{planDraft=null;planGroups=[];planReview=null;$('plan-name').value='';$('plan-target').value='';$('plan-reviewed').checked=false;message('Unsaved plan draft discarded. Saved plans are unchanged.');},{recovery:true});
$('plan-add').onclick=()=>run(async()=>{if(!planDraft)throw new Error('Start or load a plan draft.');const moduleId=$('plan-module').value,activityId=$('plan-activity').value,module=project?.modules.find(m=>m.id===moduleId);if(!module?.activities.some(a=>a.id===activityId))throw new Error('Select an existing activity in that module.');if(planGroups.some(g=>g.activityIds.includes(activityId)))throw new Error('That activity is already selected.');let group=planGroups.find(g=>g.moduleId===moduleId);if(!group){group={moduleId,activityIds:[]};planGroups.push(group);}group.activityIds.push(activityId);planChanged();});
$('plan-save').onclick=()=>run(async()=>{
  const draft=planDraft;if(!draft)throw new Error('Start or load a plan draft.');
  const state=planState();if(!$('plan-reviewed').checked||planReview!==state)throw new Error('Review this exact plan name, target, activities, and order before saving.');
  const raw=$('plan-target').value.trim();if(raw&&!/^(?:\d{1,5})(?:\.\d{1,3})?$/.test(raw))throw new Error('Enter target minutes as a positive number with up to three decimals.');
  const change={kind:draft.id?'update':'create',name:$('plan-name').value,targetMinutes:raw?Number(raw):null,groups:structuredClone(planGroups)};if(draft.id)change.id=draft.id;
  if(JSON.stringify(draft.project)!==JSON.stringify(project))throw new Error('The course changed. Keep this draft and reload the saved plan before saving.');
  // Pure validation happens before the first settings write.
  editDeliveryPlan(project,change);
  const beforeIds=new Set(planScope().plans.map(p=>p.id));const next=demo?editDeliveryPlan(project,change).next:await word.updateDeliveryPlan(draft.project,change);
  const unchanged=planDraft===draft&&state===planState(),savedId=draft.id||deliveryPlans(next).plans.find(p=>!beforeIds.has(p.id))?.id;
  project=next;invalidateOutputs();if(unchanged){planDraft=null;planGroups=[];planReview=null;$('plan-name').value='';$('plan-target').value='';$('plan-reviewed').checked=false;if(savedId){renderPlanEditor();$('plan-select').value=savedId;}}
  message('Named delivery plan saved in the master. Save your Word document. Newer unfinished plan work remains in the editor.');
});
function showAgenda(){
  agendaDraft=null;
  $('agenda-reviewed').checked=false;
  const start=$('agenda-start').value,planId=$('plan-select').value,planned=!!planId,stamp=planned?new Date().toISOString():null;
  let a,preview;clearAgendaCoverage();
  if(planned){preview=planAgendaPreview(project,planId,start);a=preview;
    $('agenda-summary').textContent=`Selected plan ${a.planName} · Plan ID: ${a.planId} · Project ID: ${a.projectId} · Revision ${a.revision} · Generated ${stamp} · ${a.minutes} minutes · ${a.start}–${a.end} · ${a.targetStatus} Exact selected scope: ${a.scope.map(g=>`${g.moduleId}: ${g.activityIds.join(', ')}`).join(' | ')}. ${a.coverageNotice}`;
    const report=selectedPlanObjectiveReport(project,planId,start),list=$('agenda-coverage');for(const row of report.rows){const item=document.createElement('li'),role=kind=>{const r=row.roles[kind];return `${kind}: ${r.selectedCurrent} current selected / ${r.courseCurrent} current whole-course, ${r.selectedStale} stale selected${r.lostBySelection?', all current coverage lost by this selection':''}${r.preexistingGap?', pre-existing gap':''}${r.currentCourseGapNote?', full-course gap note (not plan approval)':''}${r.noteConflict?', note conflicts with current course links':''}`;};item.textContent=`${row.moduleTitle} (${row.moduleId}) — ${row.blank?'Objective blank':row.objective}. ${role('practice')}. ${role('assessment')}.`;list.append(item);}list.hidden=!list.children.length;$('agenda-coverage-heading').hidden=list.hidden;
  }else{a=agendaView(project,start);preview=agendaPreview(project,start);const issues=reviewCourse(project);
    $('agenda-summary').textContent=`${AGENDA_SCOPE} · Project ID: ${project.id} · ${a.minutes} minutes · ${a.start}–${a.end} · Generation revision ${a.revision} · ${issues.filter(x=>x.severity==='error').length} items need attention; ${issues.filter(x=>x.severity!=='error').length} suggestions. Review Guide lists the unresolved items.`;}
  const body=$('agenda-body');body.replaceChildren();let last=null;
  for(const row of a.rows){
    if(row.moduleId!==last){const tr=document.createElement('tr'),th=document.createElement('th');th.colSpan=3;th.scope='rowgroup';th.textContent=row.moduleTitle;tr.className='agenda-module';tr.append(th);body.append(tr);last=row.moduleId;}
    const tr=document.createElement('tr');for(const text of [`${row.start}–${row.end}`,row.title,String(row.minutes)]){const td=document.createElement('td');td.textContent=text;tr.append(td);}body.append(tr);
  }
  $('agenda-table').hidden=!a.rows.length;$('agenda-empty').hidden=!!a.rows.length;
  if(a.rows.length&&[a.title,...a.rows.flatMap(r=>[r.moduleTitle,r.title])].every(s=>s.trim()))agendaDraft={kind:planned?'plan':'full',planId,stamp,revision:project.revision,startInput:start,snapshot:JSON.stringify(project),preview:JSON.stringify(preview),displayed:agendaDisplayed()};
  render();
}
function agendaDisplayed(){return JSON.stringify({summary:$('agenda-summary').textContent,privateHeading:$('agenda-coverage-heading').textContent,privateHeadingHidden:$('agenda-coverage-heading').hidden,coverage:$('agenda-coverage').textContent,coverageHidden:$('agenda-coverage').hidden,table:$('agenda-body').textContent});}
function showReview(){
  showCoverageReport();
  const issues=reviewCourse(project),errors=issues.filter(x=>x.severity==='error').length;
  $('review-summary').textContent=issues.length?`${errors} items need attention · ${issues.length-errors} suggestions · Revision ${project.revision}`:`No issues found in these checks · Revision ${project.revision}`;
  const root=$('review-list');root.replaceChildren();
  for(const issue of issues){
    const m=project.modules.find(m=>m.id===issue.id||m.activities.some(a=>a.id===issue.id)),a=m?.activities.find(a=>a.id===issue.id);
    const li=document.createElement('li'),tag=document.createElement('strong'),p=document.createElement('p'),b=document.createElement('button');
    tag.textContent=issue.severity==='error'?'Needs attention':'Suggestion';p.textContent=issue.message;b.textContent=a?.title||m?.title||project.title;
    b.onclick=()=>issue.type==='course'?view('course'):run(()=>demo?Promise.resolve(message('Item selected in browser preview.')):word.jump(issue.id,issue.type));
    li.append(tag,p,b);root.append(li);
  }
}
$('agenda-form').onsubmit=e=>{e.preventDefault();run(async()=>{await refresh();showAgenda();message('Agenda refreshed from the course.');});};
$('agenda-start').oninput=()=>{agendaDraft=null;$('agenda-reviewed').checked=false;clearAgendaCoverage();$('agenda-summary').textContent='Start time changed. Refresh the agenda before downloading.';render();};
$('agenda-reviewed').onchange=render;
$('agenda-download').onclick=()=>run(async()=>{
  const draft=agendaDraft,start=$('agenda-start').value;
  if(!draft||draft.startInput!==start||draft.planId!==$('plan-select').value||draft.kind==='plan'&&!$('agenda-reviewed').checked)throw new Error('Refresh and review the exact selected agenda before downloading.');
  await refresh();
  const preview=draft.kind==='plan'?planAgendaPreview(project,draft.planId,start):agendaPreview(project,start);
  if(draft.snapshot!==JSON.stringify(project)||draft.preview!==JSON.stringify(preview)||draft.displayed!==agendaDisplayed())throw new Error('The course changed or the displayed agenda changed. Refresh and review it before downloading.');
  if(!globalThis.JSZip)throw new Error('The agenda exporter has not loaded. Reopen OpenStage Guide.');
  const snapshot=structuredClone(project),blob=draft.kind==='plan'?await planAgendaBlob(snapshot,draft.planId,start,draft.stamp,globalThis.JSZip):await agendaBlob(snapshot,start,globalThis.JSZip);
  await refresh();
  if(agendaDraft!==draft||$('agenda-start').value!==start||$('plan-select').value!==draft.planId||draft.kind==='plan'&&!$('agenda-reviewed').checked||JSON.stringify(project)!==draft.snapshot||draft.displayed!==agendaDisplayed())throw new Error('The agenda changed. Refresh and review it before downloading.');
  downloadBlob(blob,draft.kind==='plan'?planAgendaName(snapshot.title,preview.planName):agendaName(snapshot.title));
  message('New agenda file prepared. Save it as a new copy. Replacing or regenerating an existing output requires a verified backup and is not available here.');
});
$('review-refresh').onclick=()=>run(async()=>{await refresh();showReview();message('Review refreshed from the course.');});
function downloadBlob(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
$('backup-save').onclick=()=>run(async()=>{
  if(demo)throw new Error('Save a full backup from Word. Browser preview has no Word master.');
  if(!globalThis.JSZip)throw new Error('The backup tools have not loaded. Reopen OpenStage Guide.');
  const snapshot=await word.captureMaster();project=snapshot.project;invalidateOutputs();
  const backup=await createCheckpoint(snapshot.master,project,globalThis.JSZip);
  downloadBlob(new Blob([backup.data],{type:'application/zip'}),checkpointName(project.title));
  message('Backup prepared and its package checked. After saving it, choose the saved file below to verify that copy.');
});
$('backup-file').onchange=()=>{backupSelection++;verifiedBackup=null;$('backup-details').textContent='';render();};
$('backup-inspect').onclick=()=>run(async()=>{
  verifiedBackup=null;$('backup-details').textContent='';
  const file=$('backup-file').files?.[0];
  const selection=backupSelection;
  if(!file)throw new Error('Choose an OpenStage backup first.');
  if(file.size>MAX_PACKAGE_BYTES)throw new Error('Choose an OpenStage backup up to 80 MB.');
  if(!globalThis.JSZip)throw new Error('The backup tools have not loaded. Reopen OpenStage Guide.');
  const data=new Uint8Array(await file.arrayBuffer()),checked=await readCheckpoint(data,globalThis.JSZip);
  if(selection!==backupSelection)throw new Error('The selected file changed. Verify the new backup.');
  verifiedBackup=data;
  $('backup-details').textContent=`Verified: ${checked.project.title} · Revision ${checked.project.revision} · ${checked.assetCount} supporting assets\nCaptured ${checked.manifest.created}\nWord master SHA-256: ${checked.masterHash}`;
  message('Saved backup verified. You can now prepare a restored Word copy.');
},{recovery:true});
$('backup-restore').onclick=()=>run(async()=>{
  if(!verifiedBackup)throw new Error('Verify the saved backup before restoring it.');
  const selection=backupSelection;
  const checked=await readCheckpoint(verifiedBackup,globalThis.JSZip);
  if(selection!==backupSelection)throw new Error('The selected file changed. Verify the new backup.');
  downloadBlob(new Blob([checked.master],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}),checkpointName(checked.project.title,true));
  message('Restored Word copy prepared. Open that downloaded copy in Word; this document has not been replaced.');
},{recovery:true});
async function init(){
  if(projectMode){
    ready=true;$('connection').textContent='Project file · opening';render();
    // The pane reads the guide only after the project file has been read; a file that cannot be opened is explained and left alone.
    // (showSave has already put the reason on the page, in the words the adapter chose.)
    try{await guideAdapter.ready;}catch{return;}
    if(programMode)globalThis.openstageProgram={openTool,updateGuide:()=>$('guide-update').click()};
    message('Your guide is saved to its project file as you work.');await guidePane.refresh();guidePane.watch();toProgram({type:'osg-ready'});return;
  }
  if(demo){ready=true;$('connection').textContent='Browser preview · not connected to Word';message('Preview mode. Changes stay in this browser session.');render();await guidePane.refresh();guidePane.watch();return;}
  if(!globalThis.Office){error(new Error('Open this panel from the OpenStage Guide ribbon in Word.'));render();return;}
  await Office.onReady();
  if(Office.context.host!==Office.HostType.Word||!word.supported()){error(new Error('This prototype needs desktop Word with WordApi 1.3 support.'));$('connection').textContent='Word support required';render();return;}
  ready=true;$('connection').textContent='Connected to Word · Prototype 0.1';
  // The guide does not wait for the older course tools: if they stall or fail, the guide still shows and can be edited.
  globalThis.addEventListener?.('focus',()=>guidePane.quietRefresh());
  const guideRead=guidePane.refresh().then(()=>guidePane.watch());
  try{project=await word.loadProject();if(project)await refresh();}catch(e){blocked=true;error(e);}
  render();
  await guideRead;
  if(word.acceptanceEligible()){
    const {attachAcceptancePanel}=await import('./acceptance-panel.mjs');
    attachAcceptancePanel({word,run,refresh,commit,getProject:()=>project,isBusy:()=>busy});
  }
}
assessmentPanel=mountAssessmentPanel(document,{isBusy:()=>busy,download:downloadBlob,getZip:()=>globalThis.JSZip});
reloadBrandPresets();
const SAVE_WORDS={new:'no changes yet',saving:'saving',saved:'saved',unsaved:'changes not saved yet',error:'not saved'};
const showSave=s=>{
  if(!projectMode||s.state==='opening')return;
  toProgram({type:'osg-status',state:s.state,message:s.message});
  $('connection').textContent=`Project file · ${SAVE_WORDS[s.state]||s.state}`;
  if(s.state==='error')error(new Error(s.message));
};
const guideAdapter=projectMode?createProjectAdapter(createDemoAdapter(),{...webProjectStore(),onStatus:showSave}):demo?createDemoAdapter():createWordAdapter();
if(projectMode){
  // Save whatever is waiting when the page is hidden or closed.
  globalThis.addEventListener?.('pagehide',()=>{guideAdapter.flush();});
  document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='hidden')guideAdapter.flush();});
}
// A file box does nothing inside Word's pane, so in Word every one of them opens the pane's file window instead (file-input-bridge.mjs).
installFileInputBridge({document,pick:opts=>guideAdapter.pickFiles(opts),onError:error,enabled:!demo});
let toolStrip;
const guidePane=mountGuidePane(document,{adapter:guideAdapter,onTool:id=>toolStrip?.show(id)});
toolStrip=mountToolStrip(document,{open:openTool});
// The pane starts on one tool, the guide itself, not on the whole guide page.
guidePane.useTool('openDesigner');
view(new URLSearchParams(location.search).get('view')||'guide');render();
const initialization=init();initialization.catch(error);
if(!demo&&globalThis.Office?.actions){
  registerRibbonActions(Office,{ready:initialization,go:openTool,onError:error});
  registerNavigationActions(Office,{ready:initialization,adapter:guideAdapter,say:text=>guidePane.say(text),onError:error});
}
