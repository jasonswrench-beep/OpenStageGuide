import {rawControl} from './activity-draft.mjs';
import {WORKBOOK_COVER_FIELDS,WORKBOOK_COPYRIGHT_FIELDS,workbookFrontMatterRequest,workbookCopyrightNoticeStarter} from './workbook-public-frontmatter.mjs';

// Only session raw draft controls. No Word/course/store/localStorage dependency.
export function mountWorkbookFrontMatterPanel(document,{isUnavailable=()=>false,onChange=()=>{},context=()=>''}={}){
  const n=id=>document.getElementById('family-front-'+id);
  const fields=[...WORKBOOK_COVER_FIELDS.map(k=>['cover',k]),...WORKBOOK_COPYRIGHT_FIELDS.map(k=>['copyright',k])];
  let pending=null;
  const choice=id=>({checked:!!n(id).checked,...rawControl(n(id))});
  const raw=()=>({schemaVersion:1,includeCover:choice('include-cover'),includeCopyright:choice('include-copyright'),
    cover:Object.fromEntries(WORKBOOK_COVER_FIELDS.map(k=>[k,rawControl(n('cover-'+k))])),copyright:Object.fromEntries(WORKBOOK_COPYRIGHT_FIELDS.map(k=>[k,rawControl(n('copyright-'+k))]))});
  const fingerprint=()=>JSON.stringify({raw:raw(),context:context()});
  const status=message=>{n('status').textContent=message;};
  const report=error=>{n('error').hidden=false;n('error').textContent=error.message||String(error);n('error').scrollIntoView?.({block:'nearest'});};
  const clearError=()=>{n('error').hidden=true;n('error').textContent='';};
  const visible=target=>{if(!target||target.isConnected===false)return false;for(let node=target;node;node=node.parentElement||node.parent){if(node.hidden||/display\s*:\s*none|visibility\s*:\s*hidden/u.test(node.getAttribute?.('style')||node.style?.cssText||''))return false;}return true;};
  const ready=target=>{if(document.getElementById('view-output')?.hidden||!visible(n('controls'))||!visible(target))throw new Error('Open the workbook public front-matter controls before changing this draft.');if(isUnavailable())throw new Error('Wait until the current course operation finishes before changing public front matter.');};
  const confirmationState=()=>{const shape=node=>({tag:node.tagName||'',text:node.textContent,hidden:!!node.hidden,class:node.className||'',style:node.getAttribute?.('style')||node.style?.cssText||'',children:Array.from(node.children||[],shape)});return JSON.stringify({root:shape(n('confirm')),text:shape(n('confirm-text')),keep:shape(n('keep')),replace:shape(n('replace'))});};
  const dismiss=()=>{const target=pending?.returnFocus;pending=null;n('confirm').hidden=true;n('confirm-text').textContent='';if(target&&!target.disabled&&visible(target))target.focus?.();else n('status').focus?.();};
  const action=(fn,target)=>()=>{try{ready(target);clearError();fn();}catch(error){report(error);}finally{renderBusy();}};
  const request=(kind,candidate,message)=>{
    if(pending)throw new Error('Keep or replace the pending public front-matter choice first.');
    pending={kind,candidate,fingerprint:fingerprint(),returnFocus:document.activeElement||document.focused||null};
    n('confirm-text').textContent=message;n('confirm').hidden=false;n('confirm').scrollIntoView?.({block:'nearest'});n('keep').focus?.();
    onChange(false);pending.display=confirmationState();
  };
  n('reset').onclick=action(()=>request('reset',null,'Reset the cover and copyright inclusion choices and all their typed text? Your branding, course and other drafts stay unchanged. Keep retains the current public draft and review.'),n('reset'));
  n('starter').onclick=action(()=>{
    const r=raw(),inactive={...r,includeCover:{...r.includeCover,checked:false},includeCopyright:{...r.includeCopyright,checked:false}};
    const metadata=workbookFrontMatterRequest(inactive),candidate=workbookCopyrightNoticeStarter(metadata);
    if(r.copyright.noticeChoice.invalid.customError||r.copyright.noticeChoice.invalid.valueMissing)throw new Error('Correct the notice choice before requesting a starter.');
    if(metadata.copyright.noticeChoice==='named-license'&&['licenseName','licenseVersion','licenseUrl'].some(k=>Object.values(r.copyright[k].invalid).some(Boolean)))throw new Error('Correct the named license fields before requesting a starter.');
    if(candidate===n('copyright-noticeText').value){status('This notice already matches the requested starter. Your draft and review were kept.');return;}
    request('starter',candidate,'Replace only the editable notice text with a starter for your selected notice choice? Review and edit that wording before sharing. No rights, holder, year or license are selected for you.');
  },n('starter'));
  n('keep').onclick=action(()=>{dismiss();status('Public cover/copyright draft and any still-current workbook review were kept.');onChange(false);},n('keep'));
  n('replace').onclick=action(()=>{
    const intent=pending;if(!intent)throw new Error('No public front-matter replacement is pending.');
    if(intent.fingerprint!==fingerprint()||intent.display!==confirmationState()){dismiss();onChange(false);throw new Error('The public draft, course, scope or design changed. Nothing was replaced; start the intended action again.');}
    if(intent.kind==='reset'){
      for(const [group,key]of fields)n(group+'-'+key).value='';n('include-cover').checked=false;n('include-copyright').checked=false;
    }else n('copyright-noticeText').value=intent.candidate;
    dismiss();onChange(true);status(intent.kind==='reset'?'Public cover/copyright draft reset. Branding and source content were preserved.':'Editable notice starter applied only to this draft. Review it and refresh the workbook copy preview.');
  },n('replace'));
  const changed=()=>{clearError();onChange(true);renderBusy();};
  for(const [group,key]of fields)n(group+'-'+key).oninput=n(group+'-'+key).onchange=changed;
  for(const id of ['include-cover','include-copyright'])n(id).oninput=n(id).onchange=changed;
  n('confirm').onkeydown=e=>{if(e.key==='Escape'){e.preventDefault?.();n('keep').onclick();}};
  function renderBusy(){
    const unavailable=!!isUnavailable();
    for(const [group,key]of fields)n(group+'-'+key).disabled=unavailable;
    for(const id of ['include-cover','include-copyright'])n(id).disabled=unavailable;
    for(const id of ['reset','starter'])n(id).disabled=unavailable||!!pending;
    n('keep').disabled=n('replace').disabled=unavailable||!pending;
  }
  for(const [group,key]of fields)n(group+'-'+key).value='';n('include-cover').checked=false;n('include-copyright').checked=false;
  n('confirm').hidden=true;clearError();status('Cover/copyright choices and text are session-only. They are not saved in the course or local brand presets.');renderBusy();
  const hasSessionWork=()=>{const r=raw();return r.includeCover.checked||r.includeCopyright.checked||[...Object.values(r.cover),...Object.values(r.copyright)].some(x=>x.value!==''||Object.values(x.invalid).some(Boolean))||Object.values(r.includeCover.invalid).some(Boolean)||Object.values(r.includeCopyright.invalid).some(Boolean);};
  const beforeUnload=e=>{if(hasSessionWork()){e.preventDefault?.();e.returnValue='';return '';}};
  globalThis.addEventListener?.('beforeunload',beforeUnload);
  return Object.freeze({raw,hasSessionWork,beforeUnload,isPending:()=>!!pending,renderBusy,reportPending:()=>report(new Error('Keep or replace the pending public front-matter choice before previewing or applying.'))});
}
