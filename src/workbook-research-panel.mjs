import {rawControl} from './activity-draft.mjs';
import {WORKBOOK_RESEARCH_FIELDS} from './workbook-research.mjs';

// Session-only authored public fields; no Word/course/store/network dependency.
export function mountWorkbookResearchPanel(document,{isUnavailable=()=>false,onChange=()=>{},context=()=>''}={}){
  const n=id=>document.getElementById('family-research-'+id);let pending=null;
  const raw=()=>({schemaVersion:1,include:{checked:!!n('include').checked,...rawControl(n('include'))},...Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(k=>[k,rawControl(n(k))]))});
  const fingerprint=()=>JSON.stringify({raw:raw(),context:context()});
  const state=node=>({tag:node.tagName||'',text:node.textContent,hidden:!!node.hidden,class:node.className||'',style:node.getAttribute?.('style')||node.style?.cssText||'',children:Array.from(node.children||[],state)});
  const confirmation=()=>JSON.stringify({root:state(n('confirm')),text:state(n('confirm-text')),keep:state(n('keep')),replace:state(n('replace'))});
  const visible=target=>{if(!target||target.isConnected===false)return false;for(let node=target;node;node=node.parentElement||node.parent){if(node.hidden||/display\s*:\s*none|visibility\s*:\s*(?:hidden|collapse)/u.test(node.getAttribute?.('style')||node.style?.cssText||''))return false;try{if(typeof globalThis.getComputedStyle==='function'&&node.nodeType===1){const style=globalThis.getComputedStyle(node);if(style.display==='none'||['hidden','collapse'].includes(style.visibility))return false;}}catch{}}return true;};
  const ready=target=>{if(document.getElementById('view-output')?.hidden||!visible(n('controls'))||!visible(target))throw new Error('Open the workbook research controls before changing this draft.');if(isUnavailable())throw new Error('Finish the current operation or other replacement choice before changing research.');};
  const status=text=>{n('status').textContent=text;};
  const report=e=>{n('error').hidden=false;n('error').textContent=e.message||String(e);n('error').scrollIntoView?.({block:'nearest'});};
  const clear=()=>{n('error').hidden=true;n('error').textContent='';};
  const dismiss=()=>{const target=pending?.returnFocus;pending=null;n('confirm').hidden=true;n('confirm-text').textContent='';if(target&&!target.disabled&&visible(target))target.focus?.();else n('status').focus?.();};
  const action=(fn,target)=>()=>{try{ready(target);clear();fn();}catch(e){report(e);}finally{renderBusy();}};
  n('reset').onclick=action(()=>{
    if(pending)throw new Error('Keep or replace the pending research reset first.');
    pending={fingerprint:fingerprint(),returnFocus:document.activeElement||document.focused||null};
    n('confirm-text').textContent='Reset only research inclusion, rationale, evidence, application, limitations and references? Other drafts, branding, cover/copyright and contents are kept.';
    n('confirm').hidden=false;n('confirm').scrollIntoView?.({block:'nearest'});n('keep').focus?.();onChange(false);pending.display=confirmation();
  },n('reset'));
  n('keep').onclick=action(()=>{dismiss();status('Research draft and any still-current review were kept.');onChange(false);},n('keep'));
  n('replace').onclick=action(()=>{
    if(!pending)throw new Error('No research reset is pending.');
    if(pending.fingerprint!==fingerprint()||pending.display!==confirmation()){dismiss();onChange(false);throw new Error('Research, source, options or the displayed confirmation changed. Nothing was reset; start the intended action again.');}
    for(const k of WORKBOOK_RESEARCH_FIELDS)n(k).value='';n('include').checked=false;dismiss();onChange(true);status('Only research text and inclusion were reset; other drafts were kept.');
  },n('replace'));
  const changed=()=>{clear();onChange(true);renderBusy();};
  for(const k of WORKBOOK_RESEARCH_FIELDS)n(k).oninput=n(k).onchange=changed;
  n('include').oninput=n('include').onchange=changed;
  n('confirm').onkeydown=e=>{if(e.key==='Escape'){e.preventDefault?.();n('keep').onclick();}};
  function renderBusy(){const blocked=!!isUnavailable();for(const k of WORKBOOK_RESEARCH_FIELDS)n(k).disabled=blocked;n('include').disabled=blocked;n('reset').disabled=blocked||!!pending;n('keep').disabled=n('replace').disabled=blocked||!pending;}
  for(const k of WORKBOOK_RESEARCH_FIELDS)n(k).value='';n('include').checked=false;n('confirm').hidden=true;clear();status('Research text is session-only and public when included. Claims, references and applicability are authored and have not been verified.');renderBusy();
  const hasSessionWork=()=>{const r=raw();return r.include.checked||Object.values(r.include.invalid).some(Boolean)||WORKBOOK_RESEARCH_FIELDS.some(k=>r[k].value!==''||Object.values(r[k].invalid).some(Boolean));};
  const beforeUnload=e=>{if(hasSessionWork()){e.preventDefault?.();e.returnValue='';return '';}};globalThis.addEventListener?.('beforeunload',beforeUnload);
  return Object.freeze({raw,isPending:()=>!!pending,renderBusy,hasSessionWork,beforeUnload,reportPending:()=>report(new Error('Keep or replace the pending research reset before reviewing or applying.'))});
}
